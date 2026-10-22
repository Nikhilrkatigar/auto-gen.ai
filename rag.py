"""
rag.py - "Ask the audit reports": retrieval-augmented Q&A over generated
daily / weekly / monthly concurrent audit reports (.docx).

Pipeline::

    .docx reports -> table rows -> mask customer data -> chunks
    chunks -> BM25 (keywords) + bge-small embeddings (meaning) -> index on disk
    question -> hybrid search (Reciprocal Rank Fusion) -> top chunks
    top chunks -> Claude, answer only from the sources, cite [n]

Design rules:
  * **Customer data never leaves unmasked.** Account / CIF / mobile numbers
    keep only their last 4 digits and person names become initials, at
    ingest time, so neither the index nor the Claude prompt sees raw values.
  * **One table row = one chunk.** Reports are mostly tables; a row plus its
    branch, report type, period and annexure title is a self-contained fact.
  * **Retrieval works offline.** Only the final answer step calls the API;
    without credentials the CLI prints the retrieved sources instead.

Usage::

    python rag.py index sample_reports
    python rag.py ask "Which staff accounts had transactions above 45000 at LAKEVIEW?"
    python rag.py ask --no-llm "interest rate variation in HILLTOP"
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path

import docx
from docx.table import Table

INDEX_DIR = Path(".rag_index")
EMBED_MODEL = "BAAI/bge-small-en-v1.5"
CLAUDE_MODEL = "claude-opus-5-5"
RRF_K = 60  # standard Reciprocal Rank Fusion constant

# Sections whose "name" columns hold people rather than ledger/BGL heads.
PERSON_SECTION = re.compile(
    r"staff|cif|mobile|sb/ ?cd|inoperative|account holders|attendance|activation", re.I
)
LONG_NUMBER = re.compile(r"\d{9,}")


@dataclass
class Chunk:
    text: str
    source: str
    branch: str
    report_type: str
    period: str
    section: str


# ---------------------------------------------------------------- masking

def mask_numbers(text: str) -> str:
    """Account, CIF and mobile numbers -> keep the last 4 digits."""
    return LONG_NUMBER.sub(lambda m: "X" * (len(m.group()) - 4) + m.group()[-4:], text)


def initials(name: str) -> str:
    words = [w for w in re.split(r"[\s.]+", name) if w and w.lower() not in ("mr", "mrs", "miss", "ms", "dr")]
    return " ".join(w[0].upper() + "." for w in words) if words else name


# ---------------------------------------------------------------- ingest

def _text(el) -> str:
    return "".join(t.text or "" for t in el.iter() if t.tag.endswith("}t")).strip()


def _cells(row) -> list[str]:
    out = []
    for c in row.cells:
        t = " ".join(c.text.split())
        if not out or t != out[-1]:  # merged cells repeat their text
            out.append(t)
    return out


def _report_type(name: str) -> str:
    low = name.lower()
    return "Monthly" if "monthly" in low else "Weekly" if "weekly" in low else "Daily" if "daily" in low else "Report"


def chunk_report(path: Path) -> list[Chunk]:
    d = docx.Document(str(path))
    rtype = _report_type(path.name)
    # Branch comes from the file name ("Daily AR - 2.1102 - HILLTOP.docx"):
    # batch monthly outputs can carry another branch's name in their headers.
    branch = path.stem.split("-")[-1].strip().lstrip("0123456789. ").strip()
    period, section, chunks = "", "", []

    for el in d.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            t = _text(el)
            if m := re.search(r"period (\d\d/\d\d/\d{4} to \d\d/\d\d/\d{4})|(?:Reporting )?Month:\s*([A-Z]+ - \d{4})", t):
                period = m.group(1) or m.group(2)
            elif not period and (m := re.search(r"Date of report:\s*(\d\d[/-]\d\d[/-]\d{4})", t)):
                period = m.group(1)
            if re.match(r"(Annexure|REPORT ON|CONSOLIDATED REPORT|Daily Concurrent Audit Report on)", t):
                section = t[:160]
            elif t.startswith("(DL/TL)"):
                section = re.sub(r"\s*\(CCOD\)\s*$", "", section) + " (DL/TL)"
        elif tag == "tbl":
            rows = [_cells(r) for r in Table(el, d).rows]
            if len(rows) < 2:
                continue
            header = rows[0]
            person = bool(PERSON_SECTION.search(section))
            for cells in rows[1:]:
                values = [v for v in cells[1:] if v] if len(cells) > 1 else cells
                # cells[0] == header[0] catches the second line of two-row headers
                if cells[0] == header[0] or not values or all(re.fullmatch(r"NIL|-+", v, re.I) for v in values):
                    continue
                pairs = []
                for h, v in zip(header, cells):
                    if not v or v == h:
                        continue
                    if person and "name" in h.lower():
                        v = initials(v)
                    pairs.append(f"{h}: {mask_numbers(v)}")
                if pairs:
                    head = f"{rtype} report | Branch: {branch} | Period: {period or 'n/a'} | {section or 'Summary'}"
                    chunks.append(Chunk(head + "\n" + "; ".join(pairs), path.name, branch, rtype, period, section))
    return chunks


def ingest(folders: list[str]) -> list[Chunk]:
    chunks = []
    for folder in folders:
        for path in sorted(Path(folder).rglob("*.docx")):
            if not path.name.startswith("~$"):
                chunks += chunk_report(path)
    return chunks


# ---------------------------------------------------------------- retrieval

def expand_amounts(question: str) -> str:
    """Reports store plain rupees: add '360000' for '3.6 lakh' (and crore)."""
    def to_rupees(m):
        value = float(m.group(1).replace(",", "")) * (10**7 if m.group(2).lower().startswith("crore") else 10**5)
        return f"{m.group()} {value:.0f}"
    return re.sub(r"(\d[\d,]*(?:\.\d+)?)\s*(lakhs?|crores?)\b", to_rupees, question, flags=re.I)


def tokenize(text: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", text.lower())


class BM25:
    """Okapi BM25 keyword ranking (k1=1.5, b=0.75)."""

    def __init__(self, docs: list[str], k1: float = 1.5, b: float = 0.75):
        self.toks = [Counter(tokenize(d)) for d in docs]
        self.lens = [sum(t.values()) for t in self.toks]
        self.avg = sum(self.lens) / max(len(self.lens), 1)
        df = Counter(w for t in self.toks for w in t)
        n = len(docs)
        self.idf = {w: math.log(1 + (n - f + 0.5) / (f + 0.5)) for w, f in df.items()}
        self.k1, self.b = k1, b

    def scores(self, query: str) -> list[float]:
        q = tokenize(query)
        out = []
        for tf, ln in zip(self.toks, self.lens):
            s = 0.0
            for w in q:
                if w in tf:
                    f = tf[w]
                    s += self.idf[w] * f * (self.k1 + 1) / (f + self.k1 * (1 - self.b + self.b * ln / self.avg))
            out.append(s)
        return out


class Embedder:
    def __init__(self):
        from fastembed import TextEmbedding  # heavy import, only when needed
        self.model = TextEmbedding(EMBED_MODEL)

    def docs(self, texts):
        import numpy as np
        return np.array(list(self.model.passage_embed(texts)), dtype="float32")

    def query(self, text):
        import numpy as np
        return np.array(next(iter(self.model.query_embed([text]))), dtype="float32")


class Index:
    def __init__(self, chunks: list[Chunk], vectors, embedder=None):
        self.chunks, self.vectors = chunks, vectors
        self.bm25 = BM25([c.text for c in chunks])
        self.branches = sorted({c.branch for c in chunks}, key=len, reverse=True)
        self._embedder = embedder

    @property
    def embedder(self):
        if self._embedder is None:
            self._embedder = Embedder()
        return self._embedder

    # -- persistence
    @classmethod
    def build(cls, chunks: list[Chunk], embedder=None) -> "Index":
        embedder = embedder or Embedder()
        return cls(chunks, embedder.docs([c.text for c in chunks]), embedder)

    def save(self, folder: Path = INDEX_DIR):
        import numpy as np
        folder.mkdir(exist_ok=True)
        (folder / "chunks.json").write_text(json.dumps([asdict(c) for c in self.chunks], indent=1))
        np.save(folder / "vectors.npy", self.vectors)

    @classmethod
    def load(cls, folder: Path = INDEX_DIR) -> "Index":
        import numpy as np
        chunks = [Chunk(**c) for c in json.loads((folder / "chunks.json").read_text())]
        return cls(chunks, np.load(folder / "vectors.npy"))

    # -- search
    def search(self, question: str, k: int = 8, mode: str = "hybrid") -> list[tuple[Chunk, float]]:
        question = expand_amounts(question)
        # A branch named in the question is a hard filter: it is metadata, not similarity.
        q = question.upper()
        branch = next((b for b in self.branches if re.search(rf"\b{re.escape(b)}\b", q)), None)
        allowed = [i for i, c in enumerate(self.chunks) if branch is None or c.branch == branch]

        rankings = []
        if mode in ("bm25", "hybrid"):
            s = self.bm25.scores(question)
            rankings.append(sorted(allowed, key=lambda i: -s[i]))
        if mode in ("dense", "hybrid"):
            s = self.vectors @ self.embedder.query(question)  # vectors are L2-normalised
            rankings.append(sorted(allowed, key=lambda i: -s[i]))

        fused = Counter()
        for ranking in rankings:
            for rank, i in enumerate(ranking):
                fused[i] += 1 / (RRF_K + rank + 1)
        return [(self.chunks[i], score) for i, score in fused.most_common(k)]


# ---------------------------------------------------------------- answer

SYSTEM = """You answer questions about bank concurrent audit reports for an auditor.
Use only the numbered sources provided. Cite every fact with its source number, like [2].
If the sources do not contain the answer, say so plainly and do not guess.
Account and mobile numbers are masked to their last 4 digits and customer names to initials; keep them that way.
Amounts are in Indian rupees. Be brief: a short answer, then a compact list or table if useful."""


def build_prompt(question: str, hits: list[tuple[Chunk, float]]) -> str:
    sources = "\n\n".join(f"[{n}] ({c.source})\n{c.text}" for n, (c, _) in enumerate(hits, 1))
    return f"<sources>\n{sources}\n</sources>\n\nQuestion: {question}"


def answer(question: str, hits: list[tuple[Chunk, float]], client=None) -> str:
    import anthropic
    client = client or anthropic.Anthropic()
    response = client.beta.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=16000,
        system=SYSTEM,
        output_config={"effort": "medium"},
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        messages=[{"role": "user", "content": build_prompt(question, hits)}],
    )
    if response.stop_reason == "refusal":
        return "The model declined to answer this question."
    return "".join(b.text for b in response.content if b.type == "text")


def has_credentials() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"))


# ---------------------------------------------------------------- CLI

def main(argv=None):
    ap = argparse.ArgumentParser(description="Ask questions about generated audit reports.")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("index", help="build the index from report folders")
    p.add_argument("folders", nargs="+")
    p = sub.add_parser("ask", help="ask a question")
    p.add_argument("question")
    p.add_argument("-k", type=int, default=8, help="number of sources to retrieve")
    p.add_argument("--no-llm", action="store_true", help="only show retrieved sources")
    args = ap.parse_args(argv)

    if args.cmd == "index":
        chunks = ingest(args.folders)
        Index.build(chunks).save()
        print(f"Indexed {len(chunks)} chunks from {len({c.source for c in chunks})} reports into {INDEX_DIR}/")
        return

    hits = Index.load().search(args.question, k=args.k)
    if args.no_llm or not has_credentials():
        if not args.no_llm:
            print("(No ANTHROPIC_API_KEY set - showing retrieved sources only.)\n")
        for n, (c, score) in enumerate(hits, 1):
            print(f"[{n}] {c.source}  (score {score:.4f})\n{c.text}\n")
        return
    print(answer(args.question, hits))
    print("\nSources: " + ", ".join(f"[{n}] {c.source}" for n, (c, _) in enumerate(hits, 1)))


if __name__ == "__main__":
    sys.exit(main())
