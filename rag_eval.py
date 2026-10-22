"""
rag_eval.py - measure retrieval quality of rag.py on a fixed question set.

A retrieved chunk is relevant when it comes from the expected branch (if one
is given) and contains the expected text. For each mode we report:

  * Hit@k - share of questions with at least one relevant chunk in the top k
  * MRR   - mean of 1/rank of the first relevant chunk (0 if not in top 10)

Usage::

    python rag.py index <folders...>     # build the index first
    python rag_eval.py
"""
import json
from pathlib import Path

from rag import Index

QUESTIONS = Path(__file__).with_name("rag_eval_questions.json")


def relevant(chunk, case) -> bool:
    return case["needle"] in chunk.text and case["branch"] in (None, chunk.branch)


def evaluate(index: Index, cases: list[dict], mode: str) -> dict:
    hits1 = hits5 = rr = 0.0
    misses = []
    for case in cases:
        ranked = [c for c, _ in index.search(case["q"], k=10, mode=mode)]
        rank = next((n for n, c in enumerate(ranked, 1) if relevant(c, case)), None)
        hits1 += rank == 1
        hits5 += rank is not None and rank <= 5
        rr += 1 / rank if rank else 0
        if not rank or rank > 5:
            misses.append(case["q"])
    n = len(cases)
    return {"mode": mode, "hit@1": hits1 / n, "hit@5": hits5 / n, "mrr": rr / n, "misses": misses}


def main():
    index = Index.load()
    cases = json.loads(QUESTIONS.read_text())
    for case in cases:  # guard against a stale question set
        assert any(relevant(c, case) for c in index.chunks), f"no relevant chunk for: {case['q']}"

    print(f"{len(cases)} questions, {len(index.chunks)} chunks\n")
    print(f"{'mode':8} {'Hit@1':>6} {'Hit@5':>6} {'MRR':>6}")
    results = [evaluate(index, cases, mode) for mode in ("bm25", "dense", "hybrid")]
    for r in results:
        print(f"{r['mode']:8} {r['hit@1']:6.0%} {r['hit@5']:6.0%} {r['mrr']:6.2f}")
    for r in results:
        for q in r["misses"]:
            print(f"  [{r['mode']}] missed top 5: {q}")


if __name__ == "__main__":
    main()
