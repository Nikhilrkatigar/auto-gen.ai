# auto-gen.ai: Ask the Audit Reports

Ask plain-English questions about bank concurrent audit reports and get answers with source citations.

> "Which staff account at HILLTOP was debited 3.6 lakh?"
> → *Staff account XXXXXX7618 (P. D.) was debited ₹3,60,000 on 19/05/2026 [1].*

This is a retrieval-augmented generation (RAG) layer that I built on top of a report generator I wrote for chartered accountant firms. That generator turns daily core-banking exports into daily, weekly and monthly concurrent audit reports in Word format. Auditors then need to find things across dozens of these reports ("which staff accounts crossed ₹45,000 this month?"). This project lets them ask instead of scrolling.

The real reports contain customer data and are **not** in this repository. `make_sample_reports.py` generates synthetic reports in the same layout, so everything below runs end to end.

## How it works

```
.docx reports ──► table rows ──► mask customer data ──► chunks (1 row = 1 chunk)
                                                          │
                          BM25 keyword index ◄────────────┤
                 bge-small embeddings (local) ◄───────────┘
                                                          
question ──► branch filter ──► BM25 + vector search ──► Reciprocal Rank Fusion ──► top 8 chunks
                                                                                      │
                                                     Claude answers only from sources, cites [n]
```

1. **Ingest:** each report is walked in order. Every table row becomes one chunk, prefixed with report type, branch, period and annexure title, so each chunk is a self-contained fact.
2. **Mask:** account, CIF and mobile numbers keep only their last 4 digits, and customer names become initials. This happens *before* indexing, so no raw customer data reaches the index or the LLM API.
3. **Index:** a BM25 keyword index plus `BAAI/bge-small-en-v1.5` embeddings, which run locally through fastembed and are stored with NumPy.
4. **Retrieve:** a branch named in the question becomes a hard metadata filter. Keyword and vector rankings are merged with Reciprocal Rank Fusion. Indian amounts are normalised ("3.6 lakh" → `360000`).
5. **Answer:** the top chunks go to Claude (`claude-opus-5-5`) with instructions to answer only from the sources, cite them as `[n]`, and say so when the answer isn't there.

## Quick start

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

python make_sample_reports.py        # 12 synthetic reports -> sample_reports/
python rag.py index sample_reports   # build the index -> .rag_index/
python rag.py ask --no-llm "Which staff accounts at MARKET YARD had transactions of 74,000?"

export ANTHROPIC_API_KEY=...          # optional: generated answers with citations
python rag.py ask "Which housing loans at GREENFIELD are charged above the product rate?"

python rag_eval.py                   # retrieval metrics
python -m pytest test_rag.py         # offline unit tests (no model download, no API calls)
```

Without an API key, `rag.py ask` prints the retrieved sources instead of an answer.

## Evaluation

Retrieval is scored on a fixed question set (`rag_eval_questions.json`). A hit means a chunk from the right branch that contains the expected fact.

**Synthetic sample reports** (20 questions, 216 chunks, reproducible from this repo):

| Retrieval | Hit@1 | Hit@5 | MRR |
|---|---|---|---|
| BM25 only | 65% | 90% | 0.76 |
| Embeddings only | 50% | 80% | 0.61 |
| Hybrid (RRF) | 60% | 90% | 0.73 |

**Real client reports** (22 questions, 1,831 chunks from 14 reports; private data, not included):

| Retrieval | Hit@1 | Hit@5 | MRR |
|---|---|---|---|
| BM25 only | 77% | 91% | 0.83 |
| Embeddings only | 68% | 82% | 0.74 |
| Hybrid (RRF) | 77% | 91% | 0.81 |

**What this shows:** audit reports are dense with exact names, codes and amounts, so keyword search is very strong here. Embeddings alone miss exact values, but they help with paraphrased questions. The hybrid keeps keyword-level recall while staying robust to wording. Normalising Indian amounts ("3.6 lakh") raised keyword Hit@1 on the real data from 68% to 77%.

## Limitations and next steps

- Counting or totalling across many rows ("how many…", "total…") needs a structured query over the tables, not top-k retrieval.
- Banking synonyms (dormant ↔ inoperative) are still missed by both retrievers; query expansion would fix this.
- Generation quality (faithfulness, citation accuracy) is not yet measured; next is an LLM-graded eval set.

## Files

| File | Purpose |
|---|---|
| `rag.py` | Ingest, masking, chunking, hybrid index, search, Claude answering, CLI |
| `rag_eval.py`, `rag_eval_questions.json` | Retrieval evaluation (Hit@1, Hit@5, MRR) for BM25, dense and hybrid |
| `test_rag.py` | Unit tests: masking, BM25, branch filter, report parsing, answer and refusal handling |
| `make_sample_reports.py` | Synthetic report generator (fixed seed) |
| `sample_reports/` | Generated synthetic reports |

Built with Python, python-docx, NumPy, fastembed and the Anthropic SDK.
