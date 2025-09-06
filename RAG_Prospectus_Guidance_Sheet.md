# Local CPU‑Only RAG for Prospectus Q&A — Guidance Sheet for ChatGPT Codex

**Purpose:** Build an offline, CPU‑only Retrieval‑Augmented Generation (RAG) system that extracts six fund fields from prospectuses and related documents (PDF, DOCX, XLSX, MSG). Codex should follow this as an **actionable build plan**: create folders, write scripts, wire a CLI/UI, and produce evaluable outputs with citations.

**Target fields (per fund):** Investment Objective; Permitted Investments (+major asset class); Restricted Investments; Benchmark; Global Exposure Method & Maximum Derivative Leverage; Redemption Frequency.

---

## 1) Global Constraints & Success Criteria
- **Privacy:** 100% local; no internet access required at runtime.
- **Hardware:** CPU‑only (Windows‑friendly). Avoid GPU dependencies.
- **Models:** Embedding Gemma (Sentence‑Transformers or ONNX); Gemma 3n quantized via `llama-cpp-python`.
- **Accuracy policy:** Prefer **deterministic extraction** (regex/section anchors) over LLM. If absent, return **“Not disclosed in the provided pages.”**
- **Citations:** Every extracted field must include `{doc_id, page}`. Keep a provenance log (file path, page range, char span if available).
- **Success KPI:** ≥85% exact match on a gold set; zero hallucinations; stable latency on 10k chunks (hybrid top‑k under 2s on typical desktop CPU).

---

## 2) Repository & Folder Structure (Codex: create automatically)
```
rag_prospectus/
  app/
    cli.py
    ui.py
  configs/
    config.yaml
  data/
    raw/            # pdf, docx, xlsx, msg drop-zone
    processed/      # normalized jsonl, table snapshots
  index/
    faiss/
    bm25/
  logs/
  models/           # local .gguf and embedding assets
  src/
    ingest/
      pdf.py
      docx.py
      xlsx.py
      msg.py
    chunk/
      rules.py
      tables.py
    embed/
      embedding_gemma.py
    index/
      faiss_store.py
      bm25.py
      hybrid.py
    extract/
      regex.py
      llm.py
      normalize.py
      validate.py
    qa/
      prompt.py
      answer.py
    eval/
      dataset.csv
      eval.py
      report.py
  tests/
    test_regex.py
    test_ingest.py
    test_pipeline.py
  README.md
```

> **Instruction to Codex:** Detect current working directory and create this structure if missing. Use relative paths from repo root.

---

## 3) Environment Setup (Codex: generate and run commands)
- Python ≥ 3.11; set up venv.
- Install minimal dependencies:
  - `pymupdf` (PDF), `python-docx` (DOCX), `pandas`, `openpyxl` (XLSX), `extract-msg` (MSG)
  - `faiss-cpu`, `rank-bm25` (or `tantivy`), `chromadb` (optional alt store)
  - `sentence-transformers`, `onnxruntime`
  - `llama-cpp-python`
  - `pydantic`, `typer`, `rich`, `pyyaml`, `rapidfuzz`
  - Optional UI: `fastapi`, `uvicorn` or `streamlit`

Codex should write a `requirements.txt` and a Windows‑friendly `setup.ps1` that creates the venv and installs packages.

---

## 4) Configuration (Codex: create `configs/config.yaml`)
```yaml
paths:
  raw: "./data/raw"
  processed: "./data/processed"
  index: "./index"
  models: "./models"
models:
  embedding:
    name: "embedding-gemma"
    provider: "sentence-transformers"     # "onnx" also supported
    batch_size: 64
  generator:
    name: "gemma-3n"
    backend: "llama-cpp-python"
    model_path: "./models/gemma-3n-q4.gguf"
    n_ctx: 4096
    n_threads: 8
    n_batch: 128
    temperature: 0.0
index:
  vector_store: "faiss"    # alt: "chroma"
  top_k: 12
  hybrid:
    enable: true
    bm25_weight: 0.55
    vector_weight: 0.45
  rerank:
    enable: false
chunking:
  max_chars: 2800
  min_chars: 800
  overlap: 200
  table_as_chunk: true
schema:
  require_fields: ["investment_objective","permitted_investments","restricted_investments","benchmark","global_exposure","redemption_frequency"]
logging:
  level: "INFO"
```

---

## 5) Ingestion Layer (Codex: implement `src/ingest/*.py`)
**Goals:** Parse content into normalized JSONL with page numbers, section hints, and metadata.

- **PDF (`pdf.py`)**
  - Use PyMuPDF to iterate pages, extract text and blocks; capture headings via font size/style heuristics.
  - Attempt table capture: PyMuPDF’s structured blocks; if weak, add a simple grid detector by whitespace columns; flag table chunks with `table_flag=true`.
- **DOCX (`docx.py`)**
  - `python-docx` to read paragraphs and heading levels; join soft line breaks; keep list/bullet structure.
- **XLSX (`xlsx.py`)**
  - Load each sheet with `pandas`; produce row‑wise JSON and store as table‑chunks with sheet, row, and header metadata.
- **MSG (`msg.py`)**
  - `extract_msg` to parse subject, sent date, body; export attachments into `data/raw`; append an entry that references attachment paths.

**Output JSONL Record (all ingestors):**
```json
{
  "doc_id": "basename.ext",
  "page": 12,
  "section_path": "Investment Objective",
  "text": "…",
  "table_flag": false,
  "metadata": {
    "fund_name": "",
    "isin": "",
    "lei": "",
    "doc_type": "Prospectus|KID|Addendum|Other",
    "effective_date": ""
  }
}
```

Codex should write a small header parser (first 3 pages) to guess `fund_name`, possible `isin/lei`, and `effective_date` using regex and store them in `metadata` for all chunks from that document.

---

## 6) Chunking Strategy (Codex: implement `src/chunk/rules.py` & `tables.py`)
- Split by **top headings**: `Investment Objective`, `Investment Policy`, `Eligible/Permitted Investments`, `Investment Restrictions`, `Global Exposure`, `Subscriptions`, `Redemptions`, `Benchmark`, `Risk Management` (case‑insensitive).
- Keep tables as independent chunks; limit chunk sizes to config bounds; maintain an `overlap` between adjacent text chunks.
- Preserve `section_path` of the parent heading for each chunk.

---

## 7) Embeddings & Vector Store (Codex: implement `src/embed/embedding_gemma.py` & `src/index/faiss_store.py`)
- Embed using Embedding Gemma from Sentence‑Transformers (or ONNXRuntime if configured).
- FAISS index: start with **FlatIP** for < 200k chunks; expose save/load to disk.
- Store metadata (document id, page, section_path, table_flag) alongside vectors (separate SQLite/JSON sidecar is acceptable).

---

## 8) BM25 & Hybrid Retrieval (Codex: implement `src/index/bm25.py` & `src/index/hybrid.py`)
- Build BM25 over tokenized chunks; support metadata filter.
- Hybrid search: run BM25 and FAISS (3×`top_k` each), normalize scores, fuse by weights from config, return final `top_k` ranked chunks.
- Add simple filters: by `section_path`, `fund_name`, `doc_type`, `effective_date` range.

---

## 9) Deterministic Extraction (Codex: implement `src/extract/regex.py` & `normalize.py`)
**Regex anchors (examples; make them configurable):**
- Objective: `(?i)^\s*investment objective\b[\s:\-]*([\s\S]{0,1200})`
- Permitted: `(?i)\b(permitted|eligible) investments?\b[\s:\-]*([\s\S]{0,1600})`
- Restricted: `(?i)\b(restricted|prohibited) investments?\b[\s:\-]*([\s\S]{0,1600})`
- Benchmark: `(?i)\b(benchmark|reference index)\b[\s:\-]*([^\n]+)`
- Method: `(?i)\b(commitment approach|relative\s*var|absolute\s*var|gross method)\b`
- Leverage: `(?i)\b(max(imum)? (level of )?leverage|var limit|sum of notionals)\b.*?(\d+(\.\d+)?\s?%|\d+(\.\d+)?\s*x)`
- Redemption: `(?i)\bredemption(s)?\b.*?(daily|weekly|monthly|quarterly|annually|on demand)`

**Normalization rules:**
- Map method → `{commitment|relative_var|absolute_var|gross}`.
- Benchmark: extract index name and suffixes `(TR|NR|PR)`, currency `(USD|EUR|GBP)`, and `hedged` flag if present.
- Frequency whitelist; else return `Other` plus raw text.
- Always attach `source: {doc_id, page}`.

---

## 10) LLM Fallback (Codex: implement `src/extract/llm.py` & `src/qa/prompt.py`)
- Backend: `llama-cpp-python` loading **Gemma 3n** quantized (`./models/gemma-3n-q4.gguf`).
- Generation settings: deterministic (temperature 0.0), context limited to retrieved chunks.
- **System prompt (strict):** “Use ONLY the provided context. If a field is absent, reply exactly: ‘Not disclosed in the provided pages.’ Return JSON matching the schema.”
- **User prompt template:** inject `fund_name`, `isin` and the concatenated top‑K hybrid chunks (with doc/page markers).

**Precedence rule:** If regex found a field, **do not override** with LLM. Use LLM only to fill missing fields or to structure bullet lists without altering content.

---

## 11) Validation & Guardrails (Codex: implement `src/extract/validate.py`)
- Method must be one of `{commitment, relative_var, absolute_var, gross}`; else set to “Not disclosed …”
- If method is VaR, expect leverage disclosure or a VaR limit phrase; otherwise flag `needs_review=true`.
- Benchmark cannot be empty for an ETF‑type prospectus; if missing, flag `needs_review=true`.
- Always allow abstention over conjecture.

---

## 12) End‑to‑End QA Pipeline (Codex: implement `src/qa/answer.py`)
**Flow:**
1. Build a **query** for the six fields filtered by `fund_name/ISIN` (if provided).
2. Hybrid retrieve top‑K chunks; apply deterministic extraction.
3. Run LLM fallback for missing fields; merge results.
4. Return JSON object with all fields + `source` per field and a `provenance_log` entry.
5. Optionally emit a CSV row for downstream reporting.

**JSON Contract:**
```json
{
  "fund_name":"",
  "isin":"",
  "lei":"",
  "doc_id":"",
  "doc_version_date":"YYYY-MM-DD",
  "sections":{
    "investment_objective":{"text":"", "source":{"doc_id":"", "page":0}},
    "permitted_investments":{"list":[], "major_asset_class":"", "source":{"doc_id":"", "page":0}},
    "restricted_investments":{"list":[], "source":{"doc_id":"", "page":0}},
    "benchmark":{"name":"", "type":"", "currency":"", "hedged":false, "source":{"doc_id":"", "page":0}},
    "global_exposure":{"method":"", "max_leverage":{"metric":"", "value":"", "unit":""}, "source":{"doc_id":"", "page":0}},
    "redemption_frequency":{"frequency":"", "notice":"", "cutoff":"", "source":{"doc_id":"", "page":0}}
  },
  "provenance_log":[{"field":"benchmark","doc_id":"","page":0,"span":[100,240]}],
  "needs_review": false
}
```

---

## 13) CLI & Optional UI (Codex: implement `app/cli.py`, `app/ui.py`)
- **CLI (Typer):**
  - `ingest` → scan `data/raw`, write JSONL to `data/processed`.
  - `index` → build/update FAISS + BM25.
  - `ask --fund "Name" --isin "ID"` → produce JSON (and `--csv` for a row output).
  - `eval` → run gold dataset and write `eval/report.csv` with metrics.
- **UI (Streamlit or FastAPI):**
  - Input fund/ISIN; show answer JSON and inline source snippets with page numbers.
  - Color badges: **GREEN** (regex), **AMBER** (LLM fallback), **RED** (Not disclosed).

---

## 14) Evaluation Plan (Codex: implement `src/eval/*`)
- Create `eval/dataset.csv`: 50–100 labeled examples across UCITS, US mutual funds, ETFs.
- Metrics: Hit@K (retrieval), Exact‑Match (lists string‑set), Citation Correctness (doc/page), Abstention rate.
- Output `report.csv` and a markdown summary with failure cases and field‑wise accuracy.

---

## 15) Performance & CPU Optimizations
- `llama-cpp-python`: set `n_threads` to logical CPUs; `n_batch` 64–128; enable `mmap=True` if supported.
- Prefer **Q4_K_M** quantization for Gemma 3n; upgrade to Q5 if RAM allows.
- Batch embedding (64–128) and use ONNX Runtime with `intra_op_num_threads` for speed.
- Start with FAISS **FlatIP**; migrate to IVF/Flat only when corpus exceeds ~200k chunks.

---

## 16) Logging & Audit
- Write `logs/pipeline.log` with timestamps for each step.
- Maintain a JSONL **audit trail** with conflicts (regex vs LLM), missing fields, and version upgrades.
- Provide a `--dry-run` mode for ingestion to preview detected section headings and tables.

---

## 17) What Codex Must Generate Next
1. Create the folder tree and configuration file.
2. Produce **ingestion scripts** for PDF/DOCX/XLSX/MSG and write normalized JSONL.
3. Implement **FAISS**, **BM25**, and **hybrid retrieval** utilities.
4. Implement **regex‑first extraction**, **LLM fallback**, and **validators** with the exact JSON contract.
5. Wire the **CLI** commands and (optionally) a simple UI.
6. Write **unit tests** and a tiny **gold dataset**; run the evaluation job and print the metrics table.
7. Produce a `README.md` with run commands and troubleshooting tips.

---

## 18) Runbook (Quick Commands for Users)
```
# 1) Create venv and install
python -m venv .venv && .\.venv\Scriptsctivate
pip install -r requirements.txt

# 2) Place documents in ./data/raw and models in ./models
#    e.g., gemma-3n-q4.gguf

# 3) Ingest & index
python -m app.cli ingest
python -m app.cli index

# 4) Ask for a fund
python -m app.cli ask --fund "DWS Global Natural Resources Equity Typ O" --isin "DE..." --csv out.csv

# 5) Evaluate
python -m app.cli eval
```

---

## 19) Professional Notes & Caveats
- UCITS documents normally disclose **Global Exposure Method & leverage**; SEC US funds may **not**. Prefer abstention to inference.
- For omnibus prospectuses containing multiple sub‑funds, always filter by **fund name/ISIN/LEI** to avoid cross‑fund leakage.
- Keep all data **on‑disk**; avoid sending content to remote services.

---

*End of guidance sheet. Codex should now create the repository structure, generate scripts and configuration files, and proceed through ingestion → indexing → extraction → evaluation as specified.*
