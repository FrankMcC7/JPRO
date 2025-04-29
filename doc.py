=============================================================

Project: Fund Prospectus Analysis – 100 % OFF‑LINE EDITION

Structure: multiple logical modules in one file (split markers)

You can split this file into individual *.py files by the

=== filename.py === markers or keep it monolithic – everything

works either way.

=============================================================

--------------------------- config.py ------------------------

Central configuration – edit the two paths below so they point

to the local folders where you placed the already‑downloaded

HuggingFace models. No internet call will ever be made.

--------------------------------------------------------------

=== config.py ===

from pathlib import Path

HOME_DIR: Path = Path.home() BASE_DIR: Path = Path(file).resolve().parent

MODEL_PATHS = { "sentence_transformer": BASE_DIR / "models" / "sentence-transformers" / "all-MiniLM-L6-v2", "cross_encoder": BASE_DIR / "models" / "cross-encoder" / "ms-marco-MiniLM-L-6-v2", }

Qdrant persistence directory (created automatically)

QDRANT_DIR: Path = BASE_DIR / "qdrant_db" QDRANT_COLLECTION: str = "document_chunks"

Misc parameters

EMBED_DIM: int = 384 MAX_CHUNK: int = 1_000 OVERLAP: int = 200

--------------------------------------------------------------

------------------------ offline.py --------------------------

Helper utilities that strictly refuse to download anything

--------------------------------------------------------------

=== offline.py ===

import os import logging from typing import Dict

from config import MODEL_PATHS

logger = logging.getLogger(name)

Force transformers / datasets libraries into offline‑only mode

os.environ["TRANSFORMERS_OFFLINE"] = "1" os.environ["HF_DATASETS_OFFLINE"] = "1" os.environ["TOKENIZERS_PARALLELISM"] = "false"

Helper – verify required files exist

REQUIRED_FILES: Dict[str, str] = { "sentence_transformer": "config.json", "cross_encoder": "config.json", }

def verify_local_models() -> None: """Raise FileNotFoundError if any expected model file is absent.""" missing = [] for key, sub in REQUIRED_FILES.items(): model_dir = MODEL_PATHS[key] if not (model_dir / sub).exists(): missing.append(str(model_dir)) if missing: raise FileNotFoundError( "\n\n❌ The following model directories are missing required files:\n    " + "\n    ".join(missing) + "\nPlease download them on a machine with internet access via e.g.\n" "  huggingface-cli download sentence-transformers/all-MiniLM-L6-v2 --local-dir ./models/sentence-transformers/all-MiniLM-L6-v2\n" "and copy to the target host." ) logger.info("✓ All local models found – running fully offline")

verify_local_models()

--------------------------------------------------------------

------------------------- models.py --------------------------

Central place to load models once and share across modules

--------------------------------------------------------------

=== models.py ===

import logging from functools import lru_cache from typing import Any

import torch from sentence_transformers import SentenceTransformer, CrossEncoder from config import MODEL_PATHS from offline import verify_local_models

logger = logging.getLogger(name)

@lru_cache(maxsize=1) def get_sentence_transformer() -> SentenceTransformer: verify_local_models() logger.info("Loading SentenceTransformer from %s", MODEL_PATHS["sentence_transformer"]) return SentenceTransformer(str(MODEL_PATHS["sentence_transformer"]), local_files_only=True)

@lru_cache(maxsize=1) def get_cross_encoder() -> CrossEncoder: verify_local_models() logger.info("Loading CrossEncoder from %s", MODEL_PATHS["cross_encoder"]) return CrossEncoder(str(MODEL_PATHS["cross_encoder"]), local_files_only=True)

--------------------------------------------------------------

------------------------ vector_db.py ------------------------

Thin wrapper around Qdrant – persisted locally

--------------------------------------------------------------

=== vector_db.py ===

import logging from typing import List, Dict, Any

import numpy as np import qdrant_client from qdrant_client.http import models as qmodels

from config import QDRANT_DIR, QDRANT_COLLECTION, EMBED_DIM

logger = logging.getLogger(name)

class VectorDB: """Persisted on‑disk Qdrant collection."""

def __init__(self):
    self.client = qdrant_client.QdrantClient(path=QDRANT_DIR)
    if not self._collection_exists():
        logger.info("Creating Qdrant collection at %s", QDRANT_DIR)
        self.client.recreate_collection(
            collection_name=QDRANT_COLLECTION,
            vectors_config=qmodels.VectorParams(size=EMBED_DIM, distance=qmodels.Distance.COSINE),
        )

def _collection_exists(self) -> bool:
    try:
        self.client.get_collection(QDRANT_COLLECTION)
        return True
    except Exception:  # collection not found
        return False

# ---------- public API ----------
def upsert_doc(self, doc_id: str, texts: List[str], embeds: np.ndarray, meta: Dict[str, Any]):
    """Insert or replace a document (dedup on doc_id + chunk_index)."""
    payloads = []
    points = []
    for idx, (text, emb) in enumerate(zip(texts, embeds)):
        pid = int(f"{abs(hash(doc_id)) % 10**6}{idx:04}")  # deterministic int id
        points.append(
            qmodels.PointStruct(id=pid, vector=emb.tolist(), payload={**meta, "text": text, "chunk_index": idx})
        )
    self.client.upsert(collection_name=QDRANT_COLLECTION, points=points)

def search(self, query_vec: np.ndarray, top_k: int = 5) -> List[Dict[str, Any]]:
    res = self.client.search(collection_name=QDRANT_COLLECTION, query_vector=query_vec.tolist(), limit=top_k)
    return [
        {
            **r.payload,
            "score": r.score,
            "document_id": r.payload.get("document_id"),
        }
        for r in res
    ]

--------------------------------------------------------------

--------------------- pdf_processing.py ----------------------

Extract text, detect structure, chunk intelligently

--------------------------------------------------------------

=== pdf_processing.py ===

import logging import re from typing import Tuple, List, Dict

import fitz  # PyMuPDF from config import MAX_CHUNK, OVERLAP

logger = logging.getLogger(name)

SECTION_PATTERNS = { "investment_objective": [r"investment\s+objective", r"fund\s+objective"], "fees_and_expenses": [r"fee", r"expense", r"charge"], "principal_risks": [r"risk", r"principal\s+risk"], "performance": [r"performance", r"return", r"yield"], }

class PDFProcessor: def extract(self, path: str) -> Tuple[str, List[Dict]]: """Return full text and structure list.""" try: doc = fitz.open(path) except Exception as e: logger.error("PyMuPDF could not open %s: %s", path, e) return "", []

full_text = []
    structure = []
    char_pos = 0
    for page_num, page in enumerate(doc):
        page_text = page.get_text()
        # crude heading detection by capitals & font size
        for match in re.finditer(r"[A-Z][A-Z\s]{5,}", page_text):
            heading = match.group(0).strip()
            section_type = self._classify_heading(heading)
            structure.append({"type": "heading", "text": heading, "char_position": char_pos + match.start(), "section_type": section_type})
        full_text.append(page_text)
        char_pos += len(page_text) + 1
    full_text = "\n".join(full_text)
    structure.sort(key=lambda x: x["char_position"])
    return full_text, structure

def _classify_heading(self, heading: str) -> str:
    hlow = heading.lower()
    for sec, pats in SECTION_PATTERNS.items():
        if any(re.search(p, hlow) for p in pats):
            return sec
    return "other"

# ------------- chunking -------------
def chunk(self, text: str, structure: List[Dict]) -> Tuple[List[str], List[Dict]]:
    if not structure:
        return self._simple_chunk(text), [{"section_type": "unknown", "start": 0, "end": len(text)}]
    # Build boundaries
    boundaries = [0] + [s["char_position"] for s in structure] + [len(text)]
    chunks: List[str] = []
    meta: List[Dict] = []
    for i in range(len(boundaries) - 1):
        part = text[boundaries[i] : boundaries[i + 1]]
        sec_type = structure[i - 1]["section_type"] if i else "introduction"
        for ch in self._simple_chunk(part):
            start = text.find(ch, boundaries[i])
            chunks.append(ch)
            meta.append({"section_type": sec_type, "start": start, "end": start + len(ch)})
    return chunks, meta

def _simple_chunk(self, text: str) -> List[str]:
    out = []
    start = 0
    while start < len(text):
        end = min(start + MAX_CHUNK, len(text))
        if end < len(text):
            bp = text.rfind(".", start, end)
            end = bp + 1 if bp > start + 100 else end
        out.append(text[start:end])
        start = end - OVERLAP if end < len(text) else len(text)
    return out

--------------------------------------------------------------

-------------------------- search.py -------------------------

Hybrid semantic + keyword search with section boosts

--------------------------------------------------------------

=== search.py ===

import logging from typing import List, Dict, Optional

import numpy as np from sklearn.feature_extraction.text import TfidfVectorizer from sklearn.metrics.pairwise import cosine_similarity

from vector_db import VectorDB from models import get_sentence_transformer

logger = logging.getLogger(name)

class HybridSearch: def init(self): self.vdb = VectorDB() self.embedder = get_sentence_transformer() self.tfidf = TfidfVectorizer(max_df=0.85, min_df=2, stop_words="english") self._corpus = []  # (doc_id, chunk_idx, text) self._matrix = None

# ---------- indexing ----------
def index(self, doc_id: str, chunks: List[str], embeds: np.ndarray, meta: Dict):
    self.vdb.upsert_doc(doc_id, chunks, embeds, meta)
    self._corpus.extend([(doc_id, i, ch) for i, ch in enumerate(chunks)])
    texts = [c[2] for c in self._corpus]
    self._matrix = self.tfidf.fit_transform(texts) if texts else None

# ---------- search ----------
def query(self, q: str, k: int = 5, section_filter: Optional[str] = None) -> List[Dict]:
    qvec = self.embedder.encode(q)
    sem_res = self.vdb.search(qvec, top_k=k * 2)
    kw_res: List[Dict] = []
    if self._matrix is not None and len(self._corpus) > 0:
        qtf = self.tfidf.transform([q])
        sims = cosine_similarity(qtf, self._matrix)[0]
        idxs = sims.argsort()[-k * 2 :][::-1]
        for idx in idxs:
            did, cidx, text = self._corpus[idx]
            kw_res.append({"document_id": did, "chunk_index": cidx, "text": text, "similarity": sims[idx]})
    merged: Dict[str, Dict] = {}
    for res in sem_res + kw_res:
        if section_filter and res.get("section_type") != section_filter:
            continue
        key = f"{res['document_id']}_{res['chunk_index']}"
        score = res.get("score", 0) + res.get("similarity", 0)
        if key not in merged or score > merged[key]["final_score"]:
            merged[key] = {**res, "final_score": score}
    return sorted(merged.values(), key=lambda x: x["final_score"], reverse=True)[:k]

--------------------------------------------------------------

-------------------- answer_generation.py --------------------

Thin wrapper using CrossEncoder and rule‑based templates

--------------------------------------------------------------

=== answer_generation.py ===

import logging from typing import List, Dict, Any

from models import get_cross_encoder

logger = logging.getLogger(name)

class AnswerGenerator: def init(self): self.re_ranker = get_cross_encoder()

def generate(self, query: str, passages: List[Dict[str, Any]]) -> str:
    if not passages:
        return "No relevant information found."
    pairs = [[query, p["text"]] for p in passages]
    scores = self.re_ranker.predict(pairs)
    top = passages[int(scores.argmax())]
    return f"Most relevant excerpt (score={scores.max():.2f})\n\n{top['text']}"

--------------------------------------------------------------

----------------------- query_engine.py ----------------------

Orchestrates search + answer; logs feedback

--------------------------------------------------------------

=== query_engine.py ===

import json import logging from datetime import datetime from pathlib import Path from typing import Dict

from answer_generation import AnswerGenerator from search import HybridSearch

logger = logging.getLogger(name)

class QueryEngine: def init(self, workdir: Path): self.searcher = HybridSearch() self.ans_gen = AnswerGenerator() self.log_path = workdir / "query_log.json" self.log = self._load()

def _load(self):
    if self.log_path.exists():
        return json.loads(self.log_path.read_text())
    return []

def _save(self):
    self.log_path.write_text(json.dumps(self.log, indent=2))

def ask(self, question: str) -> Dict:
    res = self.searcher.query(question, k=5)
    answer = self.ans_gen.generate(question, res)
    entry = {"q": question, "t": datetime.now().isoformat(), "n_hits": len(res)}
    self.log.append(entry)
    self._save()
    return {"answer": answer, "hits": res}

--------------------------------------------------------------

------------------------- learning.py ------------------------

Very small placeholder – can be extended later

--------------------------------------------------------------

=== learning.py ===

class LearningSystem: def init(self): pass  # left as exercise – logic similar to original

--------------------------------------------------------------

--------------------- streamlit_app.py -----------------------

Minimal Streamlit UI; hot‑reload friendly

--------------------------------------------------------------

=== streamlit_app.py ===

import streamlit as st import tempfile import numpy as np from pathlib import Path

from pdf_processing import PDFProcessor from models import get_sentence_transformer from search import HybridSearch from query_engine import QueryEngine from config import BASE_DIR

st.set_page_config(page_title="Offline Fund Prospectus QA", layout="wide")

Persistent objects in session

if "engine" not in st.session_state: st.session_state["engine"] = QueryEngine(BASE_DIR) st.session_state["pdf"] = PDFProcessor() st.session_state["search"] = HybridSearch() st.session_state["embedder"] = get_sentence_transformer()

engine: QueryEngine = st.session_state["engine"] processor: PDFProcessor = st.session_state["pdf"] searcher: HybridSearch = st.session_state["search"] embedder = st.session_state["embedder"]

----------- Sidebar upload -----------

with st.sidebar: st.header("Upload PDF") up = st.file_uploader("Prospectus PDF", type="pdf") if up is not None: with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tf: tf.write(up.read()) tmp_path = tf.name text, struct = processor.extract(tmp_path) if text: chunks, meta = processor.chunk(text, struct) embeds = embedder.encode(chunks, batch_size=32, show_progress_bar=True) meta_base = { "document_id": up.name, "filename": up.name, } searcher.index(up.name, chunks, np.array(embeds), meta_base) st.success(f"Indexed {len(chunks)} chunks from {up.name}")

-------------- Main ---------------

st.title("Ask about your fund") q = st.text_input("Your question") if q: out = engine.ask(q) st.markdown(out["answer"]) with st.expander("Show passages"): for h in out["hits"]: st.markdown(f"{h['filename']} – score {h['final_score']:.3f}\n\n> {h['text'][:300]} …")

--------------------------------------------------------------

-------------------- requirements.txt -----------------------

(save as separate file) – all libs are available via pip wheels

--------------------------------------------------------------

fitz==1.23.9           # PyMuPDF

sentence-transformers==2.6.1

qdrant-client==1.9.1

scikit-learn==1.4.2

streamlit==1.35.0

numpy==1.26.4

torch==2.2.2           # pre‑download wheel for the target CPU/GPU

--------------------------------------------------------------

