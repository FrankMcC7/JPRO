=============================================================

Fund Prospectus Analysis (100 % OFF‑LINE) – Source Files

=============================================================

Copy–paste each block into a separate .py file (name shown

in the header) inside one project folder.  A minimal

README.md with setup instructions is included at the end.

-------------------------------------------------------------

──────────────────── 1. config.py ────────────────────────────

from pathlib import Path

# Root of your project (adjust if you place the repo elsewhere)
BASE_DIR = Path(__file__).resolve().parent

MODEL_PATHS = {
    "sentence_transformer": BASE_DIR / "models" / "sentence-transformers" / "all-MiniLM-L6-v2",
    "cross_encoder": BASE_DIR / "models" / "cross-encoder" / "ms-marco-MiniLM-L-6-v2",
}

QDRANT_DIR = BASE_DIR / "qdrant_db"
QDRANT_COLLECTION = "document_chunks"

EMBED_DIM = 384
MAX_CHUNK = 1_000
OVERLAP   = 200

──────────────────── 2. offline.py ──────────────────────────

"""Guarantee fully‑offline execution (no downloads allowed)."""
import os, logging
from typing import Dict
from config import MODEL_PATHS

logger = logging.getLogger(__name__)

os.environ["TRANSFORMERS_OFFLINE"]   = "1"
os.environ["HF_DATASETS_OFFLINE"]    = "1"
os.environ["TOKENIZERS_PARALLELISM"] = "false"

REQUIRED: Dict[str,str] = {
    "sentence_transformer": "config.json",
    "cross_encoder":        "config.json",
}

def verify_local_models() -> None:
    missing = []
    for key, f in REQUIRED.items():
        if not (MODEL_PATHS[key] / f).exists():
            missing.append(str(MODEL_PATHS[key]))
    if missing:
        raise FileNotFoundError(
            "\n\n❌ Missing offline model files:\n  " + "\n  ".join(missing) +
            "\nDownload them on a machine with internet using: \n"
            "  huggingface-cli download <repo> --local-dir <dest>\n"
        )
    logger.info("✓ Local models found – running offline")

verify_local_models()

──────────────────── 3. models.py ───────────────────────────

import logging
from functools import lru_cache
from sentence_transformers import SentenceTransformer, CrossEncoder
from config import MODEL_PATHS
from offline import verify_local_models

logger = logging.getLogger(__name__)

@lru_cache(maxsize=1)
def get_sentence_transformer():
    verify_local_models()
    return SentenceTransformer(str(MODEL_PATHS["sentence_transformer"]), local_files_only=True)

@lru_cache(maxsize=1)
def get_cross_encoder():
    verify_local_models()
    return CrossEncoder(str(MODEL_PATHS["cross_encoder"]), local_files_only=True)

──────────────────── 4. vector_db.py ────────────────────────

import numpy as np, logging
from typing import List, Dict, Any
import qdrant_client
from qdrant_client.http import models as qmodels
from config import QDRANT_DIR, QDRANT_COLLECTION, EMBED_DIM

logger = logging.getLogger(__name__)

class VectorDB:
    def __init__(self):
        self.client = qdrant_client.QdrantClient(path=QDRANT_DIR)
        if not self._exists():
            logger.info("Creating Qdrant collection …")
            self.client.recreate_collection(
                collection_name=QDRANT_COLLECTION,
                vectors_config=qmodels.VectorParams(size=EMBED_DIM, distance=qmodels.Distance.COSINE),
            )

    def _exists(self):
        try:
            self.client.get_collection(QDRANT_COLLECTION); return True
        except Exception:
            return False

    def upsert(self, doc_id:str, texts:List[str], embeds:np.ndarray, meta:Dict[str,Any]):
        pts=[]
        for i,(t,e) in enumerate(zip(texts, embeds)):
            pid=int(f"{abs(hash(doc_id))%10**6}{i:04}")
            pts.append(qmodels.PointStruct(id=pid, vector=e.tolist(), payload={**meta,"text":t,"chunk_index":i}))
        self.client.upsert(collection_name=QDRANT_COLLECTION, points=pts)

    def search(self, vec:np.ndarray, k:int=5):
        res=self.client.search(collection_name=QDRANT_COLLECTION, query_vector=vec.tolist(), limit=k)
        return [{**r.payload,"score":r.score} for r in res]

──────────────────── 5. pdf_processing.py ───────────────────

import fitz, re, logging
from typing import List, Dict, Tuple
from config import MAX_CHUNK, OVERLAP

logger=logging.getLogger(__name__)
PATTERNS={"investment_objective":[r"investment\s+objective"],"fees_and_expenses":[r"fee",r"expense"],"principal_risks":[r"risk"],"performance":[r"performance",r"return"]}

class PDFProcessor:
    def extract(self, path:str)->Tuple[str,List[Dict]]:
        try: doc=fitz.open(path)
        except Exception as e:
            logger.error("Cannot open %s: %s",path,e); return "",[]
        full=[]; struct=[]; pos=0
        for pg in doc:
            txt=pg.get_text();
            for m in re.finditer(r"[A-Z][A-Z\s]{5,}",txt):
                heading=m.group(0).strip()
                sec=self._cls(heading)
                struct.append({"type":"heading","text":heading,"char_position":pos+m.start(),"section_type":sec})
            full.append(txt); pos+=len(txt)+1
        struct.sort(key=lambda x:x["char_position"])
        return "\n".join(full),struct

    def _cls(self,h:str):
        l=h.lower()
        for s,ps in PATTERNS.items():
            if any(re.search(p,l) for p in ps): return s
        return "other"

    def chunk(self,text:str,struct:List[Dict]])->Tuple[List[str],List[Dict]]:
        if not struct: return self._simple(text),[{"section_type":"unknown","start":0,"end":len(text)}]
        bounds=[0]+[s["char_position"] for s in struct]+[len(text)]
        chks=[];meta=[]
        for i in range(len(bounds)-1):
            part=text[bounds[i]:bounds[i+1]]; sec=struct[i-1]["section_type"] if i else "intro"
            for c in self._simple(part):
                st=text.find(c,bounds[i]); chks.append(c); meta.append({"section_type":sec,"start":st,"end":st+len(c)})
        return chks,meta

    def _simple(self,txt:str)->List[str]:
        out=[]; s=0
        while s<len(txt):
            e=min(s+MAX_CHUNK,len(txt))
            if e<len(txt):
                p=txt.rfind('.',s,e); e=p+1 if p>s+100 else e
            out.append(txt[s:e]); s=e-OVERLAP if e<len(txt) else len(txt)
        return out

──────────────────── 6. search.py ───────────────────────────

import numpy as np, logging
from typing import List, Dict, Optional
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from models import get_sentence_transformer
from vector_db import VectorDB

logger=logging.getLogger(__name__)

class HybridSearch:
    def __init__(self):
        self.embed=get_sentence_transformer(); self.vdb=VectorDB();
        self.tfidf=TfidfVectorizer(max_df=0.85,min_df=2,stop_words='english')
        self.corpus=[]; self.mat=None

    def index(self,doc_id:str,chunks:List[str],embeds:np.ndarray,meta:Dict):
        self.vdb.upsert(doc_id,chunks,embeds,meta); self.corpus+=[(doc_id,i,c) for i,c in enumerate(chunks)]
        self.mat=self.tfidf.fit_transform([c[2] for c in self.corpus]) if self.corpus else None

    def query(self,q:str,k:int=5,section_filter:Optional[str]=None)->List[Dict]:
        qv=self.embed.encode(q); sem=self.vdb.search(qv,k*2); kw=[]
        if self.mat is not None:
            sv=self.tfidf.transform([q]); sims=cosine_similarity(sv,self.mat)[0];
            for idx in sims.argsort()[-k*2:][::-1]:
                did,i,t=self.corpus[idx]; kw.append({"document_id":did,"chunk_index":i,"text":t,"similarity":sims[idx]})
        merged={}
        for r in sem+kw:
            if section_filter and r.get("section_type")!=section_filter: continue
            key=f"{r['document_id']}_{r['chunk_index']}"; score=r.get('score',0)+r.get('similarity',0)
            if key not in merged or score>merged[key]['score']:
                merged[key]={**r,'score':score}
        return sorted(merged.values(),key=lambda x:x['score'],reverse=True)[:k]

──────────────────── 7. answer_generation.py ────────────────

from typing import List, Dict
from models import get_cross_encoder

class AnswerGenerator:
    def __init__(self):
        self.rank=get_cross_encoder()
    def generate(self,q:str,passages:List[Dict]):
        if not passages: return "No relevant info found."
        pairs=[[q,p['text']] for p in passages]; scores=self.rank.predict(pairs)
        best=passages[int(scores.argmax())]
        return f"**Top excerpt (score={scores.max():.2f})**\n\n{best['text']}"

──────────────────── 8. query_engine.py ────────────────────

import json
from datetime import datetime
from pathlib import Path
from typing import Dict
from search import HybridSearch
from answer_generation import AnswerGenerator

class QueryEngine:
    def __init__(self,workdir:Path):
        self.hs=HybridSearch(); self.ag=AnswerGenerator(); self.logp=workdir/'query_log.json';
        self.log=json.loads(self.logp.read_text()) if self.logp.exists() else []
    def _save(self): self.logp.write_text(json.dumps(self.log,indent=2))

    def ask(self,q:str)->Dict:
        hits=self.hs.query(q,5); ans=self.ag.generate(q,hits)
        self.log.append({"q":q,"t":datetime.now().isoformat(),"hits":len(hits)}); self._save()
        return {"answer":ans,"hits":hits}

──────────────────── 9. streamlit_app.py ───────────────────

import streamlit as st, tempfile, numpy as np
from pathlib import Path
from pdf_processing import PDFProcessor
from models import get_sentence_transformer
from search import HybridSearch
from query_engine import QueryEngine
from config import BASE_DIR

st.set_page_config(page_title="Offline Fund QA",layout="wide")
if 'qe' not in st.session_state:
    st.session_state['qe']=QueryEngine(BASE_DIR)
    st.session_state['pp']=PDFProcessor()
    st.session_state['hs']=HybridSearch()
    st.session_state['emb']=get_sentence_transformer()
qe,pp,hs,emb=st.session_state.values()

with st.sidebar:
    up=st.file_uploader("Upload PDF",type="pdf")
    if up:
        with tempfile.NamedTemporaryFile(delete=False,suffix='.pdf') as tf: tf.write(up.read()); path=tf.name
        txt,struct=pp.extract(path)
        if txt:
            chks,meta=pp.chunk(txt,struct); em=emb.encode(chks,batch_size=32,show_progress_bar=True)
            hs.index(up.name,chks,np.array(em),{"document_id":up.name,"filename":up.name})
            st.success(f"Indexed {len(chks)} chunks")

st.title("Ask about the fund")
q=st.text_input("Question")
if q:
    out=qe.ask(q); st.markdown(out['answer'])
    with st.expander("Passages"):
        for h in out['hits']:
            st.markdown(f"**{h.get('filename','') }** • score {h['score']:.3f}\n> {h['text'][:250]} …")

──────────────────── 10. requirements.txt ──────────────────

fitz==1.23.9
sentence-transformers==2.6.1
qdrant-client==1.9.1
scikit-learn==1.4.2
streamlit==1.35.0
numpy==1.26.4
torch==2.2.2   # download wheel matching your CPU/GPU

──────────────────── 11. README.md ─────────────────────────

# Fund Prospectus Analysis (Offline‑Only)

## 1. Folder structure

📁 project-root/ ├─ models/ │  ├─ sentence-transformers/all-MiniLM-L6-v2/ (HF files) │  └─ cross-encoder/ms-marco-MiniLM-L-6-v2/   (HF files) ├─ qdrant_db/          # auto‑created ├─ config.py ├─ offline.py ├─ models.py ├─ vector_db.py ├─ pdf_processing.py ├─ search.py ├─ answer_generation.py ├─ query_engine.py ├─ streamlit_app.py └─ requirements.txt

## 2. Install (offline machine)
```bash
# inside a venv
pip install --no-index --find-links /path/to/offline/wheels -r requirements.txt

> --find-links should point to a folder where you pre‑downloaded the wheels for the packages above (plus dependencies).



3. Run

streamlit run streamlit_app.py --server.port 8501

Open http://localhost:8501.

4. Preparing the model folders (one‑time on an online box)

huggingface-cli download sentence-transformers/all-MiniLM-L6-v2 \
  --local-dir ./models/sentence-transformers/all-MiniLM-L6-v2
huggingface-cli download cross-encoder/ms-marco-MiniLM-L-6-v2 \
  --local-dir ./models/cross-encoder/ms-marco-MiniLM-L-6-v2

Copy the models/ directory to the offline server.

> The code refuses to start if any of the two directories are missing.



Enjoy fully‑offline question answering over your own prospectuses!



