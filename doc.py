# app.py
import os
import json
import logging
import numpy as np
import pickle
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional, Union

# Force libraries to work in offline mode
os.environ['TRANSFORMERS_OFFLINE'] = '1'
os.environ['HF_DATASETS_OFFLINE'] = '1'
os.environ['TOKENIZERS_PARALLELISM'] = 'false'

# Import remaining libraries
import fitz  # PyMuPDF for better PDF extraction
import torch
from sentence_transformers import SentenceTransformer, CrossEncoder
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# Vector database
import qdrant_client
from qdrant_client.http import models as qmodels

# UI
import streamlit as st

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def verify_models_exist():
    """Verify that the required models exist in the offline cache."""
    home_dir = Path.home()
    cache_dir = home_dir / ".cache" / "huggingface" / "hub"
    
    sentence_transformer_dir = cache_dir / "models--sentence-transformers--all-MiniLM-L6-v2"
    cross_encoder_dir = cache_dir / "models--cross-encoder--ms-marco-MiniLM-L-6-v2"
    
    models_exist = True
    
    if not sentence_transformer_dir.exists():
        logger.error(f"Sentence transformer model not found in cache! Expected location: {sentence_transformer_dir}")
        models_exist = False
        
    if not cross_encoder_dir.exists():
        logger.error(f"Cross-encoder model not found in cache! Expected location: {cross_encoder_dir}")
        models_exist = False
    
    if models_exist:
        logger.info("✓ Model files found in the expected locations")
    
    return models_exist

class VectorDatabase:
    """Vector database for efficient storage and retrieval of document embeddings."""
    
    def __init__(self, collection_name: str = "document_chunks", dimension: int = 384):
        """Initialize the vector database."""
        self.collection_name = collection_name
        self.dimension = dimension
        
        # Initialize Qdrant client (using in-memory storage for simplicity)
        # In production, use persist_directory for disk storage
        self.client = qdrant_client.QdrantClient(location=":memory:")
        
        # Create collection if it doesn't exist
        try:
            self.client.get_collection(collection_name=self.collection_name)
        except Exception:
            self.client.create_collection(
                collection_name=self.collection_name,
                vectors_config=qmodels.VectorParams(
                    size=self.dimension,
                    distance=qmodels.Distance.COSINE
                )
            )
            
    def add_documents(self, 
                      document_id: str, 
                      chunks: List[str], 
                      embeddings: np.ndarray, 
                      metadata: Dict[str, Any]):
        """Add document chunks and their embeddings to the database."""
        points = []
        
        for i, (chunk, embedding) in enumerate(zip(chunks, embeddings)):
            chunk_id = f"{document_id}_{i}"
            
            # Create point metadata
            point_metadata = {
                "document_id": document_id,
                "chunk_index": i,
                "text": chunk,
                **metadata  # Include document metadata
            }
            
            # Create point
            points.append(
                qmodels.PointStruct(
                    id=chunk_id,
                    vector=embedding.tolist(),
                    payload=point_metadata
                )
            )
        
        # Batch insert points
        self.client.upsert(
            collection_name=self.collection_name,
            points=points
        )
        
        return len(points)
    
    def search(self, query_vector: np.ndarray, limit: int = 5) -> List[Dict]:
        """Search for similar vectors in the database."""
        results = self.client.search(
            collection_name=self.collection_name,
            query_vector=query_vector.tolist(),
            limit=limit
        )
        
        # Format results
        formatted_results = []
        for res in results:
            formatted_results.append({
                "document_id": res.payload["document_id"],
                "chunk_index": res.payload["chunk_index"],
                "text": res.payload["text"],
                "filename": res.payload.get("filename", "Unknown"),
                "similarity": res.score,
                "section_type": res.payload.get("section_type", None),
                "financial_data": res.payload.get("financial_data", None)
            })
            
        return formatted_results
    
    def count_documents(self) -> int:
        """Count the number of documents in the database."""
        collection_info = self.client.get_collection(collection_name=self.collection_name)
        return collection_info.points_count


class DocumentProcessor:
    """Enhanced document processor optimized for fund prospectuses."""
    
    def __init__(self, storage_dir: str = "document_storage"):
        self.storage_dir = Path(storage_dir)
        self.storage_dir.mkdir(exist_ok=True)
        
        # Document metadata storage
        self.metadata_path = self.storage_dir / "metadata.json"
        self.metadata = self._load_metadata()
        
        # Initialize sentence embedding model - configured for offline use
        logger.info("Loading sentence transformer model from local cache...")
        self.embedding_model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2', 
                                                  use_auth_token=False)
        
        # Initialize vector database
        self.vector_db = VectorDatabase()
        
        # Initialize TF-IDF vectorizer for keyword search
        self.tfidf_vectorizer = TfidfVectorizer(
            max_df=0.85,
            min_df=2,
            stop_words='english'
        )
        self.tfidf_matrix = None
        self.tfidf_documents = []
        
        # Financial term dictionary for fund prospectuses
        self.financial_terms = {
            "expense ratio", "management fee", "nav", "net asset value", 
            "front-end load", "back-end load", "redemption fee", "12b-1 fee",
            "benchmark", "yield", "distribution", "dividend", "capital gain",
            "shareholder", "prospectus", "fund", "etf", "mutual fund",
            "performance", "risk", "volatility", "beta", "alpha",
            "asset allocation", "portfolio", "investment objective",
            "advisor", "manager", "custodian", "trustee", "administrator",
            "fiscal year", "class a", "class b", "class c", "class i",
            "institutional", "investor", "retail", "annual report",
            "semi-annual", "tax", "turnover", "liquidity", "derivative",
            "principal", "interest", "maturity", "duration", "credit quality",
            "leverage", "leveraged", "borrowing", "margin", "debt", 
            "gearing", "exposure", "notional", "130/30", "short", "long/short",
            "gross exposure", "net exposure", "derivatives", "swaps", "futures"
        }
        
        # Define financial document section types and their patterns
        self.section_patterns = {
            "investment_objective": [r"investment\s+objective", r"fund\s+objective", r"objective\s+and\s+goal"],
            "fees_and_expenses": [r"fee", r"expense", r"charge", r"load", r"commission", r"transaction\s+cost"],
            "principal_risks": [r"risk", r"principal\s+risk", r"risk\s+factor"],
            "performance": [r"performance", r"return", r"yield", r"history", r"historical\s+result"],
            "management": [r"management", r"advisor", r"portfolio\s+manager", r"investment\s+team"],
            "purchase_and_sale": [r"purchase", r"sale", r"buy", r"sell", r"redemption", r"exchange"],
            "tax_information": [r"tax", r"dividend", r"distribution", r"capital\s+gain"],
            "asset_allocation": [r"asset\s+allocation", r"portfolio\s+composition", r"holding", r"sector"],
            "leverage": [r"leverage", r"borrowing", r"debt", r"derivative", r"notional\s+exposure", r"short", r"long/short"]
        }
        
    def _load_metadata(self) -> Dict:
        """Load document metadata from storage."""
        if self.metadata_path.exists():
            with open(self.metadata_path, 'r') as f:
                return json.load(f)
        return {"documents": {}}
    
    def _save_metadata(self):
        """Save document metadata to storage."""
        with open(self.metadata_path, 'w') as f:
            json.dump(self.metadata, f, indent=2)
    
    def process_pdf(self, pdf_file, filename: str) -> str:
        """
        Process a fund prospectus PDF with specialized extraction.
        Returns the document_id for the processed document.
        """
        # Generate a unique ID for the document
        document_id = f"doc_{len(self.metadata['documents']) + 1}"
        
        # Create a temporary file to work with PyMuPDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            temp_file.write(pdf_file.read())
            temp_path = temp_file.name
        
        try:
            # Extract text and structure from PDF
            text, structure = self._extract_from_pdf(temp_path)
            
            # Create document chunks with structural information
            chunks, chunk_metadata = self._chunk_prospectus_with_structure(text, structure)
            
            # Create embeddings for chunks
            embeddings = self._embed_chunks(chunks)
            
            # Extract financial data from text
            financial_data = self.extract_financial_data(text)
            
            # Store document data
            doc_dir = self.storage_dir / document_id
            doc_dir.mkdir(exist_ok=True)
            
            # Save text content
            with open(doc_dir / "content.txt", 'w', encoding='utf-8') as f:
                f.write(text)
            
            # Save chunks
            with open(doc_dir / "chunks.json", 'w', encoding='utf-8') as f:
                json.dump(chunks, f, ensure_ascii=False, indent=2)
            
            # Save structure
            with open(doc_dir / "structure.json", 'w', encoding='utf-8') as f:
                json.dump(structure, f, ensure_ascii=False, indent=2)
                
            # Save financial data
            with open(doc_dir / "financial_data.json", 'w', encoding='utf-8') as f:
                json.dump(financial_data, f, ensure_ascii=False, indent=2)
            
            # Update metadata
            doc_metadata = {
                "filename": filename,
                "upload_date": datetime.now().isoformat(),
                "num_chunks": len(chunks),
                "structure_elements": len(structure),
                "financial_data_points": len(financial_data),
                "queries": []  # Track queries about this document
            }
            
            self.metadata["documents"][document_id] = doc_metadata
            self._save_metadata()
            
            # Add to vector database
            # Assign financial data to appropriate chunks
            for i, chunk in enumerate(chunks):
                chunk_financial_data = []
                chunk_start = chunk_metadata[i].get("start_pos", 0)
                chunk_end = chunk_metadata[i].get("end_pos", len(text))
                
                # Find financial data points in this chunk
                for data_point in financial_data:
                    if chunk_start <= data_point["position"] < chunk_end:
                        chunk_financial_data.append(data_point)
                
                # Create payload with metadata
                payload = {
                    "filename": filename,
                    "section_type": chunk_metadata[i].get("section_type", "unknown"),
                    "financial_data": chunk_financial_data,
                    **doc_metadata
                }
                
                # Add to vector database
                self.vector_db.add_documents(
                    document_id=document_id,
                    chunks=[chunk],
                    embeddings=embeddings[i:i+1],
                    metadata=payload
                )
            
            # Update TF-IDF model
            self.tfidf_documents.extend([(document_id, i, chunk) for i, chunk in enumerate(chunks)])
            self._update_tfidf_model()
            
            return document_id
            
        finally:
            # Clean up temporary file
            os.unlink(temp_path)
    
    def extract_financial_data(self, text):
        """Extract financial values with their context."""
        financial_values = []
        
        # Match percentage values (e.g., 1.25%)
        percentage_matches = re.finditer(r'(\d+\.\d+|\d+)%', text)
        for match in percentage_matches:
            value = float(match.group(1))
            # Get surrounding text (50 chars before, 50 after)
            start = max(0, match.start() - 50)
            end = min(len(text), match.end() + 50)
            context = text[start:end]
            
            # Classify the type of percentage
            data_type = "generic_percentage"
            context_lower = context.lower()
            
            if any(term in context_lower for term in ["expense", "fee", "ratio"]):
                data_type = "expense_ratio"
            elif any(term in context_lower for term in ["yield", "return", "performance"]):
                data_type = "yield_return"
            elif any(term in context_lower for term in ["turnover", "portfolio turnover"]):
                data_type = "turnover_rate"
            elif any(term in context_lower for term in ["leverage", "borrowing", "exposure"]):
                data_type = "leverage_ratio"
            
            financial_values.append({
                "value": value,
                "unit": "%",
                "type": data_type,
                "context": context,
                "position": match.start()
            })
        
        # Match dollar values (e.g., $1,234.56)
        dollar_matches = re.finditer(r'\$(\d{1,3}(?:,\d{3})*(?:\.\d+)?)', text)
        for match in dollar_matches:
            # Remove commas for conversion
            value_str = match.group(1).replace(',', '')
            value = float(value_str)
            start = max(0, match.start() - 50)
            end = min(len(text), match.end() + 50)
            context = text[start:end]
            
            # Classify the type of dollar value
            data_type = "generic_amount"
            context_lower = context.lower()
            
            if any(term in context_lower for term in ["minimum", "investment", "initial"]):
                data_type = "minimum_investment"
            elif any(term in context_lower for term in ["nav", "net asset value"]):
                data_type = "nav"
            elif any(term in context_lower for term in ["fee", "expense", "charge"]):
                data_type = "fee_amount"
            elif any(term in context_lower for term in ["leverage", "borrowing", "debt"]):
                data_type = "leverage_amount"
            
            financial_values.append({
                "value": value,
                "unit": "$",
                "type": data_type,
                "context": context,
                "position": match.start()
            })
            
        # Match number ranges (e.g., 0-5 years)
        year_range_matches = re.finditer(r'(\d+)(?:\s*-\s*|\s+to\s+)(\d+)\s+years?', text)
        for match in year_range_matches:
            start_val = int(match.group(1))
            end_val = int(match.group(2))
            context_start = max(0, match.start() - 50)
            context_end = min(len(text), match.end() + 50)
            context = text[context_start:context_end]
            
            financial_values.append({
                "value": [start_val, end_val],
                "unit": "years",
                "type": "time_range",
                "context": context,
                "position": match.start()
            })
            
        # Match standalone numbers that might be ratios, etc.
        ratio_matches = re.finditer(r'(?:ratio|multiple|factor)(?:\s+of)?\s+(\d+\.\d+|\d+)', text, re.IGNORECASE)
        for match in ratio_matches:
            value = float(match.group(1))
            context_start = max(0, match.start() - 50)
            context_end = min(len(text), match.end() + 50)
            context = text[context_start:context_end]
            
            financial_values.append({
                "value": value,
                "unit": "ratio",
                "type": "financial_ratio",
                "context": context,
                "position": match.start()
            })
            
        # Specifically match leverage ratios and values
        leverage_patterns = [
            # Match common leverage formats like "2x", "3x leverage", etc.
            (r'(\d+)(?:x|\s*times|\s*×)\s*(?:leverage|exposure)?', 'multiplier'),
            # Match leverage ratios like "150% of NAV", "130/30 strategy"
            (r'(\d+)(?:/\d+|\s*percent|\s*%)\s*(?:leverage|exposure|strategy|of\s+nav)', 'percentage'),
            # Match notional exposure values
            (r'notional\s+(?:exposure|value)\s+of\s+(?:\$\s*)?(\d[\d,.]*)', 'notional')
        ]
        
        for pattern, unit in leverage_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                value_str = match.group(1).replace(',', '')
                try:
                    value = float(value_str)
                    context_start = max(0, match.start() - 50)
                    context_end = min(len(text), match.end() + 50)
                    context = text[context_start:context_end]
                    
                    financial_values.append({
                        "value": value,
                        "unit": unit,
                        "type": "leverage",
                        "context": context,
                        "position": match.start()
                    })
                except ValueError:
                    pass  # Skip if we can't convert to float
        
        return financial_values
    
    def _update_tfidf_model(self):
        """Update the TF-IDF model with all current documents."""
        if not self.tfidf_documents:
            return
            
        # Extract just the text for the vectorizer
        texts = [doc[2] for doc in self.tfidf_documents]
        
        # Fit the vectorizer
        self.tfidf_matrix = self.tfidf_vectorizer.fit_transform(texts)
    
    def _extract_from_pdf(self, pdf_path: str) -> Tuple[str, List[Dict]]:
        """
        Extract text and structure from a fund prospectus PDF.
        Returns a tuple of (full_text, structure_info).
        """
        try:
            # Open the PDF
            doc = fitz.open(pdf_path)
            full_text = ""
            structure = []
            
            for page_num, page in enumerate(doc):
                # Extract text with layout information
                blocks = page.get_text("dict")["blocks"]
                
                page_text = ""
                
                for block_num, block in enumerate(blocks):
                    if block["type"] == 0:  # Text block
                        for line in block["lines"]:
                            line_text = ""
                            for span in line["spans"]:
                                line_text += span["text"]
                                
                                # Track structural elements (headings, etc.)
                                if span["size"] > 12 or span["flags"] & 16:  # Larger font or bold text
                                    heading_text = span["text"].strip()
                                    if heading_text and len(heading_text) > 3:  # Avoid short headings
                                        # Determine section type based on text
                                        section_type = self._classify_section_heading(heading_text)
                                        
                                        structure.append({
                                            "type": "heading",
                                            "text": heading_text,
                                            "section_type": section_type,
                                            "page": page_num + 1,
                                            "position": {
                                                "x0": span["bbox"][0],
                                                "y0": span["bbox"][1],
                                                "x1": span["bbox"][2],
                                                "y1": span["bbox"][3]
                                            },
                                            "char_position": len(full_text) + len(page_text)
                                        })
                            
                            page_text += line_text + "\n"
                    
                    elif block["type"] == 1:  # Image block
                        structure.append({
                            "type": "image",
                            "page": page_num + 1,
                            "position": {
                                "x0": block["bbox"][0],
                                "y0": block["bbox"][1],
                                "x1": block["bbox"][2],
                                "y1": block["bbox"][3]
                            },
                            "char_position": len(full_text) + len(page_text)
                        })
                
                # Detect tables (important for fee tables, performance tables)
                table_info = self._detect_tables(page)
                if table_info:
                    for table in table_info:
                        table["page"] = page_num + 1
                        table["char_position"] = len(full_text) + len(page_text)
                        structure.append(table)
                
                full_text += page_text + "\n\n"
            
            # Extract potential section headers using common prospectus patterns
            prospectus_headers = [
                r"INVESTMENT\s+OBJECTIVE",
                r"FEES\s+AND\s+EXPENSES",
                r"PRINCIPAL\s+(?:INVESTMENT\s+)?STRATEGIES",
                r"PRINCIPAL\s+RISKS",
                r"PERFORMANCE",
                r"PORTFOLIO\s+MANAGEMENT",
                r"PURCHASE\s+AND\s+SALE\s+OF\s+FUND\s+SHARES",
                r"TAX\s+INFORMATION",
                r"PAYMENTS\s+TO\s+BROKER-DEALERS"
            ]
            
            for pattern in prospectus_headers:
                for match in re.finditer(pattern, full_text, re.IGNORECASE):
                    heading_text = match.group(0)
                    section_type = self._classify_section_heading(heading_text)
                    
                    # Check if this heading is already in structure (avoid duplicates)
                    already_exists = False
                    for s in structure:
                        if s.get("type") == "heading" and s.get("char_position", -1) == match.start():
                            already_exists = True
                            break
                    
                    if not already_exists:
                        structure.append({
                            "type": "heading",
                            "text": heading_text,
                            "section_type": section_type,
                            "char_position": match.start()
                        })
            
            # Sort structure elements by position in document
            structure.sort(key=lambda x: x.get("char_position", 0))
            
            return full_text, structure
            
        except Exception as e:
            logger.error(f"Error extracting from PDF: {e}")
            return "", []
    
    def _classify_section_heading(self, heading_text):
        """Classify a heading into a prospectus section type."""
        heading_lower = heading_text.lower()
        
        for section_type, patterns in self.section_patterns.items():
            for pattern in patterns:
                if re.search(pattern, heading_lower, re.IGNORECASE):
                    return section_type
        
        return "other"
    
    def _detect_tables(self, page) -> List[Dict]:
        """
        Enhanced table detection for fund prospectuses.
        Returns a list of detected tables with metadata.
        """
        # Get horizontal and vertical lines
        drawings = page.get_drawings()
        if not drawings:
            return []
            
        # Count horizontal and vertical lines
        h_lines = 0
        v_lines = 0
        
        for drawing in drawings:
            for item in drawing["items"]:
                if item["type"] == "l":  # Line
                    x0, y0, x1, y1 = item["rect"]
                    if abs(y1 - y0) < 2:  # Horizontal line
                        h_lines += 1
                    if abs(x1 - x0) < 2:  # Vertical line
                        v_lines += 1
        
        # If we have multiple horizontal and vertical lines, likely a table
        if h_lines > 2 and v_lines > 2:
            # Get text in the vicinity of the table
            page_text = page.get_text()
            
            # Try to classify the table type based on its content
            table_type = "unknown"
            
            if re.search(r"fee|expense|charge", page_text, re.IGNORECASE):
                table_type = "fee_table"
            elif re.search(r"performance|return|average|annualized", page_text, re.IGNORECASE):
                table_type = "performance_table"
            elif re.search(r"allocation|holding|portfolio|sector", page_text, re.IGNORECASE):
                table_type = "allocation_table"
            elif re.search(r"financial\s+highlight|data\s+per\s+share", page_text, re.IGNORECASE):
                table_type = "financial_highlights"
                
            return [{
                "type": "table",
                "table_type": table_type,
                "lines": {"horizontal": h_lines, "vertical": v_lines}
            }]
        
        return []
    
    def _chunk_prospectus_with_structure(self, text: str, structure: List[Dict], 
                                        max_chunk_size: int = 1000, 
                                        overlap: int = 200) -> Tuple[List[str], List[Dict]]:
        """
        Split prospectus into chunks, respecting document structure.
        Returns both chunks and metadata about each chunk.
        """
        # Extract section boundaries from headings
        section_boundaries = []
        
        for item in structure:
            if item.get("type") == "heading" and "char_position" in item:
                section_boundaries.append({
                    "position": item["char_position"],
                    "text": item["text"],
                    "section_type": item.get("section_type", "unknown")
                })
        
        # Sort boundaries
        section_boundaries.sort(key=lambda x: x["position"])
        
        # If no structure found, fall back to regular chunking
        if not section_boundaries:
            chunks = self._chunk_text(text, max_chunk_size, overlap)
            chunk_metadata = [{"section_type": "unknown", "start_pos": 0, "end_pos": len(text)} for _ in chunks]
            return chunks, chunk_metadata
        
        # Create chunks based on section boundaries
        chunks = []
        chunk_metadata = []
        
        # Add document start to boundaries if not already present
        if section_boundaries[0]["position"] > 0:
            section_boundaries.insert(0, {
                "position": 0,
                "text": "Document Start",
                "section_type": "introduction"
            })
        
        # Process each section
        for i in range(len(section_boundaries) - 1):
            start = section_boundaries[i]["position"]
            end = section_boundaries[i+1]["position"]
            section_text = text[start:end]
            section_type = section_boundaries[i]["section_type"]
            
            # Skip very small sections
            if len(section_text) < 50:
                continue
                
            # If section is too big, use regular chunking
            if len(section_text) > max_chunk_size:
                section_chunks = self._chunk_text(section_text, max_chunk_size, overlap)
                
                # Add each chunk with section metadata
                for chunk in section_chunks:
                    chunks.append(chunk)
                    # Calculate approximate position
                    chunk_start = start + section_text.find(chunk[:50])
                    chunk_end = chunk_start + len(chunk)
                    chunk_metadata.append({
                        "section_type": section_type,
                        "section_heading": section_boundaries[i]["text"],
                        "start_pos": chunk_start,
                        "end_pos": chunk_end
                    })
            else:
                # Use the whole section as one chunk
                chunks.append(section_text)
                chunk_metadata.append({
                    "section_type": section_type,
                    "section_heading": section_boundaries[i]["text"],
                    "start_pos": start,
                    "end_pos": end
                })
        
        # Handle the last section
        last_start = section_boundaries[-1]["position"]
        last_section_text = text[last_start:]
        last_section_type = section_boundaries[-1]["section_type"]
        
        if len(last_section_text) > max_chunk_size:
            last_section_chunks = self._chunk_text(last_section_text, max_chunk_size, overlap)
            
            # Add each chunk with section metadata
            for chunk in last_section_chunks:
                chunks.append(chunk)
                # Calculate approximate position
                chunk_start = last_start + last_section_text.find(chunk[:50])
                chunk_end = chunk_start + len(chunk)
                chunk_metadata.append({
                    "section_type": last_section_type,
                    "section_heading": section_boundaries[-1]["text"],
                    "start_pos": chunk_start,
                    "end_pos": chunk_end
                })
        else:
            # Use the whole last section as one chunk
            chunks.append(last_section_text)
            chunk_metadata.append({
                "section_type": last_section_type,
                "section_heading": section_boundaries[-1]["text"],
                "start_pos": last_start,
                "end_pos": last_start + len(last_section_text)
            })
        
        return chunks, chunk_metadata
    
    def _chunk_text(self, text: str, max_chunk_size: int = 1000, overlap: int = 200) -> List[str]:
        """Split text into overlapping chunks, trying to break at sentence boundaries."""
        chunks = []
        start = 0
        text_len = len(text)
        
        while start < text_len:
            end = min(start + max_chunk_size, text_len)
            if end < text_len and end - start == max_chunk_size:
                # Find the last sentence boundary
                last_period = text.rfind('.', start, end)
                last_newline = text.rfind('\n', start, end)
                break_point = max(last_period, last_newline)
                
                if break_point > start + 100:  # Ensure chunk isn't too small
                    end = break_point + 1
            
            chunks.append(text[start:end])
            start = end - overlap if end < text_len else text_len
        
        return chunks
    
    def _embed_chunks(self, chunks: List[str]) -> np.ndarray:
        """Create embeddings for text chunks."""
        return self.embedding_model.encode(chunks)
    
    def hybrid_search(self, query: str, top_k: int = 5, 
                      semantic_weight: float = 0.7,
                      section_filter: Optional[str] = None) -> List[Dict]:
        """
        Perform hybrid search optimized for fund prospectus queries.
        
        Args:
            query: The search query
            top_k: Number of results to return
            semantic_weight: Weight given to semantic search (0-1)
            section_filter: Optionally filter by section type
            
        Returns:
            List of search results with document info
        """
        # 1. Semantic search
        query_embedding = self.embedding_model.encode(query)
        semantic_results = self.vector_db.search(query_embedding, limit=top_k*2)
        
        # 2. Keyword search (TF-IDF)
        keyword_results = []
        if self.tfidf_matrix is not None:
            # Add financial terms to the query if relevant
            enhanced_query = query
            for term in self.financial_terms:
                if term in query.lower():
                    # Boost query with financial term by repetition
                    enhanced_query += f" {term}"
                    
            query_vector = self.tfidf_vectorizer.transform([enhanced_query])
            tfidf_similarities = cosine_similarity(query_vector, self.tfidf_matrix)[0]
            
            # Get top TF-IDF matches
            top_indices = tfidf_similarities.argsort()[-top_k*2:][::-1]
            
            for idx in top_indices:
                doc_id, chunk_idx, chunk_text = self.tfidf_documents[idx]
                keyword_results.append({
                    "document_id": doc_id,
                    "chunk_index": chunk_idx,
                    "text": chunk_text,
                    "filename": self.metadata["documents"][doc_id]["filename"],
                    "similarity": float(tfidf_similarities[idx])
                })
        
        # 3. Boost scores for specific section types based on query content
        query_lower = query.lower()
        section_boosts = {
            "fees_and_expenses": 0.0,
            "performance": 0.0,
            "investment_objective": 0.0,
            "principal_risks": 0.0,
            "management": 0.0,
            "purchase_and_sale": 0.0,
            "tax_information": 0.0,
            "asset_allocation": 0.0,
            "leverage": 0.0
        }
        
        # Determine which section types are relevant to the query
        if any(term in query_lower for term in ["fee", "expense", "ratio", "cost", "charge"]):
            section_boosts["fees_and_expenses"] = 0.3
        
        if any(term in query_lower for term in ["performance", "return", "yield", "history"]):
            section_boosts["performance"] = 0.3
            
        if any(term in query_lower for term in ["objective", "goal", "aim", "purpose"]):
            section_boosts["investment_objective"] = 0.3
            
        if any(term in query_lower for term in ["risk", "danger", "drawdown", "loss"]):
            section_boosts["principal_risks"] = 0.3
            
        if any(term in query_lower for term in ["manager", "management", "advisor", "team"]):
            section_boosts["management"] = 0.3
            
        if any(term in query_lower for term in ["buy", "sell", "purchase", "redemption", "minimum"]):
            section_boosts["purchase_and_sale"] = 0.3
            
        if any(term in query_lower for term in ["tax", "dividend", "distribution"]):
            section_boosts["tax_information"] = 0.3
            
        if any(term in query_lower for term in ["allocation", "portfolio", "holding", "sector"]):
            section_boosts["asset_allocation"] = 0.3
            
        if any(term in query_lower for term in ["leverage", "borrowing", "debt", "exposure"]):
            section_boosts["leverage"] = 0.3
            section_boosts["principal_risks"] = 0.2  # Also boost risk section for leverage queries
        
        # 4. Merge results with weighting and section boosting
        merged_results = {}
        
        # Apply section filter if specified
        filter_results = section_filter is not None
        
        # Add semantic results with weight
        for result in semantic_results:
            # Apply section filter if specified
            if filter_results and result.get("section_type") != section_filter:
                continue
                
            key = f"{result['document_id']}_{result['chunk_index']}"
            section_boost = section_boosts.get(result.get("section_type", "unknown"), 0.0)
            
            merged_results[key] = {
                **result,
                "final_score": (result["similarity"] * semantic_weight) + section_boost
            }
        
        # Add keyword results with weight
        for result in keyword_results:
            key = f"{result['document_id']}_{result['chunk_index']}"
            
            # For keyword results, we need to find the section type
            if key in merged_results:
                section_type = merged_results[key].get("section_type", "unknown")
            else:
                # Try to find section type from semantic results
                section_type = "unknown"
                for sem_result in semantic_results:
                    if sem_result["document_id"] == result["document_id"] and sem_result["chunk_index"] == result["chunk_index"]:
                        section_type = sem_result.get("section_type", "unknown")
                        break
            
            # Apply section filter if specified
            if filter_results and section_type != section_filter:
                continue
                
            section_boost = section_boosts.get(section_type, 0.0)
            
            if key in merged_results:
                # Combine scores
                merged_results[key]["final_score"] += (result["similarity"] * (1 - semantic_weight)) + section_boost
            else:
                merged_results[key] = {
                    **result,
                    "section_type": section_type,
                    "final_score": (result["similarity"] * (1 - semantic_weight)) + section_boost
                }
        
        # Convert to list and sort by final score
        results_list = list(merged_results.values())
        results_list.sort(key=lambda x: x["final_score"], reverse=True)
        
        return results_list[:top_k]


class LocalAnswerGenerator:
    """
    Generates answers from fund prospectus chunks without external LLMs.
    Enhanced for financial data extraction.
    """
    
    def __init__(self):
        # Use a cross-encoder model for relevance scoring - configured for offline use
        logger.info("Loading cross-encoder model from local cache...")
        self.cross_encoder = CrossEncoder('cross-encoder/ms-marco-MiniLM-L-6-v2', 
                                         use_auth_token=False)
        
    def generate_answer(self, query: str, search_results: List[Dict]) -> dict:
        """
        Generate an answer based on search results, optimized for fund prospectus queries.
        Returns a dictionary with answer text and extracted financial data.
        """
        if not search_results:
            return {
                "text": "No relevant information found in the prospectuses.",
                "financial_data": []
            }
        
        # 1. Score passages with cross-encoder for more accurate relevance
        passages = [result["text"] for result in search_results]
        passage_pairs = [[query, passage] for passage in passages]
        relevance_scores = self.cross_encoder.predict(passage_pairs)
        
        # Combine with search results
        for i, result in enumerate(search_results):
            result["relevance_score"] = float(relevance_scores[i])
        
        # Re-rank based on cross-encoder scores
        search_results.sort(key=lambda x: x["relevance_score"], reverse=True)
        
        # 2. Extract and collect financial data from relevant chunks
        financial_data_points = []
        for result in search_results[:3]:  # Focus on top 3 results
            if "financial_data" in result and result["financial_data"]:
                for data_point in result["financial_data"]:
                    # Only include relevant financial data based on query
                    if self._is_financial_data_relevant(query, data_point):
                        # Add source information
                        data_point["source"] = {
                            "document": result["filename"],
                            "section_type": result.get("section_type", "unknown")
                        }
                        financial_data_points.append(data_point)
        
        # 3. Generate specialized answer based on query type
        query_lower = query.lower()
        
        # Check if this is a query about specific financial data
        if any(term in query_lower for term in ["fee", "expense", "ratio", "cost"]):
            return self._generate_fee_answer(query, search_results, financial_data_points)
            
        elif any(term in query_lower for term in ["performance", "return", "yield"]):
            return self._generate_performance_answer(query, search_results, financial_data_points)
            
        elif any(term in query_lower for term in ["risk", "principal risk"]):
            return self._generate_risk_answer(query, search_results)
            
        elif any(term in query_lower for term in ["objective", "goal", "strategy"]):
            return self._generate_objective_answer(query, search_results)
            
        elif any(term in query_lower for term in ["minimum", "investment", "purchase"]):
            return self._generate_purchase_answer(query, search_results, financial_data_points)
            
        elif any(term in query_lower for term in ["leverage", "borrowing", "debt", "gearing", "exposure"]):
            return self._generate_leverage_answer(query, search_results, financial_data_points)
        
        # 4. For general queries, extract relevant sentences
        answer_sentences = []
        
        for result in search_results[:3]:  # Focus on top 3 results
            # Split into sentences
            sentences = re.split(r'(?<=[.!?])\s+', result["text"])
            
            # Score each sentence
            sentence_pairs = [[query, sentence] for sentence in sentences]
            sentence_scores = self.cross_encoder.predict(sentence_pairs)
            
            # Get top sentences
            for i, score in enumerate(sentence_scores):
                if score > 0.5:  # Only keep relevant sentences
                    answer_sentences.append({
                        "text": sentences[i],
                        "score": float(score),
                        "document": result["filename"],
                        "section_type": result.get("section_type", "unknown")
                    })
        
        # Sort by relevance
        answer_sentences.sort(key=lambda x: x["score"], reverse=True)
        
        # 5. Construct the answer
        if not answer_sentences:
            # Fall back to using top chunks
            return self._construct_from_chunks(query, search_results[:3], financial_data_points)
        
        # Deduplicate sentences
        seen_text = set()
        unique_sentences = []
        
        for sentence in answer_sentences:
            # Simple deduplication - check if similar text is already included
            normalized = re.sub(r'\s+', ' ', sentence["text"].lower())
            if normalized not in seen_text:
                unique_sentences.append(sentence)
                seen_text.add(normalized)
        
        # Format the answer
        answer_text = f"Based on the fund prospectuses, here's what I found:\n\n"
        
        for sentence in unique_sentences[:5]:  # Top 5 sentences
            answer_text += f"• {sentence['text']} [From: {sentence['document']}, {sentence['section_type']} section]\n\n"
        
        # Add financial data if found
        if financial_data_points:
            answer_text += "\nRelevant financial data points:\n"
            for data in financial_data_points[:3]:  # Limit to top 3 points
                value_str = f"{data['value']}" if isinstance(data['value'], (int, float)) else f"{data['value']}"
                answer_text += f"• {value_str}{data['unit']} - {data['context']}\n"
        
        return {
            "text": answer_text,
            "financial_data": financial_data_points
        }
    
    def _is_financial_data_relevant(self, query, data_point):
        """Determine if a financial data point is relevant to the query."""
        query_lower = query.lower()
        
        # Check data type against query
        data_type = data_point.get("type", "")
        
        # Fee-related queries
        if any(term in query_lower for term in ["fee", "expense", "ratio", "cost"]):
            return data_type in ["expense_ratio", "fee_amount"]
        
        # Performance-related queries
        if any(term in query_lower for term in ["performance", "return", "yield"]):
            return data_type in ["yield_return"]
        
        # Minimum investment queries
        if any(term in query_lower for term in ["minimum", "investment", "purchase"]):
            return data_type in ["minimum_investment"]
            
        # Leverage-related queries
        if any(term in query_lower for term in ["leverage", "borrowed", "borrowing", "debt", "gearing", "exposure"]):
            return data_type in ["leverage", "leverage_ratio", "leverage_amount"]
        
        # General relevance - check context against query terms
        context = data_point.get("context", "").lower()
        query_terms = query_lower.split()
        
        return any(term in context for term in query_terms if len(term) > 3)
    
    def _generate_fee_answer(self, query, search_results, financial_data):
        """Generate specialized answer for fee-related queries."""
        # Filter for fee-related data points
        fee_data = [data for data in financial_data 
                   if data.get("type") in ["expense_ratio", "fee_amount"]]
        
        # Sort fee data by type
        fee_data.sort(key=lambda x: x.get("type", ""))
        
        # Extract relevant fee sentences from top results
        fee_sentences = []
        for result in search_results[:3]:
            if result.get("section_type") == "fees_and_expenses":
                # Split into sentences
                sentences = re.split(r'(?<=[.!?])\s+', result["text"])
                for sentence in sentences:
                    if any(term in sentence.lower() for term in ["fee", "expense", "ratio", "cost", "charge"]):
                        fee_sentences.append({
                            "text": sentence,
                            "document": result["filename"]
                        })
        
        # Construct answer
        answer_text = f"I found the following fee information in the fund prospectuses:\n\n"
        
        if fee_data:
            answer_text += "Fee data:\n"
            for data in fee_data:
                value_str = f"{data['value']}" if isinstance(data['value'], (int, float)) else f"{data['value']}"
                answer_text += f"• {value_str}{data['unit']} - {data['context']}\n"
            answer_text += "\n"
        
        if fee_sentences:
            answer_text += "Fee descriptions:\n"
            for sentence in fee_sentences[:5]:
                answer_text += f"• {sentence['text']} [From: {sentence['document']}]\n"
        
        return {
            "text": answer_text,
            "financial_data": fee_data
        }
    
    def _generate_performance_answer(self, query, search_results, financial_data):
        """Generate specialized answer for performance-related queries."""
        # Filter for performance-related data points
        perf_data = [data for data in financial_data 
                    if data.get("type") in ["yield_return"]]
        
        # Extract relevant performance sentences from top results
        perf_sentences = []
        for result in search_results[:3]:
            if result.get("section_type") == "performance":
                # Split into sentences
                sentences = re.split(r'(?<=[.!?])\s+', result["text"])
                for sentence in sentences:
                    if any(term in sentence.lower() for term in ["performance", "return", "yield", "history"]):
                        perf_sentences.append({
                            "text": sentence,
                            "document": result["filename"]
                        })
        
        # Construct answer
        answer_text = f"I found the following performance information in the fund prospectuses:\n\n"
        
        if perf_data:
            answer_text += "Performance data:\n"
            for data in perf_data:
                value_str = f"{data['value']}" if isinstance(data['value'], (int, float)) else f"{data['value']}"
                answer_text += f"• {value_str}{data['unit']} - {data['context']}\n"
            answer_text += "\n"
        
        if perf_sentences:
            answer_text += "Performance descriptions:\n"
            for sentence in perf_sentences[:5]:
                answer_text += f"• {sentence['text']} [From: {sentence['document']}]\n"
        
        return {
            "text": answer_text,
            "financial_data": perf_data
        }
    
    def _generate_risk_answer(self, query, search_results):
        """Generate specialized answer for risk-related queries."""
        # Extract risk factors from risk sections
        risk_sections = [result for result in search_results 
                        if result.get("section_type") == "principal_risks"]
        
        if not risk_sections:
            # Fall back to general results
            risk_sections = search_results[:3]
        
        # Extract risk factors
        risk_factors = []
        for section in risk_sections:
            # Try to identify bullet points or numbered risks
            text = section["text"]
            
            # Look for bullet point patterns
            bullet_risks = re.findall(r'[•\-\*]\s*([^•\-\*\n]+)', text)
            if bullet_risks:
                for risk in bullet_risks:
                    if len(risk.strip()) > 20:  # Avoid short fragments
                        risk_factors.append({
                            "text": risk.strip(),
                            "document": section["filename"]
                        })
            
            # Look for numbered risks
            numbered_risks = re.findall(r'\d+\.\s*([^\d\.\n]+)', text)
            if numbered_risks:
                for risk in numbered_risks:
                    if len(risk.strip()) > 20:  # Avoid short fragments
                        risk_factors.append({
                            "text": risk.strip(),
                            "document": section["filename"]
                        })
            
            # If no structured risks found, extract sentences with risk keywords
            if not bullet_risks and not numbered_risks:
                sentences = re.split(r'(?<=[.!?])\s+', text)
                for sentence in sentences:
                    if any(term in sentence.lower() for term in ["risk", "may cause", "could result", "volatility"]):
                        if len(sentence.strip()) > 20:  # Avoid short fragments
                            risk_factors.append({
                                "text": sentence.strip(),
                                "document": section["filename"]
                            })
        
        # Construct answer
        answer_text = f"I found the following principal risks in the fund prospectuses:\n\n"
        
        if risk_factors:
            for i, risk in enumerate(risk_factors[:8]):  # Limit to top 8 risks
                answer_text += f"{i+1}. {risk['text']} [From: {risk['document']}]\n\n"
        else:
            answer_text = "I couldn't find specific risk factors in the prospectuses. You may want to check the Principal Risks section directly."
        
        return {
            "text": answer_text,
            "financial_data": []
        }
    
    def _generate_objective_answer(self, query, search_results):
        """Generate specialized answer for investment objective queries."""
        # Extract from investment objective sections
        objective_sections = [result for result in search_results 
                            if result.get("section_type") == "investment_objective"]
        
        if not objective_sections:
            # Fall back to general results
            objective_sections = search_results[:3]
        
        # Extract objective statements
        objectives = []
        for section in objective_sections:
            text = section["text"]
            
            # Look for statements that typically describe objectives
            sentences = re.split(r'(?<=[.!?])\s+', text)
            for sentence in sentences:
                lower_sentence = sentence.lower()
                if any(phrase in lower_sentence for phrase in ["objective is", "seeks to", "aims to", "goal is"]):
                    objectives.append({
                        "text": sentence.strip(),
                        "document": section["filename"]
                    })
        
        # If we didn't find specific objective statements, use the first few sentences
        if not objectives and objective_sections:
            text = objective_sections[0]["text"]
            sentences = re.split(r'(?<=[.!?])\s+', text)
            objectives = [{
                "text": sentence.strip(),
                "document": objective_sections[0]["filename"]
            } for sentence in sentences[:3] if len(sentence.strip()) > 20]
        
        # Construct answer
        answer_text = f"I found the following investment objectives in the fund prospectuses:\n\n"
        
        if objectives:
            for obj in objectives[:3]:  # Limit to top 3
                answer_text += f"• {obj['text']} [From: {obj['document']}]\n\n"
        else:
            answer_text = "I couldn't find specific investment objectives in the prospectuses. You may want to check the Investment Objective section directly."
        
        return {
            "text": answer_text,
            "financial_data": []
        }
    
    def _generate_purchase_answer(self, query, search_results, financial_data):
        """Generate specialized answer for purchase and investment queries."""
        # Filter for minimum investment data points
        investment_data = [data for data in financial_data 
                          if data.get("type") in ["minimum_investment"]]
        
        # Extract relevant purchase sentences from top results
        purchase_sentences = []
        for result in search_results[:3]:
            if result.get("section_type") == "purchase_and_sale":
                # Split into sentences
                sentences = re.split(r'(?<=[.!?])\s+', result["text"])
                for sentence in sentences:
                    if any(term in sentence.lower() for term in ["minimum", "purchase", "buy", "invest", "account"]):
                        purchase_sentences.append({
                            "text": sentence,
                            "document": result["filename"]
                        })
        
        # Construct answer
        answer_text = f"I found the following purchase information in the fund prospectuses:\n\n"
        
        if investment_data:
            answer_text += "Minimum investment amounts:\n"
            for data in investment_data:
                value_str = f"{data['value']}" if isinstance(data['value'], (int, float)) else f"{data['value']}"
                answer_text += f"• {value_str}{data['unit']} - {data['context']}\n"
            answer_text += "\n"
        
        if purchase_sentences:
            answer_text += "Purchase details:\n"
            for sentence in purchase_sentences[:5]:
                answer_text += f"• {sentence['text']} [From: {sentence['document']}]\n"
        
        return {
            "text": answer_text,
            "financial_data": investment_data
        }
        
    def _generate_leverage_answer(self, query, search_results, financial_data):
        """Generate specialized answer for leverage-related queries."""
        # Filter for leverage-related data points
        leverage_data = [data for data in financial_data 
                        if data.get("type") in ["leverage", "leverage_ratio", "leverage_amount"]]
        
        # Extract relevant leverage sentences from top results
        leverage_sentences = []
        for result in search_results[:3]:
            # Check specifically for leverage section or risk section (often contains leverage info)
            if result.get("section_type") in ["leverage", "principal_risks"]:
                # Split into sentences
                sentences = re.split(r'(?<=[.!?])\s+', result["text"])
                for sentence in sentences:
                    sentence_lower = sentence.lower()
                    # Look for leverage-related terms
                    if any(term in sentence_lower for term in ["leverage", "borrowing", "debt", "gearing", 
                                                             "exposure", "130/30", "long/short", "derivative", 
                                                             "swap", "future", "notional"]):
                        leverage_sentences.append({
                            "text": sentence,
                            "document": result["filename"],
                            "section": result.get("section_type", "unknown")
                        })
        
        # Construct answer
        answer_text = f"I found the following leverage information in the fund prospectuses:\n\n"
        
        if leverage_data:
            answer_text += "Leverage metrics:\n"
            for data in leverage_data:
                value_str = f"{data['value']}" if isinstance(data['value'], (int, float)) else f"{data['value']}"
                unit = data['unit']
                
                # Format based on unit type
                if unit == "multiplier":
                    formatted_value = f"{value_str}x leverage"
                elif unit == "percentage":
                    formatted_value = f"{value_str}% exposure"
                elif unit == "notional":
                    formatted_value = f"Notional value: {value_str}"
                else:
                    formatted_value = f"{value_str} {unit}"
                
                answer_text += f"• {formatted_value} - {data['context']}\n"
            answer_text += "\n"
        
        if leverage_sentences:
            answer_text += "Leverage details:\n"
            # Group by document to maintain context
            by_document = {}
            for sentence in leverage_sentences:
                if sentence["document"] not in by_document:
                    by_document[sentence["document"]] = []
                by_document[sentence["document"]].append(sentence)
            
            # Display up to 2 sentences per document, maximum 3 documents
            for doc, sentences in list(by_document.items())[:3]:
                answer_text += f"\nFrom {doc}:\n"
                for sentence in sentences[:2]:
                    answer_text += f"• {sentence['text']}\n"
        else:
            # If no specific leverage sentences but we have search results
            if search_results:
                answer_text += "\nRelated information from prospectuses:\n"
                # Take a paragraph from the most relevant result
                top_result = search_results[0]
                text_sample = top_result["text"][:300] + "..." if len(top_result["text"]) > 300 else top_result["text"]
                answer_text += f"{text_sample}\n\n[From: {top_result['filename']}]"
        
        return {
            "text": answer_text,
            "financial_data": leverage_data
        }
    
    def _construct_from_chunks(self, query: str, chunks: List[Dict], financial_data: List = None) -> Dict:
        """Fallback method to construct an answer from whole chunks."""
        answer_text = f"Based on the fund prospectuses, I found these relevant passages:\n\n"
        
        for i, chunk in enumerate(chunks):
            section_type = chunk.get("section_type", "unknown").replace("_", " ").title()
            answer_text += f"From {chunk['filename']} ({section_type} section):\n"
            answer_text += f"{chunk['text'][:300]}...\n\n" if len(chunk['text']) > 300 else f"{chunk['text']}\n\n"
        
        # Add financial data if found
        if financial_data:
            answer_text += "\nRelevant financial data points:\n"
            for data in financial_data[:3]:  # Limit to top 3 points
                value_str = f"{data['value']}" if isinstance(data['value'], (int, float)) else f"{data['value']}"
                answer_text += f"• {value_str}{data['unit']} - {data['context']}\n"
            
        return {
            "text": answer_text,
            "financial_data": financial_data if financial_data else []
        }


class QueryEngine:
    """Enhanced query engine optimized for fund prospectus questions."""
    
    def __init__(self, document_processor: DocumentProcessor):
        self.document_processor = document_processor
        self.answer_generator = LocalAnswerGenerator()
        
        # Query log for learning
        self.query_log_path = Path(document_processor.storage_dir) / "query_log.json"
        self.query_log = self._load_query_log()
        
        # Feedback records
        self.feedback_path = Path(document_processor.storage_dir) / "feedback.json"
        self.feedback = self._load_feedback()
        
        # Financial query patterns
        self.financial_query_patterns = {
            "fees": r"(?:expense|fee|ratio|cost|charge)",
            "performance": r"(?:performance|return|yield|history)",
            "risk": r"(?:risk|danger|drawdown|loss)",
            "objective": r"(?:objective|goal|purpose|strategy)",
            "purchase": r"(?:minimum|investment|purchase|buy|sell)",
            "tax": r"(?:tax|dividend|distribution)",
            "management": r"(?:manager|management|advisor|team)",
            "allocation": r"(?:allocation|portfolio|holding|sector)",
            "leverage": r"(?:leverage|borrowing|debt|gearing|exposure|130/30|derivative|swap|future|short)"
        }
    
    def _load_query_log(self) -> List[Dict]:
        """Load the query log from storage."""
        if self.query_log_path.exists():
            with open(self.query_log_path, 'r') as f:
                return json.load(f)
        return []
    
    def _save_query_log(self):
        """Save the query log to storage."""
        with open(self.query_log_path, 'w') as f:
            json.dump(self.query_log, f, indent=2)
            
    def _load_feedback(self) -> Dict:
        """Load feedback data from storage."""
        if self.feedback_path.exists():
            with open(self.feedback_path, 'r') as f:
                return json.load(f)
        return {"positive": [], "negative": []}
    
    def _save_feedback(self):
        """Save feedback data to storage."""
        with open(self.feedback_path, 'w') as f:
            json.dump(self.feedback, f, indent=2)
    
    def expand_query(self, query: str) -> str:
        """
        Expand the query using context from previous similar queries and feedback.
        Optimized for financial queries.
        """
        # Determine query type using financial patterns
        query_type = None
        query_lower = query.lower()
        
        for q_type, pattern in self.financial_query_patterns.items():
            if re.search(pattern, query_lower):
                query_type = q_type
                break
        
        # Find similar queries in the log that received positive feedback
        expanded_terms = set()
        
        # Look through positive feedback
        for entry in self.feedback["positive"]:
            original_query = entry["query"]
            
            # For financial queries, prioritize expanding with similar query types
            if query_type:
                # Check if the original query matches the same pattern
                original_matches_type = False
                for q_type, pattern in self.financial_query_patterns.items():
                    if q_type == query_type and re.search(pattern, original_query.lower()):
                        original_matches_type = True
                        break
                
                if original_matches_type:
                    # Add terms from this query with higher priority
                    original_terms = set(original_query.lower().split())
                    query_terms = set(query_lower.split())
                    expanded_terms.update(original_terms - query_terms)
                    continue
            
            # Simple similarity - check for overlapping terms
            query_terms = set(query_lower.split())
            original_terms = set(original_query.lower().split())
            
            # If queries share key terms, extract useful terms from the original
            overlap = query_terms.intersection(original_terms)
            if overlap and len(overlap) / len(query_terms) > 0.3:  # Some overlap
                # Add non-overlapping terms from successful query
                expanded_terms.update(original_terms - query_terms)
        
        # Add financial terms relevant to the query
        if query_type:
            if query_type == "fees":
                expanded_terms.update(["expense", "ratio", "fee", "charge"])
            elif query_type == "performance":
                expanded_terms.update(["return", "history", "yield"])
            elif query_type == "risk":
                expanded_terms.update(["principal", "risks", "volatility"])
            elif query_type == "objective":
                expanded_terms.update(["investment", "objective", "strategy"])
            elif query_type == "leverage":
                expanded_terms.update(["borrowing", "debt", "exposure", "derivatives"])
        
        # Construct expanded query
        if expanded_terms:
            expanded_query = f"{query} {' '.join(expanded_terms)}"
            logger.info(f"Expanded query from '{query}' to '{expanded_query}'")
            return expanded_query
            
        return query
    
    def process_query(self, query: str) -> Dict:
        """
        Process a user query about fund prospectuses and generate a response.
        """
        # Log the query
        query_id = len(self.query_log)
        query_entry = {
            "id": query_id,
            "query": query,
            "timestamp": datetime.now().isoformat(),
            "was_successful": False,
            "query_type": None
        }
        
        # Determine query type for specialized handling
        query_lower = query.lower()
        for q_type, pattern in self.financial_query_patterns.items():
            if re.search(pattern, query_lower):
                query_entry["query_type"] = q_type
                break
        
        # Expand query if we have historical data
        expanded_query = self.expand_query(query)
        
        # Determine the appropriate section filter
        section_filter = None
        if query_entry["query_type"] == "fees":
            section_filter = "fees_and_expenses"
        elif query_entry["query_type"] == "performance":
            section_filter = "performance"
        elif query_entry["query_type"] == "risk":
            section_filter = "principal_risks"
        elif query_entry["query_type"] == "objective":
            section_filter = "investment_objective"
        elif query_entry["query_type"] == "purchase":
            section_filter = "purchase_and_sale"
        elif query_entry["query_type"] == "tax":
            section_filter = "tax_information"
        elif query_entry["query_type"] == "management":
            section_filter = "management"
        elif query_entry["query_type"] == "allocation":
            section_filter = "asset_allocation"
        elif query_entry["query_type"] == "leverage":
            section_filter = "leverage"
        
        # First try with section filter if applicable
        if section_filter:
            search_results = self.document_processor.hybrid_search(
                expanded_query, 
                top_k=5,
                semantic_weight=0.7,
                section_filter=section_filter
            )
            
            # If we don't get enough results, try without filter
            if len(search_results) < 2:
                search_results = self.document_processor.hybrid_search(
                    expanded_query, 
                    top_k=5,
                    semantic_weight=0.7
                )
        else:
            # No specific section filter
            search_results = self.document_processor.hybrid_search(
                expanded_query, 
                top_k=5,
                semantic_weight=0.7
            )
        
        if not search_results:
            response = {
                "answer": {
                    "text": "I couldn't find relevant information in the prospectuses. Please try rephrasing your query or upload more fund documents.",
                    "financial_data": []
                },
                "sources": [],
                "query_id": query_id,
                "query_type": query_entry["query_type"]
            }
        else:
            # Generate response based on retrieved chunks
            answer = self.answer_generator.generate_answer(query, search_results)
            
            response = {
                "answer": answer,
                "sources": [{
                    "filename": result["filename"],
                    "document_id": result["document_id"],
                    "text_snippet": result["text"][:200] + "..." if len(result["text"]) > 200 else result["text"],
                    "relevance": result.get("relevance_score", result.get("similarity", 0)),
                    "section_type": result.get("section_type", "unknown")
                } for result in search_results[:3]]  # Include top 3 sources
            }
            
            query_entry["was_successful"] = True
            query_entry["document_ids"] = [r["document_id"] for r in search_results[:3]]
        
        # Add query_id to response
        response["query_id"] = query_id
        response["query_type"] = query_entry["query_type"]
            
        # Update query log
        self.query_log.append(query_entry)
        self._save_query_log()
        
        # Update document metadata with this query
        for result in search_results:
            doc_id = result["document_id"]
            if doc_id in self.document_processor.metadata["documents"]:
                self.document_processor.metadata["documents"][doc_id]["queries"].append({
                    "query": query,
                    "timestamp": datetime.now().isoformat(),
                    "query_type": query_entry["query_type"]
                })
        self.document_processor._save_metadata()
        
        return response
    
    def incorporate_feedback(self, query_id: int, was_helpful: bool, user_feedback: str = None):
        """
        Incorporate user feedback to improve future responses.
        """
        if query_id < 0 or query_id >= len(self.query_log):
            return False
            
        # Update the query log with feedback
        self.query_log[query_id]["was_successful"] = was_helpful
        if user_feedback:
            self.query_log[query_id]["feedback"] = user_feedback
            
        self._save_query_log()
        
        # Add to feedback collection for query expansion learning
        query_info = self.query_log[query_id]
        feedback_entry = {
            "query": query_info["query"],
            "query_type": query_info.get("query_type"),
            "timestamp": datetime.now().isoformat(),
            "document_ids": query_info.get("document_ids", []),
            "user_feedback": user_feedback
        }
        
        if was_helpful:
            self.feedback["positive"].append(feedback_entry)
        else:
            self.feedback["negative"].append(feedback_entry)
            
        self._save_feedback()
        return True


class LearningSystem:
    """Learning system enhanced for fund prospectus analysis."""
    
    def __init__(self, document_processor: DocumentProcessor, query_engine: QueryEngine):
        self.document_processor = document_processor
        self.query_engine = query_engine
        
        # Track metrics for improvement
        self.performance_log_path = Path(document_processor.storage_dir) / "performance_log.json"
        self.performance_log = self._load_performance_log()
        
    def _load_performance_log(self) -> Dict:
        """Load performance metrics from storage."""
        if self.performance_log_path.exists():
            with open(self.performance_log_path, 'r') as f:
                return json.load(f)
        return {
            "query_success_rate": [],
            "document_coverage": {},
            "query_patterns": {},
            "financial_query_performance": {},
            "improvement_suggestions": []
        }
    
    def _save_performance_log(self):
        """Save performance metrics to storage."""
        with open(self.performance_log_path, 'w') as f:
            json.dump(self.performance_log, f, indent=2)
    
    def analyze_performance(self) -> Dict:
        """
        Analyze system performance based on query logs and feedback.
        Enhanced for financial queries.
        """
        query_log = self.query_engine.query_log
        
        if not query_log:
            return {"status": "No queries processed yet"}
            
        # Calculate success rate
        success_count = sum(1 for q in query_log if q.get("was_successful", False))
        success_rate = success_count / len(query_log) if query_log else 0
        
        # Analyze document coverage
        doc_query_counts = {}
        for doc_id in self.document_processor.metadata["documents"]:
            doc_query_counts[doc_id] = sum(
                1 for q in query_log if "document_ids" in q and doc_id in q.get("document_ids", [])
            )
        
        # Analyze query patterns
        query_patterns = {}
        for query in query_log:
            # Extract key terms
            terms = [term.lower() for term in re.findall(r'\b\w{3,}\b', query["query"])]
            for term in terms:
                if term not in query_patterns:
                    query_patterns[term] = {
                        "count": 0,
                        "success_count": 0
                    }
                query_patterns[term]["count"] += 1
                if query.get("was_successful", False):
                    query_patterns[term]["success_count"] += 1
        
        # Calculate success rate for each term
        for term, stats in query_patterns.items():
            stats["success_rate"] = stats["success_count"] / stats["count"]
        
        # Sort terms by count
        sorted_terms = sorted(query_patterns.items(), key=lambda x: x[1]["count"], reverse=True)
        top_terms = dict(sorted_terms[:10])  # Keep top 10 terms
        
        # Analyze performance by financial query type
        financial_query_performance = {}
        
        for query in query_log:
            query_type = query.get("query_type")
            if query_type:
                if query_type not in financial_query_performance:
                    financial_query_performance[query_type] = {
                        "count": 0,
                        "success_count": 0
                    }
                financial_query_performance[query_type]["count"] += 1
                if query.get("was_successful", False):
                    financial_query_performance[query_type]["success_count"] += 1
        
        # Calculate success rate for each query type
        for query_type, stats in financial_query_performance.items():
            stats["success_rate"] = stats["success_count"] / stats["count"] if stats["count"] > 0 else 0
        
        # Track metrics
        self.performance_log["query_success_rate"].append({
            "timestamp": datetime.now().isoformat(),
            "rate": success_rate,
            "total_queries": len(query_log)
        })
        
        self.performance_log["document_coverage"] = doc_query_counts
        self.performance_log["query_patterns"] = top_terms
        self.performance_log["financial_query_performance"] = financial_query_performance
        
        # Generate improvement suggestions
        suggestions = []
        
        # Suggest processing more documents if few are available
        if len(self.document_processor.metadata["documents"]) < 3:
            suggestions.append("Upload more fund prospectuses to improve the knowledge base")
        
        # Identify underperforming financial query types
        for query_type, stats in financial_query_performance.items():
            if stats["count"] > 2 and stats["success_rate"] < 0.5:
                suggestions.append(f"Improve handling of {query_type} queries, current success rate: {stats['success_rate']:.1%}")
        
        # Identify problematic terms
        problem_terms = [
            term for term, stats in query_patterns.items() 
            if stats["count"] > 2 and stats["success_rate"] < 0.5
        ]
        if problem_terms:
            suggestions.append(f"Low success rate for queries containing: {', '.join(problem_terms[:3])}")
        
        # Identify underutilized documents
        unused_docs = [
            self.document_processor.metadata["documents"][doc_id]["filename"]
            for doc_id, count in doc_query_counts.items() if count == 0
        ]
        if unused_docs:
            suggestions.append(f"Fund prospectuses never referenced in queries: {', '.join(unused_docs[:3])}")
        
        self.performance_log["improvement_suggestions"] = suggestions
        self._save_performance_log()
        
        return {
            "success_rate": success_rate,
            "processed_queries": len(query_log),
            "document_coverage": doc_query_counts,
            "popular_terms": dict(sorted_terms[:5]),
            "financial_query_performance": financial_query_performance,
            "suggestions": suggestions
        }
    
    def improve_system(self):
        """
        Apply automated improvements to the system based on collected data.
        Optimized for fund prospectus queries.
        """
        # Analyze performance to get latest data
        self.analyze_performance()
        
        # 1. Adjusting semantic vs keyword search weights based on financial query type
        financial_query_performance = self.performance_log.get("financial_query_performance", {})
        
        # Default weights
        weights_by_type = {
            "fees": 0.7,
            "performance": 0.7,
            "risk": 0.8,
            "objective": 0.8,
            "purchase": 0.7,
            "tax": 0.6,
            "management": 0.8,
            "allocation": 0.7,
            "leverage": 0.7
        }
        
        # Adjust weights based on performance
        for query_type, stats in financial_query_performance.items():
            if stats["count"] > 3:  # Only adjust if we have enough data
                if stats["success_rate"] < 0.5:
                    # If current weight is primarily semantic (>0.5), reduce semantic weight
                    if weights_by_type.get(query_type, 0.7) > 0.5:
                        weights_by_type[query_type] = max(0.5, weights_by_type.get(query_type, 0.7) - 0.1)
                    else:
                        # Otherwise increase semantic weight
                        weights_by_type[query_type] = min(0.9, weights_by_type.get(query_type, 0.7) + 0.1)
        
        logger.info(f"Adjusted weights by query type: {weights_by_type}")
        
        # 2. Identify high-value document sections
        section_importance = {
            "fees_and_expenses": 0,
            "performance": 0,
            "investment_objective": 0,
            "principal_risks": 0,
            "management": 0,
            "purchase_and_sale": 0,
            "tax_information": 0,
            "asset_allocation": 0,
            "leverage": 0
        }
        
        for entry in self.query_engine.feedback["positive"]:
            query_type = entry.get("query_type")
            if query_type == "fees":
                section_importance["fees_and_expenses"] += 1
            elif query_type == "performance":
                section_importance["performance"] += 1
            elif query_type == "risk":
                section_importance["principal_risks"] += 1
            elif query_type == "objective":
                section_importance["investment_objective"] += 1
            elif query_type == "purchase":
                section_importance["purchase_and_sale"] += 1
            elif query_type == "tax":
                section_importance["tax_information"] += 1
            elif query_type == "management":
                section_importance["management"] += 1
            elif query_type == "allocation":
                section_importance["asset_allocation"] += 1
            elif query_type == "leverage":
                section_importance["leverage"] += 1
        
        # Log which sections are most valuable
        sorted_sections = sorted(section_importance.items(), key=lambda x: x[1], reverse=True)
        logger.info(f"Most valuable sections: {sorted_sections[:3]}")
        
        # 3. Record improvement attempt
        improvement_record = {
            "timestamp": datetime.now().isoformat(),
            "metrics_before": self.performance_log["query_success_rate"][-1] if self.performance_log["query_success_rate"] else None,
            "actions_taken": [
                "Adjusted search weights by query type",
                "Identified valuable prospectus sections",
                "Analyzed financial query performance"
            ],
            "weight_adjustments": weights_by_type,
            "section_importance": dict(sorted_sections)
        }
        
        if "improvement_history" not in self.performance_log:
            self.performance_log["improvement_history"] = []
            
        self.performance_log["improvement_history"].append(improvement_record)
        self._save_performance_log()
        
        logger.info("System improvement cycle completed for fund prospectus analysis")
        return True


def main():
    st.title("Fund Prospectus Analysis System")
    
    # Check if models exist before initializing the application
    models_exist = verify_models_exist()
    
    if not models_exist:
        st.error("""
        Required models not found in the expected locations. Please ensure you've downloaded and placed 
        the models in the correct directories as described in the setup guide.
        
        Models needed:
        1. sentence-transformers/all-MiniLM-L6-v2
        2. cross-encoder/ms-marco-MiniLM-L-6-v2
        """)
        return
    
    # Initialize components
    if 'initialized' not in st.session_state:
        st.session_state.document_processor = DocumentProcessor()
        st.session_state.query_engine = QueryEngine(st.session_state.document_processor)
        st.session_state.learning_system = LearningSystem(
            st.session_state.document_processor, 
            st.session_state.query_engine
        )
        st.session_state.initialized = True
    
    # Sidebar for document upload
    with st.sidebar:
        st.header("Upload Fund Prospectuses")
        uploaded_file = st.file_uploader("Choose a PDF prospectus", type="pdf")
        
        if uploaded_file is not None:
            with st.spinner("Processing fund prospectus..."):
                doc_id = st.session_state.document_processor.process_pdf(
                    uploaded_file, uploaded_file.name
                )
                st.success(f"Prospectus processed: {uploaded_file.name}")
        
        # Document list
        st.header("Your Fund Documents")
        docs = st.session_state.document_processor.metadata["documents"]
        for doc_id, info in docs.items():
            st.write(f"{info['filename']} - {info['upload_date'][:10]}")
            
        # Common financial questions
        st.header("Common Questions")
        common_questions = [
            "What is the expense ratio?",
            "What is the investment objective?",
            "What are the principal risks?",
            "What is the performance history?",
            "What is the minimum investment?",
            "Who is the fund manager?",
            "What are the tax implications?",
            "Does the fund use leverage?",
            "How much leverage does the fund employ?"
        ]
        
        for question in common_questions:
            if st.button(question):
                st.session_state.query = question
    
    # Main area for query input
    st.header("Ask about the Fund Prospectus")
    query = st.text_input("Enter your question:", key="query")
    
    if query:
        with st.spinner("Analyzing fund prospectuses..."):
            response = st.session_state.query_engine.process_query(query)
            
        st.subheader("Answer")
        st.write(response["answer"]["text"])
        
        # Display financial data if available
        if response["answer"].get("financial_data"):
            st.subheader("Financial Data Points")
            for data in response["answer"]["financial_data"]:
                with st.expander(f"{data['value']}{data['unit']} ({data['type']})"):
                    st.write(data['context'])
        
        if response["sources"]:
            st.subheader("Sources")
            for src in response["sources"]:
                section_type = src["section_type"].replace("_", " ").title() if src["section_type"] else "Unknown"
                with st.expander(f"From {src['filename']} ({section_type} section, Relevance: {src['relevance']:.2f})"):
                    st.write(src["text_snippet"])
        
        # Store query ID for feedback
        st.session_state.last_query_id = response["query_id"]
        
        # Feedback mechanism
        st.subheader("Was this answer helpful?")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("👍 Yes"):
                st.session_state.query_engine.incorporate_feedback(
                    st.session_state.last_query_id, True
                )
                st.success("Thanks for your feedback!")
        with col2:
            if st.button("👎 No"):
                feedback = st.text_input("What could be improved?")
                if st.button("Submit Feedback"):
                    st.session_state.query_engine.incorporate_feedback(
                        st.session_state.last_query_id, False, feedback
                    )
                    st.info("Thanks for your feedback. We'll try to improve!")
    
    # System analytics section
    if st.checkbox("Show System Analytics"):
        st.header("System Performance")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Analyze Performance"):
                with st.spinner("Analyzing system performance..."):
                    performance = st.session_state.learning_system.analyze_performance()
                    st.session_state.performance = performance
        with col2:
            if st.button("Improve System"):
                with st.spinner("Improving system based on feedback..."):
                    success = st.session_state.learning_system.improve_system()
                    if success:
                        st.success("System improvement complete!")
                    else:
                        st.info("Not enough data yet for improvement.")
        
        # Display performance metrics if available
        if hasattr(st.session_state, 'performance'):
            performance = st.session_state.performance
            
            if "status" in performance:
                st.write(performance["status"])
            else:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Query Success Rate", f"{performance['success_rate']:.2%}")
                with col2:
                    st.metric("Prospectuses Processed", len(docs))
                with col3:
                    st.metric("Queries Processed", performance["processed_queries"])
                
                # Document coverage
                st.subheader("Prospectus Coverage")
                coverage_data = []
                for doc_id, count in performance["document_coverage"].items():
                    if doc_id in docs:
                        coverage_data.append({
                            "Document": docs[doc_id]["filename"],
                            "Queries": count
                        })
                
                if coverage_data:
                    st.bar_chart(
                        data={row["Document"]: row["Queries"] for row in coverage_data},
                        use_container_width=True
                    )
                
                # Financial query performance
                if "financial_query_performance" in performance:
                    st.subheader("Financial Query Performance")
                    for query_type, stats in performance["financial_query_performance"].items():
                        query_type_display = query_type.replace("_", " ").title()
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric(f"{query_type_display} Queries", stats["count"])
                        with col2:
                            st.metric(f"{query_type_display} Success Rate", f"{stats['success_rate']:.1%}")
                
                # Popular search terms
                if "popular_terms" in performance:
                    st.subheader("Popular Search Terms")
                    for term, stats in performance["popular_terms"].items():
                        st.write(f"• {term}: {stats['count']} queries, {stats['success_rate']:.1%} success rate")
                
                # Improvement suggestions
                if performance["suggestions"]:
                    st.subheader("Improvement Suggestions")
                    for suggestion in performance["suggestions"]:
                        st.write(f"• {suggestion}")

if __name__ == "__main__":
    main()
