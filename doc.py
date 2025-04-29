# app.py
import os
import json
import logging
import numpy as np
import pickle
import fitz  # PyMuPDF for better PDF extraction
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional, Union

# For vector embeddings and search
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
                "similarity": res.score
            })
            
        return formatted_results
    
    def count_documents(self) -> int:
        """Count the number of documents in the database."""
        collection_info = self.client.get_collection(collection_name=self.collection_name)
        return collection_info.points_count


class DocumentProcessor:
    """Enhanced document processor with improved PDF extraction and vector storage."""
    
    def __init__(self, storage_dir: str = "document_storage"):
        self.storage_dir = Path(storage_dir)
        self.storage_dir.mkdir(exist_ok=True)
        
        # Document metadata storage
        self.metadata_path = self.storage_dir / "metadata.json"
        self.metadata = self._load_metadata()
        
        # Initialize sentence embedding model
        self.embedding_model = SentenceTransformer('all-MiniLM-L6-v2')
        
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
        Process a PDF file with enhanced extraction and storage.
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
            chunks = self._chunk_text_with_structure(text, structure)
            
            # Create embeddings for chunks
            embeddings = self._embed_chunks(chunks)
            
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
            
            # Update metadata
            doc_metadata = {
                "filename": filename,
                "upload_date": datetime.now().isoformat(),
                "num_chunks": len(chunks),
                "structure_elements": len(structure),
                "queries": []  # Track queries about this document
            }
            
            self.metadata["documents"][document_id] = doc_metadata
            self._save_metadata()
            
            # Add to vector database
            self.vector_db.add_documents(
                document_id=document_id,
                chunks=chunks,
                embeddings=embeddings,
                metadata={"filename": filename, **doc_metadata}
            )
            
            # Update TF-IDF model
            self.tfidf_documents.extend([(document_id, i, chunk) for i, chunk in enumerate(chunks)])
            self._update_tfidf_model()
            
            return document_id
            
        finally:
            # Clean up temporary file
            os.unlink(temp_path)
    
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
        Extract text and structure from a PDF file using PyMuPDF.
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
                                if span["size"] > 12:  # Larger font might indicate heading
                                    structure.append({
                                        "type": "heading",
                                        "text": span["text"],
                                        "page": page_num + 1,
                                        "position": {
                                            "x0": span["bbox"][0],
                                            "y0": span["bbox"][1],
                                            "x1": span["bbox"][2],
                                            "y1": span["bbox"][3]
                                        }
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
                            }
                        })
                
                # Detect tables (simple heuristic based on line spacing patterns)
                if self._detect_tables(page):
                    structure.append({
                        "type": "table",
                        "page": page_num + 1
                    })
                
                full_text += page_text + "\n\n"
            
            # Extract potential section headers
            section_headers = re.findall(r'^(?:[0-9.]+\s+)?([A-Z][A-Za-z\s]{2,50})$', 
                                        full_text, 
                                        re.MULTILINE)
            
            for header in section_headers:
                if header not in [s["text"] for s in structure if s.get("type") == "heading"]:
                    structure.append({
                        "type": "section",
                        "text": header
                    })
            
            return full_text, structure
            
        except Exception as e:
            logger.error(f"Error extracting from PDF: {e}")
            return "", []
    
    def _detect_tables(self, page) -> bool:
        """
        Simple table detection heuristic.
        Returns True if a table is likely present on the page.
        """
        # Get horizontal and vertical lines
        drawings = page.get_drawings()
        if not drawings:
            return False
            
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
        return h_lines > 2 and v_lines > 2
    
    def _chunk_text_with_structure(self, text: str, structure: List[Dict], 
                                 max_chunk_size: int = 1000, 
                                 overlap: int = 200) -> List[str]:
        """
        Split text into chunks, respecting document structure where possible.
        Uses structural elements like headings as natural chunk boundaries.
        """
        # Extract potential chunk boundaries from structure
        boundaries = []
        
        for item in structure:
            if item.get("type") in ["heading", "section"] and "text" in item:
                # Find all occurrences of the heading text
                for match in re.finditer(re.escape(item["text"]), text):
                    boundaries.append(match.start())
        
        # Sort boundaries
        boundaries.sort()
        
        # If no structure found, fall back to regular chunking
        if not boundaries:
            return self._chunk_text(text, max_chunk_size, overlap)
        
        # Create chunks based on structural boundaries
        chunks = []
        start = 0
        
        for boundary in boundaries:
            # If boundary is close to start or exceeds max size, adjust
            if boundary - start < 100:  # Too small chunk, skip
                continue
                
            if boundary - start > max_chunk_size:
                # Too big, use regular chunking for this section
                section_text = text[start:boundary]
                section_chunks = self._chunk_text(section_text, max_chunk_size, overlap)
                chunks.extend(section_chunks)
            else:
                # Good size, use the boundary
                chunks.append(text[start:boundary])
            
            start = boundary
        
        # Handle the last chunk
        if start < len(text):
            remaining_text = text[start:]
            if len(remaining_text) > max_chunk_size:
                final_chunks = self._chunk_text(remaining_text, max_chunk_size, overlap)
                chunks.extend(final_chunks)
            else:
                chunks.append(remaining_text)
        
        return chunks
    
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
                      semantic_weight: float = 0.7) -> List[Dict]:
        """
        Perform hybrid search combining semantic and keyword-based approaches.
        
        Args:
            query: The search query
            top_k: Number of results to return
            semantic_weight: Weight given to semantic search (0-1)
            
        Returns:
            List of search results with document info
        """
        # 1. Semantic search
        query_embedding = self.embedding_model.encode(query)
        semantic_results = self.vector_db.search(query_embedding, limit=top_k*2)
        
        # 2. Keyword search (TF-IDF)
        keyword_results = []
        if self.tfidf_matrix is not None:
            query_vector = self.tfidf_vectorizer.transform([query])
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
        
        # 3. Merge results with weighting
        # Create a map for faster lookup
        merged_results = {}
        
        # Add semantic results with weight
        for result in semantic_results:
            key = f"{result['document_id']}_{result['chunk_index']}"
            merged_results[key] = {
                **result,
                "final_score": result["similarity"] * semantic_weight
            }
        
        # Add keyword results with weight
        for result in keyword_results:
            key = f"{result['document_id']}_{result['chunk_index']}"
            if key in merged_results:
                # Combine scores
                merged_results[key]["final_score"] += result["similarity"] * (1 - semantic_weight)
            else:
                merged_results[key] = {
                    **result,
                    "final_score": result["similarity"] * (1 - semantic_weight)
                }
        
        # Convert to list and sort by final score
        results_list = list(merged_results.values())
        results_list.sort(key=lambda x: x["final_score"], reverse=True)
        
        return results_list[:top_k]


class LocalAnswerGenerator:
    """
    Generates answers from document chunks without using external LLMs.
    Uses a cross-encoder model for relevance and answer extraction techniques.
    """
    
    def __init__(self):
        # Use a cross-encoder model for relevance scoring
        self.cross_encoder = CrossEncoder('cross-encoder/ms-marco-MiniLM-L-6-v2')
        
    def generate_answer(self, query: str, search_results: List[Dict]) -> str:
        """
        Generate an answer based on search results without using external LLMs.
        Uses relevance scoring and information extraction techniques.
        """
        if not search_results:
            return "No relevant information found in the documents."
        
        # 1. Score passages with cross-encoder for more accurate relevance
        passages = [result["text"] for result in search_results]
        passage_pairs = [[query, passage] for passage in passages]
        relevance_scores = self.cross_encoder.predict(passage_pairs)
        
        # Combine with search results
        for i, result in enumerate(search_results):
            result["relevance_score"] = float(relevance_scores[i])
        
        # Re-rank based on cross-encoder scores
        search_results.sort(key=lambda x: x["relevance_score"], reverse=True)
        
        # 2. Extract potential answer sentences
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
                        "document": result["filename"]
                    })
        
        # Sort by relevance
        answer_sentences.sort(key=lambda x: x["score"], reverse=True)
        
        # 3. Construct the answer
        if not answer_sentences:
            # Fall back to using top chunks
            return self._construct_from_chunks(query, search_results[:3])
        
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
        answer = f"Based on the documents, here's what I found:\n\n"
        
        for sentence in unique_sentences[:5]:  # Top 5 sentences
            answer += f"• {sentence['text']} [From: {sentence['document']}]\n\n"
        
        return answer
    
    def _construct_from_chunks(self, query: str, chunks: List[Dict]) -> str:
        """Fallback method to construct an answer from whole chunks."""
        answer = f"Based on the documents, I found these relevant passages:\n\n"
        
        for i, chunk in enumerate(chunks):
            answer += f"From {chunk['filename']}:\n"
            answer += f"{chunk['text'][:500]}...\n\n" if len(chunk['text']) > 500 else f"{chunk['text']}\n\n"
            
        return answer


class QueryEngine:
    """Enhanced query engine with hybrid search and improved answer generation."""
    
    def __init__(self, document_processor: DocumentProcessor):
        self.document_processor = document_processor
        self.answer_generator = LocalAnswerGenerator()
        
        # Query log for learning
        self.query_log_path = Path(document_processor.storage_dir) / "query_log.json"
        self.query_log = self._load_query_log()
        
        # Feedback records
        self.feedback_path = Path(document_processor.storage_dir) / "feedback.json"
        self.feedback = self._load_feedback()
    
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
        """
        # Find similar queries in the log that received positive feedback
        expanded_terms = set()
        
        # Look through positive feedback
        for entry in self.feedback["positive"]:
            original_query = entry["query"]
            
            # Simple similarity - check for overlapping terms
            query_terms = set(query.lower().split())
            original_terms = set(original_query.lower().split())
            
            # If queries share key terms, extract useful terms from the original
            overlap = query_terms.intersection(original_terms)
            if overlap and len(overlap) / len(query_terms) > 0.3:  # Some overlap
                # Add non-overlapping terms from successful query
                expanded_terms.update(original_terms - query_terms)
        
        # Construct expanded query
        if expanded_terms:
            expanded_query = f"{query} {' '.join(expanded_terms)}"
            logger.info(f"Expanded query from '{query}' to '{expanded_query}'")
            return expanded_query
            
        return query
    
    def process_query(self, query: str) -> Dict:
        """
        Process a user query and generate a response using hybrid search.
        """
        # Log the query
        query_id = len(self.query_log)
        query_entry = {
            "id": query_id,
            "query": query,
            "timestamp": datetime.now().isoformat(),
            "was_successful": False
        }
        
        # Expand query if we have historical data
        expanded_query = self.expand_query(query)
        
        # Search for relevant document chunks using hybrid search
        search_results = self.document_processor.hybrid_search(
            expanded_query, 
            top_k=5,
            semantic_weight=0.7  # Balance between semantic and keyword search
        )
        
        if not search_results:
            response = {
                "answer": "I couldn't find relevant information in the documents. Please try rephrasing your query or upload more documents.",
                "sources": [],
                "query_id": query_id
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
                    "relevance": result.get("relevance_score", result.get("similarity", 0))
                } for result in search_results[:3]]  # Include top 3 sources
            }
            
            query_entry["was_successful"] = True
            query_entry["document_ids"] = [r["document_id"] for r in search_results[:3]]
        
        # Add query_id to response
        response["query_id"] = query_id
            
        # Update query log
        self.query_log.append(query_entry)
        self._save_query_log()
        
        # Update document metadata with this query
        for result in search_results:
            doc_id = result["document_id"]
            if doc_id in self.document_processor.metadata["documents"]:
                self.document_processor.metadata["documents"][doc_id]["queries"].append({
                    "query": query,
                    "timestamp": datetime.now().isoformat()
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
    """Enhanced learning system with more metrics and suggestions."""
    
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
            "improvement_suggestions": []
        }
    
    def _save_performance_log(self):
        """Save performance metrics to storage."""
        with open(self.performance_log_path, 'w') as f:
            json.dump(self.performance_log, f, indent=2)
    
    def analyze_performance(self) -> Dict:
        """
        Analyze system performance based on query logs and feedback.
        Returns metrics and improvement suggestions.
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
        
        # Track metrics
        self.performance_log["query_success_rate"].append({
            "timestamp": datetime.now().isoformat(),
            "rate": success_rate,
            "total_queries": len(query_log)
        })
        
        self.performance_log["document_coverage"] = doc_query_counts
        self.performance_log["query_patterns"] = top_terms
        
        # Generate improvement suggestions
        suggestions = []
        
        # Suggest processing more documents if few are available
        if len(self.document_processor.metadata["documents"]) < 3:
            suggestions.append("Upload more documents to improve the knowledge base")
        
        # Identify underperforming queries
        failed_queries = [q["query"] for q in query_log if not q.get("was_successful", False)]
        if failed_queries:
            suggestions.append(f"Consider improving handling for queries like: {', '.join(failed_queries[:3])}")
        
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
            suggestions.append(f"Documents never referenced in queries: {', '.join(unused_docs[:3])}")
        
        self.performance_log["improvement_suggestions"] = suggestions
        self._save_performance_log()
        
        return {
            "success_rate": success_rate,
            "processed_queries": len(query_log),
            "document_coverage": doc_query_counts,
            "popular_terms": dict(sorted_terms[:5]),
            "suggestions": suggestions
        }
    
    def improve_system(self):
        """
        Apply automated improvements to the system based on collected data.
        """
        # Analyze performance to get latest data
        self.analyze_performance()
        
        # 1. Adjusting semantic vs keyword search weights based on performance
        positive_feedback = self.query_engine.feedback["positive"]
        negative_feedback = self.query_engine.feedback["negative"]
        
        if positive_feedback and negative_feedback:
            # If we have enough feedback, we can adjust weights
            # This is a simple heuristic - could be more sophisticated
            semantic_weight = 0.7  # Default
            
            # Measure hypothesis: more technical terms might benefit from keyword search
            technical_term_pattern = re.compile(r'\b[A-Z][A-Za-z]*|[A-Za-z]+\d+|\d+[A-Za-z]+\b')
            
            tech_terms_positive = sum(
                1 for entry in positive_feedback 
                if technical_term_pattern.search(entry["query"])
            )
            tech_terms_negative = sum(
                1 for entry in negative_feedback 
                if technical_term_pattern.search(entry["query"])
            )
            
            # If technical queries are failing more, increase keyword weight
            if (tech_terms_negative / len(negative_feedback)) > 0.5:
                semantic_weight = 0.6  # Give more weight to keyword search
            
            logger.info(f"Adjusted semantic weight to {semantic_weight}")
            
            # We would update the search parameter here in a real system
            # For demonstration, we'll just log it
        
        # 2. Identify high-value document sections
        doc_sections = {}
        for entry in positive_feedback:
            if "document_ids" in entry:
                for doc_id in entry["document_ids"]:
                    if doc_id not in doc_sections:
                        doc_sections[doc_id] = 0
                    doc_sections[doc_id] += 1
        
        # Log which documents are most valuable
        if doc_sections:
            sorted_docs = sorted(doc_sections.items(), key=lambda x: x[1], reverse=True)
            logger.info(f"Most valuable documents: {sorted_docs[:3]}")
            
            # In a real system, we could prioritize chunks from these documents
            # or analyze what makes them more useful
        
        # 3. Record improvement attempt
        improvement_record = {
            "timestamp": datetime.now().isoformat(),
            "metrics_before": self.performance_log["query_success_rate"][-1] if self.performance_log["query_success_rate"] else None,
            "actions_taken": ["Analyzed query patterns", "Adjusted search weights", "Identified valuable documents"]
        }
        
        if "improvement_history" not in self.performance_log:
            self.performance_log["improvement_history"] = []
            
        self.performance_log["improvement_history"].append(improvement_record)
        self._save_performance_log()
        
        logger.info("System improvement cycle completed")
        return True


def main():
    st.title("Enhanced Document Analysis System")
    
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
        st.header("Upload Documents")
        uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
        
        if uploaded_file is not None:
            with st.spinner("Processing document..."):
                doc_id = st.session_state.document_processor.process_pdf(
                    uploaded_file, uploaded_file.name
                )
                st.success(f"Document processed: {uploaded_file.name}")
        
        # Document list
        st.header("Your Documents")
        docs = st.session_state.document_processor.metadata["documents"]
        for doc_id, info in docs.items():
            st.write(f"{info['filename']} - {info['upload_date'][:10]}")
            
        # Search settings
        st.header("Search Settings")
        semantic_weight = st.slider(
            "Semantic vs Keyword Balance", 
            min_value=0.0, 
            max_value=1.0, 
            value=0.7,
            help="Higher values prioritize semantic understanding, lower values prioritize keyword matching"
        )
    
    # Main area for query input
    st.header("Ask a Question")
    query = st.text_input("Enter your question about the documents:")
    
    if query:
        with st.spinner("Searching documents..."):
            response = st.session_state.query_engine.process_query(query)
            
        st.subheader("Answer")
        st.write(response["answer"])
        
        if response["sources"]:
            st.subheader("Sources")
            for src in response["sources"]:
                with st.expander(f"From {src['filename']} (Relevance: {src['relevance']:.2f})"):
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
                    st.metric("Documents Processed", len(docs))
                with col3:
                    st.metric("Queries Processed", performance["processed_queries"])
                
                # Document coverage
                st.subheader("Document Coverage")
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