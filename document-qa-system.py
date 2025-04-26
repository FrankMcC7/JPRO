import os
import logging
from typing import List, Dict, Any, Optional
import tempfile

# Document processing
import fitz  # PyMuPDF for PDF processing
import re
from pathlib import Path

# NLP and semantic processing
import nltk
from nltk.tokenize import sent_tokenize
nltk.download('punkt', quiet=True)
import numpy as np

# Vector database 
import chromadb
from chromadb.utils import embedding_functions

# Model dependencies
from transformers import AutoTokenizer, AutoModel, pipeline
import torch

# API 
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DocumentProcessor:
    """Handles document loading, parsing, and chunking."""
    
    def __init__(self, chunk_size: int = 1000, chunk_overlap: int = 200):
        """
        Initialize the document processor.
        
        Args:
            chunk_size: The target size of each text chunk
            chunk_overlap: The overlap between consecutive chunks
        """
        self.chunk_size = chunk_size
        self.chunk_overlap = chunk_overlap
        logger.info(f"Initialized DocumentProcessor with chunk_size={chunk_size}, chunk_overlap={chunk_overlap}")
    
    def process_pdf(self, file_path: str) -> List[str]:
        """
        Extract text from a PDF file and split it into chunks.
        
        Args:
            file_path: Path to the PDF file
            
        Returns:
            List of text chunks
        """
        logger.info(f"Processing PDF: {file_path}")
        try:
            # Open the PDF
            doc = fitz.open(file_path)
            full_text = ""
            
            # Extract text from each page
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                full_text += page.get_text()
            
            # Clean the text
            full_text = self._clean_text(full_text)
            
            # Split into chunks
            chunks = self._create_chunks(full_text)
            
            logger.info(f"Successfully processed PDF, created {len(chunks)} chunks")
            return chunks
            
        except Exception as e:
            logger.error(f"Error processing PDF {file_path}: {str(e)}")
            raise
    
    def _clean_text(self, text: str) -> str:
        """
        Clean extracted text by removing extra whitespace, headers, footers, etc.
        
        Args:
            text: Raw extracted text
            
        Returns:
            Cleaned text
        """
        # Remove excessive newlines
        text = re.sub(r'\n{3,}', '\n\n', text)
        
        # Remove excessive spaces
        text = re.sub(r' {2,}', ' ', text)
        
        # Handle typical PDF extraction artifacts
        text = re.sub(r'(\w+)-\s*\n\s*(\w+)', r'\1\2', text)
        
        return text.strip()
    
    def _create_chunks(self, text: str) -> List[str]:
        """
        Split text into overlapping chunks based on chunk_size and chunk_overlap.
        
        Args:
            text: Text to be split into chunks
            
        Returns:
            List of text chunks
        """
        # Split text into sentences
        sentences = sent_tokenize(text)
        
        chunks = []
        current_chunk = []
        current_chunk_size = 0
        
        for sentence in sentences:
            sentence_size = len(sentence)
            
            # If adding this sentence would exceed chunk_size, 
            # save the current chunk and start a new one
            if current_chunk_size + sentence_size > self.chunk_size and current_chunk:
                chunks.append(" ".join(current_chunk))
                
                # Create overlap with previous chunk
                overlap_size = 0
                overlap_chunk = []
                
                # Add sentences from the end of the previous chunk for overlap
                for s in reversed(current_chunk):
                    if overlap_size + len(s) <= self.chunk_overlap:
                        overlap_chunk.insert(0, s)
                        overlap_size += len(s)
                    else:
                        break
                        
                # Start new chunk with overlap sentences
                current_chunk = overlap_chunk
                current_chunk_size = overlap_size
            
            # Add the current sentence to the chunk
            current_chunk.append(sentence)
            current_chunk_size += sentence_size
        
        # Add the last chunk if it's not empty
        if current_chunk:
            chunks.append(" ".join(current_chunk))
            
        return chunks


class SemanticProcessor:
    """
    Handles semantic processing using transformer models for embeddings.
    """
    
    def __init__(self, model_name: str = "sentence-transformers/all-MiniLM-L6-v2"):
        """
        Initialize the semantic processor with a transformer model.
        
        Args:
            model_name: The name of the pre-trained model to use
        """
        logger.info(f"Initializing SemanticProcessor with model: {model_name}")
        self.model_name = model_name
        
        # Load model and tokenizer
        self.tokenizer = AutoTokenizer.from_pretrained(model_name)
        self.model = AutoModel.from_pretrained(model_name)
        
        # Move model to GPU if available
        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        logger.info(f"Using device: {self.device}")
        self.model = self.model.to(self.device)
    
    def get_embeddings(self, texts: List[str]) -> np.ndarray:
        """
        Generate embeddings for a list of text chunks.
        
        Args:
            texts: List of text chunks
            
        Returns:
            Array of embeddings
        """
        embeddings = []
        
        for text in texts:
            # Tokenize and prepare for model
            inputs = self.tokenizer(text, return_tensors="pt", 
                                   padding=True, truncation=True, max_length=512)
            inputs = {k: v.to(self.device) for k, v in inputs.items()}
            
            # Generate embeddings
            with torch.no_grad():
                outputs = self.model(**inputs)
                
            # Use mean pooling to get sentence embeddings
            attention_mask = inputs['attention_mask']
            token_embeddings = outputs.last_hidden_state
            
            # Mask tokens that are padding
            input_mask_expanded = attention_mask.unsqueeze(-1).expand(token_embeddings.size()).float()
            mask_embeddings = token_embeddings * input_mask_expanded
            
            # Sum and average
            sum_embeddings = torch.sum(mask_embeddings, 1)
            sum_mask = torch.clamp(input_mask_expanded.sum(1), min=1e-9)
            embedding = (sum_embeddings / sum_mask).squeeze().cpu().numpy()
            
            embeddings.append(embedding)
        
        return np.array(embeddings)


class VectorStore:
    """
    Manages document embeddings storage and retrieval using a vector database.
    """
    
    def __init__(self, collection_name: str = "document_chunks"):
        """
        Initialize the vector store.
        
        Args:
            collection_name: Name of the collection in the vector database
        """
        logger.info(f"Initializing VectorStore with collection: {collection_name}")
        self.collection_name = collection_name
        
        # Use sentence-transformers as the embedding function
        self.embedding_function = embedding_functions.SentenceTransformerEmbeddingFunction(
            model_name="sentence-transformers/all-MiniLM-L6-v2"
        )
        
        # Initialize ChromaDB client (persistent)
        self.client = chromadb.PersistentClient(path="./chroma_db")
        
        # Get or create collection
        try:
            self.collection = self.client.get_collection(
                name=collection_name,
                embedding_function=self.embedding_function
            )
            logger.info(f"Retrieved existing collection: {collection_name}")
        except:
            self.collection = self.client.create_collection(
                name=collection_name,
                embedding_function=self.embedding_function
            )
            logger.info(f"Created new collection: {collection_name}")
    
    def add_documents(self, 
                     document_id: str, 
                     chunks: List[str], 
                     metadata: Optional[List[Dict[str, Any]]] = None) -> None:
        """
        Add document chunks to the vector database.
        
        Args:
            document_id: ID of the document
            chunks: List of text chunks
            metadata: Optional metadata for each chunk
        """
        # Create unique IDs for each chunk
        chunk_ids = [f"{document_id}_{i}" for i in range(len(chunks))]
        
        # Create metadata for each chunk if not provided
        if metadata is None:
            metadata = [{"document_id": document_id, "chunk_index": i} for i in range(len(chunks))]
        
        # Add chunks to the vector store
        logger.info(f"Adding {len(chunks)} chunks to vector store for document {document_id}")
        self.collection.add(
            ids=chunk_ids,
            documents=chunks,
            metadatas=metadata
        )
    
    def search(self, query: str, n_results: int = 5) -> Dict[str, Any]:
        """
        Search for relevant document chunks based on a query.
        
        Args:
            query: The search query
            n_results: Number of results to return
            
        Returns:
            Search results containing chunks, relevance scores, and metadata
        """
        logger.info(f"Searching for: '{query}' with n_results={n_results}")
        
        results = self.collection.query(
            query_texts=[query],
            n_results=n_results
        )
        
        return {
            "ids": results.get("ids", [[]])[0],
            "documents": results.get("documents", [[]])[0],
            "distances": results.get("distances", [[]])[0],
            "metadatas": results.get("metadatas", [[]])[0]
        }


class QuestionAnswerer:
    """
    Handles question answering using a transformer model and retrieved context.
    """
    
    def __init__(self, model_name: str = "deepset/roberta-base-squad2"):
        """
        Initialize the question answerer with a QA model.
        
        Args:
            model_name: Name of the pre-trained QA model
        """
        logger.info(f"Initializing QuestionAnswerer with model: {model_name}")
        self.qa_pipeline = pipeline(
            "question-answering",
            model=model_name,
            tokenizer=model_name,
            device=0 if torch.cuda.is_available() else -1
        )
    
    def answer_question(self, question: str, contexts: List[str]) -> Dict[str, Any]:
        """
        Answer a question based on provided contexts.
        
        Args:
            question: The question to answer
            contexts: List of relevant document chunks
            
        Returns:
            Answer with confidence score and source context
        """
        logger.info(f"Answering question: '{question}'")
        
        # Combine contexts
        combined_context = " ".join(contexts)
        
        # Get answer using the QA pipeline
        result = self.qa_pipeline(
            question=question,
            context=combined_context,
            handle_impossible_answer=True
        )
        
        # Identify which context the answer came from
        source_context = None
        for context in contexts:
            if result["answer"] in context:
                source_context = context
                break
        
        return {
            "answer": result["answer"],
            "confidence": result["score"],
            "source_context": source_context,
            "start": result["start"],
            "end": result["end"]
        }


class DocumentQASystem:
    """
    Integrates all components for a complete document QA system.
    """
    
    def __init__(self):
        """Initialize the document QA system with all required components."""
        logger.info("Initializing DocumentQASystem")
        self.document_processor = DocumentProcessor()
        self.semantic_processor = SemanticProcessor()
        self.vector_store = VectorStore()
        self.question_answerer = QuestionAnswerer()
        
        # Track processed documents
        self.processed_documents = set()
    
    def process_document(self, file_path: str, document_id: Optional[str] = None) -> str:
        """
        Process a document and store it in the vector database.
        
        Args:
            file_path: Path to the document file
            document_id: Optional custom document ID
            
        Returns:
            Document ID
        """
        # Generate document ID if not provided
        if document_id is None:
            document_id = Path(file_path).stem
        
        logger.info(f"Processing document: {file_path} with ID: {document_id}")
        
        # Extract text chunks from the document
        chunks = self.document_processor.process_pdf(file_path)
        
        # Add document chunks to the vector store
        self.vector_store.add_documents(document_id, chunks)
        
        # Mark document as processed
        self.processed_documents.add(document_id)
        
        return document_id
    
    def answer_question(self, question: str, n_results: int = 5) -> Dict[str, Any]:
        """
        Answer a question based on processed documents.
        
        Args:
            question: The question to answer
            n_results: Number of relevant chunks to retrieve
            
        Returns:
            Answer with metadata
        """
        if not self.processed_documents:
            logger.warning("No documents have been processed yet")
            return {
                "answer": "No documents have been processed yet. Please upload a document first.",
                "confidence": 0.0,
                "source_contexts": []
            }
        
        # Search for relevant chunks
        search_results = self.vector_store.search(question, n_results=n_results)
        relevant_chunks = search_results["documents"]
        
        # Answer the question using retrieved chunks
        answer = self.question_answerer.answer_question(question, relevant_chunks)
        
        # Enhance response with search metadata
        return {
            "answer": answer["answer"],
            "confidence": answer["confidence"],
            "source_contexts": relevant_chunks,
            "source_document_ids": [meta.get("document_id") for meta in search_results["metadatas"]],
            "relevance_scores": [1.0 - dist for dist in search_results["distances"]]
        }


# API Models
class QuestionRequest(BaseModel):
    question: str
    document_id: Optional[str] = None
    n_results: int = 5

class QuestionResponse(BaseModel):
    answer: str
    confidence: float
    source_contexts: List[str]
    source_document_ids: List[str]
    relevance_scores: List[float]

# Create FastAPI app
app = FastAPI(title="Document QA System API")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For development - restrict in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize the document QA system
qa_system = DocumentQASystem()

@app.post("/upload_document")
async def upload_document(file: UploadFile = File(...)):
    """
    Upload and process a document.
    """
    if not file.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are supported")
    
    # Save the uploaded file to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
        temp_file_path = temp_file.name
        contents = await file.read()
        temp_file.write(contents)
    
    try:
        # Process the document
        document_id = qa_system.process_document(temp_file_path)
        
        return {"message": "Document processed successfully", "document_id": document_id}
    except Exception as e:
        logger.error(f"Error processing document: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error processing document: {str(e)}")
    finally:
        # Clean up the temporary file
        os.unlink(temp_file_path)

@app.post("/ask_question", response_model=QuestionResponse)
async def ask_question(request: QuestionRequest):
    """
    Answer a question based on processed documents.
    """
    try:
        # Get answer from QA system
        answer = qa_system.answer_question(
            question=request.question,
            n_results=request.n_results
        )
        
        return answer
    except Exception as e:
        logger.error(f"Error answering question: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error answering question: {str(e)}")

@app.get("/health")
async def health_check():
    """
    Health check endpoint.
    """
    return {"status": "healthy"}

if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
