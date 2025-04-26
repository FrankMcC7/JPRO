import unittest
import os
import tempfile
import shutil
from pathlib import Path

# Import the system components
from document_qa_system import (
    DocumentProcessor, 
    SemanticProcessor,
    VectorStore,
    QuestionAnswerer,
    DocumentQASystem
)

class TestDocumentQASystem(unittest.TestCase):
    """Test cases for the Document QA System components."""
    
    @classmethod
    def setUpClass(cls):
        """Set up test environment."""
        # Create a temporary directory for test files
        cls.test_dir = tempfile.mkdtemp()
        
        # Create a simple test PDF (this is just a placeholder - in real testing you'd use a real PDF)
        cls.test_pdf_path = os.path.join(cls.test_dir, "test_document.pdf")
        with open(cls.test_pdf_path, "w") as f:
            f.write("This is a test PDF file.")
        
        # Initialize components for testing
        cls.doc_processor = DocumentProcessor()
        cls.semantic_processor = SemanticProcessor()
        cls.vector_store = VectorStore(collection_name="test_collection")
        cls.qa_system = DocumentQASystem()
    
    @classmethod
    def tearDownClass(cls):
        """Clean up after tests."""
        # Remove the temporary directory and its contents
        shutil.rmtree(cls.test_dir)
        
        # Clean up vector store test collection
        try:
            cls.vector_store.client.delete_collection("test_collection")
        except:
            pass
    
    def test_document_processor_initialization(self):
        """Test DocumentProcessor initialization."""
        processor = DocumentProcessor(chunk_size=500, chunk_overlap=100)
        self.assertEqual(processor.chunk_size, 500)
        self.assertEqual(processor.chunk_overlap, 100)
    
    def test_text_cleaning(self):
        """Test text cleaning function."""
        dirty_text = "This is a  test\n\n\nwith   multiple spaces\n and newlines."
        cleaned_text = self.doc_processor._clean_text(dirty_text)
        self.assertNotEqual(dirty_text, cleaned_text)
        self.assertNotIn("  ", cleaned_text)  # No double spaces
        self.assertNotIn("\n\n\n", cleaned_text)  # No triple newlines
    
    def test_chunk_creation(self):
        """Test text chunking function."""
        long_text = " ".join(["This is sentence number " + str(i) + "." for i in range(100)])
        chunks = self.doc_processor._create_chunks(long_text)
        
        # Should create multiple chunks
        self.assertGreater(len(chunks), 1)
        
        # First chunk should contain the beginning of the text
        self.assertTrue(chunks[0].startswith("This is sentence number 0"))
    
    def test_embedding_generation(self):
        """Test embedding generation."""
        test_texts = ["This is a test sentence.", "This is another test sentence."]
        embeddings = self.semantic_processor.get_embeddings(test_texts)
        
        # Should return the right number of embeddings
        self.assertEqual(len(embeddings), len(test_texts))
        
        # Embeddings should be numerical arrays with the expected dimensions
        self.assertEqual(embeddings.ndim, 2)
        
        # Different texts should have different embeddings
        self.assertFalse(all(embeddings[0] == embeddings[1]))
    
    def test_vector_store_operations(self):
        """Test vector store operations."""
        # Test adding documents
        test_chunks = ["This is chunk 1.", "This is chunk 2.", "This is chunk 3."]
        test_doc_id = "test_doc_123"
        
        self.vector_store.add_documents(test_doc_id, test_chunks)
        
        # Test searching
        search_results = self.vector_store.search("chunk 2", n_results=1)
        
        # Should return the right number of results
        self.assertEqual(len(search_results["documents"]), 1)
        
        # Should return the most relevant chunk
        self.assertEqual(search_results["documents"][0], "This is chunk 2.")
    
    def test_question_answerer(self):
        """Test question answering functionality."""
        qa = QuestionAnswerer()
        
        context = ["Paris is the capital of France.", 
                  "Berlin is the capital of Germany.",
                  "Rome is the capital of Italy."]
        
        result = qa.answer_question("What is the capital of France?", context)
        
        # Should return an answer
        self.assertIsNotNone(result["answer"])
        
        # Should return "Paris" with high confidence
        self.assertEqual(result["answer"], "Paris")
        self.assertGreater(result["confidence"], 0.5)
    
    def test_integrated_system(self):
        """Test the integrated document QA system."""
        # This is a simplified integration test - in a real scenario, you would use actual PDFs
        
        # First, mock the process_pdf method to return predefined chunks
        original_process_pdf = self.qa_system.document_processor.process_pdf
        
        def mock_process_pdf(file_path):
            return [
                "Albert Einstein was born in Ulm, Germany.",
                "Einstein developed the theory of relativity.",
                "Einstein won the Nobel Prize in Physics in 1921."
            ]
        
        self.qa_system.document_processor.process_pdf = mock_process_pdf
        
        try:
            # Process a document
            doc_id = self.qa_system.process_document(self.test_pdf_path, document_id="einstein_bio")
            
            # Ask a question
            result = self.qa_system.answer_question("Where was Einstein born?")
            
            # Validate results
            self.assertTrue("Ulm" in result["answer"])
            self.assertTrue("Germany" in result["answer"])
            self.assertGreater(result["confidence"], 0.5)
            
        finally:
            # Restore the original method
            self.qa_system.document_processor.process_pdf = original_process_pdf


if __name__ == "__main__":
    unittest.main()
