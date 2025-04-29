def verify_models_exist():
    """Verify that the required models exist in the offline cache."""
    import os
    from pathlib import Path
    
    home_dir = str(Path.home())  # Get home directory as string
    cache_dir = os.path.join(home_dir, ".cache", "huggingface", "hub")
    
    sentence_transformer_dir = os.path.join(cache_dir, "models--sentence-transformers--all-MiniLM-L6-v2")
    cross_encoder_dir = os.path.join(cache_dir, "models--cross-encoder--ms-marco-MiniLM-L-6-v2")
    
    models_exist = True
    
    if not os.path.exists(sentence_transformer_dir):
        logger.error(f"Sentence transformer model not found in cache! Expected location: {sentence_transformer_dir}")
        models_exist = False
        
    if not os.path.exists(cross_encoder_dir):
        logger.error(f"Cross-encoder model not found in cache! Expected location: {cross_encoder_dir}")
        models_exist = False
    
    if models_exist:
        logger.info("✓ Model files found in the expected locations")
    
    return models_exist
