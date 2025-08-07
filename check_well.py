#!/usr/bin/env python3
"""
chat_qwen_hf.py

Interactive Q&A/summarization over documents using Qwen3-4B-Thinking-2507 (CPU-only).
Prereqs:
  pip install transformers accelerate torch pypdf python-docx
"""

import os
import torch
from transformers import AutoTokenizer, AutoModelForCausalLM
from pypdf import PdfReader
from docx import Document

# === CONFIGURATION ===
MODEL_PATH = r"D:\Prospectus\Qwen\Qwen3-4B-Thinking-2507"
MAX_TOKENS = 512
TEMP       = 0.7

def extract_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        reader = PdfReader(path)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    if ext in (".docx", ".doc"):
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def main():
    print(f"Loading model from {MODEL_PATH} …")
    tokenizer = AutoTokenizer.from_pretrained(MODEL_PATH, trust_remote_code=True)
    model = AutoModelForCausalLM.from_pretrained(
        MODEL_PATH, 
        trust_remote_code=True, 
        device_map="cpu",
        low_cpu_mem_usage=True
    )
    # Apply dynamic INT8 quantization to reduce RAM
    model = torch.quantization.quantize_dynamic(
        model, {torch.nn.Linear}, dtype=torch.qint8
    )
    model.eval()
    print("Model ready on CPU (dynamic quantization enabled).")

    # Document attachment
    while True:
        path = input("\nEnter document path (.pdf/.docx/.txt) or 'exit': ").strip()
        if path.lower() in ("exit", "quit"):
            return
        if not os.path.isfile(path):
            print("❌ File not found.")
            continue
        text = extract_text(path)
        if not text.strip():
            print("❌ No extractable text.")
            continue
        print(f"✅ Extracted {len(text)} characters.")
        break

    # Chat loop
    print("\nAsk questions about the document (type 'exit' to quit).")
    while True:
        query = input("\nQuestion: ").strip()
        if query.lower() in ("exit", "quit"):
            print("Session ended."); break
        prompt = f"Document:\n{text}\n\nQuestion: {query}\nAnswer:"
        inputs = tokenizer(prompt, return_tensors="pt")
        outputs = model.generate(
            **inputs, 
            max_new_tokens=MAX_TOKENS, 
            temperature=TEMP
        )
        ans = tokenizer.decode(outputs[0], skip_special_tokens=True)
        # strip prompt
        print("\nAnswer:\n" + ans.split("Answer:")[-1].strip())

if __name__ == "__main__":
    main()
