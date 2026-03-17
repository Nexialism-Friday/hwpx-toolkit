#!/usr/bin/env python3
"""
Example: Vectorize HWP/HWPX documents for RAG pipelines

Prerequisites:
    pip install qdrant-client ollama
    # Or set OPENAI_API_KEY for OpenAI embeddings

    # Run Qdrant locally:
    docker run -p 6333:6333 qdrant/qdrant
"""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from hwpx_toolkit.vectorizer import HWPVectorizer

QDRANT_URL = os.environ.get("QDRANT_URL", "http://localhost:6333")
COLLECTION  = os.environ.get("QDRANT_COLLECTION", "hwpx_docs")
EMBED_MODEL = os.environ.get("EMBED_MODEL", "nomic-embed-text")  # via Ollama


def main():
    if len(sys.argv) < 2:
        print("Usage: python vectorize_docs.py <folder_or_file>")
        print()
        print("Environment variables:")
        print("  QDRANT_URL         Qdrant server URL (default: http://localhost:6333)")
        print("  QDRANT_COLLECTION  Collection name (default: hwpx_docs)")
        print("  EMBED_MODEL        Ollama model for embeddings (default: nomic-embed-text)")
        sys.exit(1)

    target = Path(sys.argv[1])
    files = []

    if target.is_file():
        files = [target]
    elif target.is_dir():
        files = list(target.glob("**/*.hwp")) + list(target.glob("**/*.hwpx"))
    else:
        print(f"Error: {target} does not exist")
        sys.exit(1)

    print(f"Found {len(files)} HWP/HWPX files")

    vectorizer = HWPVectorizer(
        qdrant_url=QDRANT_URL,
        collection_name=COLLECTION,
        embed_model=EMBED_MODEL,
    )
    vectorizer.ensure_collection()

    for i, fp in enumerate(files, 1):
        print(f"[{i}/{len(files)}] {fp.name}")
        try:
            n = vectorizer.vectorize_file(str(fp))
            print(f"  → {n} chunks indexed")
        except Exception as e:
            print(f"  ✗ Error: {e}")

    print(f"\nDone. {len(files)} files processed.")
    print(f"Collection '{COLLECTION}' at {QDRANT_URL}")


if __name__ == "__main__":
    main()
