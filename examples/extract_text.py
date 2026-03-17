#!/usr/bin/env python3
"""
Example: Extract text and tables from HWP/HWPX files
"""
import sys
from pathlib import Path

# Add parent to path if running from examples/
sys.path.insert(0, str(Path(__file__).parent.parent))

from hwpx_toolkit.extractor import extract_text, extract_tables

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_text.py <file.hwp|file.hwpx>")
        print()
        print("Example:")
        print("  python extract_text.py document.hwpx")
        sys.exit(1)

    filepath = sys.argv[1]

    print(f"Processing: {filepath}")
    print("=" * 60)

    # Extract text
    print("\n[TEXT EXTRACTION]")
    text = extract_text(filepath)
    print(text[:2000])  # Show first 2000 chars
    if len(text) > 2000:
        print(f"\n... ({len(text) - 2000} more characters)")

    # Extract tables
    print("\n[TABLE EXTRACTION]")
    tables = extract_tables(filepath)
    if tables:
        for i, table in enumerate(tables):
            print(f"\nTable {i+1}:")
            for row in table:
                print(" | ".join(str(cell) for cell in row))
    else:
        print("No tables found.")


if __name__ == "__main__":
    main()
