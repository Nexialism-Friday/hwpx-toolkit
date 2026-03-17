# hwpx-toolkit

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://python.org)
[![Korean](https://img.shields.io/badge/lang-한국어-red)](README.ko.md)

**Python toolkit for HWP/HWPX document processing, vectorization, and AI integration.**

HWP and HWPX are the dominant document formats used in South Korea (government, legal, academic). This toolkit provides:

- **Text & table extraction** from HWP and HWPX files
- **Document creation & editing** via Windows COM bridge (WSL-compatible)
- **Vectorization** for RAG (Retrieval-Augmented Generation) pipelines
- **Template-based HWPX generation** without requiring Hancom Office
- **MCP server** for AI assistant integration (Claude, etc.)

> HWP/HWPX are used by 95%+ of Korean government agencies and businesses. English-language open-source tooling has been virtually nonexistent — until now.

---

## Features

### 1. Extractor (`hwpx_toolkit.extractor`)
- Extracts text and tables from both `.hwp` (binary) and `.hwpx` (ZIP+XML) formats
- Handles complex nested tables, footnotes, and section structures
- Falls back gracefully between parsing methods

### 2. Writer (`hwpx_toolkit.writer`)
- Creates HWP documents with full formatting control (fonts, margins, tables, images)
- Edits existing documents: find/replace, regex substitution, cell formatting
- Runs as a Windows COM bridge — fully controllable from WSL/Linux
- Supports headers, footers, footnotes, page breaks, and page setup

### 3. Vectorizer (`hwpx_toolkit.vectorizer`)
- Chunks extracted text for embedding
- Integrates with Qdrant vector database
- Supports batch processing of large document collections
- Designed for ODA/government document RAG pipelines

### 4. Generator (`hwpx_toolkit.generator`)
- Generates HWPX documents from structured data (no COM/Office required)
- Template-based document creation
- Useful for automated report generation

### 5. MCP Server (`hwpx_toolkit.mcp_server`)
- Exposes HWP/HWPX tools as [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) endpoints
- Enables direct integration with Claude and other AI assistants
- DOCX/MD → HWPX conversion pipeline

---

## Installation

```bash
pip install hwpx-toolkit
```

Or from source:

```bash
git clone https://github.com/YOUR_USERNAME/hwpx-toolkit.git
cd hwpx-toolkit
pip install -e .
```

**Requirements:**
- Python 3.10+
- `hwp5` (for legacy `.hwp` parsing): `pip install pyhwp`
- For Writer: Windows + Hancom HWP installed (COM automation)
- For Vectorizer: Qdrant instance

---

## Quick Start

### Extract text from HWP/HWPX

```python
from hwpx_toolkit.extractor import extract_text

text = extract_text("document.hwpx")
print(text)
```

### Create a new HWP document

```python
import json, subprocess

commands = [
    {"action": "page_setup", "paper": "A4"},
    {"action": "text", "content": "Hello, World!", "font_size": 12},
    {"action": "table", "rows": 3, "cols": 3, "data": [["A","B","C"],["1","2","3"],["x","y","z"]]},
]

with open("commands.json", "w") as f:
    json.dump({"mode": "create", "output": "output.hwp", "commands": commands}, f)

subprocess.run(["python", "hwpx_toolkit/writer.py", "commands.json"])
```

### Vectorize documents for RAG

```python
from hwpx_toolkit.vectorizer import HWPVectorizationEngine

engine = HWPVectorizationEngine(qdrant_host="localhost", qdrant_port=6333)
engine.process_file("document.hwpx", collection_name="my_docs")
```

---

## Background

This toolkit was developed to support an AI-powered document processing pipeline for Korean ODA (Official Development Assistance) projects. Korean government documents are almost exclusively HWP/HWPX format, yet existing Python tools for this format are either unmaintained or incomplete.

Key gaps this project fills:
- **No maintained HWPX writer** existed for Python
- **No vectorization pipeline** for HWP documents
- **No MCP integration** for AI-assisted document workflows
- **WSL/Linux compatibility** for HWP automation

---

## Project Structure

```
hwpx-toolkit/
├── hwpx_toolkit/
│   ├── extractor.py      # HWP/HWPX text & table extraction
│   ├── writer.py         # Document creation & editing (v10)
│   ├── vectorizer.py     # Vector embedding pipeline
│   ├── generator.py      # Template-based HWPX generation
│   └── mcp_server.py     # MCP server for AI integration
├── examples/             # Usage examples
├── docs/                 # Documentation
├── LICENSE               # Apache-2.0
└── NOTICE                # Attribution notices
```

---

## Contributing

Contributions welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) before submitting PRs.

Areas where help is appreciated:
- Pure Python HWPX writer (eliminating Windows COM dependency)
- Additional document format support (HML, OWPML)
- Test coverage

---

## License

Apache License 2.0 — see [LICENSE](LICENSE) for details.

This project uses the publicly documented HWPX specification ([KS X 6101](https://www.hancom.com/support/downloadCenter/hwpOwpml)) and does not include any proprietary Hancom code.

---

## Related Projects

- [pyhwp](https://github.com/mete0r/pyhwp) — HWP5 binary format parser (AGPL-3.0)
- [hwp5](https://pypi.org/project/pyhwp/) — CLI tools for HWP5
