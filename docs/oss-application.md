# Claude for Open Source — Application Draft

## Project Information

- **Repository**: https://github.com/Nexialism-Friday/hwpx-toolkit
- **License**: Apache-2.0
- **Language**: Python
- **PyPI**: hwpx-toolkit (pending publish)

---

## Project Description (for application form)

**hwpx-toolkit** is a Python toolkit for processing HWP and HWPX documents — the dominant document formats used by 95%+ of South Korean government agencies, businesses, and academic institutions.

HWP/HWPX (Hangul Word Processor) are proprietary formats developed by Hancom Inc. Despite their ubiquity in Korea, English-language open-source tooling has been virtually nonexistent. Existing libraries (`pyhwp`, `python-hwpx`) are either unmaintained or limited in scope.

### What this toolkit does:
1. **Text & table extraction** from both `.hwp` (binary) and `.hwpx` (ZIP+XML) formats — handles complex nested tables, footnotes, sections
2. **Document creation & editing** via Windows COM bridge — fully controllable from WSL/Linux
3. **Vectorization pipeline** — chunks and embeds HWP documents for RAG systems using Qdrant
4. **Template-based generation** — creates HWPX from structured data without COM/Office
5. **MCP server** — exposes HWP tools as Model Context Protocol endpoints for Claude/AI assistants

### Why AI assistance is critical:
This project integrates directly with Claude via MCP. I use Claude Code extensively to:
- Develop and iterate on the HWPX XML parsing logic
- Design the vectorization pipeline for Korean government documents
- Write the MCP server that connects HWP processing to AI assistants
- Build RAG pipelines for ODA (Official Development Assistance) procurement documents

### Ecosystem Impact:
- Enables AI-powered processing of Korean government documents (procurement, ODA, legal)
- First toolkit to support HWP → RAG pipeline in English-language ecosystem
- Bridges the gap between Korea's dominant document format and modern AI tooling
- Open-source under Apache-2.0 to maximize reuse across Korean civil society, academia, NGOs

### Maintainer Background:
I am an ODA/international development consultant working with Korean government agencies. I built this toolkit to process procurement documents (나라장터) and ODA project files — and am now open-sourcing it to benefit the broader Korean developer community.

---

## Application Submission

**URL**: https://claude.ai/contact-sales/claude-for-oss

**Form fields to fill:**
- Project name: hwpx-toolkit
- Repository URL: https://github.com/Nexialism-Friday/hwpx-toolkit
- Your role: Maintainer/Creator
- How you use Claude: Claude Code for development, MCP integration, RAG pipeline design
- Expected usage: Heavy development usage during v1.0 buildout (3-6 months)

---

## Appeal Letter (if needed)

> Dear Anthropic team,
>
> I am the creator of hwpx-toolkit, an open-source Python library for processing HWP/HWPX documents. While our GitHub Stars are currently modest (the project was just published), I believe this project has significant ecosystem impact that warrants consideration.
>
> HWP is the document format used by 95%+ of Korean government agencies — the Korean equivalent of .docx for the Western world. Despite millions of documents being created in this format daily, English-language open-source tooling has been nearly nonexistent. Existing projects like pyhwp haven't been maintained in years.
>
> hwpx-toolkit fills this gap by providing: text/table extraction, document generation, vectorization for RAG pipelines, and an MCP server for direct AI assistant integration.
>
> Claude has been instrumental in building this toolkit — particularly Claude Code for iterating on the complex HWPX XML parsing logic and designing the MCP server architecture. The Claude for OSS program would allow me to continue this development and eventually make HWP processing accessible to all Korean developers working with AI systems.
>
> I would be grateful for consideration even without the 5,000 star threshold.
>
> Thank you,
> Tommy Keum
