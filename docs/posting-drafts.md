# hwpx-toolkit 홍보 포스팅 초안

---

## 1. Hacker News (Show HN)

**Title:**
```
Show HN: hwpx-toolkit – Open-source Python library for HWP/HWPX documents (Korean .docx)
```

**Body:**
```
HWP/HWPX is the dominant document format in South Korea — used by 95%+ of government agencies,
courts, schools, and businesses. Think of it as Korea's .docx.

Despite millions of HWP documents being created daily, English-language open-source tooling
has been virtually nonexistent. Existing Python libraries (pyhwp, python-hwpx) are either
unmaintained or limited in scope.

hwpx-toolkit fills this gap:

- Text & table extraction from .hwp (binary, OLE) and .hwpx (ZIP+XML)
- Document creation and editing via Windows COM bridge (WSL-compatible)
- Vectorization pipeline for RAG (Qdrant integration)
- Template-based HWPX generation — no COM/Hancom Office required
- MCP server for direct Claude/AI assistant integration

Built while processing Korean ODA procurement documents (나라장터).
Clean-room implementation based on the publicly released HWP/HWPML specification (KS X 6101).

GitHub: https://github.com/Nexialism-Friday/hwpx-toolkit
License: Apache-2.0

Happy to answer questions about the HWP format — it's a fascinating (and sometimes infuriating)
binary format with 30 years of history.
```

---

## 2. Reddit — r/learnpython / r/Python

**Title:**
```
I open-sourced a Python toolkit for HWP/HWPX files — Korea's dominant document format that
most Western tools can't handle
```

**Body:**
```
If you've ever worked with Korean government documents, you know the pain of HWP files.

HWP (Hangul Word Processor) is used by essentially every Korean government agency, court,
and school. It's like Korea's version of .docx — except Python support has been nearly
nonexistent.

I spent months building a toolkit to process these documents for AI/RAG pipelines and
decided to open-source it:

**hwpx-toolkit** on GitHub: https://github.com/Nexialism-Friday/hwpx-toolkit

**What it can do:**
- Extract text and tables from both .hwp (binary OLE) and .hwpx (ZIP+XML)
- Create/edit HWP documents from Python (WSL ↔ Windows COM bridge)
- Vectorize documents for RAG pipelines (Qdrant)
- Generate HWPX from structured data without needing Hancom Office installed
- MCP server for AI assistant integration

**Built for:** Processing Korean ODA/procurement documents — but useful for anyone dealing
with Korean government or business documents.

**License:** Apache-2.0

Would love feedback from anyone who's dealt with HWP files before!
```

---

## 3. X (Twitter/X)

**Thread:**
```
Tweet 1:
Just open-sourced hwpx-toolkit — a Python library for HWP/HWPX documents (Korea's dominant
doc format, used by 95%+ of gov agencies).

English-language OSS for HWP has been almost nonexistent. Time to change that.

→ https://github.com/Nexialism-Friday/hwpx-toolkit

Thread 🧵

Tweet 2:
What is HWP?

Korea's equivalent of .docx — but with 30 years of history and millions of daily documents.
Every Korean court filing, government procurement, academic paper... HWP.

Yet Python support? Nearly dead projects from 5+ years ago.

Tweet 3:
hwpx-toolkit supports:

• Text + table extraction (.hwp binary & .hwpx XML)
• Document creation via COM bridge (works from WSL!)
• Vectorization → Qdrant for RAG pipelines
• Template-based HWPX generation (no Office needed)
• MCP server for Claude/AI integration

Tweet 4:
Why I built it:

I process Korean government ODA procurement docs (나라장터) for AI analysis.
Needed reliable HWP → text → vector pipeline.

Existing tools couldn't handle complex tables, nested structures, or HWPX.

So I built one. V10 is now public.

Tweet 5:
Clean-room implementation based on Hancom's published spec (KS X 6101 / HWPML).
Apache-2.0 licensed.

If you work with Korean documents, give it a ⭐ — helps more Koreans find it.

https://github.com/Nexialism-Friday/hwpx-toolkit
```

---

## 4. 국내 커뮤니티 (한국어)

### 개발자 커뮤니티 (OKKY / 개발자 오픈채팅)

```
제목: HWP/HWPX Python 처리 오픈소스 라이브러리 공개했습니다

안녕하세요.

나라장터 공문/조달 문서 AI 분석 시스템을 개발하면서 만든 HWP/HWPX 처리 라이브러리를 오픈소스로 공개했습니다.

hwpx-toolkit: https://github.com/Nexialism-Friday/hwpx-toolkit

주요 기능:
- .hwp (OLE 바이너리) / .hwpx (ZIP+XML) 텍스트·표 추출
- Python에서 HWP 문서 생성/편집 (WSL ↔ Windows COM 브릿지)
- RAG 파이프라인 연동 (Qdrant 벡터DB)
- 한컴 오피스 없이 HWPX 템플릿 기반 생성
- Claude/AI 어시스턴트 MCP 서버 연동

한국 공공문서 AI 분석이나 HWP 자동화 작업에 관심 있으신 분들께 도움이 되길 바랍니다.

라이선스: Apache-2.0
피드백/기여 환영합니다!
```

### X (한국어 개발자 계정)

```
HWP/HWPX Python 처리 라이브러리 오픈소스로 공개했습니다.

나라장터 공문 AI 분석 시스템 개발하면서 만든 건데, 쓸만한 한국어 오픈소스가 없어서 직접 만들었어요.

- HWP/HWPX 텍스트·표 추출
- Python으로 HWP 생성/편집 (WSL도 됩니다)
- RAG 파이프라인 연동
- Claude MCP 서버 포함

스타 주시면 더 많은 분들이 찾을 수 있어요 🙏
→ https://github.com/Nexialism-Friday/hwpx-toolkit
```

---

## 포스팅 타이밍 전략

| 채널 | 시간 | 주의사항 |
|------|------|---------|
| Hacker News | 월~화 오전 9-11시 (미국 동부) = 한국 밤 10시-자정 | "Show HN:" 필수 |
| Reddit r/Python | 화~목 오전 | 자기홍보 규정 확인 |
| X 영문 | 미국 동부 오전 9-11시 | 스레드로 작성 |
| X 한국어 | 한국 오전 9-11시 | 개발자 해시태그 |
| OKKY | 주중 오전 | 커뮤니티 규정 확인 |
