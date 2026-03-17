# hwpx-toolkit

[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://python.org)
[![English](https://img.shields.io/badge/lang-English-blue)](README.md)

**HWP/HWPX 문서 처리, 벡터화, AI 통합을 위한 Python 툴킷**

HWP/HWPX는 대한민국 공공기관, 법조계, 학계에서 사용하는 표준 문서 포맷입니다. 이 툴킷은 다음을 제공합니다:

- **텍스트/표 추출** — HWP 및 HWPX 파일에서 텍스트와 표 구조 추출
- **문서 생성/편집** — Windows COM 브릿지 기반 (WSL 완전 호환)
- **벡터화** — RAG(검색 증강 생성) 파이프라인용 임베딩
- **HWPX 템플릿 생성** — 한컴 오피스 없이 HWPX 문서 자동 생성
- **MCP 서버** — Claude 등 AI 어시스턴트와 직접 연동

---

## 주요 기능

### 1. Extractor — 텍스트/표 추출
- `.hwp`(바이너리) 및 `.hwpx`(ZIP+XML) 형식 모두 지원
- 복잡한 중첩 표, 각주, 섹션 구조 처리
- 자동 폴백 파싱

### 2. Writer — 문서 생성/편집 (v10)
- 용지 크기/여백, 폰트, 표, 이미지 등 완전한 서식 제어
- 찾기/바꾸기, 정규식 치환, 셀 서식 변경
- WSL/Linux에서 Windows 한컴 오피스 제어
- 머리말, 꼬리말, 각주, 페이지 나누기 지원

### 3. Vectorizer — RAG 벡터화
- 문서 청킹 + 임베딩
- Qdrant 벡터 DB 연동
- 대용량 문서 일괄 처리
- ODA/공공문서 RAG 파이프라인 최적화

### 4. Generator — HWPX 자동 생성
- 구조화 데이터로 HWPX 문서 생성 (COM 불필요)
- 템플릿 기반 보고서 자동화

### 5. MCP Server — AI 통합
- MCP(Model Context Protocol) 엔드포인트로 HWP 도구 노출
- Claude, GPT 등 AI 어시스턴트와 직접 연동
- DOCX/MD → HWPX 변환 파이프라인

---

## 설치

```bash
pip install hwpx-toolkit
```

소스에서 설치:

```bash
git clone https://github.com/YOUR_USERNAME/hwpx-toolkit.git
cd hwpx-toolkit
pip install -e .
```

---

## 라이선스

Apache License 2.0 — 자세한 내용은 [LICENSE](LICENSE) 참조

HWPX 공개 스펙([KS X 6101](https://www.hancom.com/support/downloadCenter/hwpOwpml)) 기반으로 개발되었으며, 한컴 독점 코드를 포함하지 않습니다.
