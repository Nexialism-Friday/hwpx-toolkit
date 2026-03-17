#!/usr/bin/env python3
"""
HWP MCP Server
- DOCX/MD → HWPX 자동 변환
- HWPX 직접 생성 (HWP COM 불필요)

Configuration via environment variables:
  HWPX_TEMPLATE_DIR : path to HWPX template directory (required)
  HWPX_OUTPUT_DIR   : default output directory (default: /tmp/hwpx_output)
  TELEGRAM_BOT_TOKEN: Telegram bot token for send_hwpx_telegram tool
  TELEGRAM_CHAT_ID  : Telegram chat ID for send_hwpx_telegram tool
"""

import asyncio
import os
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

TEMPLATE_DIR = os.environ.get("HWPX_TEMPLATE_DIR", str(Path.home() / "hwpx_template"))
OUTPUT_DIR = os.environ.get("HWPX_OUTPUT_DIR", "/tmp/hwpx_output")

app = Server("hwp-mcp")

# ── 스타일 매핑 ──
CHAR_PR = {"title1": "8", "title2": "9", "body": "7", "bullet": "7", "blank": "7"}
PARA_PR = {"title1": "2", "title2": "3", "body": "0", "bullet": "1", "blank": "0"}

NAMESPACES = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)


def escape_xml(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def md_to_paragraphs(md_text: str) -> list:
    """Markdown 텍스트 → 단락 리스트 변환"""
    paragraphs = []
    for line in md_text.splitlines():
        stripped = line.strip()
        if not stripped:
            paragraphs.append(("blank", ""))
        elif stripped.startswith("# "):
            paragraphs.append(("title1", stripped[2:]))
        elif stripped.startswith("## "):
            paragraphs.append(("title2", stripped[3:]))
        elif stripped.startswith("### "):
            paragraphs.append(("title2", stripped[4:]))
        elif stripped.startswith(("- ", "* ", "• ")):
            paragraphs.append(("bullet", stripped[2:]))
        elif re.match(r"^\d+\.", stripped):
            paragraphs.append(("bullet", re.sub(r"^\d+\.\s*", "", stripped)))
        else:
            paragraphs.append(("body", stripped))
    return paragraphs


def docx_to_paragraphs(docx_path: str) -> list:
    """DOCX → 단락 리스트 변환 (python-docx 사용)"""
    try:
        from docx import Document
        doc = Document(docx_path)
        paragraphs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                paragraphs.append(("blank", ""))
                continue
            style = para.style.name.lower()
            if "heading 1" in style or "제목 1" in style:
                paragraphs.append(("title1", text))
            elif "heading 2" in style or "제목 2" in style:
                paragraphs.append(("title2", text))
            elif "list" in style or text.startswith(("•", "-", "*")):
                paragraphs.append(("bullet", text.lstrip("•-* ")))
            else:
                paragraphs.append(("body", text))
        return paragraphs
    except ImportError:
        return [("body", "[python-docx 미설치: pip install python-docx]")]


def make_para_xml(para_type: str, text: str, pid: int) -> str:
    char_id = CHAR_PR.get(para_type, "7")
    para_id = PARA_PR.get(para_type, "0")
    prefix = "• " if para_type == "bullet" else ""
    content = escape_xml(f"{prefix}{text}")
    return (
        f'<hp:p id="{pid}" paraPrIDRef="{para_id}" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{char_id}">'
        f'<hp:t xml:space="preserve">{content}</hp:t>'
        f'</hp:run>'
        f'</hp:p>\n'
    )


def build_section0_xml(paragraphs: list) -> str:
    # secPr 추출 (페이지 설정 보존)
    secpr = ""
    try:
        with open(f"{TEMPLATE_DIR}/Contents/section0.xml", "r", encoding="utf-8") as f:
            orig = f.read()
        m = re.search(r'(<hp:secPr[^>]*>.*?</hp:secPr>)', orig, re.DOTALL)
        if m:
            secpr = m.group(1)
    except Exception:
        pass

    lines = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>', f'<hs:sec {NAMESPACES}>']

    # 첫 단락에 secPr 포함
    if paragraphs:
        ptype, text = paragraphs[0]
        char_id = CHAR_PR.get(ptype, "7")
        para_id = PARA_PR.get(ptype, "0")
        lines.append(
            f'<hp:p id="1000" paraPrIDRef="{para_id}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="{char_id}">{secpr}'
            f'<hp:t xml:space="preserve">{escape_xml(text)}</hp:t>'
            f'</hp:run></hp:p>'
        )
        rest = paragraphs[1:]
    else:
        rest = paragraphs

    for i, (ptype, text) in enumerate(rest, start=1001):
        lines.append(make_para_xml(ptype, text, i))

    lines.append('</hs:sec>')
    return "\n".join(lines)


def create_hwpx(paragraphs: list, output_path: str) -> str:
    """단락 리스트 → HWPX 파일 생성"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    section0 = build_section0_xml(paragraphs)
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(zipfile.ZipInfo("mimetype"), "application/hwp+zip")

        for fname in ["META-INF/container.xml", "META-INF/manifest.xml"]:
            try:
                with open(f"{TEMPLATE_DIR}/{fname}", "rb") as f:
                    zf.writestr(fname, f.read())
            except Exception:
                pass

        try:
            with open(f"{TEMPLATE_DIR}/Contents/header.xml", "rb") as f:
                zf.writestr("Contents/header.xml", f.read())
        except Exception:
            pass

        try:
            with open(f"{TEMPLATE_DIR}/Contents/content.hpf", "r", encoding="utf-8") as f:
                hpf = f.read()
            hpf = re.sub(r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z', now, hpf)
            zf.writestr("Contents/content.hpf", hpf.encode("utf-8"))
        except Exception:
            pass

        zf.writestr("Contents/section0.xml", section0.encode("utf-8"))

        for fname in ["settings.xml", "version.xml"]:
            try:
                with open(f"{TEMPLATE_DIR}/{fname}", "rb") as f:
                    zf.writestr(fname, f.read())
            except Exception:
                pass

    return output_path


# ── MCP 도구 정의 ──

@app.list_tools()
async def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="convert_to_hwpx",
            description="DOCX 또는 MD 파일을 HWPX 형식으로 변환합니다 (한글 정품 불필요)",
            inputSchema={
                "type": "object",
                "properties": {
                    "input_path": {"type": "string", "description": "변환할 파일 경로 (DOCX 또는 MD)"},
                    "output_path": {"type": "string", "description": "출력 HWPX 파일 경로 (선택, 기본: 같은 폴더)"},
                },
                "required": ["input_path"],
            },
        ),
        types.Tool(
            name="create_hwpx_from_text",
            description="마크다운 텍스트로 HWPX 파일을 직접 생성합니다",
            inputSchema={
                "type": "object",
                "properties": {
                    "markdown_text": {"type": "string", "description": "마크다운 형식 문서 내용"},
                    "output_path": {"type": "string", "description": "출력 HWPX 파일 경로"},
                    "filename": {"type": "string", "description": "파일명 (output_path 미지정 시 사용)"},
                },
                "required": ["markdown_text"],
            },
        ),
        types.Tool(
            name="send_hwpx_telegram",
            description="HWPX 파일을 Tommy 텔레그램으로 전송합니다",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "전송할 HWPX 파일 경로"},
                    "caption": {"type": "string", "description": "파일 설명 메시지"},
                },
                "required": ["file_path"],
            },
        ),
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:

    if name == "convert_to_hwpx":
        input_path = arguments["input_path"]
        output_path = arguments.get("output_path") or str(Path(input_path).with_suffix(".hwpx"))

        if not os.path.exists(input_path):
            return [types.TextContent(type="text", text=f"❌ 파일 없음: {input_path}")]

        ext = Path(input_path).suffix.lower()
        if ext == ".md":
            with open(input_path, "r", encoding="utf-8") as f:
                paragraphs = md_to_paragraphs(f.read())
        elif ext == ".docx":
            paragraphs = docx_to_paragraphs(input_path)
        else:
            return [types.TextContent(type="text", text=f"❌ 지원 형식: .md, .docx (입력: {ext})")]

        result = create_hwpx(paragraphs, output_path)
        size = os.path.getsize(result)
        return [types.TextContent(type="text", text=f"✅ 변환 완료: {result} ({size:,} bytes)\n단락 수: {len(paragraphs)}개")]

    elif name == "create_hwpx_from_text":
        md_text = arguments["markdown_text"]
        filename = arguments.get("filename", f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.hwpx")
        output_path = arguments.get("output_path") or f"{OUTPUT_DIR}/{filename}"
        if not output_path.endswith(".hwpx"):
            output_path += ".hwpx"

        paragraphs = md_to_paragraphs(md_text)
        result = create_hwpx(paragraphs, output_path)
        size = os.path.getsize(result)
        return [types.TextContent(type="text", text=f"✅ 생성 완료: {result} ({size:,} bytes)\n단락 수: {len(paragraphs)}개")]

    elif name == "send_hwpx_telegram":
        file_path = arguments["file_path"]
        caption = arguments.get("caption", "")

        if not os.path.exists(file_path):
            return [types.TextContent(type="text", text=f"❌ 파일 없음: {file_path}")]

        import requests
        TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
        CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

        if not TOKEN or not CHAT_ID:
            return [types.TextContent(type="text", text="❌ TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID 환경변수가 설정되지 않았습니다")]

        with open(file_path, "rb") as f:
            fname = Path(file_path).name
            res = requests.post(
                f"https://api.telegram.org/bot{TOKEN}/sendDocument",
                data={"chat_id": CHAT_ID, "caption": caption or fname},
                files={"document": (fname, f, "application/zip")},
            )

        if res.ok:
            return [types.TextContent(type="text", text=f"✅ 텔레그램 전송 완료: {fname}")]
        else:
            return [types.TextContent(type="text", text=f"❌ 전송 실패: {res.text}")]

    return [types.TextContent(type="text", text=f"❌ 알 수 없는 도구: {name}")]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
