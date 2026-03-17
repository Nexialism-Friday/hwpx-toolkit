#!/usr/bin/env python3
"""
개선된 HWP/HWPX 텍스트 추출기
- HWP: hwp5txt CLI + 레코드 파싱 (표 포함)
- HWPX: ZIP+XML 직접 파싱 (표 포함)
"""
import os
import re
import shutil as _shutil
import struct
import subprocess
import zipfile
import zlib
from pathlib import Path

HWP5TXT = os.environ.get("HWP5TXT_PATH") or _shutil.which("hwp5txt") or str(Path(__file__).parent / 'venv/bin/hwp5txt')
_HP_NS = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
_HP = f'{{{_HP_NS}}}'
_SECTION_RE = re.compile(r'^Contents/section\d+\.xml$')


def _log(msg):
    import datetime
    print(f'[{datetime.datetime.now():%H:%M:%S}] {msg}')


# ─── HWPX 개선 추출기 ────────────────────────────────────────

def _iter_section_xmls_from_zip(filepath):
    """ZIP에서 section XML 요소 리스트 반환"""
    try:
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(str(filepath), 'r') as zf:
            names = sorted(n for n in zf.namelist() if _SECTION_RE.match(n))
            return [ET.fromstring(zf.read(n)) for n in names]
    except Exception as e:
        _log(f'  HWPX ZIP 열기 실패: {e}')
        return []


def _get_all_t_text(root_elem):
    """모든 hp:t 텍스트를 재귀로 수집 (표 포함)"""
    import xml.etree.ElementTree as ET
    lines = []
    for p in root_elem.findall(f'{_HP}p'):
        para_parts = []
        for run in p.findall(f'{_HP}run'):
            # 직접 텍스트
            for t in run.findall(f'{_HP}t'):
                if t.text:
                    para_parts.append(t.text)
            # 표 내 텍스트
            for tbl in run.findall(f'.//{_HP}tbl'):
                for tr in tbl.findall(f'.//{_HP}tr'):
                    row_parts = []
                    for tc in tr.findall(f'.//{_HP}tc'):
                        cell_text = ''.join(t.text for t in tc.findall(f'.//{_HP}t') if t.text)
                        row_parts.append(cell_text.strip())
                    if any(row_parts):
                        lines.append('| ' + ' | '.join(row_parts) + ' |')
        if para_parts:
            lines.append(''.join(para_parts))
    return lines


def extract_hwpx_improved(filepath):
    """
    HWPX 텍스트 추출 (개선판 - 표 포함)
    방법1: 직접 ZIP+XML 파싱
    방법2: hwpx.export_markdown() (표 포함)
    방법3: hwpx.export_text() (폴백)
    """
    # 방법1: 직접 ZIP+XML
    try:
        sections = _iter_section_xmls_from_zip(filepath)
        if sections:
            all_lines = []
            for sec in sections:
                all_lines.extend(_get_all_t_text(sec))
            text = '\n'.join(all_lines).strip()
            if len(text) > 50:
                return text
    except Exception as e:
        _log(f'  HWPX 직접파싱 실패: {e}')

    # 방법2: export_markdown (표 포함)
    try:
        from hwpx import HwpxDocument
        doc = HwpxDocument.open(str(filepath))
        text = doc.export_markdown()
        doc.close()
        if text.strip():
            return text.strip()
    except Exception as e:
        _log(f'  HWPX export_markdown 실패: {e}')

    # 방법3: export_text (폴백)
    try:
        from hwpx import HwpxDocument
        doc = HwpxDocument.open(str(filepath))
        text = doc.export_text()
        doc.close()
        return text.strip()
    except Exception as e:
        _log(f'  HWPX export_text 실패: {e}')
        return ''


# ─── HWP 바이너리 개선 추출기 ─────────────────────────────────

def _parse_hwp_records(data):
    """HWP 레코드 스트림 파싱 → HWPTAG_PARA_TEXT(66) 텍스트 수집"""
    PARA_TEXT_TAG = 66
    texts = []
    pos = 0
    n = len(data)
    while pos + 4 <= n:
        header = struct.unpack_from('<I', data, pos)[0]
        pos += 4
        tag_id = header & 0x3FF
        # level = (header >> 10) & 0xF  # 사용 안 함
        size = (header >> 20) & 0xFFF
        if size == 0xFFF:
            if pos + 4 > n:
                break
            size = struct.unpack_from('<I', data, pos)[0]
            pos += 4
        if pos + size > n:
            break
        payload = data[pos:pos + size]
        pos += size
        if tag_id == PARA_TEXT_TAG:
            try:
                t = payload.decode('utf-16-le', errors='ignore')
                # 제어문자(space·newline 제외) 제거
                t = re.sub(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]', '', t)
                t = t.strip()
                if t:
                    texts.append(t)
            except Exception:
                pass
    return texts


def extract_hwp_improved(filepath):
    """
    HWP 바이너리 텍스트 추출 (개선판)
    방법1: hwp5txt CLI (venv)
    방법2: 직접 레코드 파싱 (표 포함 시도)
    방법3: PrvText (폴백)
    """
    fp = str(filepath)

    # 방법1: hwp5txt CLI
    try:
        result = subprocess.run(
            [HWP5TXT, fp],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0 and result.stdout.strip():
            raw = result.stdout
            # <표>, <그림> 등 플레이스홀더 제거
            raw = re.sub(r'<[^>]+>', '', raw)
            raw = re.sub(r'\s+', ' ', raw).strip()
            if len(raw) > 50:
                return raw
    except Exception as e:
        _log(f'  hwp5txt 실패: {e}')

    # 방법2: 직접 레코드 파싱
    try:
        import olefile
        f = olefile.OleFileIO(fp)
        hdr = f.openstream('FileHeader').read()
        is_compressed = (hdr[36] & 1) == 1

        all_texts = []
        for i in range(200):
            sname = f'BodyText/Section{i}'
            if not f.exists(sname):
                break
            raw_data = f.openstream(sname).read()
            if is_compressed:
                try:
                    raw_data = zlib.decompress(raw_data, -15)
                except Exception:
                    try:
                        raw_data = zlib.decompress(raw_data)
                    except Exception:
                        continue
            texts = _parse_hwp_records(raw_data)
            all_texts.extend(texts)
        f.close()

        if all_texts:
            return '\n'.join(all_texts)
    except Exception as e:
        _log(f'  HWP 레코드파싱 실패: {e}')

    # 방법3: PrvText (폴백)
    try:
        import olefile
        f = olefile.OleFileIO(fp)
        if f.exists('PrvText'):
            data = f.openstream('PrvText').read()
            text = data.decode('utf-16-le', errors='ignore').strip()
            f.close()
            return text
        f.close()
    except Exception as e:
        _log(f'  HWP PrvText 폴백도 실패: {e}')

    return ''


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('사용: python hwp_extractor.py <파일.hwp 또는 파일.hwpx>')
        sys.exit(1)
    fp = Path(sys.argv[1])
    if fp.suffix.lower() == '.hwpx':
        text = extract_hwpx_improved(fp)
    else:
        text = extract_hwp_improved(fp)
    print(f'=== 추출 결과 ({len(text)} 글자) ===')
    print(text[:3000])
