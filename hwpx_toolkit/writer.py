#!/usr/bin/env python3
"""
Friday HWP Writer v10 - WSL에서 Windows 한글을 제어하는 브릿지
기능:
  create 모드: 새 문서 생성
    - page_setup  : 용지 크기/방향/여백 설정 (최상단에 위치해야 함)
    - text        : 텍스트 삽입 (v9: 장평/자간/문단앞뒤간격/내어쓰기/위아래첨자/영문폰트 추가)
    - table       : 표 생성 (v9: cell_formats/cell_text_align 추가)
    - image       : 이미지 삽입 (path, width, height, align)
    - header      : 머리말 설정
    - footer      : 꼬리말 설정
    - footnote    : 각주 삽입 (anchor_text, note_text)
    - pagebreak   : 페이지 나누기
  edit 모드: 기존 문서 열어서 수정 (v9: text_format/para_format/insert_before/table_cell_format 추가)
    - replace         : 찾기/바꾸기
    - replace_regex   : 정규식 치환
    - insert_after    : 특정 텍스트 뒤에 삽입
    - insert_before   : 특정 텍스트 앞에 삽입 (v9 신규)
    - text_format     : 특정 텍스트 서식 변경 (v9 신규)
    - para_format     : 특정 문단 서식 변경 (v9 신규)
    - table_cell      : 표 셀 내용 교체
    - table_cell_color: 표 셀 배경색 변경
    - table_cell_format: 표 셀 텍스트 서식 변경 (v9 신규)
    - table_add_row   : 표에 행 추가
    - table_merge     : 표 셀 병합
    - table_row_height: 표 행 높이 변경
    - table_border    : 표 전체 테두리 스타일 변경
    - delete_line     : 특정 텍스트 포함 줄 삭제
    - append          : 문서 끝에 섹션 추가 (text/table/pagebreak 지원)
  저장:
    - output_format: "HWP"(기본) / "HWPX" / "PDF"

=== v9 신규 파라미터 (text 섹션) ===
  char_scale       : 장평 % (기본 100, 예: 90~110)
  letter_spacing   : 자간 % (기본 0, 범위 -50~50)
  space_before     : 문단 앞 간격 mm (기본 0)
  space_after      : 문단 뒤 간격 mm (기본 0)
  right_indent     : 오른쪽 들여쓰기 mm (기본 0)
  hanging          : 내어쓰기 mm (기본 0)
  line_spacing_type: 줄간격 방식 "percent"(기본)/"fixed"/"minimum"
  superscript      : 위첨자 bool
  subscript        : 아래첨자 bool
  font_latin       : 영문 폰트 별도 지정 (없으면 font와 동일)
  font_symbol      : 기호 폰트 별도 지정

=== v9 신규 파라미터 (table 섹션) ===
  cell_formats     : 셀별 서식 {"r,c": {"font":..,"size":..,"bold":..,"italic":..,"color":..,"align":"left/center/right"}}
  cell_text_align  : 셀별 가로 정렬 {"r,c": "left/center/right"} (cell_formats보다 간단한 대안)

=== v9 신규 edit 작업 ===
  insert_before    : 특정 텍스트 앞에 삽입
  text_format      : 특정 텍스트 서식만 변경 (찾아서 선택 후 CharShape 적용)
  para_format      : 특정 텍스트 포함 문단의 단락 서식 변경
  table_cell_format: 표 특정 셀 텍스트 서식 변경
"""
import subprocess
import sys
import json
import argparse
import os

WINDOWS_PYTHON = os.environ.get(
    "HWPX_WINDOWS_PYTHON",
    "/mnt/c/Users/YourUsername/AppData/Local/Programs/Python/Python314/python.exe"
)


# ─────────────────────────────────────────────────────────────────────────────
# 헬퍼: 텍스트 섹션 스크립트 생성 (create + edit/append 공용)
# ─────────────────────────────────────────────────────────────────────────────
def _text_section_lines(section):
    lines = []
    text         = section.get("text", "")
    font         = section.get("font", "맑은 고딕")
    font_latin   = section.get("font_latin", None)    # 영문 폰트 (없으면 font와 동일)
    font_symbol  = section.get("font_symbol", None)   # 기호 폰트
    size         = section.get("size", 10)
    bold         = section.get("bold", False)
    italic       = section.get("italic", False)
    underline    = section.get("underline", False)
    strikeout    = section.get("strikethrough", False)
    align        = section.get("align", "left")
    newline      = section.get("newline", True)
    color        = section.get("color", None)           # "RRGGBB"
    highlight    = section.get("highlight", None)       # "RRGGBB" 글자 배경색
    # 줄간격 (v9: type 분리)
    line_spacing      = section.get("line_spacing", None)       # 줄간격 값 (% 또는 mm)
    line_spacing_type = section.get("line_spacing_type", "percent")  # percent/fixed/minimum
    # 문단 간격 (v9 신규)
    space_before = section.get("space_before", None)    # 문단 앞 간격 mm
    space_after  = section.get("space_after", None)     # 문단 뒤 간격 mm
    # 들여쓰기/내어쓰기 (v9 강화)
    indent       = section.get("indent", None)          # 첫 줄 들여쓰기 mm
    right_indent = section.get("right_indent", None)    # 오른쪽 들여쓰기 mm
    hanging      = section.get("hanging", None)         # 내어쓰기 mm
    # 글자 간격/장평 (v9 신규)
    char_scale      = section.get("char_scale", None)     # 장평 % (기본 100)
    letter_spacing  = section.get("letter_spacing", None) # 자간 % (기본 0)
    # 위/아래 첨자 (v9 신규)
    superscript  = section.get("superscript", False)
    subscript    = section.get("subscript", False)
    # 기타
    style_name   = section.get("style_name", None)      # 단락 스타일 이름
    url          = section.get("url", None)             # 하이퍼링크 URL

    align_map = {"left": "Left", "center": "Center", "right": "Right", "justify": "Justify"}
    align_val = align_map.get(align, "Left")

    # ── CharShape (글자 모양)
    lines.append("hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
    lines.append(f"hwp.HParameterSet.HCharShape.FaceNameHangul = '{font}'")
    # 영문 폰트 (별도 지정 없으면 한글 폰트와 동일)
    latin = font_latin if font_latin else font
    lines.append(f"hwp.HParameterSet.HCharShape.FaceNameLatin  = '{latin}'")
    # 기호 폰트
    if font_symbol:
        lines.append(f"hwp.HParameterSet.HCharShape.FaceNameSymbol = '{font_symbol}'")
    lines.append(f"hwp.HParameterSet.HCharShape.Height = {int(size * 100)}")
    lines.append(f"hwp.HParameterSet.HCharShape.Bold   = {'1' if bold else '0'}")
    lines.append(f"hwp.HParameterSet.HCharShape.Italic = {'1' if italic else '0'}")
    lines.append(f"hwp.HParameterSet.HCharShape.UnderlineType = {'1' if underline else '0'}")
    lines.append(f"hwp.HParameterSet.HCharShape.StrikeOutType = {'1' if strikeout else '0'}")
    # 글자 색상
    if color:
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        lines.append(f"hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor({r},{g},{b})")
    # 글자 배경(형광펜)
    if highlight:
        r, g, b = int(highlight[0:2], 16), int(highlight[2:4], 16), int(highlight[4:6], 16)
        lines.append(f"hwp.HParameterSet.HCharShape.ShadeColor = hwp.RGBColor({r},{g},{b})")
        lines.append(f"hwp.HParameterSet.HCharShape.Shade = 100")
    # 장평 (v9)
    if char_scale is not None:
        lines.append(f"hwp.HParameterSet.HCharShape.CharScale = {int(char_scale)}")
    # 자간 (v9)
    if letter_spacing is not None:
        lines.append(f"hwp.HParameterSet.HCharShape.Spacing = {int(letter_spacing)}")
    # 위/아래 첨자 (v9)
    # SuperScript: 0=보통, 1=위첨자, 2=아래첨자
    if superscript:
        lines.append(f"hwp.HParameterSet.HCharShape.SuperScript = 1")
    elif subscript:
        lines.append(f"hwp.HParameterSet.HCharShape.SuperScript = 2")
    lines.append("hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")

    # ── ParaShape (문단 모양)
    lines.append("hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
    lines.append(f"hwp.HParameterSet.HParaShape.Alignment = hwp.HParameterSet.HParaShape.Alignment.enum{align_val}")
    # 줄간격 (v9: type 분리)
    if line_spacing is not None:
        lsmap = {"percent": 0, "fixed": 1, "minimum": 2}
        ls_type = lsmap.get(line_spacing_type, 0)
        lines.append(f"hwp.HParameterSet.HParaShape.LineSpacingType = {ls_type}")
        if line_spacing_type == "fixed":
            lines.append(f"hwp.HParameterSet.HParaShape.LineSpacing = {int(line_spacing * 283.465)}")
        else:
            lines.append(f"hwp.HParameterSet.HParaShape.LineSpacing = {int(line_spacing)}")
    # 첫 줄 들여쓰기
    if indent is not None:
        lines.append(f"hwp.HParameterSet.HParaShape.Indent = {int(indent * 283.465)}")
    # 내어쓰기 (v9)
    if hanging is not None:
        lines.append(f"hwp.HParameterSet.HParaShape.Outdent = {int(hanging * 283.465)}")
    # 오른쪽 들여쓰기 (v9)
    if right_indent is not None:
        lines.append(f"hwp.HParameterSet.HParaShape.RightIndent = {int(right_indent * 283.465)}")
    # 문단 앞 간격 (v9)
    if space_before is not None:
        lines.append(f"hwp.HParameterSet.HParaShape.SpaceBefore = {int(space_before * 283.465)}")
    # 문단 뒤 간격 (v9)
    if space_after is not None:
        lines.append(f"hwp.HParameterSet.HParaShape.SpaceAfter = {int(space_after * 283.465)}")
    lines.append("hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")

    # 단락 스타일 적용
    if style_name:
        lines.append(f"hwp.HAction.GetDefault('ApplyStyle', hwp.HParameterSet.HStyleApply.HSet)")
        lines.append(f"hwp.HParameterSet.HStyleApply.StyleName = '{style_name}'")
        lines.append(f"hwp.HAction.Execute('ApplyStyle', hwp.HParameterSet.HStyleApply.HSet)")

    # 하이퍼링크 또는 일반 텍스트 삽입
    if url:
        escaped_url  = url.replace("'", "\\'")
        escaped_text = text.replace("'", "\\'")
        lines.append(f"hwp.HAction.GetDefault('HyperLink', hwp.HParameterSet.HHyperLink.HSet)")
        lines.append(f"hwp.HParameterSet.HHyperLink.HyperLinkUrl  = '{escaped_url}'")
        lines.append(f"hwp.HParameterSet.HHyperLink.HyperLinkText = '{escaped_text}'")
        lines.append(f"hwp.HAction.Execute('HyperLink', hwp.HParameterSet.HHyperLink.HSet)")
    else:
        escaped = text.replace("'", "\\'").replace("\n", "\\n")
        lines.append("hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
        lines.append(f"hwp.HParameterSet.HInsertText.Text = '{escaped}'")
        lines.append("hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")

    if newline:
        lines.append("hwp.HAction.Run('BreakPara')")
    lines.append("")
    return lines


# ─────────────────────────────────────────────────────────────────────────────
# 헬퍼: 표 이동 후 셀 위치로 커서 이동
# ─────────────────────────────────────────────────────────────────────────────
def _table_goto_cell_lines(tbl_idx, row, col):
    lines = []
    lines.append(f"_tbl_count = 0")
    lines.append(f"_ctrl = hwp.HeadCtrl")
    lines.append(f"_target_tbl = None")
    lines.append(f"while _ctrl:")
    lines.append(f"    if _ctrl.UserDesc == '표':")
    lines.append(f"        if _tbl_count == {tbl_idx}:")
    lines.append(f"            _target_tbl = _ctrl")
    lines.append(f"            break")
    lines.append(f"        _tbl_count += 1")
    lines.append(f"    _ctrl = _ctrl.Next")
    lines.append(f"if _target_tbl:")
    lines.append(f"    _cell = _target_tbl.CellAt({row}, {col})")
    lines.append(f"    hwp.SetPosBySet(_cell)")
    return lines


# ─────────────────────────────────────────────────────────────────────────────
# 헬퍼: 표 섹션 생성 (create + edit/append 공용)
# ─────────────────────────────────────────────────────────────────────────────
def _table_section_lines(section, table_var="_last_tbl"):
    """
    표 생성 + 데이터 채우기 + 셀 배경색 + 셀 병합 + 행 높이 + 셀 수직 정렬 + 테두리
    v9 신규: cell_formats (셀별 서식), cell_text_align (셀별 가로 정렬)
    """
    lines = []
    rows        = section.get("rows", 2)
    cols        = section.get("cols", 2)
    data        = section.get("data", [])
    header_row  = section.get("header_row", False)
    font        = section.get("font", "맑은 고딕")
    size        = section.get("size", 9)
    col_widths  = section.get("col_widths", [])
    row_heights = section.get("row_heights", [])
    cell_colors = section.get("cell_colors", {})
    cell_valign = section.get("cell_valign", {})
    merge_cells = section.get("merge_cells", [])
    border      = section.get("border", None)
    # v9 신규
    cell_formats    = section.get("cell_formats", {})    # {"r,c": {"font":..,"size":..,"bold":..,"italic":..,"color":..,"align":..}}
    cell_text_align = section.get("cell_text_align", {}) # {"r,c": "left/center/right"}
    # v10 신규: 복합 표 구조
    cells_def    = section.get("cells", None)        # 셀 정의 배열 [{r,c,rowspan,colspan,text,...}]
    table_width  = section.get("table_width", None)  # 표 전체 너비 mm
    table_align  = section.get("table_align", None)  # 표 정렬 "left"/"center"/"right"
    cell_padding = section.get("cell_padding", None) # 전체 셀 내부 여백 {"all":mm}
    # cells_def로 rows/cols 자동 계산
    if cells_def:
        if not section.get("rows"):
            rows = max((c.get("r", 0) + c.get("rowspan", 1)) for c in cells_def)
        if not section.get("cols"):
            cols = max((c.get("c", 0) + c.get("colspan", 1)) for c in cells_def)

    lines.append("hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet)")
    lines.append(f"hwp.HParameterSet.HTableCreation.Rows = {rows}")
    lines.append(f"hwp.HParameterSet.HTableCreation.Cols = {cols}")
    lines.append(f"hwp.HParameterSet.HTableCreation.WidthType = 0")
    lines.append(f"hwp.HParameterSet.HTableCreation.HeightType = 0")
    lines.append(f"hwp.HParameterSet.HTableCreation.CreateItemArray('ColWidth', {cols})")

    if col_widths and len(col_widths) == cols:
        for ci, w in enumerate(col_widths):
            lines.append(f"hwp.HParameterSet.HTableCreation.ColWidth.SetItem({ci}, {int(w * 283.465)})")
    else:
        lines.append(f"_col_w = 42000 // {cols}")
        lines.append(f"for _ci in range({cols}):")
        lines.append(f"    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(_ci, _col_w)")

    lines.append("hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet)")
    lines.append("time.sleep(0.3)")
    lines.append("")

    tv = table_var
    lines.append(f"# 방금 생성한 마지막 표 참조")
    lines.append(f"{tv} = None")
    lines.append(f"_chk = hwp.HeadCtrl")
    lines.append(f"while _chk:")
    lines.append(f"    if _chk.UserDesc == '표': {tv} = _chk")
    lines.append(f"    _chk = _chk.Next")

    # ── 데이터 채우기 (v10: cells 배열 우선, 없으면 data[][])
    if cells_def:
        # === v10: cells 배열 기반 복합 표 처리 (rowspan/colspan 완전 지원) ===
        lines.append(f"# v10 cells 기반 복합 표 처리")
        lines.append(f"_cells_def_data = {json.dumps(cells_def, ensure_ascii=False)}")
        lines.append(f"_cells_dfont = '{font}'")
        lines.append(f"_cells_dsize = {size}")
        lines.append(f"_cells_bsmap = {{'solid':1,'dashed':3,'dotted':2,'double':6,'none':0}}")
        lines.append(f"_cells_amap  = {{'left':'Left','center':'Center','right':'Right','justify':'Justify'}}")
        lines.append(f"_cells_vamap = {{'top':0,'center':1,'bottom':2}}")
        lines.append(f"if {tv}:")
        lines.append(f"    # Phase 1: 데이터 채우기 (병합 전)")
        lines.append(f"    for _cdef in _cells_def_data:")
        lines.append(f"        _cr  = _cdef.get('r', 0)")
        lines.append(f"        _cc  = _cdef.get('c', 0)")
        lines.append(f"        _ctext = _cdef.get('text', '')")
        lines.append(f"        if _ctext is None: _ctext = ''")
        lines.append(f"        _cfont  = _cdef.get('font', _cells_dfont)")
        lines.append(f"        _csize  = _cdef.get('size', _cells_dsize)")
        lines.append(f"        _cbold  = 1 if _cdef.get('bold',  False) else 0")
        lines.append(f"        _cital  = 1 if _cdef.get('italic',False) else 0")
        lines.append(f"        _cund   = 1 if _cdef.get('underline',False) else 0")
        lines.append(f"        _ccolor = _cdef.get('color', None)")
        lines.append(f"        _calign = _cdef.get('align', None)")
        lines.append(f"        _cval   = _cdef.get('valign', None)")
        lines.append(f"        _cbg    = _cdef.get('bg_color', None)")
        lines.append(f"        _cbords = _cdef.get('borders', None)")
        lines.append(f"        _cpad   = _cdef.get('padding', None)")
        lines.append(f"        _ccs    = _cdef.get('char_scale', None)")
        lines.append(f"        _cls    = _cdef.get('letter_spacing', None)")
        lines.append(f"        _cspb   = _cdef.get('space_before', None)")
        lines.append(f"        _cspa   = _cdef.get('space_after', None)")
        lines.append(f"        _clsp   = _cdef.get('line_spacing', None)")
        lines.append(f"        _clspt  = _cdef.get('line_spacing_type', 'percent')")
        lines.append(f"        _cind   = _cdef.get('indent', None)")
        lines.append(f"        _chang  = _cdef.get('hanging', None)")
        lines.append(f"        try:")
        lines.append(f"            _ccell = {tv}.CellAt(_cr, _cc)")
        lines.append(f"            if not _ccell: continue")
        lines.append(f"            hwp.SetPosBySet(_ccell)")
        lines.append(f"        except Exception as _ce: continue")
        lines.append(f"        hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
        lines.append(f"        hwp.HParameterSet.HCharShape.FaceNameHangul = _cfont")
        lines.append(f"        hwp.HParameterSet.HCharShape.FaceNameLatin  = _cfont")
        lines.append(f"        hwp.HParameterSet.HCharShape.Height = int(_csize * 100)")
        lines.append(f"        hwp.HParameterSet.HCharShape.Bold   = _cbold")
        lines.append(f"        hwp.HParameterSet.HCharShape.Italic = _cital")
        lines.append(f"        hwp.HParameterSet.HCharShape.UnderlineType = _cund")
        lines.append(f"        if _ccolor:")
        lines.append(f"            _cr2,_cg2,_cb2 = int(_ccolor[0:2],16),int(_ccolor[2:4],16),int(_ccolor[4:6],16)")
        lines.append(f"            hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor(_cr2,_cg2,_cb2)")
        lines.append(f"        if _ccs is not None: hwp.HParameterSet.HCharShape.CharScale = int(_ccs)")
        lines.append(f"        if _cls is not None: hwp.HParameterSet.HCharShape.Spacing  = int(_cls)")
        lines.append(f"        hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")
        lines.append(f"        hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
        lines.append(f"        if _calign:")
        lines.append(f"            _cav = _cells_amap.get(_calign, 'Left')")
        lines.append(f"            hwp.HParameterSet.HParaShape.Alignment = getattr(hwp.HParameterSet.HParaShape.Alignment, 'enum' + _cav)")
        lines.append(f"        if _clsp is not None:")
        lines.append(f"            _clst = {{'percent':0,'fixed':1,'minimum':2}}.get(_clspt, 0)")
        lines.append(f"            hwp.HParameterSet.HParaShape.LineSpacingType = _clst")
        lines.append(f"            hwp.HParameterSet.HParaShape.LineSpacing = int(_clsp*283.465) if _clspt=='fixed' else int(_clsp)")
        lines.append(f"        if _cspb is not None: hwp.HParameterSet.HParaShape.SpaceBefore = int(_cspb*283.465)")
        lines.append(f"        if _cspa is not None: hwp.HParameterSet.HParaShape.SpaceAfter  = int(_cspa*283.465)")
        lines.append(f"        if _cind  is not None: hwp.HParameterSet.HParaShape.Indent  = int(_cind *283.465)")
        lines.append(f"        if _chang is not None: hwp.HParameterSet.HParaShape.Outdent = int(_chang*283.465)")
        lines.append(f"        hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
        lines.append(f"        if isinstance(_ctext, list):")
        lines.append(f"            for _pi, _pt in enumerate(_ctext):")
        lines.append(f"                hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
        lines.append(f"                hwp.HParameterSet.HInsertText.Text = str(_pt)")
        lines.append(f"                hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
        lines.append(f"                if _pi < len(_ctext)-1: hwp.HAction.Run('BreakPara')")
        lines.append(f"        else:")
        lines.append(f"            hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
        lines.append(f"            hwp.HParameterSet.HInsertText.Text = str(_ctext)")
        lines.append(f"            hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
        lines.append(f"        _need_cbf = _cbg or _cval or _cpad or _cbords")
        lines.append(f"        if _need_cbf:")
        lines.append(f"            hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"            _cbfs = hwp.HParameterSet.HCellBorderFill")
        lines.append(f"            if _cbg:")
        lines.append(f"                _bgr,_bgg,_bgb = int(_cbg[0:2],16),int(_cbg[2:4],16),int(_cbg[4:6],16)")
        lines.append(f"                _cbfs.FillColorRGB = hwp.RGBColor(_bgr,_bgg,_bgb); _cbfs.FillType = 1")
        lines.append(f"            if _cval: _cbfs.VertAlign = _cells_vamap.get(_cval, 1)")
        lines.append(f"            if _cpad:")
        lines.append(f"                if 'top'    in _cpad: _cbfs.PaddingTop    = int(_cpad['top']*283.465)")
        lines.append(f"                if 'bottom' in _cpad: _cbfs.PaddingBottom = int(_cpad['bottom']*283.465)")
        lines.append(f"                if 'left'   in _cpad: _cbfs.PaddingLeft   = int(_cpad['left']*283.465)")
        lines.append(f"                if 'right'  in _cpad: _cbfs.PaddingRight  = int(_cpad['right']*283.465)")
        lines.append(f"            if _cbords:")
        lines.append(f"                for _bsd, _bdt in _cbords.items():")
        lines.append(f"                    _btp = _cells_bsmap.get(_bdt.get('style','solid'), 1)")
        lines.append(f"                    _bsc = _bsd.capitalize()")
        lines.append(f"                    if _btp == 0: setattr(_cbfs, f'Border{{_bsc}}Type', 0)")
        lines.append(f"                    else:")
        lines.append(f"                        _bw  = int(_bdt.get('width',0.5)*283.465)")
        lines.append(f"                        _bcl = _bdt.get('color','000000')")
        lines.append(f"                        _bcr,_bcg,_bcb = int(_bcl[0:2],16),int(_bcl[2:4],16),int(_bcl[4:6],16)")
        lines.append(f"                        setattr(_cbfs, f'Border{{_bsc}}Type',  _btp)")
        lines.append(f"                        setattr(_cbfs, f'Border{{_bsc}}Width', _bw)")
        lines.append(f"                        setattr(_cbfs, f'Border{{_bsc}}Color', hwp.RGBColor(_bcr,_bcg,_bcb))")
        lines.append(f"            hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"    # Phase 2: 병합 처리 (상단→하단, 좌→우 순서)")
        lines.append(f"    _merges_todo = sorted(")
        lines.append(f"        [(_d['r'],_d['c'],_d['r']+_d.get('rowspan',1)-1,_d['c']+_d.get('colspan',1)-1)")
        lines.append(f"         for _d in _cells_def_data if _d.get('rowspan',1)>1 or _d.get('colspan',1)>1],")
        lines.append(f"        key=lambda x:(x[0],x[1]))")
        lines.append(f"    for _mr1,_mc1,_mr2,_mc2 in _merges_todo:")
        lines.append(f"        try:")
        lines.append(f"            _mca = {tv}.CellAt(_mr1, _mc1)")
        lines.append(f"            _mcb = {tv}.CellAt(_mr2, _mc2)")
        lines.append(f"            if not _mca or not _mcb: continue")
        lines.append(f"            hwp.SetPosBySet(_mca)")
        lines.append(f"            hwp.HAction.Run('TableCellBlock')")
        lines.append(f"            hwp.MovePosBySet(_mcb)")
        lines.append(f"            hwp.HAction.Run('TableMergeCell')")
        lines.append(f"            time.sleep(0.15)")
        lines.append(f"        except Exception as _me: print(f'merge err: {{_me}}')")
        lines.append(f"")
    elif data:
        lines.append(f"if {tv}:")
        lines.append(f"    _tbl_data = {json.dumps(data, ensure_ascii=False)}")
        lines.append(f"    _cell_fmts = {json.dumps(cell_formats, ensure_ascii=False)}")
        lines.append(f"    _align_map_tbl = {{'left':'Left','center':'Center','right':'Right','justify':'Justify'}}")
        lines.append(f"    for _r, _row_data in enumerate(_tbl_data):")
        lines.append(f"        for _c, _cell_text in enumerate(_row_data):")
        lines.append(f"            _cell = {tv}.CellAt(_r, _c)")
        lines.append(f"            hwp.SetPosBySet(_cell)")
        # 셀별 서식 우선, 없으면 기본값
        lines.append(f"            _fmt = _cell_fmts.get(f'{{_r}},{{_c}}', {{}})")
        lines.append(f"            _cfont = _fmt.get('font', '{font}')")
        lines.append(f"            _csize = _fmt.get('size', {size})")
        lines.append(f"            _cbold = _fmt.get('bold', {1 if header_row else 0})")
        lines.append(f"            if _r == 0 and {1 if header_row else 0}: _cbold = 1")
        lines.append(f"            _cbold = _fmt.get('bold', _cbold)")
        lines.append(f"            _citalic = _fmt.get('italic', 0)")
        lines.append(f"            _ccolor = _fmt.get('color', None)")
        lines.append(f"            hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
        lines.append(f"            hwp.HParameterSet.HCharShape.FaceNameHangul = _cfont")
        lines.append(f"            hwp.HParameterSet.HCharShape.FaceNameLatin  = _cfont")
        lines.append(f"            hwp.HParameterSet.HCharShape.Height = int(_csize * 100)")
        lines.append(f"            hwp.HParameterSet.HCharShape.Bold   = 1 if _cbold else 0")
        lines.append(f"            hwp.HParameterSet.HCharShape.Italic = 1 if _citalic else 0")
        lines.append(f"            if _ccolor:")
        lines.append(f"                _cr2,_cg2,_cb2 = int(_ccolor[0:2],16),int(_ccolor[2:4],16),int(_ccolor[4:6],16)")
        lines.append(f"                hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor(_cr2,_cg2,_cb2)")
        lines.append(f"            hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")
        # 셀 내 텍스트 정렬 (cell_formats 또는 cell_text_align)
        lines.append(f"            _calign_str = _fmt.get('align', None)")
        lines.append(f"            if _calign_str:")
        lines.append(f"                _calign_v = _align_map_tbl.get(_calign_str, 'Left')")
        lines.append(f"                hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
        lines.append(f"                hwp.HParameterSet.HParaShape.Alignment = getattr(hwp.HParameterSet.HParaShape.Alignment, 'enum' + _calign_v)")
        lines.append(f"                hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
        lines.append(f"            hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
        lines.append(f"            hwp.HParameterSet.HInsertText.Text = str(_cell_text)")
        lines.append(f"            hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
        lines.append(f"")

    # ── 셀별 텍스트 정렬 (cell_text_align, data와 별도) (v9)
    if cell_text_align:
        lines.append(f"# 셀 텍스트 가로 정렬 (cell_text_align)")
        lines.append(f"_cta_data = {json.dumps(cell_text_align, ensure_ascii=False)}")
        lines.append(f"_cta_amap = {{'left':'Left','center':'Center','right':'Right','justify':'Justify'}}")
        lines.append(f"if {tv}:")
        lines.append(f"    for _cta_key, _cta_val in _cta_data.items():")
        lines.append(f"        _cta_rc = _cta_key.split(',')")
        lines.append(f"        _cta_r, _cta_c = int(_cta_rc[0]), int(_cta_rc[1])")
        lines.append(f"        _cta_cell = {tv}.CellAt(_cta_r, _cta_c)")
        lines.append(f"        hwp.SetPosBySet(_cta_cell)")
        lines.append(f"        _cta_align = _cta_amap.get(_cta_val, 'Left')")
        lines.append(f"        hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
        lines.append(f"        hwp.HParameterSet.HParaShape.Alignment = getattr(hwp.HParameterSet.HParaShape.Alignment, 'enum' + _cta_align)")
        lines.append(f"        hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
        lines.append(f"")

    # ── 행 높이 지정
    if row_heights:
        lines.append(f"# 행 높이 설정")
        lines.append(f"if {tv}:")
        for ri, h in enumerate(row_heights):
            lines.append(f"    _rh_cell = {tv}.CellAt({ri}, 0)")
            lines.append(f"    hwp.SetPosBySet(_rh_cell)")
            lines.append(f"    hwp.HAction.GetDefault('TableCellProperty', hwp.HParameterSet.HShapeObject.HSet)")
            lines.append(f"    hwp.HParameterSet.HShapeObject.Height = {int(h * 283.465)}")
            lines.append(f"    hwp.HParameterSet.HShapeObject.HeightRelTo = 0")
            lines.append(f"    hwp.HAction.Execute('TableCellProperty', hwp.HParameterSet.HShapeObject.HSet)")
        lines.append(f"")

    # ── 셀 수직 정렬
    if cell_valign:
        lines.append(f"# 셀 수직 정렬")
        lines.append(f"_cv_data = {json.dumps(cell_valign, ensure_ascii=False)}")
        lines.append(f"if {tv}:")
        lines.append(f"    for _key, _va in _cv_data.items():")
        lines.append(f"        _rc = _key.split(',')")
        lines.append(f"        _vr, _vc = int(_rc[0]), int(_rc[1])")
        lines.append(f"        _va_map = {{'top':0,'center':1,'bottom':2}}")
        lines.append(f"        _va_val = _va_map.get(_va, 1)")
        lines.append(f"        _vacell = {tv}.CellAt(_vr, _vc)")
        lines.append(f"        hwp.SetPosBySet(_vacell)")
        lines.append(f"        hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"        hwp.HParameterSet.HCellBorderFill.VertAlign = _va_val")
        lines.append(f"        hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"")

    # ── 셀 배경색
    if cell_colors:
        lines.append(f"# 셀 배경색 설정")
        lines.append(f"_cell_colors = {json.dumps(cell_colors, ensure_ascii=False)}")
        lines.append(f"if {tv}:")
        lines.append(f"    for _key, _hex in _cell_colors.items():")
        lines.append(f"        _rc = _key.split(',')")
        lines.append(f"        _r2, _c2 = int(_rc[0]), int(_rc[1])")
        lines.append(f"        _cr, _cg, _cb = int(_hex[0:2],16), int(_hex[2:4],16), int(_hex[4:6],16)")
        lines.append(f"        _ccell = {tv}.CellAt(_r2, _c2)")
        lines.append(f"        hwp.SetPosBySet(_ccell)")
        lines.append(f"        hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"        hwp.HParameterSet.HCellBorderFill.FillColorRGB = hwp.RGBColor(_cr, _cg, _cb)")
        lines.append(f"        hwp.HParameterSet.HCellBorderFill.FillType = 1")
        lines.append(f"        hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"")

    # ── 셀 병합 (cells_def가 없을 때만 - cells_def는 내부에서 처리)
    if merge_cells and not cells_def:
        lines.append(f"# 셀 병합")
        for mg in merge_cells:
            r1, c1 = mg["from"][0], mg["from"][1]
            r2, c2 = mg["to"][0], mg["to"][1]
            lines.append(f"if {tv}:")
            lines.append(f"    _mc1 = {tv}.CellAt({r1}, {c1})")
            lines.append(f"    hwp.SetPosBySet(_mc1)")
            lines.append(f"    hwp.HAction.Run('TableCellBlock')")
            lines.append(f"    _mc2 = {tv}.CellAt({r2}, {c2})")
            lines.append(f"    hwp.MovePosBySet(_mc2)")
            lines.append(f"    hwp.HAction.Run('TableMergeCell')")
            lines.append(f"    time.sleep(0.1)")
        lines.append("")

    # ── 표 전체 테두리
    if border:
        b_style = border.get("style", "solid")
        b_width = border.get("width", 0.5)
        b_color = border.get("color", "000000")
        style_map = {"solid": 1, "dashed": 3, "dotted": 2, "double": 6}
        b_type = style_map.get(b_style, 1)
        br, bg, bb = int(b_color[0:2], 16), int(b_color[2:4], 16), int(b_color[4:6], 16)
        b_hwp_width = int(b_width * 283.465)
        lines.append(f"# 표 전체 테두리 설정")
        lines.append(f"if {tv}:")
        lines.append(f"    _br_rows = {tv}.Rows")
        lines.append(f"    _br_cols = {tv}.Cols")
        lines.append(f"    for _br in range(_br_rows):")
        lines.append(f"        for _bc in range(_br_cols):")
        lines.append(f"            _brcell = {tv}.CellAt(_br, _bc)")
        lines.append(f"            hwp.SetPosBySet(_brcell)")
        lines.append(f"            hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"            _bset = hwp.HParameterSet.HCellBorderFill")
        for side in ["Left", "Right", "Top", "Bottom"]:
            lines.append(f"            _bset.Border{side}Type  = {b_type}")
            lines.append(f"            _bset.Border{side}Width = {b_hwp_width}")
            lines.append(f"            _bset.Border{side}Color = hwp.RGBColor({br},{bg},{bb})")
        lines.append(f"            hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
        lines.append(f"")

    # ── v10: 전체 셀 내부 여백 (cell_padding)
    if cell_padding:
        all_pd = cell_padding.get("all", None)
        if all_pd is not None:
            lines.append(f"# 전체 셀 내부 여백 설정")
            lines.append(f"if {tv}:")
            lines.append(f"    _gpd_rows = {tv}.Rows")
            lines.append(f"    _gpd_cols = {tv}.Cols")
            lines.append(f"    _gpd_val  = {int(all_pd * 283.465)}")
            lines.append(f"    for _gpd_r in range(_gpd_rows):")
            lines.append(f"        for _gpd_c in range(_gpd_cols):")
            lines.append(f"            try:")
            lines.append(f"                _gpd_cell = {tv}.CellAt(_gpd_r, _gpd_c)")
            lines.append(f"                if not _gpd_cell: continue")
            lines.append(f"                hwp.SetPosBySet(_gpd_cell)")
            lines.append(f"                hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
            lines.append(f"                hwp.HParameterSet.HCellBorderFill.PaddingTop    = _gpd_val")
            lines.append(f"                hwp.HParameterSet.HCellBorderFill.PaddingBottom = _gpd_val")
            lines.append(f"                hwp.HParameterSet.HCellBorderFill.PaddingLeft   = _gpd_val")
            lines.append(f"                hwp.HParameterSet.HCellBorderFill.PaddingRight  = _gpd_val")
            lines.append(f"                hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
            lines.append(f"            except: pass")
            lines.append(f"")

    # ── v10: 표 전체 너비 / 정렬
    if table_width or table_align:
        align_map_t = {"left": 0, "center": 1, "right": 2}
        lines.append(f"# 표 너비/정렬 설정")
        lines.append(f"if {tv}:")
        lines.append(f"    try:")
        lines.append(f"        hwp.SetPosBySet({tv})")
        lines.append(f"        hwp.HAction.GetDefault('TableCellProperty', hwp.HParameterSet.HShapeObject.HSet)")
        if table_width:
            lines.append(f"        hwp.HParameterSet.HShapeObject.Width = {int(table_width * 283.465)}")
            lines.append(f"        hwp.HParameterSet.HShapeObject.WidthRelTo = 0")
        if table_align:
            av_t = align_map_t.get(table_align, 1)
            lines.append(f"        hwp.HParameterSet.HShapeObject.HorzAlign = {av_t}")
        lines.append(f"        hwp.HAction.Execute('TableCellProperty', hwp.HParameterSet.HShapeObject.HSet)")
        lines.append(f"    except: pass")
        lines.append(f"")

    return lines


def generate_hwp_script(doc):
    lines = []
    lines.append("import win32com.client")
    lines.append("import threading")
    lines.append("import ctypes")
    lines.append("import time")
    lines.append("import os")
    lines.append("")

    # 보안 경고 자동 허용 스레드
    lines.append("def auto_dismiss():")
    lines.append("    user32 = ctypes.windll.user32")
    lines.append("    INPUT_KEYBOARD = 1")
    lines.append("    KEYEVENTF_KEYUP = 2")
    lines.append("    class KEYBDINPUT(ctypes.Structure):")
    lines.append("        _fields_ = [('wVk',ctypes.c_ushort),('wScan',ctypes.c_ushort),('dwFlags',ctypes.c_ulong),('time',ctypes.c_ulong),('dwExtraInfo',ctypes.POINTER(ctypes.c_ulong))]")
    lines.append("    class INPUT(ctypes.Structure):")
    lines.append("        _fields_ = [('type',ctypes.c_ulong),('ki',KEYBDINPUT),('padding',ctypes.c_ubyte*8)]")
    lines.append("    VK_MENU = 0x12")
    lines.append("    VK_Y    = 0x59")
    lines.append("    def send_key(vk, up=False):")
    lines.append("        inp = INPUT()")
    lines.append("        inp.type = INPUT_KEYBOARD")
    lines.append("        inp.ki.wVk = vk")
    lines.append("        inp.ki.dwFlags = KEYEVENTF_KEYUP if up else 0")
    lines.append("        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))")
    lines.append("    for _ in range(6):")
    lines.append("        time.sleep(0.8)")
    lines.append("        send_key(VK_MENU); time.sleep(0.05)")
    lines.append("        send_key(VK_Y);    time.sleep(0.05)")
    lines.append("        send_key(VK_Y, up=True); time.sleep(0.05)")
    lines.append("        send_key(VK_MENU, up=True)")
    lines.append("")
    lines.append("t = threading.Thread(target=auto_dismiss, daemon=True)")
    lines.append("t.start()")
    lines.append("")
    lines.append("hwp = win32com.client.Dispatch('HWPFrame.HwpObject')")
    lines.append("hwp.XHwpWindows.Item(0).Visible = True")
    lines.append("")

    mode = doc.get("mode", "create")

    # ──────────────────────────────────────────────
    # CREATE 모드: 새 문서 생성
    # ──────────────────────────────────────────────
    if mode == "create":
        lines.append("hwp.HAction.Run('FileNew')")
        lines.append("time.sleep(0.5)")
        lines.append("")

        sections = doc.get("sections", [])
        tbl_counter = [0]

        for section in sections:
            stype = section.get("type", "text")

            # ── 페이지 설정
            if stype == "page_setup":
                paper         = section.get("paper", "A4")
                orient        = section.get("orient", "portrait")
                margin_top    = section.get("margin_top", 30)
                margin_bottom = section.get("margin_bottom", 30)
                margin_left   = section.get("margin_left", 25)
                margin_right  = section.get("margin_right", 25)
                margin_header = section.get("margin_header", 15)
                margin_footer = section.get("margin_footer", 15)

                paper_map  = {"A4": 5, "B5": 4, "A5": 6, "Letter": 0, "Legal": 1, "A3": 8}
                paper_code = paper_map.get(paper.upper(), 5)
                orient_code = 1 if orient.lower() == "landscape" else 0

                lines.append(f"# 페이지 설정")
                lines.append(f"hwp.HAction.GetDefault('PageSetup', hwp.HParameterSet.HSecDef.HSet)")
                lines.append(f"hwp.HParameterSet.HSecDef.PaperProp.PaperCode = {paper_code}")
                lines.append(f"hwp.HParameterSet.HSecDef.PaperProp.Landscape = {orient_code}")
                lines.append(f"hwp.HParameterSet.HSecDef.MarginTop    = {int(margin_top    * 283.465)}")
                lines.append(f"hwp.HParameterSet.HSecDef.MarginBottom = {int(margin_bottom * 283.465)}")
                lines.append(f"hwp.HParameterSet.HSecDef.MarginLeft   = {int(margin_left   * 283.465)}")
                lines.append(f"hwp.HParameterSet.HSecDef.MarginRight  = {int(margin_right  * 283.465)}")
                lines.append(f"hwp.HParameterSet.HSecDef.MarginHeader = {int(margin_header * 283.465)}")
                lines.append(f"hwp.HParameterSet.HSecDef.MarginFooter = {int(margin_footer * 283.465)}")
                lines.append(f"hwp.HAction.Execute('PageSetup', hwp.HParameterSet.HSecDef.HSet)")
                lines.append("")

            # ── 텍스트
            elif stype == "text":
                lines.extend(_text_section_lines(section))

            # ── 표
            elif stype == "table":
                tv = f"_tbl_{tbl_counter[0]}"
                tbl_counter[0] += 1
                lines.extend(_table_section_lines(section, table_var=tv))
                lines.append("hwp.HAction.Run('MoveDocEnd')")
                lines.append("hwp.HAction.Run('BreakPara')")
                lines.append("")

            # ── 이미지 삽입
            elif stype == "image":
                img_path  = section.get("path", "")
                width_mm  = section.get("width", 0)
                height_mm = section.get("height", 0)
                align     = section.get("align", "center")

                if not img_path:
                    lines.append("# 이미지 경로 누락 - 건너뜀")
                else:
                    if img_path.startswith("/mnt/c/"):
                        win_img = "C:\\" + img_path[7:].replace("/", "\\")
                    elif img_path.startswith("/mnt/"):
                        drive = img_path[5].upper()
                        win_img = drive + ":\\" + img_path[7:].replace("/", "\\")
                    else:
                        win_img = img_path

                    align_map = {"left": "Left", "center": "Center", "right": "Right"}
                    align_val = align_map.get(align, "Center")

                    lines.append("hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                    lines.append(f"hwp.HParameterSet.HParaShape.Alignment = hwp.HParameterSet.HParaShape.Alignment.enum{align_val}")
                    lines.append("hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                    lines.append(f"hwp.HAction.GetDefault('InsertPicture', hwp.HParameterSet.HInsertPicture.HSet)")
                    lines.append(f"hwp.HParameterSet.HInsertPicture.FileName = r'{win_img}'")
                    lines.append(f"hwp.HParameterSet.HInsertPicture.Embedded = 1")
                    lines.append(f"hwp.HParameterSet.HInsertPicture.sizetype = 0")
                    if width_mm and height_mm:
                        lines.append(f"hwp.HParameterSet.HInsertPicture.Width  = {int(width_mm  * 283.465)}")
                        lines.append(f"hwp.HParameterSet.HInsertPicture.Height = {int(height_mm * 283.465)}")
                    lines.append(f"hwp.HAction.Execute('InsertPicture', hwp.HParameterSet.HInsertPicture.HSet)")
                    lines.append("hwp.HAction.Run('BreakPara')")
                    lines.append("")

            # ── 머리말
            elif stype == "header":
                text      = section.get("text", "").replace("'", "\\'")
                font      = section.get("font", "맑은 고딕")
                size      = section.get("size", 9)
                align     = section.get("align", "right")
                align_map = {"left": "Left", "center": "Center", "right": "Right"}
                align_val = align_map.get(align, "Right")

                lines.append(f"# 머리말 설정")
                lines.append(f"hwp.HAction.Run('HeaderFooterEdit')")
                lines.append(f"time.sleep(0.3)")
                lines.append(f"hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                lines.append(f"hwp.HParameterSet.HParaShape.Alignment = hwp.HParameterSet.HParaShape.Alignment.enum{align_val}")
                lines.append(f"hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                lines.append(f"hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                lines.append(f"hwp.HParameterSet.HCharShape.FaceNameHangul = '{font}'")
                lines.append(f"hwp.HParameterSet.HCharShape.Height = {int(size * 100)}")
                lines.append(f"hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                lines.append(f"hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"hwp.HParameterSet.HInsertText.Text = '{text}'")
                lines.append(f"hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"hwp.HAction.Run('CloseEx')")
                lines.append("")

            # ── 꼬리말
            elif stype == "footer":
                text      = section.get("text", "").replace("'", "\\'")
                font      = section.get("font", "맑은 고딕")
                size      = section.get("size", 9)
                align     = section.get("align", "center")
                align_map = {"left": "Left", "center": "Center", "right": "Right"}
                align_val = align_map.get(align, "Center")
                page_num  = section.get("page_number", False)

                lines.append(f"# 꼬리말 설정")
                lines.append(f"hwp.HAction.Run('FooterEdit')")
                lines.append(f"time.sleep(0.3)")
                lines.append(f"hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                lines.append(f"hwp.HParameterSet.HParaShape.Alignment = hwp.HParameterSet.HParaShape.Alignment.enum{align_val}")
                lines.append(f"hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                lines.append(f"hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                lines.append(f"hwp.HParameterSet.HCharShape.FaceNameHangul = '{font}'")
                lines.append(f"hwp.HParameterSet.HCharShape.Height = {int(size * 100)}")
                lines.append(f"hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                if text:
                    lines.append(f"hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                    lines.append(f"hwp.HParameterSet.HInsertText.Text = '{text}'")
                    lines.append(f"hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                if page_num:
                    lines.append(f"hwp.HAction.Run('InsertPageNumber')")
                lines.append(f"hwp.HAction.Run('CloseEx')")
                lines.append("")

            # ── 각주
            elif stype == "footnote":
                anchor_text = section.get("anchor_text", "").replace("'", "\\'")
                note_text   = section.get("note_text", "").replace("'", "\\'")
                lines.append(f"# 각주 삽입: '{anchor_text}'")
                if anchor_text:
                    lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                    lines.append(f"hwp.HParameterSet.HFindReplace.FindString = '{anchor_text}'")
                    lines.append(f"hwp.HParameterSet.HFindReplace.IgnoreCase = 1")
                    lines.append(f"_fn_found = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                    lines.append(f"if _fn_found: hwp.HAction.Run('MoveRight')")
                lines.append(f"hwp.HAction.Run('InsertFootnote')")
                lines.append(f"time.sleep(0.2)")
                lines.append(f"hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"hwp.HParameterSet.HInsertText.Text = '{note_text}'")
                lines.append(f"hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"hwp.HAction.Run('CloseEx')")
                lines.append("")

            elif stype == "pagebreak":
                lines.append("hwp.HAction.Run('BreakPage')")
                lines.append("")

    # ──────────────────────────────────────────────
    # EDIT 모드: 기존 문서 열어서 수정
    # ──────────────────────────────────────────────
    elif mode == "edit":
        input_path = doc.get("input", "")
        if not input_path:
            lines.append("raise ValueError('edit 모드: input 경로가 없습니다')")
        else:
            win_path = input_path.replace("/", "\\\\")
            lines.append(f"hwp.Open('{win_path}')")
            lines.append(f"time.sleep(0.8)")
            lines.append("")

        operations = doc.get("operations", [])
        tbl_counter = [0]

        for op in operations:
            op_type = op.get("type", "")

            # ── 텍스트 찾기/바꾸기
            if op_type == "replace":
                find    = op.get("find", "").replace("'", "\\'")
                replace = op.get("replace", "").replace("'", "\\'")
                all_    = op.get("all", True)
                lines.append(f"# 찾기/바꾸기: '{find}' → '{replace}'")
                lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"hwp.HParameterSet.HFindReplace.FindString   = '{find}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.ReplaceString = '{replace}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.IgnoreCase   = 1")
                lines.append(f"hwp.HParameterSet.HFindReplace.WholeWordOnly = 0")
                lines.append(f"hwp.HParameterSet.HFindReplace.RegExp = 0")
                if all_:
                    lines.append(f"hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet)")
                else:
                    lines.append(f"hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append("")

            # ── 정규식 찾기/바꾸기
            elif op_type == "replace_regex":
                find    = op.get("find", "").replace("'", "\\'")
                replace = op.get("replace", "").replace("'", "\\'")
                lines.append(f"# 정규식 치환: '{find}' → '{replace}'")
                lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"hwp.HParameterSet.HFindReplace.FindString   = '{find}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.ReplaceString = '{replace}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.RegExp = 1")
                lines.append(f"hwp.HAction.Execute('AllReplace', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append("")

            # ── 특정 텍스트 뒤에 삽입
            elif op_type == "insert_after":
                find    = op.get("find", "").replace("'", "\\'")
                text    = op.get("text", "").replace("'", "\\'")
                newline = op.get("newline", True)
                lines.append(f"# insert_after: '{find}' 뒤에 삽입")
                lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"hwp.HParameterSet.HFindReplace.FindString = '{find}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.IgnoreCase = 1")
                lines.append(f"_found = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"if _found:")
                lines.append(f"    hwp.HAction.Run('MoveRight')")
                lines.append(f"    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"    hwp.HParameterSet.HInsertText.Text = '{text}'")
                lines.append(f"    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                if newline:
                    lines.append(f"    hwp.HAction.Run('BreakPara')")
                lines.append("")

            # ── 특정 텍스트 앞에 삽입 (v9 신규)
            elif op_type == "insert_before":
                find    = op.get("find", "").replace("'", "\\'")
                text    = op.get("text", "").replace("'", "\\'")
                newline = op.get("newline", True)
                lines.append(f"# insert_before: '{find}' 앞에 삽입")
                lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"hwp.HParameterSet.HFindReplace.FindString = '{find}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.IgnoreCase = 1")
                lines.append(f"_found_b = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"if _found_b:")
                lines.append(f"    hwp.HAction.Run('MoveLeft')")
                lines.append(f"    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"    hwp.HParameterSet.HInsertText.Text = '{text}'")
                lines.append(f"    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                if newline:
                    lines.append(f"    hwp.HAction.Run('BreakPara')")
                lines.append("")

            # ── 특정 텍스트 서식 변경 (v9 신규)
            elif op_type == "text_format":
                find         = op.get("find", "").replace("'", "\\'")
                font         = op.get("font", None)
                size         = op.get("size", None)
                bold         = op.get("bold", None)
                italic       = op.get("italic", None)
                underline    = op.get("underline", None)
                color        = op.get("color", None)
                char_scale   = op.get("char_scale", None)
                letter_spacing = op.get("letter_spacing", None)
                superscript  = op.get("superscript", None)
                subscript    = op.get("subscript", None)
                find_all     = op.get("all", False)   # 전체 찾아 서식 변경 여부

                lines.append(f"# text_format: '{find}' 서식 변경")
                if find_all:
                    # 문서 시작으로 이동 후 전체 검색
                    lines.append(f"hwp.HAction.Run('MoveDocBegin')")
                    lines.append(f"_tf_continue = True")
                    lines.append(f"while _tf_continue:")
                    lines.append(f"    hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                    lines.append(f"    hwp.HParameterSet.HFindReplace.FindString = '{find}'")
                    lines.append(f"    hwp.HParameterSet.HFindReplace.IgnoreCase = 1")
                    lines.append(f"    _tf_found = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                    lines.append(f"    if not _tf_found:")
                    lines.append(f"        _tf_continue = False")
                    lines.append(f"        break")
                    lines.append(f"    hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                    if font:
                        lines.append(f"    hwp.HParameterSet.HCharShape.FaceNameHangul = '{font}'")
                        lines.append(f"    hwp.HParameterSet.HCharShape.FaceNameLatin  = '{font}'")
                    if size is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Height = {int(size * 100)}")
                    if bold is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Bold = {'1' if bold else '0'}")
                    if italic is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Italic = {'1' if italic else '0'}")
                    if underline is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.UnderlineType = {'1' if underline else '0'}")
                    if color:
                        cr, cg, cb = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
                        lines.append(f"    hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor({cr},{cg},{cb})")
                    if char_scale is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.CharScale = {int(char_scale)}")
                    if letter_spacing is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Spacing = {int(letter_spacing)}")
                    if superscript:
                        lines.append(f"    hwp.HParameterSet.HCharShape.SuperScript = 1")
                    elif subscript:
                        lines.append(f"    hwp.HParameterSet.HCharShape.SuperScript = 2")
                    lines.append(f"    hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                else:
                    lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                    lines.append(f"hwp.HParameterSet.HFindReplace.FindString = '{find}'")
                    lines.append(f"hwp.HParameterSet.HFindReplace.IgnoreCase = 1")
                    lines.append(f"_tf_found = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                    lines.append(f"if _tf_found:")
                    lines.append(f"    hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                    if font:
                        lines.append(f"    hwp.HParameterSet.HCharShape.FaceNameHangul = '{font}'")
                        lines.append(f"    hwp.HParameterSet.HCharShape.FaceNameLatin  = '{font}'")
                    if size is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Height = {int(size * 100)}")
                    if bold is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Bold = {'1' if bold else '0'}")
                    if italic is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Italic = {'1' if italic else '0'}")
                    if underline is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.UnderlineType = {'1' if underline else '0'}")
                    if color:
                        cr, cg, cb = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
                        lines.append(f"    hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor({cr},{cg},{cb})")
                    if char_scale is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.CharScale = {int(char_scale)}")
                    if letter_spacing is not None:
                        lines.append(f"    hwp.HParameterSet.HCharShape.Spacing = {int(letter_spacing)}")
                    if superscript:
                        lines.append(f"    hwp.HParameterSet.HCharShape.SuperScript = 1")
                    elif subscript:
                        lines.append(f"    hwp.HParameterSet.HCharShape.SuperScript = 2")
                    lines.append(f"    hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                lines.append("")

            # ── 특정 문단 서식 변경 (v9 신규)
            elif op_type == "para_format":
                find         = op.get("find", "").replace("'", "\\'")
                align        = op.get("align", None)
                line_spacing = op.get("line_spacing", None)
                line_spacing_type = op.get("line_spacing_type", "percent")
                space_before = op.get("space_before", None)
                space_after  = op.get("space_after", None)
                indent       = op.get("indent", None)
                right_indent = op.get("right_indent", None)
                hanging      = op.get("hanging", None)

                lines.append(f"# para_format: '{find}' 포함 문단 서식 변경")
                lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"hwp.HParameterSet.HFindReplace.FindString = '{find}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.IgnoreCase = 1")
                lines.append(f"_pf_found = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"if _pf_found:")
                lines.append(f"    hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                if align:
                    align_map2 = {"left": "Left", "center": "Center", "right": "Right", "justify": "Justify"}
                    av = align_map2.get(align, "Left")
                    lines.append(f"    hwp.HParameterSet.HParaShape.Alignment = hwp.HParameterSet.HParaShape.Alignment.enum{av}")
                if line_spacing is not None:
                    lsmap = {"percent": 0, "fixed": 1, "minimum": 2}
                    ls_type = lsmap.get(line_spacing_type, 0)
                    lines.append(f"    hwp.HParameterSet.HParaShape.LineSpacingType = {ls_type}")
                    if line_spacing_type == "fixed":
                        lines.append(f"    hwp.HParameterSet.HParaShape.LineSpacing = {int(line_spacing * 283.465)}")
                    else:
                        lines.append(f"    hwp.HParameterSet.HParaShape.LineSpacing = {int(line_spacing)}")
                if space_before is not None:
                    lines.append(f"    hwp.HParameterSet.HParaShape.SpaceBefore = {int(space_before * 283.465)}")
                if space_after is not None:
                    lines.append(f"    hwp.HParameterSet.HParaShape.SpaceAfter = {int(space_after * 283.465)}")
                if indent is not None:
                    lines.append(f"    hwp.HParameterSet.HParaShape.Indent = {int(indent * 283.465)}")
                if right_indent is not None:
                    lines.append(f"    hwp.HParameterSet.HParaShape.RightIndent = {int(right_indent * 283.465)}")
                if hanging is not None:
                    lines.append(f"    hwp.HParameterSet.HParaShape.Outdent = {int(hanging * 283.465)}")
                lines.append(f"    hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                lines.append("")

            # ── 표 특정 셀 내용 교체
            elif op_type == "table_cell":
                tbl_idx = op.get("table_index", 0)
                row     = op.get("row", 0)
                col     = op.get("col", 0)
                text    = op.get("text", "").replace("'", "\\'")
                append  = op.get("append", False)
                lines.append(f"# table_cell: 표[{tbl_idx}] ({row},{col}) 셀 수정")
                lines.extend(_table_goto_cell_lines(tbl_idx, row, col))
                lines.append(f"if _target_tbl:")
                if not append:
                    lines.append(f"    hwp.HAction.Run('SelectAll')")
                    lines.append(f"    hwp.HAction.Run('Delete')")
                lines.append(f"    hwp.HAction.GetDefault('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"    hwp.HParameterSet.HInsertText.Text = '{text}'")
                lines.append(f"    hwp.HAction.Execute('InsertText', hwp.HParameterSet.HInsertText.HSet)")
                lines.append(f"else:")
                lines.append(f"    print('경고: 표 인덱스 {tbl_idx}를 찾지 못했습니다')")
                lines.append("")

            # ── 표 셀 텍스트 서식 변경 (v9 신규)
            elif op_type == "table_cell_format":
                tbl_idx      = op.get("table_index", 0)
                row          = op.get("row", 0)
                col          = op.get("col", 0)
                font         = op.get("font", None)
                size         = op.get("size", None)
                bold         = op.get("bold", None)
                italic       = op.get("italic", None)
                underline    = op.get("underline", None)
                color        = op.get("color", None)
                char_scale   = op.get("char_scale", None)
                letter_spacing = op.get("letter_spacing", None)
                align        = op.get("align", None)
                space_before = op.get("space_before", None)
                space_after  = op.get("space_after", None)

                lines.append(f"# table_cell_format: 표[{tbl_idx}] ({row},{col}) 셀 서식 변경")
                lines.extend(_table_goto_cell_lines(tbl_idx, row, col))
                lines.append(f"if _target_tbl:")
                lines.append(f"    _tcf_cell = _target_tbl.CellAt({row}, {col})")
                lines.append(f"    hwp.SetPosBySet(_tcf_cell)")
                lines.append(f"    hwp.HAction.Run('SelectAll')")
                lines.append(f"    hwp.HAction.GetDefault('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                if font:
                    lines.append(f"    hwp.HParameterSet.HCharShape.FaceNameHangul = '{font}'")
                    lines.append(f"    hwp.HParameterSet.HCharShape.FaceNameLatin  = '{font}'")
                if size is not None:
                    lines.append(f"    hwp.HParameterSet.HCharShape.Height = {int(size * 100)}")
                if bold is not None:
                    lines.append(f"    hwp.HParameterSet.HCharShape.Bold = {'1' if bold else '0'}")
                if italic is not None:
                    lines.append(f"    hwp.HParameterSet.HCharShape.Italic = {'1' if italic else '0'}")
                if underline is not None:
                    lines.append(f"    hwp.HParameterSet.HCharShape.UnderlineType = {'1' if underline else '0'}")
                if color:
                    cr, cg, cb = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
                    lines.append(f"    hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor({cr},{cg},{cb})")
                if char_scale is not None:
                    lines.append(f"    hwp.HParameterSet.HCharShape.CharScale = {int(char_scale)}")
                if letter_spacing is not None:
                    lines.append(f"    hwp.HParameterSet.HCharShape.Spacing = {int(letter_spacing)}")
                lines.append(f"    hwp.HAction.Execute('CharShape', hwp.HParameterSet.HCharShape.HSet)")
                if align:
                    align_map3 = {"left": "Left", "center": "Center", "right": "Right", "justify": "Justify"}
                    av3 = align_map3.get(align, "Left")
                    lines.append(f"    hwp.HAction.GetDefault('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                    lines.append(f"    hwp.HParameterSet.HParaShape.Alignment = hwp.HParameterSet.HParaShape.Alignment.enum{av3}")
                    if space_before is not None:
                        lines.append(f"    hwp.HParameterSet.HParaShape.SpaceBefore = {int(space_before * 283.465)}")
                    if space_after is not None:
                        lines.append(f"    hwp.HParameterSet.HParaShape.SpaceAfter = {int(space_after * 283.465)}")
                    lines.append(f"    hwp.HAction.Execute('ParagraphShape', hwp.HParameterSet.HParaShape.HSet)")
                lines.append("")

            # ── 표 셀 배경색 변경
            elif op_type == "table_cell_color":
                tbl_idx = op.get("table_index", 0)
                row     = op.get("row", 0)
                col     = op.get("col", 0)
                color   = op.get("color", "FFFFFF")
                r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
                lines.append(f"# table_cell_color: 표[{tbl_idx}] ({row},{col}) 배경색 #{color}")
                lines.extend(_table_goto_cell_lines(tbl_idx, row, col))
                lines.append(f"if _target_tbl:")
                lines.append(f"    hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
                lines.append(f"    hwp.HParameterSet.HCellBorderFill.FillColorRGB = hwp.RGBColor({r},{g},{b})")
                lines.append(f"    hwp.HParameterSet.HCellBorderFill.FillType = 1")
                lines.append(f"    hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
                lines.append("")

            # ── 표 행 추가
            elif op_type == "table_add_row":
                tbl_idx  = op.get("table_index", 0)
                position = op.get("position", "end")
                count    = op.get("count", 1)
                lines.append(f"# table_add_row: 표[{tbl_idx}] 행 {count}개 추가 (위치={position})")
                lines.extend(_table_goto_cell_lines(tbl_idx, 0, 0))
                lines.append(f"if _target_tbl:")
                if position == "end":
                    lines.append(f"    _row_cnt = _target_tbl.Rows")
                    lines.append(f"    _col_cnt = _target_tbl.Cols")
                    lines.append(f"    _last_cell = _target_tbl.CellAt(_row_cnt-1, _col_cnt-1)")
                    lines.append(f"    hwp.SetPosBySet(_last_cell)")
                    lines.append(f"    for _ in range({count}):")
                    lines.append(f"        hwp.HAction.Run('TableAppendRow')")
                    lines.append(f"        time.sleep(0.05)")
                else:
                    lines.append(f"    _ins_cell = _target_tbl.CellAt({position}, 0)")
                    lines.append(f"    hwp.SetPosBySet(_ins_cell)")
                    lines.append(f"    for _ in range({count}):")
                    lines.append(f"        hwp.HAction.Run('TableInsertRow')")
                    lines.append(f"        time.sleep(0.05)")
                lines.append("")

            # ── 표 셀 병합
            elif op_type == "table_merge":
                tbl_idx = op.get("table_index", 0)
                r1, c1  = op.get("from", [0, 0])
                r2, c2  = op.get("to", [0, 1])
                lines.append(f"# table_merge: 표[{tbl_idx}] ({r1},{c1})~({r2},{c2}) 병합")
                lines.extend(_table_goto_cell_lines(tbl_idx, r1, c1))
                lines.append(f"if _target_tbl:")
                lines.append(f"    _mc1 = _target_tbl.CellAt({r1}, {c1})")
                lines.append(f"    hwp.SetPosBySet(_mc1)")
                lines.append(f"    hwp.HAction.Run('TableCellBlock')")
                lines.append(f"    _mc2 = _target_tbl.CellAt({r2}, {c2})")
                lines.append(f"    hwp.MovePosBySet(_mc2)")
                lines.append(f"    hwp.HAction.Run('TableMergeCell')")
                lines.append(f"    time.sleep(0.1)")
                lines.append("")

            # ── 표 행 높이 변경
            elif op_type == "table_row_height":
                tbl_idx = op.get("table_index", 0)
                row     = op.get("row", 0)
                height  = op.get("height", 10)
                lines.append(f"# table_row_height: 표[{tbl_idx}] 행{row} 높이={height}mm")
                lines.extend(_table_goto_cell_lines(tbl_idx, row, 0))
                lines.append(f"if _target_tbl:")
                lines.append(f"    _rh_cell = _target_tbl.CellAt({row}, 0)")
                lines.append(f"    hwp.SetPosBySet(_rh_cell)")
                lines.append(f"    hwp.HAction.GetDefault('TableCellProperty', hwp.HParameterSet.HShapeObject.HSet)")
                lines.append(f"    hwp.HParameterSet.HShapeObject.Height = {int(height * 283.465)}")
                lines.append(f"    hwp.HParameterSet.HShapeObject.HeightRelTo = 0")
                lines.append(f"    hwp.HAction.Execute('TableCellProperty', hwp.HParameterSet.HShapeObject.HSet)")
                lines.append("")

            # ── 표 테두리 일괄 변경
            elif op_type == "table_border":
                tbl_idx = op.get("table_index", 0)
                b_style = op.get("style", "solid")
                b_width = op.get("width", 0.5)
                b_color = op.get("color", "000000")
                style_map = {"solid": 1, "dashed": 3, "dotted": 2, "double": 6}
                b_type = style_map.get(b_style, 1)
                br, bg, bb = int(b_color[0:2], 16), int(b_color[2:4], 16), int(b_color[4:6], 16)
                b_hwp_width = int(b_width * 283.465)
                lines.append(f"# table_border: 표[{tbl_idx}] 전체 테두리 설정")
                lines.extend(_table_goto_cell_lines(tbl_idx, 0, 0))
                lines.append(f"if _target_tbl:")
                lines.append(f"    _tbr_rows = _target_tbl.Rows")
                lines.append(f"    _tbr_cols = _target_tbl.Cols")
                lines.append(f"    for _tbr in range(_tbr_rows):")
                lines.append(f"        for _tbc in range(_tbr_cols):")
                lines.append(f"            _tbcell = _target_tbl.CellAt(_tbr, _tbc)")
                lines.append(f"            hwp.SetPosBySet(_tbcell)")
                lines.append(f"            hwp.HAction.GetDefault('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
                lines.append(f"            _bset = hwp.HParameterSet.HCellBorderFill")
                for side in ["Left", "Right", "Top", "Bottom"]:
                    lines.append(f"            _bset.Border{side}Type  = {b_type}")
                    lines.append(f"            _bset.Border{side}Width = {b_hwp_width}")
                    lines.append(f"            _bset.Border{side}Color = hwp.RGBColor({br},{bg},{bb})")
                lines.append(f"            hwp.HAction.Execute('CellBorderFill', hwp.HParameterSet.HCellBorderFill.HSet)")
                lines.append("")

            # ── 특정 텍스트 줄 삭제
            elif op_type == "delete_line":
                find = op.get("find", "").replace("'", "\\'")
                lines.append(f"# delete_line: '{find}' 포함 줄 삭제")
                lines.append(f"hwp.HAction.GetDefault('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"hwp.HParameterSet.HFindReplace.FindString = '{find}'")
                lines.append(f"hwp.HParameterSet.HFindReplace.IgnoreCase = 1")
                lines.append(f"_found = hwp.HAction.Execute('RepeatFind', hwp.HParameterSet.HFindReplace.HSet)")
                lines.append(f"if _found:")
                lines.append(f"    hwp.HAction.Run('SelectLine')")
                lines.append(f"    hwp.HAction.Run('Delete')")
                lines.append("")

            # ── 문서 끝에 섹션 추가
            elif op_type == "append":
                lines.append("hwp.HAction.Run('MoveDocEnd')")
                lines.append("hwp.HAction.Run('BreakPara')")
                for sub in op.get("sections", []):
                    stype2 = sub.get("type", "text")
                    if stype2 == "text":
                        lines.extend(_text_section_lines(sub))
                    elif stype2 == "table":
                        tv = f"_tbl_{tbl_counter[0]}"
                        tbl_counter[0] += 1
                        lines.extend(_table_section_lines(sub, table_var=tv))
                        lines.append("hwp.HAction.Run('MoveDocEnd')")
                        lines.append("hwp.HAction.Run('BreakPara')")
                    elif stype2 == "image":
                        img_path  = sub.get("path", "")
                        width_mm  = sub.get("width", 0)
                        height_mm = sub.get("height", 0)
                        if img_path:
                            if img_path.startswith("/mnt/c/"):
                                win_img = "C:\\" + img_path[7:].replace("/", "\\")
                            else:
                                win_img = img_path
                            lines.append(f"hwp.HAction.GetDefault('InsertPicture', hwp.HParameterSet.HInsertPicture.HSet)")
                            lines.append(f"hwp.HParameterSet.HInsertPicture.FileName = r'{win_img}'")
                            lines.append(f"hwp.HParameterSet.HInsertPicture.Embedded = 1")
                            if width_mm and height_mm:
                                lines.append(f"hwp.HParameterSet.HInsertPicture.Width  = {int(width_mm  * 283.465)}")
                                lines.append(f"hwp.HParameterSet.HInsertPicture.Height = {int(height_mm * 283.465)}")
                            lines.append(f"hwp.HAction.Execute('InsertPicture', hwp.HParameterSet.HInsertPicture.HSet)")
                            lines.append("hwp.HAction.Run('BreakPara')")
                    elif stype2 == "pagebreak":
                        lines.append("hwp.HAction.Run('BreakPage')")
                    lines.append("")
                lines.append("")

    # ──────────────────────────────────────────────
    # 저장 (output_format HWP/HWPX/PDF 지원)
    # ──────────────────────────────────────────────
    output_path   = doc.get("output", "")
    output_format = doc.get("output_format", "HWP").upper()

    if not output_path:
        lines.append("hwp.HAction.Run('FileSave')")
    else:
        win_out = output_path.replace("/", "\\\\")
        if output_format == "PDF":
            lines.append(f"# PDF로 저장")
            lines.append(f"hwp.HAction.GetDefault('FileSaveAs_S', hwp.HParameterSet.HFileSaveAs.HSet)")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.filename = '{win_out}'")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.Format   = 'PDF'")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.lthread  = 1")
            lines.append(f"hwp.HAction.Execute('FileSaveAs_S', hwp.HParameterSet.HFileSaveAs.HSet)")
        elif output_format == "HWPX":
            lines.append(f"# HWPX로 저장")
            lines.append(f"hwp.HAction.GetDefault('FileSaveAs_S', hwp.HParameterSet.HFileSaveAs.HSet)")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.filename = '{win_out}'")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.Format   = 'HWPX'")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.lthread  = 1")
            lines.append(f"hwp.HAction.Execute('FileSaveAs_S', hwp.HParameterSet.HFileSaveAs.HSet)")
        else:
            lines.append(f"# HWP로 저장")
            lines.append(f"hwp.HAction.GetDefault('FileSaveAs_S', hwp.HParameterSet.HFileSaveAs.HSet)")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.filename = '{win_out}'")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.Format   = 'HWP'")
            lines.append(f"hwp.HParameterSet.HFileSaveAs.lthread  = 1")
            lines.append(f"hwp.HAction.Execute('FileSaveAs_S', hwp.HParameterSet.HFileSaveAs.HSet)")

    lines.append("time.sleep(0.5)")
    lines.append("print('HWP 작업 완료')")

    return "\n".join(lines)


def run_on_windows(script: str):
    tmp = "/tmp/hwp_run_script.py"
    win_tmp = r"C:\Windows\Temp\hwp_run_script.py"

    with open(tmp, "w", encoding="utf-8") as f:
        f.write(script)

    result = subprocess.run(
        [WINDOWS_PYTHON, win_tmp],
        capture_output=True,
        text=True,
        timeout=120
    )

    if result.returncode != 0:
        print("=== 오류 ===")
        print(result.stderr)
        sys.exit(1)
    else:
        print(result.stdout)


def main():
    parser = argparse.ArgumentParser(description="Friday HWP Writer v9")
    parser.add_argument("--json",  help="JSON 파일 또는 JSON 문자열")
    parser.add_argument("--stdin", action="store_true", help="stdin에서 JSON 읽기")
    parser.add_argument("--dump",  action="store_true", help="생성 스크립트만 출력 (실행 안 함)")
    args = parser.parse_args()

    if args.stdin:
        doc = json.load(sys.stdin)
    elif args.json:
        if os.path.exists(args.json) or (args.json.startswith("/") or args.json.startswith("C:")):
            path = args.json
            if path.startswith("C:") or path.startswith("c:"):
                path = "/mnt/c" + path[2:].replace("\\", "/")
            with open(path, "r", encoding="utf-8") as f:
                doc = json.load(f)
        else:
            doc = json.loads(args.json)
    else:
        print("Friday HWP Writer v9")
        print("")
        print("사용법: hwp_writer.py --json <파일 또는 JSON> | --stdin [--dump]")
        print("")
        print("=== CREATE 섹션 타입 ===")
        print('  page_setup : {"type":"page_setup","paper":"A4","orient":"portrait",')
        print('                 "margin_top":30,"margin_bottom":30,"margin_left":25,"margin_right":25}')
        print('  text       : {"type":"text","text":"안녕","font":"맑은 고딕","size":12,')
        print('                 "bold":true,"italic":true,"underline":true,')
        print('                 "color":"FF0000","highlight":"FFFF00",')
        print('                 "char_scale":95,          ← 장평 % (v9)')
        print('                 "letter_spacing":-3,      ← 자간 % (v9)')
        print('                 "space_before":3,         ← 문단 앞 간격 mm (v9)')
        print('                 "space_after":3,          ← 문단 뒤 간격 mm (v9)')
        print('                 "indent":5,               ← 첫 줄 들여쓰기 mm')
        print('                 "hanging":5,              ← 내어쓰기 mm (v9)')
        print('                 "right_indent":5,         ← 오른쪽 들여쓰기 mm (v9)')
        print('                 "line_spacing":160,       ← 줄간격 값')
        print('                 "line_spacing_type":"percent", ← percent/fixed/minimum (v9)')
        print('                 "superscript":true,       ← 위첨자 (v9)')
        print('                 "subscript":true,         ← 아래첨자 (v9)')
        print('                 "font_latin":"Arial",     ← 영문 폰트 별도 지정 (v9)')
        print('                 "url":"https://example.com"}')
        print('  table      : {"type":"table","rows":3,"cols":2,')
        print('                 "col_widths":[40,80],"row_heights":[10,15,10],')
        print('                 "header_row":true,')
        print('                 "cell_colors":{"0,0":"C5D9F1"},')
        print('                 "cell_valign":{"0,0":"center"},')
        print('                 "cell_text_align":{"0,0":"center","1,0":"left"},  ← v9')
        print('                 "cell_formats":{"0,0":{"font":"나눔고딕","size":10,"bold":true,')
        print('                                  "color":"FF0000","align":"center"}},  ← v9')
        print('                 "merge_cells":[{"from":[2,0],"to":[2,1]}],')
        print('                 "border":{"style":"solid","width":0.5,"color":"000000"},')
        print('                 "data":[["헤더A","헤더B"],["값1","값2"],["병합",""]]}')
        print('  image      : {"type":"image","path":"/mnt/c/Users/YourUsername/img.png",')
        print('                 "width":80,"height":60,"align":"center"}')
        print('  header     : {"type":"header","text":"문서 제목","align":"right","size":9}')
        print('  footer     : {"type":"footer","text":"","page_number":true,"align":"center"}')
        print('  footnote   : {"type":"footnote","anchor_text":"참조어","note_text":"각주 내용"}')
        print('  pagebreak  : {"type":"pagebreak"}')
        print("")
        print("=== EDIT 작업 타입 ===")
        print('  replace         : {"type":"replace","find":"구","replace":"신","all":true}')
        print('  replace_regex   : {"type":"replace_regex","find":"\\\\d+","replace":"##"}')
        print('  insert_after    : {"type":"insert_after","find":"기준텍스트","text":"삽입내용"}')
        print('  insert_before   : {"type":"insert_before","find":"기준텍스트","text":"삽입내용"}  ← v9')
        print('  text_format     : {"type":"text_format","find":"대상텍스트",   ← v9')
        print('                      "font":"나눔고딕","size":12,"bold":true,')
        print('                      "color":"FF0000","char_scale":90,"letter_spacing":-5,')
        print('                      "all":false}')
        print('  para_format     : {"type":"para_format","find":"대상텍스트",   ← v9')
        print('                      "align":"center","line_spacing":160,')
        print('                      "line_spacing_type":"percent",')
        print('                      "space_before":3,"space_after":3,')
        print('                      "indent":10,"hanging":5,"right_indent":5}')
        print('  table_cell      : {"type":"table_cell","table_index":0,"row":0,"col":0,"text":"내용"}')
        print('  table_cell_color: {"type":"table_cell_color","table_index":0,"row":0,"col":0,"color":"C5D9F1"}')
        print('  table_cell_format: {"type":"table_cell_format","table_index":0,"row":0,"col":0,  ← v9')
        print('                       "font":"나눔고딕","size":10,"bold":true,')
        print('                       "char_scale":95,"letter_spacing":-3,')
        print('                       "align":"center","space_before":2,"space_after":2}')
        print('  table_add_row   : {"type":"table_add_row","table_index":0,"position":"end","count":1}')
        print('  table_merge     : {"type":"table_merge","table_index":0,"from":[0,0],"to":[0,2]}')
        print('  table_row_height: {"type":"table_row_height","table_index":0,"row":0,"height":15}')
        print('  table_border    : {"type":"table_border","table_index":0,')
        print('                      "style":"solid","width":0.5,"color":"000000"}')
        print('  delete_line     : {"type":"delete_line","find":"삭제할 텍스트"}')
        print('  append          : {"type":"append","sections":[<text|table|image|pagebreak>]}')
        print("")
        print("=== 저장 옵션 ===")
        print('  output_format: "HWP"(기본) | "HWPX" | "PDF"')
        sys.exit(0)

    script = generate_hwp_script(doc)

    if args.dump:
        print(script)
        return

    run_on_windows(script)


if __name__ == "__main__":
    main()
