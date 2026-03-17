#!/usr/bin/env python3
"""
HWPX 문서 생성 예제 (Generator)
- Markdown 단락 리스트 → HWPX 파일 직접 생성
- 한글(HWP) 설치 없이 동작

환경변수:
  HWPX_TEMPLATE_DIR : HWPX 템플릿 디렉토리 경로 (필수)
  HWPX_OUTPUT_DIR   : 출력 디렉토리 (기본: /tmp/hwpx_output)
"""

import zipfile
import shutil
import os
from datetime import datetime

# ──────────────────────────────────────────────
# 출력 경로
# ──────────────────────────────────────────────
OUTPUT_DIR = os.environ.get("HWPX_OUTPUT_DIR", "/tmp/hwpx_output")
OUTPUT_PATH = os.path.join(OUTPUT_DIR, f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.hwpx")
TEMPLATE_DIR = os.environ.get("HWPX_TEMPLATE_DIR", str(os.path.expanduser("~/hwpx_template")))

# ──────────────────────────────────────────────
# 문서 내용 (단락 리스트)
# format: ("type", "text")
#   type = "title1" | "title2" | "title3" | "body" | "bullet" | "blank"
# ──────────────────────────────────────────────
PARAGRAPHS = [
    ("title1", "DR콩고 쾅고 주 태양광 발전 및 역량강화 사업"),
    ("title2", "PMC 특수제안서"),
    ("body", "작성일: 2026년 3월 4일"),
    ("body", "작성기관: [PMC 수행기관명]"),
    ("blank", ""),

    # ── 사업 개요 ──
    ("title1", "I. 사업 개요"),
    ("body", "본 사업은 KOICA 예산과 UNDP의 현지 조달 역량을 결합한 국제기구 협력사업으로, "
             "기술(전기공학)과 정책(신재생 제도)이 결합된 통합적 개발 컨설팅 사업입니다. "
             "특히 USAID 탈퇴로 인해 PMC의 기술 자문 및 성과관리 역할이 대폭 강화된 것이 특징입니다."),
    ("blank", ""),
    ("title2", "사업 목표 체계"),
    ("bullet", "장기(Impact): 재생에너지 전환을 통한 기후변화 대응 및 DR콩고 NDC 달성 기여"),
    ("bullet", "중기(Outcome): 쾅고 주 농촌 지역의 에너지 접근성 개선 및 탄소 중립 생활환경 조성"),
    ("bullet", "단기(Output): 500kW급 태양광 인프라 구축, O&M 체계 정립, 신재생에너지 제도 개선"),
    ("blank", ""),
    ("title2", "주요 기술 사양"),
    ("bullet", "태양광 어레이: 500kWp"),
    ("bullet", "에너지 저장장치(ESS): 1,700kWh 배터리 시스템"),
    ("bullet", "예비 전원: 180kVA 디젤 하이브리드"),
    ("bullet", "미니그리드: 켕게 HGR 반경 1km, 0.4kV 단상 그리드"),
    ("bullet", "수요 근거: 일 소비 약 1,282kWh/day → 발전 약 1,826kWh/day (보수적 설계 70%)"),
    ("blank", ""),

    # ── 특수제안 1 ──
    ("title1", "II. 특수제안 1: 블렌디드 파이낸스 연계 민간투자 유치"),
    ("blank", ""),
    ("title2", "1. 배경 및 필요성"),
    ("body", "USAID 탈퇴로 발생한 기술 자문 공백을 PMC가 주도적으로 채우는 한편, "
             "KOICA ODA 단독 재원의 한계를 극복하기 위한 민간재원 연계 구조가 필요합니다. "
             "DR콩고 에너지 접근률은 약 19%로 사하라 이남 아프리카 최저 수준이며, "
             "쾅고 주의 경우 전력 인프라가 전무한 상태입니다."),
    ("blank", ""),
    ("title2", "2. 블렌디드 파이낸스 4레이어 구조"),
    ("bullet", "1레이어 (증여/ODA): KOICA 무상원조 → 기반 인프라(태양광·ESS·미니그리드) 구축 비용"),
    ("bullet", "2레이어 (양허성 차관): EDCF 또는 AKCF → O&M 초기 운영비 및 역량강화 프로그램"),
    ("bullet", "3레이어 (민간투자): 한국 에너지 기업 SPC 구성 → 수익형 운영권(15~20년) 취득"),
    ("bullet", "4레이어 (탄소크레딧): VCS/GS 방법론 기반 탄소배출권 수익 → 운영 지속가능성 확보"),
    ("blank", ""),
    ("title2", "3. PMC 역할 및 액션플랜"),
    ("bullet", "F/S 단계: 민간투자 타당성 평가 및 투자 구조 설계 지원"),
    ("bullet", "조달 단계: EPC 입찰 기술 검토 시 민간 참여 조건 사전 반영"),
    ("bullet", "운영 단계: SPC 운영 모니터링 및 KOICA 보고체계 연계"),
    ("bullet", "탄소크레딧 등록: 방법론 선정(AMS-I.D 또는 AMS-I.F) → MRV 체계 구축 지원"),
    ("blank", ""),
    ("title2", "4. 법적·제도적 위험 관리 [신규 - Perplexity 보완]"),
    ("body", "DR콩고의 불안정한 법·제도 환경을 고려한 위험 관리 체계가 필수적입니다."),
    ("bullet", "투자 보호: MIGA(세계은행 다자투자보증기구) 또는 K-SURE 정치적 위험 보험 적용"),
    ("bullet", "계약 안정성: OHADA(아프리카 사업법 조화기구) 기반 계약 구조 채택"),
    ("bullet", "토지 리스크: UNDP 현지 네트워크를 통한 부지 사전 법적 검토 및 지역사회 동의 확보"),
    ("bullet", "규제 리스크: 수자원전력부(MRHE) 정책 변화 모니터링 체계 구축 및 PSC 의제 정기 상정"),
    ("bullet", "분쟁 해결: ICC 국제중재 또는 ICSID 투자분쟁해결 조항 계약 포함"),
    ("blank", ""),
    ("title2", "5. 탄소크레딧 방법론 구체화 [신규 - Perplexity 보완]"),
    ("body", "단순 탄소크레딧 언급에서 구체적 방법론·절차로 고도화합니다."),
    ("bullet", "적용 방법론: VCS AMS-I.D (전력망 연결 재생에너지) 또는 AMS-I.F (독립형 재생에너지)"),
    ("bullet", "기준 배출계수: DR콩고 국가 전력망 배출계수 적용 (IEA 데이터 기준)"),
    ("bullet", "MRV 체계: 원격 모니터링 시스템(IoT 센서) → 실시간 발전량 기록 → 연간 검증(VVB)"),
    ("bullet", "등록 기관: Verra(VCS) 또는 Gold Standard Foundation에 프로젝트 등록"),
    ("bullet", "수익 배분: 탄소크레딧 수익 30% → O&M 기금, 30% → 주민위원회 운영, 40% → SPC 수익"),
    ("blank", ""),
    ("title2", "6. 실용 ODA 정책 부합성"),
    ("body", "본 제안은 2023년 ODA법 개정 및 국제개발협력 종합기본계획(2021-2025)의 "
             "민관협력(PPP) 확대 방침에 부합하며, KOICA의 임팩트 투자 확대 전략과 직접 연계됩니다."),
    ("blank", ""),

    # ── 특수제안 2 ──
    ("title1", "III. 특수제안 2: 한국 기업 현지 진출 지원 로드맵"),
    ("blank", ""),
    ("title2", "1. 단계별 진출 로드맵"),
    ("bullet", "Phase 1 (2026~2027): PMC 사업 수행 중 현지 파트너 발굴 및 시장 조사"),
    ("bullet", "Phase 2 (2027~2028): 한국 에너지 기업 컨소시엄 구성 → 파일럿 투자(500kW 운영권)"),
    ("bullet", "Phase 3 (2029~): SPC 설립 → 쾅고 주 추가 태양광 사업 확장(5MW 이상 목표)"),
    ("blank", ""),
    ("title2", "2. 지역사회 참여 메커니즘 [신규 - Perplexity 보완]"),
    ("body", "사업의 사회적 지속가능성과 지역사회 수용성 확보를 위한 체계적 참여 구조가 필요합니다."),
    ("bullet", "주민위원회 구성: UNDP 조직, 12명 내외 선출 (사용자-운영자 가교 역할)"),
    ("bullet", "2단계 교육: ① 거버넌스(위원회 운영·리더십) → ② 인식개선(에너지 효율·요금 납부)"),
    ("bullet", "수익 공유: 탄소크레딧 수익 일부 및 전력 판매 수익의 일정 비율 지역사회 환원"),
    ("bullet", "갈등 관리: 정기 주민 간담회(분기별) + 민원 처리 절차 명문화"),
    ("bullet", "젠더 포용: 여성 위원 최소 30% 포함, 여성 에너지 기업인 양성 프로그램 연계"),
    ("blank", ""),
    ("title2", "3. 유사 사례 벤치마킹 [신규 - Perplexity 보완]"),
    ("body", "성공적인 유사 사례를 통해 제안의 현실성과 신뢰도를 높입니다."),
    ("bullet", "KOICA-ENGIE 케냐 태양광(2019): 블렌디드 파이낸스 구조, 탄소크레딧 병행 → 지속가능성 입증"),
    ("bullet", "ADB 솔로몬제도 미니그리드(2022): 도서 오지 500kW급, PMC 주도 O&M 체계 정립"),
    ("bullet", "World Bank 마다가스카르 전력화(2021): 주민위원회 수익공유 모델, 요금 체계 성공 사례"),
    ("bullet", "한국중부발전 르완다(2020): 한국 에너지 기업 ODA 연계 진출 선례 — Phase 3 벤치마크"),
    ("blank", ""),
    ("title2", "4. 기대 효과"),
    ("bullet", "경제적: 한국 기업 약 3~5개사 DR콩고 진출, 수출액 약 500만~1,000만 달러 예상"),
    ("bullet", "개발: 쾅고 주 650가구 이상 전력 접근, 켕게 HGR 의료서비스 정상화"),
    ("bullet", "환경: 연간 약 800~1,000톤 CO₂ 감축(디젤 발전 대체 기준)"),
    ("blank", ""),

    # ── M&E 프레임워크 ──
    ("title1", "IV. 모니터링·평가(M&E) 프레임워크 [신규 - Perplexity 보완]"),
    ("blank", ""),
    ("title2", "1. M&E 체계 개요"),
    ("body", "PMC는 KOICA 위임자로서 사업 전 기간 성과관리를 총괄하며, "
             "UNDP 산출물에 대한 1차 기술 검토 및 품질관리를 수행합니다."),
    ("blank", ""),
    ("title2", "2. 성과지표(KPI) 체계"),
    ("bullet", "인프라: 태양광 시스템 가동률 ≥95%, ESS 충방전 효율 ≥90%"),
    ("bullet", "에너지 접근: 수혜 가구 수 ≥650가구, 켕게 HGR 24시간 전력 공급"),
    ("bullet", "역량강화: 정부 전문가 320명 교육 이수, O&M팀 독립 운영 달성"),
    ("bullet", "제도 개선: 신재생에너지 관련 법령 개정안 제출, 액션플랜 이행률 ≥80%"),
    ("bullet", "환경: 연간 탄소 감축량 검증 완료, VCS/GS 등록 완료"),
    ("blank", ""),
    ("title2", "3. 모니터링 주기 및 보고 체계"),
    ("bullet", "월간: PMC FM 현장 점검 보고서 → KOICA 사무소 제출"),
    ("bullet", "분기: PIU 실무 조정회의 → UNDP·MRHE 공동 검토"),
    ("bullet", "연간: PSC 전체 회의 → 연차 성과보고서 KOICA 본부 제출"),
    ("bullet", "최종: 사업 완료 보고서 + 지속가능성 평가 + 교훈 도출"),
    ("blank", ""),
    ("title2", "4. 위험 감지 및 대응 체계"),
    ("bullet", "원격 모니터링 시스템: IoT 센서 → 실시간 이상 징후 알림"),
    ("bullet", "조기경보: KPI 목표 대비 80% 미달 시 즉시 PMC-KOICA 협의"),
    ("bullet", "환경사회위험: ESIA 권고사항 이행 여부 분기별 점검"),
    ("blank", ""),

    # ── 거버넌스 ──
    ("title1", "V. PMC 거버넌스 및 수행 체계"),
    ("blank", ""),
    ("title2", "1. 거버넌스 구조"),
    ("bullet", "PSC (최고의사결정): KOICA 사무소장 + MRHE 차관 공동의장, 연1회 이상"),
    ("bullet", "PIU (실무조정): PMC 리더 + UNDP + MRHE 과장급, 분기별"),
    ("bullet", "현장 책임자(FM): 사업 말기 1년 켕게 현장 상주, 기술 전수 및 O&M 감독"),
    ("blank", ""),
    ("title2", "2. UNDP 협력 관계"),
    ("body", "PMC는 UNDP의 상급 기술 검토자로서 기능합니다. "
             "UNDP가 수행하는 F/S, ESIA, EPC 조달 등 모든 주요 산출물에 대해 "
             "PMC가 1차 기술 검토 및 품질관리를 수행한 후 KOICA 최종 승인을 지원합니다."),
    ("blank", ""),

    # ── 종합 건의 ──
    ("title1", "VI. 종합 건의사항"),
    ("blank", ""),
    ("body", "본 특수제안서는 다음 3가지 핵심 가치를 기반으로 합니다:"),
    ("bullet", "① 혁신성: 블렌디드 파이낸스 + 탄소크레딧 수익 연계로 ODA 단독 재원의 한계 극복"),
    ("bullet", "② 지속가능성: 기술(O&M) + 정책(제도개선) + 사회(주민참여) 3축 통합 접근"),
    ("bullet", "③ 확장성: 쾅고 주 500kW 파일럿 → 킨샤사 광역 미니그리드 사업으로 단계적 확장"),
    ("blank", ""),
    ("body", "KOICA PMC 수행기관으로서 본 기관은 상기 특수제안 이행을 위한 전문인력, "
             "국제 네트워크, 탄소금융 실무 경험을 보유하고 있으며, "
             "사업의 성공적 완수와 한국 기업의 아프리카 에너지 시장 진출을 적극 지원할 것을 약속합니다."),
    ("blank", ""),
    ("body", "끝."),
]


# ──────────────────────────────────────────────
# HWPX XML 생성 함수
# ──────────────────────────────────────────────

CHAR_PR = {
    "title1":  "8",
    "title2":  "9",
    "title3":  "10",
    "body":    "7",
    "bullet":  "7",
    "blank":   "7",
}

PARA_PR = {
    "title1":  "2",
    "title2":  "3",
    "title3":  "4",
    "body":    "0",
    "bullet":  "1",
    "blank":   "0",
}

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


def escape_xml(text):
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def make_para(para_type, text, pid):
    char_id = CHAR_PR.get(para_type, "7")
    para_id = PARA_PR.get(para_type, "0")
    safe_text = escape_xml(text)

    # 불릿 접두사
    prefix = ""
    if para_type == "bullet":
        prefix = "• "

    content = f"{prefix}{safe_text}"

    return (
        f'<hp:p id="{pid}" paraPrIDRef="{para_id}" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{char_id}">'
        f'<hp:t xml:space="preserve">{content}</hp:t>'
        f'</hp:run>'
        f'</hp:p>\n'
    )


def build_section0():
    # 기존 템플릿의 섹션 설정(secPr) 재사용
    with open(f"{TEMPLATE_DIR}/Contents/section0.xml", "r", encoding="utf-8") as f:
        orig = f.read()

    # secPr 블록 추출 (첫 번째 <hp:p> 안에 있음)
    import re
    secpr_match = re.search(r'(<hp:secPr[^>]*>.*?</hp:secPr>)', orig, re.DOTALL)
    secpr = secpr_match.group(1) if secpr_match else ""

    lines = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>']
    lines.append(f'<hs:sec {NAMESPACES}>')

    # 첫 번째 단락에 secPr 포함
    first_type, first_text = PARAGRAPHS[0]
    char_id = CHAR_PR.get(first_type, "7")
    para_id = PARA_PR.get(first_type, "0")
    safe_text = escape_xml(first_text)
    lines.append(
        f'<hp:p id="1000" paraPrIDRef="{para_id}" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{char_id}">'
        f'{secpr}'
        f'<hp:t xml:space="preserve">{safe_text}</hp:t>'
        f'</hp:run>'
        f'</hp:p>\n'
    )

    for i, (ptype, text) in enumerate(PARAGRAPHS[1:], start=1001):
        lines.append(make_para(ptype, text, i))

    lines.append('</hs:sec>')
    return "\n".join(lines)


def build_hwpx():
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    section0_xml = build_section0()

    with zipfile.ZipFile(OUTPUT_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        # mimetype (압축 없이)
        zf.writestr(zipfile.ZipInfo("mimetype"), "application/hwp+zip")

        # META-INF
        with open(f"{TEMPLATE_DIR}/META-INF/container.xml", "rb") as f:
            zf.writestr("META-INF/container.xml", f.read())
        with open(f"{TEMPLATE_DIR}/META-INF/manifest.xml", "rb") as f:
            zf.writestr("META-INF/manifest.xml", f.read())

        # Contents
        with open(f"{TEMPLATE_DIR}/Contents/header.xml", "rb") as f:
            zf.writestr("Contents/header.xml", f.read())

        # content.hpf 날짜만 업데이트
        now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        with open(f"{TEMPLATE_DIR}/Contents/content.hpf", "r", encoding="utf-8") as f:
            hpf = f.read()
        hpf = hpf.replace("2026-03-03T11:50:40Z", now).replace("2026-03-03T12:09:38Z", now)
        zf.writestr("Contents/content.hpf", hpf.encode("utf-8"))

        # section0.xml (새 내용)
        zf.writestr("Contents/section0.xml", section0_xml.encode("utf-8"))

        # settings.xml, version.xml
        with open(f"{TEMPLATE_DIR}/settings.xml", "rb") as f:
            zf.writestr("settings.xml", f.read())
        with open(f"{TEMPLATE_DIR}/version.xml", "rb") as f:
            zf.writestr("version.xml", f.read())

    print(f"✅ HWPX 생성 완료: {OUTPUT_PATH}")
    size = os.path.getsize(OUTPUT_PATH)
    print(f"   파일 크기: {size:,} bytes")


if __name__ == "__main__":
    build_hwpx()
