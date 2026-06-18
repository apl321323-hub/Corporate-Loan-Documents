"""
PDF 재무제표 파서
=================
pdfplumber로 텍스트를 추출하고, pdf_mapping.py의 매핑 규칙에 따라
BS / IS / Summary 엑셀 데이터의 특정 월 컬럼에 값을 업데이트한다.

반환 형식:
  {
    "period":  "2026.05",          # 업데이트 대상 기간 (사용자 입력)
    "bs":  { label: value, ... },  # 엑셀 BS label → 원 단위 값
    "is_": { label: value, ... },  # 엑셀 IS label → 원 단위 값
    "summary": { label: value, ... }  # 엑셀 Summary label → 원 단위 값
  }
"""

import re
import pdfplumber
from pdf_mapping import (
    PDF_BS_MAPPING, PDF_IS_MAPPING, PDF_SUMMARY_MAPPING,
    SUMMARY_CALC_RULES
)


# ──────────────────────────────────────────────
# 공통 유틸
# ──────────────────────────────────────────────
def _normalize(text: str) -> str:
    """
    공백·점·괄호·로마숫자 접두사 등 제거 후 정규화 (매핑 키 비교용)
    예: 'Ⅴ 영업이익' → '영업이익'
        'Ⅰ. 매출액'  → '매출액'
        '(2) 유형자산' → '유형자산'
    """
    # 로마숫자+아라비아숫자 접두사 제거 (예: Ⅴ, Ⅰ, Ⅹ, 1., (2), ①)
    t = re.sub(r'^[Ⅰ-Ⅹⅰ-ⅹ①-⑩\dⅠ]+[\s\.\)）]*', '', text.strip())
    # 공백·점·괄호 제거
    t = re.sub(r'[\s\u3000\u00a0()（）·・.]', '', t)
    return t


def _parse_number(text: str):
    """
    금액 문자열 → float
    예: '108,541,055,308' → 108541055308.0
        '(1,234)' → -1234.0  (괄호=음수)
        '0' → 0.0
    """
    if not text:
        return None
    t = text.strip()
    negative = t.startswith('(') and t.endswith(')')
    t = re.sub(r'[(),\s]', '', t)
    t = t.replace(',', '')
    try:
        v = float(t)
        return -v if negative else v
    except ValueError:
        return None


# ──────────────────────────────────────────────
# PDF 텍스트 → 과목:금액 dict 추출
# ──────────────────────────────────────────────
def _extract_raw_values(pdf_path: str) -> dict:
    """
    PDF 전체 텍스트를 줄 단위로 파싱해
    { normalize(과목명): (당기금액, 전기금액) } 딕셔너리 반환
    
    처리 전략:
    - 줄 끝의 숫자 1~2개를 금액으로, 앞부분을 과목명으로 처리
    - 로마숫자/번호 접두사 제거 후 정규화
    - 동일 과목이 여러 줄 → 마지막 값(소계/합계 우선) 사용
    """
    raw = {}

    # 금액 패턴: 3자리 콤마 숫자 (최소 1자리)
    AMT_PATTERN = re.compile(r'[\d]{1,3}(?:,\d{3})*(?:\.\d+)?')
    # 줄 끝 금액 추출 패턴 (당기 + 선택적 전기)
    LINE_PATTERN = re.compile(
        r'^(.+?)\s+([\d,]+)\s+([\d,]+)\s*$'   # 과목 당기 전기
        r'|^(.+?)\s+([\d,]+)\s*$'              # 과목 당기만
    )

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                m = LINE_PATTERN.match(line)
                if not m:
                    continue

                if m.group(1):  # 당기+전기
                    label_raw = m.group(1).strip()
                    cur_str   = m.group(2)
                    prev_str  = m.group(3)
                else:           # 당기만
                    label_raw = m.group(4).strip()
                    cur_str   = m.group(5)
                    prev_str  = None

                label_raw = label_raw.strip()
                if not label_raw:
                    continue

                # 숫자만으로 된 과목명 제거 (항목 번호 줄 등)
                if re.match(r'^[\dⅠ-Ⅹ\s\.\(\)]+$', label_raw):
                    continue

                norm = _normalize(label_raw)
                if len(norm) < 2:
                    continue

                cur  = _parse_number(cur_str)
                prev = _parse_number(prev_str) if prev_str else None

                # 합계/총계 줄이 더 신뢰도 높으므로 무조건 덮어씌움
                raw[norm] = (cur, prev)

    # 보조: 단순 줄 파싱 (위 패턴 못 잡은 것)
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                line = line.strip()
                if not line:
                    continue
                # 숫자 추출
                nums = re.findall(r'\d{1,3}(?:,\d{3})+', line)
                if not nums:
                    continue
                label_part = re.sub(r'[\d,]+', '', line).strip()
                label_part = re.sub(r'\s+', ' ', label_part).strip()
                norm = _normalize(label_part)
                if len(norm) < 2 or norm in raw:
                    continue
                parsed = [_parse_number(n) for n in nums]
                parsed = [v for v in parsed if v is not None]
                if parsed:
                    raw[norm] = (parsed[0], parsed[1] if len(parsed) > 1 else None)

    return raw


# ──────────────────────────────────────────────
# 재무상태표(BS) 파싱
# ──────────────────────────────────────────────
def _parse_bs(raw: dict) -> dict:
    """
    raw dict + PDF_BS_MAPPING → {엑셀_label: 원단위값}
    합산 항목(현금/예치금, 기타자산, 기타부채 등)은 더해서 저장
    """
    result = {}   # label → float

    def add(label, val):
        if label is None or val is None:
            return
        result[label] = result.get(label, 0.0) + val

    for norm_key, excel_label in PDF_BS_MAPPING.items():
        if excel_label is None:
            continue
        if norm_key in raw:
            cur_val, _ = raw[norm_key]
            if cur_val is not None:
                add(excel_label, cur_val)

    # 유형자산 순장부가 = 취득가 - 감가상각누계액 → 3. 기타자산에 합산
    # PDF 구조: 비품(취득가) / 감가상각누계액XX (누계) / 시설장치(취득가) / 감가상각누계액XX (누계)
    tangible_net = 0.0
    for k, v in raw.items():
        # "비품", "시설장치" → 취득가
        if k in ("비품", "시설장치"):
            if v[0] is not None:
                tangible_net += v[0]
        # "감가상각누계액..." → 누계액 차감
        elif k.startswith("감가상각누계액"):
            if v[0] is not None:
                tangible_net -= v[0]
    if tangible_net != 0:
        add("3. 기타자산", tangible_net)

    # 무형자산 순장부가 = 소프트웨어 + 개발비 → 3. 기타자산에 합산
    for k in ("소프트웨어", "개발비"):
        if k in raw and raw[k][0] is not None:
            add("3. 기타자산", raw[k][0])

    # 대손충당금: 엑셀 원본이 양수 표기 기준이므로 절댓값으로 저장
    if "대손충당금" in result:
        result["대손충당금"] = abs(result["대손충당금"])

    return result


# ──────────────────────────────────────────────
# 손익계산서(IS) 파싱
# ──────────────────────────────────────────────
def _parse_is(raw: dict) -> dict:
    result = {}

    def add(label, val):
        if label is None or val is None:
            return
        result[label] = result.get(label, 0.0) + val

    for norm_key, excel_label in PDF_IS_MAPPING.items():
        if excel_label is None:
            continue
        if norm_key in raw:
            cur_val, _ = raw[norm_key]
            if cur_val is not None:
                add(excel_label, cur_val)

    return result


# ──────────────────────────────────────────────
# Summary 파싱 (BS + IS 결과 활용)
# ──────────────────────────────────────────────
def _parse_summary(raw: dict, bs_vals: dict, is_vals: dict) -> dict:
    result = {}

    # 1) PDF 직접 매핑
    for norm_key, summary_label in PDF_SUMMARY_MAPPING.items():
        if summary_label is None:
            continue
        # raw에서 직접 찾기
        if norm_key in raw:
            cur_val, _ = raw[norm_key]
            if cur_val is not None:
                result[summary_label] = cur_val

    # 2) BS 값에서 가져오기 (직접 매핑보다 우선)
    bs_to_summary = {
        "Ⅰ. 자산총계":  "자산총계",
        "Ⅱ. 부채총계":  "부채총계",
        "Ⅲ. 자본총계":  "자본총계",
        "1. 자본금":     "자본금",
        "2. 자본잉여금(자본조정)": "자본잉여금",
        "3. 이익잉여금": "이익잉여금",
        "사채":          "사채",
    }
    for bs_label, sum_label in bs_to_summary.items():
        if bs_label in bs_vals:
            result[sum_label] = bs_vals[bs_label]

    # 3) IS 값에서 가져오기
    is_to_summary = {
        "Ⅰ. 영업수익":              "영업수익",
        "1. 차입금 이자비용":        "지급이자",
        "중개수수료":                "중개수수료",
        "대손상각비":                "대손상각비",
        "Ⅳ. 영업이익":              "영업이익",
        "Ⅵ. 당기순이익":            "당기손익",
    }
    for is_label, sum_label in is_to_summary.items():
        if is_label in is_vals:
            result[sum_label] = is_vals[is_label]

    # raw에서 직접 키로 보완 (로마숫자 접두사 제거 후 매핑)
    raw_direct_map = [
        ("매출액",             "영업수익"),
        ("영업이익",           "영업이익"),
        ("당기순이익",         "당기손익"),
        ("판매비와관리비",     "판관비"),
        ("차입금및회사채이자", "지급이자"),
    ]
    for pdf_key, sum_label in raw_direct_map:
        nk = _normalize(pdf_key)
        if nk in raw:
            cur_v, prev_v = raw[nk]
            # cur_v가 항목번호(작은 정수)인 경우 prev_v가 실제 값
            v = None
            if cur_v is not None and cur_v > 10000:
                v = cur_v
            elif prev_v is not None and prev_v > 10000:
                v = prev_v
            if v is not None:
                result[sum_label] = v

    # 4) 판관비 = raw에서 직접 (위에서 처리되었을 것)
    판관비_key = _normalize("판매비와관리비")
    if 판관비_key in raw and "판관비" not in result:
        cur_v, prev_v = raw[판관비_key]
        v = cur_v if (cur_v and cur_v > 10000) else (prev_v if (prev_v and prev_v > 10000) else None)
        if v:
            result["판관비"] = v

    # 5) 현금및현금성자산 = 현금+보통예금+단기금융상품 (raw 합산)
    keys_cash = ["현금", "보통예금", "기타단기금융상품"]
    total_cash = 0.0
    for k in keys_cash:
        nk = _normalize(k)
        if nk in raw and raw[nk][0] is not None:
            total_cash += raw[nk][0]
    if total_cash > 0:
        result["현금및현금성자산"] = total_cash

    # 6) 대출채권, 대손충당금 (raw에서 직접)
    for pdf_key, sum_label in [("대출채권", "대출채권"), ("대손충당금", "대손충당금")]:
        nk = _normalize(pdf_key)
        if nk in raw and raw[nk][0] is not None:
            v = raw[nk][0]
            # 대손충당금: 엑셀 원본이 양수 기준이므로 절댓값으로 저장
            result[sum_label] = abs(v) if sum_label == "대손충당금" else v

    # 7) 유형자산: (항목번호, 소계) 구조 → prev_v가 실제 소계
    유형_key = _normalize("유형자산")
    if 유형_key in raw:
        cur_v, prev_v = raw[유형_key]
        if prev_v is not None and prev_v > 1000:
            result["유형자산"] = prev_v
        elif cur_v is not None and cur_v > 1000:
            result["유형자산"] = cur_v

    # 8) 무형자산: 마찬가지로 두 번째 값이 소계
    무형_key = _normalize("무형자산")
    if 무형_key in raw:
        cur_v, prev_v = raw[무형_key]
        if prev_v is not None and prev_v > 1000:
            result["무형자산"] = prev_v
        elif cur_v is not None and cur_v > 1000:
            result["무형자산"] = cur_v
        else:
            소프트_key = _normalize("소프트웨어")
            if 소프트_key in raw and raw[소프트_key][0]:
                result["무형자산"] = raw[소프트_key][0]

    # 9) 기타비유동자산 = 임차보증금 + 기타보증금
    keys_etc = ["임차보증금", "기타보증금"]
    total_etc = 0.0
    for k in keys_etc:
        nk = _normalize(k)
        if nk in raw and raw[nk][0] is not None:
            total_etc += raw[nk][0]
    if total_etc > 0:
        result["기타비유동자산"] = total_etc

    # 10) 장기차입금, 단기차입금 (Summary에 별도 항목)
    for pdf_key, sum_label in [("장기차입금", "장기차입금"), ("단기차입금", "단기차입금")]:
        nk = _normalize(pdf_key)
        if nk in raw and raw[nk][0] is not None:
            result[sum_label] = raw[nk][0]

    # 11) 미지급금, 예수금, 미지급세금, 미지급비용, 선수금 (개별)
    individual_items = [
        ("미수수익", "미수수익"),
        ("미수금", "미수금"),
        ("선급금", "선급금"),
        ("가지급금", "가지급금"),
        ("선급비용", "선급비용"),
        ("선납세금", "선납세금"),
        ("미지급금", "미지급금"),
        ("예수금", "예수금"),
        ("미지급세금", "미지급세금"),
        ("미지급비용", "미지급비용"),
        ("당기법인세부채", "당기법인세부채"),
        ("선수금", "선수금"),
    ]
    for pdf_key, sum_label in individual_items:
        nk = _normalize(pdf_key)
        if nk in raw and raw[nk][0] is not None:
            result[sum_label] = raw[nk][0]

    # 12) 부채비율 계산 = 부채총계 / 자본총계 * 100
    if "부채총계" in result and "자본총계" in result and result["자본총계"] != 0:
        result["부채비율"] = result["부채총계"] / result["자본총계"] * 100

    return result


# ──────────────────────────────────────────────
# 공개 인터페이스
# ──────────────────────────────────────────────
def apply_user_mapping(raw: dict, user_mapping: dict) -> dict:
    """
    사용자가 UI에서 설정한 매핑을 raw 값에 적용
    user_mapping: { "bs": [{pdf_key, excel_label, enabled}], "is_": [...], "summary": [...] }
    반환: { "bs": {excel_label: val}, "is_": {...}, "summary": {...} }
    """
    result = {"bs": {}, "is_": {}, "summary": {}}

    for sheet_key in ["bs", "is_", "summary"]:
        rows = user_mapping.get(sheet_key, [])
        acc = {}  # excel_label → 합산값 (여러 PDF 과목이 같은 엑셀 과목으로 매핑될 수 있음)
        for row in rows:
            if not row.get("enabled", True):
                continue
            excel_label = row.get("excel_label")
            if not excel_label:
                continue
            pdf_key = row.get("pdf_key", "")
            if pdf_key in raw:
                val = raw[pdf_key][0]  # 당기값
                if val is not None:
                    acc[excel_label] = acc.get(excel_label, 0.0) + val
        result[sheet_key] = acc

    return result


def parse_pdf_fs(pdf_path: str, user_mapping: dict = None) -> dict:
    """
    재무제표 PDF 파싱 → BS/IS/Summary 값 딕셔너리 반환

    Args:
        pdf_path:     PDF 파일 경로
        user_mapping: UI에서 저장한 매핑 설정 (없으면 기본 매핑 사용)

    Returns:
        {
          "bs":      { excel_label: float_value, ... },
          "is_":     { excel_label: float_value, ... },
          "summary": { excel_label: float_value, ... },
          "raw":     { norm_label: [cur, prev], ... }
        }
    """
    raw = _extract_raw_values(pdf_path)

    if user_mapping:
        # 사용자 정의 매핑 사용
        mapped = apply_user_mapping(raw, user_mapping)
        bs_vals  = mapped["bs"]
        is_vals  = mapped["is_"]
        sum_vals = mapped["summary"]
    else:
        # 기본 코드 매핑 사용
        bs_vals  = _parse_bs(raw)
        is_vals  = _parse_is(raw)
        sum_vals = _parse_summary(raw, bs_vals, is_vals)

    return {
        "bs":      bs_vals,
        "is_":     is_vals,
        "summary": sum_vals,
        "raw":     {k: list(v) for k, v in raw.items()},
    }


# ──────────────────────────────────────────────
# bsis.json 업데이트 함수
# ──────────────────────────────────────────────
def apply_pdf_to_bsis(bsis_data: dict, parsed: dict, period: str) -> dict:
    """
    기존 bsis_data에 PDF 파싱 결과를 특정 기간(period)에 덮어씌움

    Args:
        bsis_data: 기존 bsis.json 전체 딕셔너리
        parsed:    parse_pdf_fs() 반환값
        period:    "2026.05" 형식의 업데이트 대상 기간

    Returns:
        업데이트된 bsis_data (in-place 수정 후 반환)
    """
    sheets_updated = []

    for sheet_key, excel_map in [("bs", parsed["bs"]),
                                  ("is_", parsed["is_"]),
                                  ("summary", parsed["summary"])]:
        sheet = bsis_data.get(sheet_key)
        if not sheet:
            continue

        headers = sheet.get("headers", [])
        # period가 headers에 있는지 확인
        if period not in headers:
            continue

        col_idx = headers.index(period)
        rows    = sheet.get("rows", [])
        updated = 0

        for row in rows:
            label = row.get("label", "")
            if label in excel_map:
                val = excel_map[label]
                # values 리스트 길이 보장
                while len(row["values"]) <= col_idx:
                    row["values"].append(None)
                row["values"][col_idx] = val
                updated += 1

        if updated > 0:
            sheets_updated.append(f"{sheet_key}({updated}개 항목)")

    bsis_data["_pdf_updated"] = bsis_data.get("_pdf_updated", {})
    bsis_data["_pdf_updated"][period] = True

    return bsis_data, sheets_updated


# ──────────────────────────────────────────────
# 테스트용 실행
# ──────────────────────────────────────────────
if __name__ == "__main__":
    import json, sys

    pdf_path = sys.argv[1] if len(sys.argv) > 1 else \
        "/home/user/uploaded_files/2605재무제표_에이피엘_dl.pdf"

    print("=== PDF 파싱 테스트 ===")
    result = parse_pdf_fs(pdf_path)

    print("\n[BS 추출값]")
    for k, v in sorted(result["bs"].items()):
        print(f"  {k}: {v:,.0f}")

    print("\n[IS 추출값]")
    for k, v in sorted(result["is_"].items()):
        print(f"  {k}: {v:,.0f}")

    print("\n[Summary 추출값]")
    for k, v in sorted(result["summary"].items()):
        if isinstance(v, float):
            print(f"  {k}: {v:,.0f}")
        else:
            print(f"  {k}: {v}")
