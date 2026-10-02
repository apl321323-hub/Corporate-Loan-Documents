"""
대출자산속성 파서 (대출자산속성 (APL기준) 시트 기준)
행4: 날짜 헤더 (datetime)
각 섹션: ▶로 시작하는 행 → 하위 데이터 행들
"""
import openpyxl
import re
from datetime import datetime


# ── 섹션 정의 ────────────────────────────────────────────────────
# (섹션 키, 헤더 행번호(1-indexed), 데이터 시작 행, 데이터 끝 행, 스킵할 레이블 목록)
SECTION_DEFS = [
    # key, header_row, data_start, data_end, skip_labels
    ("대출취급건수",       3,  5, 10,  ["구분", "승인율"]),
    ("대출채권잔액",      11, 13, 21,  ["구분", "영업현황 검수"]),
    ("대출취급액",        23, 25, 30,  ["구분", "평균단가", "추가재대율"]),
    ("대출만기",          31, 33, 38,  ["구분", "자산건정성 검수"]),
    ("평균이자율",        40, 42, 44,  ["구분"]),
    ("평균이자율2",       45, 47, 49,  ["구분"]),
    ("상환방식",          50, 52, 61,  ["구분", "대출만기 검수"]),
    ("금액별",           63, 65, 77,  ["구분", "상환방식 검수"]),
    ("접수경로",         79, 81, 87,  ["구분", "금액별 검수"]),
    ("상품별",           89, 91, 121, ["구분", "접수경로 검수"]),
    ("상품구분",        122, 124, 129, ["구분", "상품별 검수"]),
    ("상품구분상세",     130, 132, 138, ["구분", "상품별 검수"]),
    ("상품구분화해채권", 139, 141, 148, ["구분", "상품별 검수"]),
    ("성별",            148, 150, 153, ["구분", "CB 검수"]),
    ("연령별",          154, 156, 161, ["구분", "성별 검수"]),
    ("직업별",          162, 164, 169, ["구분", "연령별 검수", "학생"]),
    ("지역별",          171, 173, 184, ["구분", "융잔검수"]),
    ("CB등급",          187, 189, 202, ["구분", "상품별 검수"]),
]


def _is_skip(label: str, skip_labels: list) -> bool:
    if not label or not isinstance(label, str):
        return True
    for sk in skip_labels:
        if label.strip() == sk or label.strip().startswith("▶"):
            return True
    # 검수 행 스킵
    if "검수" in label:
        return True
    return False


def _normalize_cb_product_label(label: str) -> str | None:
    text = re.sub(r"\s+", "", str(label).replace("\xa0", " ").strip())
    if not text:
        return None
    if text in ("무등급", "무분위"):
        return "무등급"
    if "융잔합계" in text or "전체융잔합계" in text:
        return "융잔합계"

    match = re.match(r"^(\d+)(?:분위|등급)\(([^)]*)\)", text)
    if not match:
        return None
    grade_no = match.group(1)
    score_range = match.group(2).replace("점", "")
    return f"{grade_no}분위({score_range})"


def _merge_cb_no_grade_into_first(cb_products: dict[str, dict[str, dict[str, float | None]]]) -> None:
    for grades in cb_products.values():
        no_grade = grades.pop("무등급", None)
        if not no_grade:
            continue
        first_key = next((key for key in grades if key.startswith("1분위")), None) or "1분위(0~724)"
        first_row = dict(grades.get(first_key) or {})
        for ym, val in no_grade.items():
            if val is None:
                continue
            base = first_row.get(ym)
            first_row[ym] = (base if base is not None else 0) + val
        grades[first_key] = first_row


def parse_asset(filepath: str) -> dict:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))

    # 행4(0-indexed: 3)에서 날짜 헤더 추출
    header = all_rows[3]
    period_cols: list[tuple[int, str]] = []
    for i, v in enumerate(header):
        if isinstance(v, datetime):
            period_cols.append((i, v.strftime("%Y-%m")))

    periods = [ym for _, ym in period_cols]

    # 섹션별 데이터 추출
    sections: dict[str, dict[str, dict[str, float | None]]] = {}

    for sec_key, hrow, dstart, dend, skip_labels in SECTION_DEFS:
        sec_data: dict[str, dict[str, float | None]] = {}
        # dstart, dend: 1-indexed → 0-indexed
        for ri in range(dstart - 1, min(dend, len(all_rows))):
            row = all_rows[ri]
            label = row[0]
            if not label or not isinstance(label, str):
                continue
            label = label.strip()
            if _is_skip(label, skip_labels):
                continue
            row_data: dict[str, float | None] = {}
            for ci, ym in period_cols:
                v = row[ci] if ci < len(row) else None
                if isinstance(v, (int, float)):
                    row_data[ym] = float(v)
                else:
                    row_data[ym] = None
            sec_data[label] = row_data
        sections[sec_key] = sec_data

    # 상품별CB등급 섹션 (행203~685)
    # 구조: 상품명 행 → 등급별 행들 (융잔합계 포함)
    cb_products: dict[str, dict[str, dict[str, float | None]]] = {}
    # 알려진 상품명 목록 (상품별 섹션에서 추출)
    known_products = [
        "N론", "V론", "담보론", "담보론(지분대출)", "레이디론", "스타론",
        "스타스위치론", "우량론", "큐브론", "테일론", "토탈론", "프리론",
        "프리미엄론", "전월세론", "플러스론", "토마토론", "토마토N론",
        "토마토토탈", "T플러스론", "OP론", "충성론", "오투론", "오투N론",
        "토마토토탈론플러스", "다이렉트론(A)", "다이렉트론(W)", "N론(하이브리드)",
        "기타", "기타N",
    ]

    current_product: str | None = None
    for ri in range(201, len(all_rows)):
        row = all_rows[ri]
        label = row[0]
        if not label:
            continue

        # 날짜값이 있는 행 = 상품명 헤더 행
        if isinstance(label, str) and label.strip() in known_products:
            current_product = label.strip()
            cb_products[current_product] = {}
            continue

        if current_product and isinstance(label, str):
            label = _normalize_cb_product_label(label)
            if label:
                row_data: dict[str, float | None] = {}
                for ci, ym in period_cols:
                    v = row[ci] if ci < len(row) else None
                    if isinstance(v, (int, float)):
                        row_data[ym] = float(v)
                    else:
                        row_data[ym] = None
                cb_products[current_product][label] = row_data

    _merge_cb_no_grade_into_first(cb_products)

    return {
        "periods": periods,
        "sections": sections,
        "cb_products": cb_products,
    }


if __name__ == "__main__":
    import json
    result = parse_asset("/home/user/uploaded_files/대출자산속성.xlsx")
    print(f"기간 수: {len(result['periods'])}, 범위: {result['periods'][0]} ~ {result['periods'][-1]}")
    print(f"섹션 수: {len(result['sections'])}")
    for k, v in result["sections"].items():
        keys = list(v.keys())
        print(f"  {k}: {len(keys)}개 항목 → {keys}")
    print(f"\n상품별CB등급 상품 수: {len(result['cb_products'])}")
    for pname, grades in result["cb_products"].items():
        print(f"  {pname}: {list(grades.keys())[:3]}...")
