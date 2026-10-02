"""
영업현황 엑셀 파서
파일: 기업여신자료 - 영업현황.xlsx  (시트: '영업현황 (APL기준)')

구조:
  Row 1   : ■ 영업현황 (헤더)
  Row 2   : (단위:백만원)
  Row 3   : 구분 | 2016-12 | ... | 2026-05  (기간 헤더, 103개)

  [섹션1] 전체 영업현황 (Row 4~18)
    Row 4  : 기초대출자산
    Row 5  : 대출 - 신규
    Row 6  :      - 추가
    Row 7  :      - 재대출
    Row 8  :      - 취급액(대출)
    Row 9  : 매입
    Row 10 : 원금회수액
    Row 11 : 이자회수액
    Row 12 : 실상각액 - 월중매각액
    Row 13 :         - 월중상각액
    Row 14 : 기말자산
    Row 15 : 대손충당 - 기초 충당금
    Row 16 :         - 설정분
    Row 17 :         - 사용분(환입)
    Row 18 : 기말 충당금

  [섹션2] 대출구분별 (Row 21~28)
    Row 21 : 신용대출 - 신규
    Row 22 :         - 추가
    Row 23 :         - 재대출
    Row 24 :         - 취급액(대출)
    Row 25 : 기업대출 - 신규
    Row 26 :         - 추가
    Row 27 :         - 재대출
    Row 28 :         - 취급액(대출)

  [섹션3] 담보구분별 (Row 31~38)
    Row 31 : 신용대출 - 신규
    Row 32 :         - 추가
    Row 33 :         - 재대출
    Row 34 :         - 취급액(대출)
    Row 35 : 담보대출 - 신규
    Row 36 :         - 추가
    Row 37 :         - 재대출
    Row 38 :         - 취급액(대출)
"""

from datetime import datetime
import openpyxl


# ── 행 정의 ───────────────────────────────────────────────────
SECTION1_ROWS = {
    4:  ('기초대출자산', None),
    5:  ('대출', '신규'),
    6:  ('대출', '추가'),
    7:  ('대출', '재대출'),
    8:  ('대출', '취급액(대출)'),
    9:  ('매입', None),
    10: ('원금회수액', None),
    11: ('이자회수액', None),
    12: ('실상각액', '월중매각액'),
    13: ('실상각액', '월중상각액'),
    14: ('기말자산', None),
    15: ('대손충당', '기초 충당금'),
    16: ('대손충당', '설정분'),
    17: ('대손충당', '사용분(환입)'),
    18: ('기말 충당금', None),
}

SECTION2_ROWS = {
    21: ('신용대출', '신규'),
    22: ('신용대출', '추가'),
    23: ('신용대출', '재대출'),
    24: ('신용대출', '취급액(대출)'),
    25: ('기업대출', '신규'),
    26: ('기업대출', '추가'),
    27: ('기업대출', '재대출'),
    28: ('기업대출', '취급액(대출)'),
}

SECTION3_ROWS_LEGACY = {
    31: ('\uc2e0\uc6a9\ub300\ucd9c', '\uc2e0\uaddc'),
    32: ('\uc2e0\uc6a9\ub300\ucd9c', '\ucd94\uac00'),
    33: ('\uc2e0\uc6a9\ub300\ucd9c', '\uc7ac\ub300\ucd9c'),
    34: ('\uc2e0\uc6a9\ub300\ucd9c', '\ucde8\uae09\uc561(\ub300\ucd9c)'),
    35: ('\ub2f4\ubcf4\ub300\ucd9c', '\uc2e0\uaddc'),
    36: ('\ub2f4\ubcf4\ub300\ucd9c', '\ucd94\uac00'),
    37: ('\ub2f4\ubcf4\ub300\ucd9c', '\uc7ac\ub300\ucd9c'),
    38: ('\ub2f4\ubcf4\ub300\ucd9c', '\ucde8\uae09\uc561(\ub300\ucd9c)'),
}

SECTION3_ROWS_CURRENT = {
    21: ('\uc2e0\uc6a9\ub300\ucd9c', '\uc2e0\uaddc'),
    22: ('\uc2e0\uc6a9\ub300\ucd9c', '\ucd94\uac00'),
    23: ('\uc2e0\uc6a9\ub300\ucd9c', '\uc7ac\ub300\ucd9c'),
    24: ('\uc2e0\uc6a9\ub300\ucd9c', '\ucde8\uae09\uc561(\ub300\ucd9c)'),
    25: ('\ub2f4\ubcf4\ub300\ucd9c', '\uc2e0\uaddc'),
    26: ('\ub2f4\ubcf4\ub300\ucd9c', '\ucd94\uac00'),
    27: ('\ub2f4\ubcf4\ub300\ucd9c', '\uc7ac\ub300\ucd9c'),
    28: ('\ub2f4\ubcf4\ub300\ucd9c', '\ucde8\uae09\uc561(\ub300\ucd9c)'),
}

SECTION3_ROWS = SECTION3_ROWS_CURRENT

# row → key 매핑 (section별)
def _row_key(group, sub):
    if sub:
        return f"{group}_{sub}"
    return group

# 섹션별 key 순서
S1_KEYS = [_row_key(*v) for v in SECTION1_ROWS.values()]
S2_KEYS = [_row_key(*v) for v in SECTION2_ROWS.values()]
S3_KEYS = [_row_key(*v) for v in SECTION3_ROWS_CURRENT.values()]


def _safe_num(v) -> float:
    """셀 값을 숫자로 변환, 백만원 단위로 반올림"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v / 1_000_000
    return None



def _has_numeric_values(values: dict) -> bool:
    for series in values.values():
        if any(v is not None for v in series):
            return True
    return False


def _is_current_collateral_layout(ws) -> bool:
    label = ws.cell(row=25, column=1).value
    return isinstance(label, str) and '\ub2f4\ubcf4' in label

def _period_from_value(value):
    if isinstance(value, datetime):
        return value.strftime('%Y-%m')
    if isinstance(value, str):
        import re
        text = value.strip().replace(".", "-")
        m = re.search(r"(\d{4})-(\d{1,2})", text)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}"
    return None

def _find_period_columns(ws):
    best_row = None
    best_cols = []
    for row_idx in range(1, min(ws.max_row, 12) + 1):
        cols = []
        for cell in ws[row_idx]:
            period = _period_from_value(cell.value)
            if period:
                cols.append((cell.column, period))
        if len(cols) > len(best_cols):
            best_row = row_idx
            best_cols = cols
    return best_row, best_cols

def _clean_label(value):
    return str(value or "").strip()

def _compact_collateral_rows(ws, header_row):
    if not header_row:
        return {}
    allowed_subs = {"신규", "추가", "재대출", "취급액(대출)"}
    rows = {}
    current_group = ""
    for row_idx in range(header_row + 1, ws.max_row + 1):
        group = _clean_label(ws.cell(row=row_idx, column=1).value)
        sub = _clean_label(ws.cell(row=row_idx, column=2).value)
        if group:
            current_group = group
        if current_group not in {"신용대출", "담보대출", "신용", "담보"}:
            continue
        if sub not in allowed_subs:
            continue
        group_key = current_group if current_group.endswith("대출") else f"{current_group}대출"
        rows[f"{group_key}_{sub}"] = row_idx
    return rows

def _is_collateral_only_sheet(ws, header_row):
    return ws.max_row <= 15 and bool(_compact_collateral_rows(ws, header_row))

def parse_business_excel(filepath: str) -> dict:
    """
    영업현황 엑셀 파싱 → JSON 직렬화 가능한 dict 반환

    Returns
    -------
    {
      "periods" : ["2016-12", "2017-12", ...],   # 103개
      "section1": {
          "기초대출자산": [v1, v2, ...],
          "대출_신규":    [v1, v2, ...],
          ...
      },
      "section2": {
          "신용대출_신규": [...],
          ...
      },
      "section3": {
          "신용대출_신규": [...],
          ...
      }
    }
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb['영업현황 (APL기준)'] if '영업현황 (APL기준)' in wb.sheetnames else None
    if ws is None:
        for sheet_name in wb.sheetnames:
            if '영업현황' in sheet_name:
                ws = wb[sheet_name]
                break
    if ws is None:
        ws = wb.active

    # ── 기간 헤더 추출 (3행) ──────────────────────────────────
    header_row, period_cols = _find_period_columns(ws)

    periods = [p for _, p in period_cols]
    col_to_idx = {col: i for i, (col, _) in enumerate(period_cols)}

    n = len(periods)

    def _extract_row(row_num) -> list:
        vals = [None] * n
        for col, idx in col_to_idx.items():
            v = ws.cell(row=row_num, column=col).value
            vals[idx] = _safe_num(v)
        return vals

    collateral_only = _is_collateral_only_sheet(ws, header_row)
    dynamic_section3_rows = _compact_collateral_rows(ws, header_row)

    if collateral_only:
        s1 = {}
        s2 = {}
        s3 = {key: _extract_row(row_num) for key, row_num in dynamic_section3_rows.items()}
        wb.close()
        return {
            "periods": periods,
            "section1": s1,
            "section2": s2,
            "section3": s3,
            "s1_keys": S1_KEYS,
            "s2_keys": S2_KEYS,
            "s3_keys": S3_KEYS,
            "is_collateral_only": True,
        }

    # ── 섹션1 ─────────────────────────────────────────────────
    s1 = {}
    for row_num, (group, sub) in SECTION1_ROWS.items():
        key = _row_key(group, sub)
        s1[key] = _extract_row(row_num)

    # ── 섹션2 ─────────────────────────────────────────────────
    s2 = {}
    for row_num, (group, sub) in SECTION2_ROWS.items():
        key = _row_key(group, sub)
        s2[key] = _extract_row(row_num)

    # ── 섹션3 ─────────────────────────────────────────────────
    s3 = {}
    if dynamic_section3_rows:
        for key, row_num in dynamic_section3_rows.items():
            s3[key] = _extract_row(row_num)
    else:
        section3_rows = SECTION3_ROWS_CURRENT if _is_current_collateral_layout(ws) else SECTION3_ROWS_LEGACY
        for row_num, (group, sub) in section3_rows.items():
            key = _row_key(group, sub)
            s3[key] = _extract_row(row_num)
        if section3_rows is SECTION3_ROWS_LEGACY and not _has_numeric_values(s3):
            s3 = {}
            for row_num, (group, sub) in SECTION3_ROWS_CURRENT.items():
                key = _row_key(group, sub)
                s3[key] = _extract_row(row_num)

    wb.close()

    return {
        "periods": periods,
        "section1": s1,
        "section2": s2,
        "section3": s3,
        "s1_keys": S1_KEYS,
        "s2_keys": S2_KEYS,
        "s3_keys": S3_KEYS,
        "is_collateral_only": False,
    }


# ── CLI 테스트 ─────────────────────────────────────────────────
if __name__ == "__main__":
    import json, sys
    filepath = sys.argv[1] if len(sys.argv) > 1 \
        else "/home/user/uploaded_files2/기업여신자료_영업현황.xlsx"

    result = parse_business_excel(filepath)
    print(f"기간: {result['periods'][0]} ~ {result['periods'][-1]} ({len(result['periods'])}개)")
    print()
    print("=== 섹션1 최근 3개월 ===")
    last3 = result['periods'][-3:]
    idxs = list(range(len(result['periods'])-3, len(result['periods'])))
    for k in result['s1_keys']:
        vals = result['section1'][k]
        row = [f"{vals[i]:>10}" if vals[i] is not None else f"{'None':>10}" for i in idxs]
        print(f"  {k:20s}: {' | '.join(row)}")
    print()
    print("=== 섹션2 최근 3개월 ===")
    for k in result['s2_keys']:
        vals = result['section2'][k]
        row = [f"{vals[i]:>10}" if vals[i] is not None else f"{'None':>10}" for i in idxs]
        print(f"  {k:20s}: {' | '.join(row)}")
    print()
    print("=== 섹션3 최근 3개월 ===")
    for k in result['s3_keys']:
        vals = result['section3'][k]
        row = [f"{vals[i]:>10}" if vals[i] is not None else f"{'None':>10}" for i in idxs]
        print(f"  {k:20s}: {' | '.join(row)}")
