"""
계약리스트 엑셀 파서
파일: 계약리스트_YYYYMM.xlsx  (시트: '계약리스트')

핵심 열:
  C(3)  : 상품명
  Q(17) : 계약구분  — '신규' | '추가대출' | '재대출' | '만기연장(전환)'
  AD(30): 계약일   — 'YYYY-MM-DD' 문자열
  T(20) : 대출액   — float (원 단위)

집계 결과:
  {
    "periods": ["2026-05", ...],          # 파일에 존재하는 월 목록 (정렬)
    "section1": {                          # 전체 영업현황 (취급액 관련)
      "신규":        {"2026-05": 1234.5, ...},   # 백만원
      "추가대출":    {...},
      "재대출":      {...},
      "취급액(대출)": {...}                # 신규+추가대출+재대출 합
    },
    "section3": {                          # 담보구분별
      "<그룹명>": {                        # product_groups 기반 (신용/담보 등)
        "신규":        {...},
        "추가대출":    {...},
        "재대출":      {...},
        "취급액(대출)": {...}
      },
      ...
    },
    "products": ["담보론", "오투N론", ...],  # 파일 내 C열 상품명 고유값
    "source_file": "계약리스트_202606.xlsx",
  }
"""

from datetime import datetime
import openpyxl
from collections import defaultdict


# 집계 대상 계약구분 (이 3개만 취급액에 포함)
DEAL_TYPES = {'신규', '추가대출', '재대출'}
# 표시명 매핑 (Q열값 → 영업현황 행키)
Q_LABEL = {
    '신규':    '신규',
    '추가대출': '추가대출',
    '재대출':  '재대출',
}


def _ym(val) -> str | None:
    """계약일 → 'YYYY-MM' 변환"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime('%Y-%m')
    s = str(val).strip()
    if len(s) >= 7:
        return s[:7]
    return None


def _safe_amt(v) -> float:
    """대출액 → float (원 단위)"""
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def parse_contract_list(filepath: str, product_groups: list) -> dict:
    """
    계약리스트 엑셀 파싱

    Parameters
    ----------
    filepath : str
        엑셀 파일 경로
    product_groups : list
        [{"name": "신용", "products": [...]}, {"name": "담보", "products": [...]}]
        — section3 그룹 분류 기준

    Returns
    -------
    dict  (JSON 직렬화 가능)
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # 시트 탐색: '계약리스트' 우선, 없으면 첫 번째
    ws = None
    for sname in wb.sheetnames:
        if '계약리스트' in sname:
            ws = wb[sname]
            break
    if ws is None:
        ws = wb.active

    # ── 상품→그룹 매핑 dict ────────────────────────────────────────
    prod_to_group: dict[str, str] = {}
    for g in (product_groups or []):
        gname = g.get('name', '')
        for p in g.get('products', []):
            prod_to_group[p] = gname

    # ── 행 순회: C(3)/Q(17)/AD(30)/T(20) ─────────────────────────
    # s1_raw[ym][q_label] = 합계(원)
    s1_raw: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    # s3_raw[group][ym][q_label] = 합계(원)
    s3_raw: dict[str, dict[str, dict[str, float]]] = \
        defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    products_set: set[str] = set()

    for row in range(2, ws.max_row + 1):
        prod = ws.cell(row=row, column=3).value    # C: 상품명
        q    = ws.cell(row=row, column=17).value   # Q: 계약구분
        ad   = ws.cell(row=row, column=30).value   # AD: 계약일
        t    = ws.cell(row=row, column=20).value   # T: 대출액

        if not q:
            continue
        q = str(q).strip()
        if q not in DEAL_TYPES:
            continue

        ym = _ym(ad)
        if not ym:
            continue

        amt = _safe_amt(t)
        label = Q_LABEL[q]

        # 섹션1: 전체 합산
        s1_raw[ym][label] += amt

        # 섹션3: 상품그룹별
        if prod:
            products_set.add(str(prod))
            gname = prod_to_group.get(str(prod).strip())
            if gname:
                s3_raw[gname][ym][label] += amt

    wb.close()

    # ── 기간 목록 ─────────────────────────────────────────────────
    periods = sorted(s1_raw.keys())

    # ── dict → 백만원 시계열 변환 헬퍼 ──────────────────────────
    def to_series(ym_label_dict: dict) -> dict:
        """
        {ym: {label: float}} → {label: {ym: val_백만원}}
        + '취급액(대출)' 자동 계산
        """
        labels = ['신규', '추가대출', '재대출']
        result = {lbl: {} for lbl in labels}
        result['취급액(대출)'] = {}

        for ym in periods:
            d = ym_label_dict.get(ym, {})
            total = 0.0
            for lbl in labels:
                v = round(d.get(lbl, 0) / 1_000_000, 1)
                result[lbl][ym] = v
                total += d.get(lbl, 0)
            result['취급액(대출)'][ym] = round(total / 1_000_000, 1)

        return result

    # ── 섹션1 ─────────────────────────────────────────────────────
    section1 = to_series(s1_raw)

    # ── 섹션3 ─────────────────────────────────────────────────────
    section3 = {}
    # product_groups 순서 유지 (그룹이 없어도 빈 그룹 포함)
    for g in (product_groups or []):
        gname = g.get('name', '')
        section3[gname] = to_series(s3_raw.get(gname, {}))

    return {
        "periods":     periods,
        "section1":    section1,
        "section3":    section3,
        "products":    sorted(products_set),
    }


# ── CLI 테스트 ─────────────────────────────────────────────────────
if __name__ == "__main__":
    import json, sys

    fpath  = sys.argv[1] if len(sys.argv) > 1 else "/tmp/contract_list.xlsx"
    groups = [
        {"name": "신용", "products": [
            "N론","V론","레이디론","스타론","스타스위치론","우량론","큐브론",
            "테일론","토탈론","프리론","프리미엄론","전월세론","플러스론",
            "토마토론","토마토N론","토마토토탈론","T플러스론","OP론","충성론",
            "오투론","오투N론","토마토토탈론플러스","다이렉트론(A)","다이렉트론(W)",
            "N론(하이브리드)","기타","기타N"
        ]},
        {"name": "담보", "products": ["담보론","담보론(지분대출)"]},
    ]

    result = parse_contract_list(fpath, groups)
    print(f"기간: {result['periods']}")
    print(f"상품: {result['products']}")
    print()
    print("=== 섹션1 전체 영업현황 ===")
    for label, series in result['section1'].items():
        print(f"  {label:15s}: {series}")
    print()
    print("=== 섹션3 담보구분별 ===")
    for gname, gs in result['section3'].items():
        print(f"  [{gname}]")
        for label, series in gs.items():
            print(f"    {label:15s}: {series}")
