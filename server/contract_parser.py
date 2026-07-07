"""
계약리스트 엑셀 파서
파일: 계약리스트_YYYYMM.xlsx  (시트: '계약리스트')

핵심 열:
  B(2)  : 계약번호
  C(3)  : 상품명
  Q(17) : 계약구분  — '신규' | '추가대출' | '재대출' | '만기연장(전환)'
  AD(30): 계약일   — 'YYYY-MM-DD' 문자열
  T(20) : 대출액   — float (원 단위)
  V(22) : 대출잔액  — float (원 단위)  ← 대출채권잔액 집계용

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
    "section_balance": {                   # 대출채권잔액 (건수/잔액, V열 기반)
      "신용채권 건수":  {"2026-05": 1234, ...},
      "신용채권 잔액":  {"2026-05": 5678.9, ...},  # 백만원
      "담보채권 건수":  {...},
      "담보채권 잔액":  {...},
      "보증채권 건수":  {...},
      "보증채권 잔액":  {...},
      "전체채권 건수":  {...},
      "전체채권 잔액":  {...},
    },
    "products": ["담보론", "오투N론", ...],  # 파일 내 C열 상품명 고유값
    "source_file": "계약리스트_202606.xlsx",
  }
"""

from datetime import date, datetime
import re
import openpyxl
from collections import defaultdict


# 집계 대상 계약구분 (이 3개만 취급액에 포함)
DEAL_TYPES = {'신규', '추가대출', '재대출'}
MATURITY_EXTENSION_TYPE = '만기연장(전환)'
# 표시명 매핑 (Q열값 → 영업현황 행키)
Q_LABEL = {
    '신규':    '신규',
    '추가대출': '추가대출',
    '재대출':  '재대출',
}


def _ym(val) -> str | None:
    """Convert a date-like value to YYYY-MM."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m')
    s = str(val).strip()
    m = re.search(r'(\d{4})\D+(\d{1,2})', s)
    if m:
        month = int(m.group(2))
        if 1 <= month <= 12:
            return f"{m.group(1)}-{month:02d}"
    m = re.search(r'(?<!\d)(\d{2})\D+(\d{1,2})(?!\d)', s)
    if m:
        month = int(m.group(2))
        if 1 <= month <= 12:
            return f"20{m.group(1)}-{month:02d}"
    return None


def normalize_period_key(val) -> str | None:
    """Normalize an upload/file period value to YYYY-MM."""
    return _ym(val)


def _date_str(val) -> str | None:
    """Convert a date-like value to YYYY-MM-DD."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    m = re.search(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', s)
    if m:
        month = int(m.group(2))
        day = int(m.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"{m.group(1)}-{month:02d}-{day:02d}"
    m = re.search(r'(?<!\d)(\d{2})\D+(\d{1,2})\D+(\d{1,2})(?!\d)', s)
    if m:
        month = int(m.group(2))
        day = int(m.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"20{m.group(1)}-{month:02d}-{day:02d}"
    return None


def _norm_header(val) -> str:
    return re.sub(r'\s+', '', str(val or '')).lower()


def _find_channel_col(ws) -> int | None:
    keywords = ('\uc811\uc218\uacbd\ub85c', '\uad11\uace0\ub9e4\uccb4', '\uad11\uace0', '\ub9e4\uccb4', '\uc5d0\uc774\uc804\ud2b8', '\uc720\uc785\uacbd\ub85c', '\ucc44\ub110', 'channel', 'agent')
    reserved = {2, 3, 17, 20, 22, 26, 30}
    for col in range(1, ws.max_column + 1):
        label = _norm_header(ws.cell(row=1, column=col).value)
        if not label or col in reserved:
            continue
        if any(k.lower() in label for k in keywords):
            return col
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

    # ── 행 순회: C(3)/Q(17)/AD(30)/T(20)/V(22) ───────────────────
    # s1_raw[ym][q_label] = 합계(원)
    channel_col = _find_channel_col(ws)

    s1_raw: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    # s3_raw[group][ym][q_label] = 합계(원)
    s3_raw: dict[str, dict[str, dict[str, float]]] = \
        defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    product_raw: dict[str, dict[str, dict[str, float]]] = \
        defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    # balance_raw[group][ym] = {"건수": int, "잔액": float(원)}  ← 모든 행 집계
    balance_raw: dict = defaultdict(lambda: defaultdict(lambda: {"건수": 0, "잔액": 0.0}))
    # balance_total[ym] = {"건수": int, "잔액": float(원)}
    balance_total: dict = defaultdict(lambda: {"건수": 0, "잔액": 0.0})
    maturity_extension_raw: dict[str, float] = defaultdict(float)
    maturity_extension_group_raw: dict[str, dict[str, float]] = \
        defaultdict(lambda: defaultdict(float))
    maturity_extension_product_raw: dict[str, dict[str, float]] = \
        defaultdict(lambda: defaultdict(float))
    raw_rows: list[dict] = []

    products_set: set[str] = set()
    # 잔액 집계용 기간 Set (계약구분 필터 없이 모든 행에서 수집)
    balance_yms: set[str] = set()

    # ── 취급대출 평균이율 집계 버킷 ──────────────────────────────
    # 가중평균이율 = Σ(대출액 × 정상이율) / Σ대출액 × 100
    # {ym: Σ(대출액 × 정상이율%)}, {ym: Σ대출액}
    rate_sum_deal: dict[str, float] = defaultdict(float)  # Σ(T × Z)
    amt_sum_deal:  dict[str, float] = defaultdict(float)  # Σ T

    for row in range(2, ws.max_row + 1):
        contract_no = ws.cell(row=row, column=2).value  # B: 계약번호
        prod = ws.cell(row=row, column=3).value    # C: 상품명
        q    = ws.cell(row=row, column=17).value   # Q: 계약구분
        ad   = ws.cell(row=row, column=30).value   # AD: 계약일
        t    = ws.cell(row=row, column=20).value   # T: 대출액
        v    = ws.cell(row=row, column=22).value   # V: 대출잔액
        z    = ws.cell(row=row, column=26).value   # Z: 정상이율(%)

        ym = _ym(ad)
        if not ym:
            continue

        # ── 대출채권잔액: 계약구분 무관, 모든 행 집계 ────────────
        bal = _safe_amt(v)
        balance_yms.add(ym)
        balance_total[ym]["건수"] += 1
        balance_total[ym]["잔액"] += bal
        if prod:
            gname_bal = prod_to_group.get(str(prod).strip())
            if gname_bal:
                balance_raw[gname_bal][ym]["건수"] += 1
                balance_raw[gname_bal][ym]["잔액"] += bal

        # ── 취급액(section1/section3): DEAL_TYPES 행만 집계 ──────
        if not q:
            continue
        q = str(q).strip()
        amt = _safe_amt(t)
        contract_no_raw = str(contract_no).strip() if contract_no is not None else ''
        prod_name = str(prod).strip() if prod else ''
        gname = prod_to_group.get(prod_name) if prod_name else ''
        contract_date = _date_str(ad)
        channel = ''
        if channel_col:
            ch_val = ws.cell(row=row, column=channel_col).value
            channel = str(ch_val).strip() if ch_val is not None else ''
        if q in DEAL_TYPES or q == MATURITY_EXTENSION_TYPE:
            raw_rows.append({
                "period": ym,
                "date": contract_date or '',
                "contract_no": contract_no_raw,
                "product": prod_name,
                "group": gname or '',
                "channel": channel,
                "deal_type": q,
                "amount": amt / 1_000_000,
                "amount_won": amt,
            })
        if q == MATURITY_EXTENSION_TYPE:
            maturity_extension_raw[ym] += amt
            if gname:
                maturity_extension_group_raw[gname][ym] += amt
            if prod_name:
                maturity_extension_product_raw[prod_name][ym] += amt
        if q not in DEAL_TYPES:
            continue

        label = Q_LABEL[q]

        # ??1: ?? ??
        s1_raw[ym][label] += amt

        # ??3: ?????
        if prod_name:
            products_set.add(prod_name)
            product_raw[prod_name][ym][label] += amt
            if gname:
                s3_raw[gname][ym][label] += amt

        rate_z = _safe_amt(z)   # 정상이율(%) — None/빈값 → 0.0
        if amt > 0 and rate_z > 0:
            rate_sum_deal[ym] += amt * rate_z
            amt_sum_deal[ym]  += amt

    wb.close()

    # ── 기간 목록 ─────────────────────────────────────────────────
    # 취급액 기간 + 잔액 기간을 합쳐 정렬 (한쪽만 있는 월도 포함)
    periods = sorted(set(s1_raw.keys()) | balance_yms)

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
                v = d.get(lbl, 0) / 1_000_000
                result[lbl][ym] = v
                total += d.get(lbl, 0)
            result['취급액(대출)'][ym] = total / 1_000_000

        return result

    # ── 섹션1 ─────────────────────────────────────────────────────
    section1 = to_series(s1_raw)

    # ── 섹션3 ─────────────────────────────────────────────────────
    section3 = {}
    # product_groups 순서 유지 (그룹이 없어도 빈 그룹 포함)
    for g in (product_groups or []):
        gname = g.get('name', '')
        section3[gname] = to_series(s3_raw.get(gname, {}))

    # ── 상품별 취급액 (마감보고 전월 그룹 fallback용) ───────────────
    section_product = {}
    for pname in sorted(product_raw.keys()):
        section_product[pname] = to_series(product_raw.get(pname, {}))

    # ── 대출취급액 (section_deal) ──────────────────────────────────
    # T열(취급액) 기준, 계약구분 매핑:
    #   신규        = Q열 '신규'
    #   추가재대출  = Q열 '추가대출' + '재대출'
    #   대출취급액 합계 = 세 계약구분 전체 합
    # 단위: 원 → 백만원 (laFmt와 동일 처리)
    section_deal: dict = {}
    for ym in periods:
        d = s1_raw.get(ym, {})
        section_deal.setdefault('신규',          {})[ym] = d.get('신규', 0) / 1_000_000
        section_deal.setdefault('추가재대출',     {})[ym] = (
            (d.get('추가대출', 0) + d.get('재대출', 0)) / 1_000_000
        )
        section_deal.setdefault('대출취급액 합계', {})[ym] = (
            (d.get('신규', 0) + d.get('추가대출', 0) + d.get('재대출', 0)) / 1_000_000
        )

    maturity_extension_amount: dict[str, float] = {
        ym: maturity_extension_raw.get(ym, 0) / 1_000_000
        for ym in periods
    }

    def maturity_series(bucket: dict) -> dict:
        return {
            name: {ym: vals.get(ym, 0) / 1_000_000 for ym in periods}
            for name, vals in bucket.items()
        }

    maturity_extension_section3 = maturity_series(maturity_extension_group_raw)
    maturity_extension_product = maturity_series(maturity_extension_product_raw)

    # ── 대출채권잔액 (section_balance) ────────────────────────────
    # 잔액은 원(raw) 단위로 저장 — asset_parser 엑셀 데이터와 단위 통일
    # (laFmt에서 /1_000_000 변환, 건수는 그대로)
    GROUP_ORDER = ['신용', '담보', '보증']
    section_balance: dict = {}
    for g in GROUP_ORDER:
        section_balance[f"{g}채권 건수"] = {
            ym: balance_raw[g][ym]["건수"] for ym in periods
        }
        section_balance[f"{g}채권 잔액"] = {
            ym: balance_raw[g][ym]["잔액"] for ym in periods   # 원 단위
        }
    section_balance["전체채권 건수"] = {
        ym: balance_total[ym]["건수"] for ym in periods
    }
    section_balance["전체채권 잔액"] = {
        ym: balance_total[ym]["잔액"] for ym in periods        # 원 단위
    }

    # ── 취급대출 가중평균이율 계산 ────────────────────────────────
    # {ym: Σ(T×Z) / Σ T} — 소수 둘째자리, 집계 없으면 None
    avg_rate_deal: dict[str, float | None] = {}
    for ym in periods:
        ws_ = rate_sum_deal.get(ym, 0.0)
        wa  = amt_sum_deal.get(ym, 0.0)
        avg_rate_deal[ym] = round(ws_ / wa, 2) if wa > 0 else None

    result = {
        "periods":          periods,
        "section1":         section1,
        "section3":         section3,
        "section_deal":     section_deal,
        "section_product":  section_product,
        "section_balance":  section_balance,
        "maturity_extension_amount": maturity_extension_amount,
        "maturity_extension_section3": maturity_extension_section3,
        "maturity_extension_product": maturity_extension_product,
        "avg_rate_deal":    avg_rate_deal,   # weighted deal average rate {ym: float|None}
        "raw_rows":         raw_rows,
        "channel_column":   channel_col,
        "products":         sorted(products_set),
    }
    return result


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
    print()
    print("=== 대출취급액 ===")
    for label, series in result['section_deal'].items():
        print(f"  {label:15s}: {series}")
    print()
    print("=== 대출채권잔액 ===")
    for label, series in result['section_balance'].items():
        print(f"  {label:15s}: {series}")
