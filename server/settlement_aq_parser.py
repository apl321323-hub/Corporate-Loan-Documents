"""
결산자료 엑셀 → 자산건전성 데이터 파서
파일: 결산자료_YYYYMMDD.xlsx  (시트: 결산자료_YYYYMMDD)

사용 열:
  H열 (index 7)  : 현재상품
  J열 (index 9)  : 연체일수  (문자열로 저장됨)
  L열 (index 11) : 잔액      (숫자)
  X열 (index 23) : 만기일    ('YYYY-MM-DD' 문자열)

연체 구간 (자산건전성 기존 테이블 기준):
  무연체  : 0일
  1~30   : 1 ~ 30일
    1~10  : 1 ~ 10일   (하위)
    11~30 : 11 ~ 30일  (하위)
  31~60  : 31 ~ 60일
  61~90  : 61 ~ 90일
  91~120 : 91 ~ 120일
  121~150: 121 ~ 150일
  151~180: 151 ~ 180일
  181~   : 181일 이상
  연체합계: 1일 이상 전체
  융잔합계: 전체 (무연체 + 연체합계)

대출만기 구간 (기준일 = 업로드 월 말일):
  12개월 미만   : 기준일 ~ +12개월 미만 (이미 만기된 건 포함)
  12~24개월 미만: +12개월 이상 ~ +24개월 미만
  24~36개월 미만: +24개월 이상 ~ +36개월 미만
  36개월 이상   : +36개월 이상
  전체융잔 합계 : 전체

섹션3 담보구분별:
  product_groups.json 의 그룹 이름으로 분류
  '신용' 그룹, '보증' 그룹, '담보' 그룹

섹션4 상품별:
  H열의 각 현재상품명으로 분류
"""

from typing import Optional
import openpyxl
from datetime import date, datetime
import calendar


# ── 연체 구간 정의 ──────────────────────────────────────────────────
# (key: 표시명, lo: 최소일, hi: 최대일, parent: 상위 구간 key or None)
OVERDUE_BANDS = [
    {"key": "무연체",   "lo": 0,   "hi": 0,   "parent": None},
    {"key": "1~30",    "lo": 1,   "hi": 30,  "parent": None},
    {"key": "1~10",    "lo": 1,   "hi": 10,  "parent": "1~30"},
    {"key": "11~30",   "lo": 11,  "hi": 30,  "parent": "1~30"},
    {"key": "31~60",   "lo": 31,  "hi": 60,  "parent": None},
    {"key": "61~90",   "lo": 61,  "hi": 90,  "parent": None},
    {"key": "91~120",  "lo": 91,  "hi": 120, "parent": None},
    {"key": "121~150", "lo": 121, "hi": 150, "parent": None},
    {"key": "151~180", "lo": 151, "hi": 180, "parent": None},
    {"key": "181~",    "lo": 181, "hi": 99999,"parent": None},
]

# 섹션1/섹션4 메인 구간 (1~10, 11~30 은 하위 집계용)
S1_KEYS = ["무연체","1~30","1~10","11~30","31~60","61~90","91~120","121~150","151~180","181~","연체합계","융잔합계"]

# 섹션3 메인 구간 (자산건전성 엑셀의 섹션3 행 레이아웃과 일치)
S3_KEYS = ["무연체","1~30","31~60","61~90","91~180","181~","연체합계","융잔합계","연체율(%)"]

# ── 대출만기 구간 정의 ──────────────────────────────────────────
MATURITY_KEYS = ["12개월 미만", "12~24개월 미만", "24~36개월 미만", "36개월 이상", "전체융잔 합계"]


def _last_day_of_month(year: int, month: int) -> date:
    """해당 연월의 마지막 날 반환"""
    return date(year, month, calendar.monthrange(year, month)[1])


def _months_remaining(maturity_val, base_date: date) -> Optional[int]:
    """
    만기일(문자열 또는 datetime/date 객체) → 기준일 기준 남은 개월 수
    - 만기일 <= 기준일  → 0 (이미 만기, 12개월 미만 처리)
    - 계산: (만기년*12 + 만기월) - (기준년*12 + 기준월) + (일 보정)
    """
    if maturity_val is None:
        return None
    try:
        # datetime / date 객체 직접 처리 (엑셀 날짜 셀)
        if isinstance(maturity_val, (datetime, date)):
            mat = maturity_val.date() if isinstance(maturity_val, datetime) else maturity_val
        elif isinstance(maturity_val, str):
            s = maturity_val.strip()
            if not s:
                return None
            mat = date.fromisoformat(s[:10])
        else:
            return None
    except (ValueError, TypeError):
        return None
    if mat <= base_date:
        return 0
    # 월 단위 차이: 만기월 - 기준월 (일수는 올림 처리)
    months = (mat.year - base_date.year) * 12 + (mat.month - base_date.month)
    # 만기일의 일(day)이 기준일의 일(day)보다 크면 +1
    if mat.day > base_date.day:
        months += 1
    return max(months, 0)


def _classify_maturity(months: int) -> str:
    """남은 개월 수 → 만기 구간 키"""
    if months < 12:
        return "12개월 미만"
    elif months < 24:
        return "12~24개월 미만"
    elif months < 36:
        return "24~36개월 미만"
    else:
        return "36개월 이상"


def _to_int(v) -> Optional[int]:
    """셀 값을 정수로 변환 (실패 시 None)"""
    if v is None:
        return None
    try:
        return int(str(v).strip())
    except (ValueError, TypeError):
        return None


def _to_float(v) -> Optional[float]:
    """셀 값을 float으로 변환 (실패 시 None)"""
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _classify(days: int) -> dict:
    """연체일수 → 각 구간 key에 True/False 매핑"""
    result = {}
    for band in OVERDUE_BANDS:
        result[band["key"]] = (band["lo"] <= days <= band["hi"])
    # 집계 키
    result["연체합계"] = days >= 1
    result["융잔합계"] = True
    return result


def _empty_s1() -> dict:
    return {k: 0 for k in S1_KEYS}


def _empty_s3() -> dict:
    return {k: 0 for k in S3_KEYS}


def parse_settlement_asset_quality(filepath: str, product_groups: list) -> dict:
    """
    결산자료 엑셀을 읽어 자산건전성 구조로 파싱.

    Parameters
    ----------
    filepath      : 결산자료 xlsx 경로
    product_groups: [{"name": str, "products": [str, ...]}, ...]
                    (서버의 product_groups.json 내용)

    Returns
    -------
    {
      "period"  : "2026-05",            # 파일명 or 시트명에서 추출
      "section1": { "무연체": 1234567, ... },   # 금액 (원)
      "section3": {
          "신용": { "무연체": ..., ..., "연체율(%)": 0.12 },
          "보증": { ... },
          "담보": { ... },
      },
      "section4": {
          "담보론": { "무연체": ..., ... },
          ...
      },
      "products": ["담보론", ...],   # 데이터에 존재하는 상품 목록
    }
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # ── 기간 추출 (시트명에서) ──────────────────────────────────────
    sheet_name = ws.title  # 예: 결산자료_20260531
    period = _extract_period(sheet_name)

    # ── 기준일: 업로드 월 말일 ──────────────────────────────────────
    base_date: Optional[date] = None
    if period and len(period) == 7:  # 'YYYY-MM'
        try:
            y, m = int(period[:4]), int(period[5:7])
            base_date = _last_day_of_month(y, m)
        except (ValueError, TypeError):
            base_date = None

    # ── 상품→그룹 매핑 ───────────────────────────────────────────────
    product_to_group: dict[str, str] = {}
    for g in product_groups:
        for p in g.get("products", []):
            product_to_group[p] = g["name"]

    # 섹션1 집계 버킷
    s1: dict = _empty_s1()

    # 섹션3 집계 버킷 (그룹명 → 버킷)
    group_names = [g["name"] for g in product_groups] if product_groups else ["신용","보증","담보"]
    s3: dict[str, dict] = {gname: _empty_s3() for gname in group_names}

    # 섹션4 집계 버킷 (상품명 → 버킷)
    s4: dict[str, dict] = {}

    # 대출만기 집계 버킷 (원 단위)
    maturity_raw: dict[str, float] = {k: 0.0 for k in MATURITY_KEYS}

    # ── 평균이율 집계 버킷 ────────────────────────────────────────
    # 가중평균이율 = Σ(잔액 × 정상이율) / Σ잔액 × 100
    # 전체(대출자산평균)
    rate_sum_total: float = 0.0   # Σ(잔액 × 정상이율)
    bal_sum_total:  float = 0.0   # Σ잔액
    # 그룹별(신용/담보 등)
    rate_sum_group: dict[str, float] = {g: 0.0 for g in group_names}
    bal_sum_group:  dict[str, float] = {g: 0.0 for g in group_names}

    # ── 행 순회 ─────────────────────────────────────────────────────
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # 열 인덱스 (0-based)
        product      = str(row[7]).strip() if row[7] is not None else None  # H열
        days         = _to_int(row[9])                                       # J열
        balance      = _to_float(row[11])                                    # L열
        rate         = _to_float(row[14]) if len(row) > 14 else None         # O열: 정상이율(%)
        maturity_str = row[23] if len(row) > 23 else None                   # X열: 만기일

        if product is None or days is None or balance is None:
            continue

        clf = _classify(days)

        # ── 섹션1: 전체 합계 ────────────────────────────────────────
        for k in S1_KEYS:
            if clf.get(k):
                s1[k] += balance

        # ── 섹션3: 담보구분별 ───────────────────────────────────────
        gname = product_to_group.get(product)
        if gname and gname in s3:
            bucket = s3[gname]
            # s3 키는 S3_KEYS 기준: 91~180 묶음, 연체율(%) 는 계산
            _add_s3(bucket, days, balance)

        # ── 섹션4: 상품별 ───────────────────────────────────────────
        if product not in s4:
            s4[product] = _empty_s1()
        for k in S1_KEYS:
            if clf.get(k):
                s4[product][k] += balance

        # ── 대출만기: X열 만기일 → 기준일 기준 남은 개월 집계 ───────
        if base_date and maturity_str:
            months = _months_remaining(maturity_str, base_date)
            if months is not None:
                band = _classify_maturity(months)
                maturity_raw[band]            += balance
                maturity_raw["전체융잔 합계"] += balance

        # ── 평균이율: O열 정상이율 × L열 잔액 집계 ─────────────────
        if rate is not None and balance > 0:
            weighted = balance * rate
            # 전체 (대출자산평균)
            rate_sum_total += weighted
            bal_sum_total  += balance
            # 그룹별 (신용/담보 등)
            if gname and gname in rate_sum_group:
                rate_sum_group[gname] += weighted
                bal_sum_group[gname]  += balance

    wb.close()

    # ── 섹션3 연체율(%) 계산 ────────────────────────────────────────
    for gname, bucket in s3.items():
        total = bucket.get("융잔합계", 0)
        overdue = bucket.get("연체합계", 0)
        bucket["연체율(%)"] = round(overdue / total * 100, 2) if total > 0 else 0.0

    # ── 섹션4 융잔합계로 연체율(%) 추가 (선택) ─────────────────────
    # (섹션4는 융잔합계만 있으므로 연체율은 별도 계산)

    # ── 원 단위 → 백만원 반올림 (표시용, 백만원 단위) ──────────────
    s1_m    = {k: round(v / 1_000_000) for k, v in s1.items()}
    s3_m    = {}
    for gname, bucket in s3.items():
        s3_m[gname] = {
            k: (round(v / 1_000_000) if k != "연체율(%)" else v)
            for k, v in bucket.items()
        }
    s4_m    = {
        prod: {k: round(v / 1_000_000) for k, v in bucket.items()}
        for prod, bucket in s4.items()
    }
    # 대출만기: 원 단위 → 백만원
    maturity_m = {k: round(v / 1_000_000) for k, v in maturity_raw.items()}

    # ── 평균이율 계산 (소수 둘째자리 반올림) ────────────────────────
    def _wavg(wsum: float, bsum: float) -> Optional[float]:
        """가중평균이율 계산. 잔액 합계가 0이면 None"""
        if bsum <= 0:
            return None
        return round(wsum / bsum, 2)

    avg_rate = {
        "대출자산평균": _wavg(rate_sum_total, bal_sum_total),
        **{gname: _wavg(rate_sum_group[gname], bal_sum_group[gname]) for gname in group_names},
    }

    products_list = sorted(s4_m.keys())

    return {
        "period"   : period,
        "section1" : s1_m,
        "section3" : s3_m,
        "section4" : s4_m,
        "maturity" : maturity_m,
        "avg_rate" : avg_rate,   # 평균이율 집계 결과
        "products" : products_list,
    }


# ── 섹션3 전용 구간 집계 헬퍼 ─────────────────────────────────────
# S3_KEYS = ["무연체","1~30","31~60","61~90","91~180","181~","연체합계","융잔합계","연체율(%)"]
def _add_s3(bucket: dict, days: int, balance: float):
    if days == 0:
        bucket["무연체"]   += balance
    elif 1 <= days <= 30:
        bucket["1~30"]    += balance
        bucket["연체합계"] += balance
    elif 31 <= days <= 60:
        bucket["31~60"]   += balance
        bucket["연체합계"] += balance
    elif 61 <= days <= 90:
        bucket["61~90"]   += balance
        bucket["연체합계"] += balance
    elif 91 <= days <= 180:
        bucket["91~180"]  += balance
        bucket["연체합계"] += balance
    elif days >= 181:
        bucket["181~"]    += balance
        bucket["연체합계"] += balance
    bucket["융잔합계"] += balance


# ── 기간 문자열 추출 ───────────────────────────────────────────────
def _extract_period(sheet_name: str) -> str:
    """
    시트명 예: '결산자료_20260531' → '2026-05'
    실패 시 빈 문자열 반환
    """
    import re
    m = re.search(r'(\d{4})(\d{2})\d{2}', sheet_name)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    # YYYYMM 형태
    m = re.search(r'(\d{4})(\d{2})', sheet_name)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return ""


# ── CLI 테스트 ─────────────────────────────────────────────────────
if __name__ == "__main__":
    import json, sys
    filepath = sys.argv[1] if len(sys.argv) > 1 else "/home/user/uploaded_files2/결산자료_5월.xlsx"
    
    # 테스트용 product_groups
    test_groups = [
        {"name": "신용", "products": [
            "N론","V론","레이디론","스타론","스타스위치론","우량론","큐브론","테일론","토탈론",
            "프리론","프리미엄론","전월세론","플러스론","토마토론","토마토N론","토마토토탈론",
            "T플러스론","OP론","충성론","오투론","오투N론","토마토토탈론플러스",
            "다이렉트론(A)","다이렉트론(W)","N론(하이브리드)","기타","기타N"
        ]},
        {"name": "담보", "products": ["담보론","담보론(지분대출)"]}
    ]
    
    result = parse_settlement_asset_quality(filepath, test_groups)
    print(f"기간: {result['period']}")
    print(f"상품 수: {len(result['products'])} 개")
    print(f"\n=== 섹션1 (금액, 백만원) ===")
    for k, v in result['section1'].items():
        print(f"  {k:12}: {v:>12,}")
    print(f"\n=== 섹션3 (담보구분별, 백만원) ===")
    for gname, bucket in result['section3'].items():
        print(f"  [{gname}]")
        for k, v in bucket.items():
            print(f"    {k:12}: {v:>12}")
    print(f"\n=== 섹션4 상품별 (일부) ===")
    for prod in list(result['products'])[:5]:
        print(f"  [{prod}]  융잔합계={result['section4'][prod]['융잔합계']:,}")
