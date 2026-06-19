"""
대손리스트 엑셀 파서
파일: 대손리스트_YYYYMM.xlsx  (시트: 첫 번째 시트, 보통 '계약리스트')

구조:
  1행: 헤더
  2행~: 실제 데이터

집계 대상:
  V(22)열  대출잔액  → 월중상각액 합산 (원 단위 → 백만원)

월 기준:
  AZ(52)열 대손일 'YYYY-MM-DD' → 'YYYY-MM'
  (비어있으면 해당 행 제외)

집계 결과:
  {
    "periods":  ["2026-03", ...],
    "section1": {
      "실상각액_월중상각액": {"2026-03": 11.1, ...},  # 백만원
    },
    "source_file": "대손리스트_202603.xlsx",
  }
"""

import os
from datetime import datetime, date
import openpyxl
from collections import defaultdict

# 고정 열 번호 (1행 헤더 기준)
COL_V   = 22   # 대출잔액
COL_AZ  = 52   # 대손일


def _ym(val) -> str | None:
    """대손일 → 'YYYY-MM'"""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m')
    s = str(val).strip()
    if len(s) >= 7:
        return s[:7]
    return None


def _safe_amt(v) -> float:
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def parse_writeoff(filepath: str) -> dict:
    """
    대손리스트 엑셀 파싱

    Parameters
    ----------
    filepath : str  엑셀 파일 경로

    Returns
    -------
    dict (JSON 직렬화 가능)
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    source_file = os.path.basename(filepath)

    # 첫 번째 시트 사용
    ws = wb.active

    # ── 행 순회 (1행 헤더 → 2행부터 데이터) ──────────────────────
    raw: dict[str, float] = defaultdict(float)  # {ym: 합계(원)}

    for row in range(2, ws.max_row + 1):
        az  = ws.cell(row=row, column=COL_AZ).value  # 대손일
        v   = ws.cell(row=row, column=COL_V).value   # 대출잔액

        ym = _ym(az)
        if not ym:
            continue

        raw[ym] += _safe_amt(v)

    wb.close()

    # ── 기간 목록 ─────────────────────────────────────────────────
    periods = sorted(raw.keys())

    # ── 백만원 변환 ───────────────────────────────────────────────
    series = {ym: round(raw[ym] / 1_000_000, 1) for ym in periods}

    return {
        "periods": periods,
        "section1": {
            "실상각액_월중상각액": series,
        },
        "source_file": source_file,
    }


# ── CLI 테스트 ──────────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    fpath = sys.argv[1] if len(sys.argv) > 1 else "/tmp/대손리스트_202603.xlsx"
    result = parse_writeoff(fpath)

    print(f"파일: {result['source_file']}")
    print(f"기간: {result['periods']}")
    print()
    print("=== 섹션1 월중상각액 ===")
    for label, series in result['section1'].items():
        print(f"  {label}: {series}")
