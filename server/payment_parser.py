"""
입금명세 엑셀 파서
파일: 입금명세_YYYYMMDD.xlsx  (시트: '입금명세')

구조:
  1행: 헤더
  2행~: 실제 데이터

집계 조건:
  K(11)열(회계) = '+' 인 행만
  AY(51)열(계약구분) = '신규' | '추가대출' | '재대출' 인 행만

집계 대상:
  N(14)열 원금입금  → 원금회수액 (원 단위 합산 → 백만원)
  Q(17)열 이자입금  → 이자회수액 (원 단위 합산 → 백만원)

기간 기준: E(5)열 입금일 'YYYY-MM-DD' → 'YYYY-MM'

집계 결과:
  {
    "periods":  ["2026-05", ...],
    "section1": {
      "원금회수액": {"2026-05": 5056.4, ...},  # 백만원
      "이자회수액": {"2026-05": 1393.5, ...},  # 백만원
    },
    "source_file": "입금명세_20260531.xlsx",
  }
"""

import os
from datetime import datetime, date
import openpyxl
from collections import defaultdict

# 집계 대상 계약구분
DEAL_TYPES = {'신규', '추가대출', '재대출'}

# 입금명세 시트 고정 열 번호
COL_K   = 11   # 회계
COL_E   =  5   # 입금일
COL_AY  = 51   # 계약구분
COL_N   = 14   # 원금입금
COL_Q   = 17   # 이자입금


def _ym(val) -> str | None:
    """입금일 → 'YYYY-MM'"""
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


def parse_payment(filepath: str) -> dict:
    """
    입금명세 엑셀 파싱

    Parameters
    ----------
    filepath : str  엑셀 파일 경로

    Returns
    -------
    dict (JSON 직렬화 가능)
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    source_file = os.path.basename(filepath)

    # '입금명세' 시트 선택 (숫자 없는 기본 시트)
    ws = None
    for sname in wb.sheetnames:
        if sname.strip() == '입금명세':
            ws = wb[sname]
            break
    if ws is None:
        # fallback: '입금명세'가 포함된 첫 번째 시트
        for sname in wb.sheetnames:
            if '입금명세' in sname:
                ws = wb[sname]
                break
    if ws is None:
        ws = wb.active

    # ── 행 순회 (1행 헤더 → 2행부터 데이터) ──────────────────────
    raw: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))

    for row in range(2, ws.max_row + 1):
        k   = ws.cell(row=row, column=COL_K).value   # 회계
        e   = ws.cell(row=row, column=COL_E).value   # 입금일
        ay  = ws.cell(row=row, column=COL_AY).value  # 계약구분
        n   = ws.cell(row=row, column=COL_N).value   # 원금입금
        q   = ws.cell(row=row, column=COL_Q).value   # 이자입금

        # K = '+' 필터
        if str(k or '').strip() != '+':
            continue

        # 계약구분 필터
        if str(ay or '').strip() not in DEAL_TYPES:
            continue

        ym = _ym(e)
        if not ym:
            continue

        raw[ym]['원금회수액'] += _safe_amt(n)
        raw[ym]['이자회수액'] += _safe_amt(q)

    wb.close()

    # ── 기간 목록 ─────────────────────────────────────────────────
    periods = sorted(raw.keys())

    # ── 백만원 변환 ───────────────────────────────────────────────
    section1 = {
        '원금회수액': {},
        '이자회수액': {},
    }
    for ym in periods:
        d = raw[ym]
        section1['원금회수액'][ym] = round(d['원금회수액'] / 1_000_000, 1)
        section1['이자회수액'][ym] = round(d['이자회수액'] / 1_000_000, 1)

    return {
        "periods":     periods,
        "section1":    section1,
        "source_file": source_file,
    }


# ── CLI 테스트 ──────────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    fpath = sys.argv[1] if len(sys.argv) > 1 else "/tmp/입금명세_20260531.xlsx"
    result = parse_payment(fpath)

    print(f"파일: {result['source_file']}")
    print(f"기간: {result['periods']}")
    print()
    print("=== 섹션1 원금/이자 회수액 ===")
    for label, series in result['section1'].items():
        print(f"  {label}: {series}")
