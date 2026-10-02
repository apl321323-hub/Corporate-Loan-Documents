"""
매각리스트 파서
- 시트명  : 계약리스트 (wb.active 사용)
- 헤더행  : 1행
- COL_V   = 22  대출잔액 → 월중매각액 합산
- COL_DATE= 30  계약일 'YYYY-MM-DD' → 월 기준 집계
- 반환: {"periods": ["YYYY-MM", ...], "section1": {"실상각액_월중매각액": {"YYYY-MM": val_백만, ...}}}
"""

import openpyxl
from collections import defaultdict

COL_V    = 22   # 대출잔액 (V열)
COL_DATE = 30   # 계약일


def parse_sale(file_path: str) -> dict:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active   # 시트명: 계약리스트

    monthly: dict[str, float] = defaultdict(float)

    for row in ws.iter_rows(min_row=2):
        try:
            date_val   = row[COL_DATE - 1].value   # 계약일
            balance    = row[COL_V - 1].value       # 대출잔액

            if not date_val or balance is None:
                continue

            # 날짜 → YYYY-MM
            date_str = str(date_val).strip()
            if len(date_str) < 7:
                continue
            ym = date_str[:7]   # 'YYYY-MM'
            if not ym[4] == '-':
                continue

            monthly[ym] += float(balance)
        except Exception:
            continue

    wb.close()

    # 백만 단위 변환
    result = {ym: v / 1_000_000 for ym, v in monthly.items()}
    periods = sorted(result.keys())

    return {
        "periods": periods,
        "section1": {
            "실상각액_월중매각액": result
        }
    }


# ── CLI 테스트 ─────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys, json
    path = sys.argv[1] if len(sys.argv) > 1 else \
        "/home/user/uploaded_files/매각리스트_20260619.xlsx"
    data = parse_sale(path)
    print("periods:", data["periods"])
    print("실상각액_월중매각액:")
    for ym, v in sorted(data["section1"]["실상각액_월중매각액"].items()):
        print(f"  {ym}: {v:.1f}백만")
    total = sum(data["section1"]["실상각액_월중매각액"].values())
    print(f"총합계: {total:.1f}백만")
