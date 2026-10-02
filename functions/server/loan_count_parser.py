"""
대출 취급건수 파서
파일: 영업분석_YYYYMMDD.xls (HTML 위장 xls)

테이블 구조 (1개 테이블):
  행1: 헤더 (구분 | 접수건수 | 거절건수 | 계약내역 | ...)
  행2: 서브헤더 (건수 | 금액 | 승인율 | 송금액 | 단가 | ...)
  행3: 신규
  행4: 재대출
  행5: 추가대출
  행6: 만기연장(전환)
  행7: 합  계

컬럼 인덱스 (행2 기준 0-indexed):
  col 0  : 구분 (신규/재대출/추가대출/만기연장/합계)
  col 1  : 접수건수 - 건수   → B열 (신청건수)
  col 3  : 계약내역 - 건수   → D열 (승인건수)

추출 항목:
  1. 신청건수(신규대출)    = 신규 행 col1
  2. 신청건수(추가재대출)  = 재대출 col1 + 추가대출 col1
  3. 승인건수(신규대출)    = 신규 행 col3
  4. 승인건수(추가재대출)  = 재대출 col3 + 추가대출 col3
"""
import re
from bs4 import BeautifulSoup


def _to_int(val: str) -> int:
    """'3,741' → 3741, 빈 문자열/비숫자 → 0"""
    cleaned = re.sub(r'[,\s]', '', str(val))
    try:
        return int(float(cleaned))
    except (ValueError, TypeError):
        return 0


def parse_loan_count(filepath: str, period: str = "") -> dict:
    """
    파일 파싱 후 결과 반환.

    Returns:
        {
          "period": "2026-05",
          "rows": {
            "신청건수(신규대출)":    3741,
            "신청건수(추가재대출)":  447,
            "승인건수(신규대출)":    303,
            "승인건수(추가재대출)":  249,
          }
        }
    """
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        html = f.read()

    # period 정규화: "2026.05" → "2026-05"
    if period and '.' in period and '-' not in period:
        period = period.replace('.', '-')

    soup = BeautifulSoup(html, 'html.parser')
    tables = soup.find_all('table')
    if not tables:
        raise ValueError("테이블을 찾을 수 없습니다.")

    tbl = tables[0]
    rows = tbl.find_all('tr')

    # 데이터 행: 행3=신규(idx2), 행4=재대출(idx3), 행5=추가대출(idx4)
    # 헤더 행(0,1) 제외하고 구분명으로 찾음
    row_map: dict[str, list[str]] = {}
    for row in rows:
        cells = [td.get_text(strip=True) for td in row.find_all(['td', 'th'])]
        if not cells:
            continue
        label = cells[0].replace('\xa0', '').strip()
        if label in ('신규', '재대출', '추가대출', '만기연장(전환)', '합  계', '합계'):
            row_map[label] = cells

    # 컬럼 위치 확인: 헤더 행2(idx1)에서 서브헤더 파싱
    # 구조: 구분(0) | 접수건수건수(1) | 접수건수금액(2) | 거절건수(3?) | 계약내역건수(?) ...
    # 실제 헤더 행1(idx0)을 보면:
    #   구분 | 접수건수 | 거절건수 | 계약내역 | 활동내역 | ...
    # 행2(idx1): 건수 | 금액 | 승인율 | 송금액 | 단가 | ...
    # → col1=접수건수(건수), col3=계약내역(건수)

    COL_APPLY  = 1   # B열: 접수건수(건수)
    COL_APPROVE = 3  # D열: 계약내역(건수)

    def get(label: str, col: int) -> int:
        cells = row_map.get(label, [])
        if col < len(cells):
            return _to_int(cells[col])
        return 0

    # 접수건수: B열(col1)
    apply_new     = get('신규', COL_APPLY)
    apply_re      = get('재대출', COL_APPLY)
    apply_add     = get('추가대출', COL_APPLY)

    # 승인건수: D열(col3) ← 헤더 구조상 계약내역 첫 번째 서브컬럼
    approve_new   = get('신규', COL_APPROVE)
    approve_re    = get('재대출', COL_APPROVE)
    approve_add   = get('추가대출', COL_APPROVE)

    return {
        "period": period,
        "rows": {
            "신청건수(신규대출)":    apply_new,
            "신청건수(추가재대출)":  apply_re + apply_add,
            "승인건수(신규대출)":    approve_new,
            "승인건수(추가재대출)":  approve_re + approve_add,
        },
        "raw": {
            "신규_접수":    apply_new,
            "재대출_접수":  apply_re,
            "추가대출_접수": apply_add,
            "신규_승인":    approve_new,
            "재대출_승인":  approve_re,
            "추가대출_승인": approve_add,
        }
    }


if __name__ == "__main__":
    import json
    result = parse_loan_count('/home/user/uploaded_files/영업분석_20260619.xls', '2026-05')
    print(json.dumps(result, ensure_ascii=False, indent=2))
