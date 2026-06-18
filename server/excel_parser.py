import openpyxl
from datetime import datetime
from typing import Any, Dict, List, Optional
import json


def safe_value(val):
    """셀 값을 JSON 직렬화 가능한 형태로 변환"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime('%Y-%m')
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, bool):
        return val
    return str(val)


def parse_company_info(ws) -> Dict:
    """기업정보 시트 파싱 - 엑셀 레이아웃과 동일하게"""
    data = {
        "company": {},
        "shareholders": [],
        "executives": [],
        "ceo_history": [],
        "company_history": [],
        "guarantor": {},
        "guarantor_shareholders": [],
        "guarantor_executives": [],
        "guarantor_history": [],
        "business_direction": "",
        "affiliates": ""
    }

    # 전체 셀을 좌표 기준으로 읽기
    cells = {}
    for row in ws.iter_rows():
        for cell in row:
            try:
                if cell.value is not None:
                    cells[cell.coordinate] = cell.value
            except AttributeError:
                continue

    def g(coord):
        return cells.get(coord, "")

    def sv(v):
        return safe_value(v) if v is not None else ""

    # ■ 차주사 기본정보 (row 4~7)
    data["company"] = {
        "name":            str(g("B4") or ""),
        "representative":  str(g("D4") or ""),
        "actual_manager":  str(g("F4") or ""),
        "business_number": str(g("B5") or ""),
        "industry":        str(g("D5") or ""),
        "company_type":    str(g("F5") or ""),
        "phone":           str(g("B6") or ""),
        "rep_phone":       str(g("D6") or ""),
        "listed":          str(g("F6") or ""),
        "employees":       str(g("B7") or ""),
        "address":         str(g("D7") or ""),
    }

    # 주주현황 (row 9~14, A8:A14 병합)
    shareholders = []
    for row_num in range(9, 15):
        name = g(f"B{row_num}")
        shares = g(f"C{row_num}")
        ratio = g(f"D{row_num}")
        amount = g(f"E{row_num}")
        relation = g(f"F{row_num}")
        if name:
            shareholders.append({
                "name": str(name),
                "shares": sv(shares),
                "ratio": sv(ratio),
                "amount": sv(amount),
                "relation": str(relation or "")
            })
    data["shareholders"] = shareholders

    # 등기임원 (row 15)
    executives = []
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
        v = g(f"{col}15")
        if v:
            executives.append(str(v))
    data["executives"] = executives

    # 대표이사(실경영자) 이력 (row 17~24, A16:A24 병합)
    ceo_history = []
    for row_num in range(17, 25):
        date_val = g(f"B{row_num}")
        content = g(f"C{row_num}")
        if date_val or content:
            # float인 경우 연도.월 형태로 변환 (예: 2004.02)
            if isinstance(date_val, float):
                year = int(date_val)
                month = round((date_val - year) * 100)
                date_str = f"{year}.{month:02d}" if month > 0 else str(year)
            else:
                date_str = str(safe_value(date_val) or "")
            ceo_history.append({
                "date": date_str,
                "content": str(content or "")
            })
    data["ceo_history"] = ceo_history

    # 연혁 (row 26~40, A25:A40 병합)
    company_history = []
    for row_num in range(26, 41):
        date_val = g(f"B{row_num}")
        content = g(f"C{row_num}")
        if date_val or content:
            company_history.append({
                "date": str(safe_value(date_val) or ""),
                "content": str(content or "")
            })
    data["company_history"] = company_history

    # ■ 보증(법)인 기본정보 (row 43~46)
    # A43=보증(법)인명 라벨, B43=이름값, C43=관계(대표이사) 라벨+값, E43=실경영자 라벨
    # A44=사업자번호 라벨, B44=값, C44=업종 라벨, D44=값, E44=기업형태 라벨
    # A45=전화번호 라벨, B45=값, C45=대표이사전화번호 라벨, E45=상장 라벨
    # A46=임직원수 라벨, C46=사업장주소 라벨
    data["guarantor"] = {
        "name":            str(g("B43") or ""),
        "representative":  str(g("C43") or ""),   # 관계 (대표이사 등)
        "actual_manager":  str(g("F43") or ""),   # 실경영자 값 (F43)
        "business_number": str(g("B44") or ""),
        "industry":        str(g("D44") or ""),
        "company_type":    str(g("F44") or ""),
        "phone":           str(g("B45") or ""),
        "rep_phone":       str(g("D45") or ""),
        "listed":          str(g("F45") or ""),
        "employees":       str(g("B46") or ""),
        "address":         str(g("D46") or ""),
    }

    # 보증인 주주현황 (row 48~53, A47:A53 병합)
    gu_shareholders = []
    for row_num in range(48, 54):
        name = g(f"B{row_num}")
        shares = g(f"C{row_num}")
        ratio = g(f"D{row_num}")
        amount = g(f"E{row_num}")
        relation = g(f"F{row_num}")
        if name:
            gu_shareholders.append({
                "name": str(name),
                "shares": sv(shares),
                "ratio": sv(ratio),
                "amount": sv(amount),
                "relation": str(relation or "")
            })
    data["guarantor_shareholders"] = gu_shareholders

    # 보증인 등기임원 (row 54)
    gu_executives = []
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
        v = g(f"{col}54")
        if v:
            gu_executives.append(str(v))
    data["guarantor_executives"] = gu_executives

    # 보증인 연혁(경력) (row 56~63, A55:A63 병합)
    gu_history = []
    for row_num in range(56, 64):
        date_val = g(f"B{row_num}")
        content = g(f"C{row_num}")
        if date_val or content:
            if isinstance(date_val, float):
                year = int(date_val)
                month = round((date_val - year) * 100)
                date_str = f"{year}.{month:02d}" if month > 0 else str(year)
            else:
                date_str = str(safe_value(date_val) or "")
            gu_history.append({
                "date": date_str,
                "content": str(content or "")
            })
    data["guarantor_history"] = gu_history

    # ■ 기타 (row 66~68)
    data["business_direction"] = str(g("B67") or "")
    data["affiliates"] = str(g("B68") or "")

    return data


def parse_bs_is(wb) -> Dict:
    """BS/IS 시트 파싱 - 최근 데이터 위주"""
    result = {"bs": {}, "is_data": {}, "summary": {}}

    # BSPL 요약 시트
    if "BSPL_요약" in wb.sheetnames:
        ws = wb["BSPL_요약"]
        headers = []
        row_headers = []
        data_matrix = {}

        for row in ws.iter_rows(min_row=3, max_row=50):
            row_label = None
            for cell in row:
                if cell.column_letter == "A" and cell.value:
                    row_label = str(cell.value)
                elif cell.value is not None and row_label:
                    col = cell.column_letter
                    if col not in data_matrix:
                        data_matrix[col] = {}
                    data_matrix[col][row_label] = safe_value(cell.value)

        # 날짜 헤더 추출
        date_row = list(ws.iter_rows(min_row=3, max_row=3))[0]
        dates = {}
        for cell in date_row:
            if cell.value and isinstance(cell.value, datetime):
                dates[cell.column_letter] = cell.value.strftime('%Y-%m')

        result["summary"]["dates"] = dates
        result["summary"]["data"] = data_matrix

    # BS 시트
    if "1. BS" in wb.sheetnames:
        ws = wb["1. BS"]
        bs_items = {}
        date_cols = {}

        date_row = list(ws.iter_rows(min_row=4, max_row=4))[0]
        for cell in date_row:
            if cell.value and isinstance(cell.value, datetime):
                date_cols[cell.column_letter] = cell.value.strftime('%Y-%m')

        for row in ws.iter_rows(min_row=5, max_row=32):
            row_label = None
            for cell in row:
                if cell.column_letter == "A" and cell.value:
                    row_label = str(cell.value)
                    bs_items[row_label] = {}
                elif cell.value is not None and row_label and cell.column_letter in date_cols:
                    bs_items[row_label][date_cols[cell.column_letter]] = safe_value(cell.value)

        result["bs"] = {
            "dates": date_cols,
            "items": bs_items
        }

    # IS 시트
    if "2. IS" in wb.sheetnames:
        ws = wb["2. IS"]
        is_items = {}
        date_cols = {}

        date_row = list(ws.iter_rows(min_row=4, max_row=4))[0]
        for cell in date_row:
            if cell.value and isinstance(cell.value, datetime):
                date_cols[cell.column_letter] = cell.value.strftime('%Y-%m')

        for row in ws.iter_rows(min_row=5, max_row=50):
            row_label = None
            for cell in row:
                if cell.column_letter == "A" and cell.value:
                    row_label = str(cell.value)
                    is_items[row_label] = {}
                elif cell.value is not None and row_label and cell.column_letter in date_cols:
                    is_items[row_label][date_cols[cell.column_letter]] = safe_value(cell.value)

        result["is_data"] = {
            "dates": date_cols,
            "items": is_items
        }

    return result


def parse_cashflow(ws) -> Dict:
    """현금흐름 시트 파싱"""
    result = {"dates": {}, "items": {}}
    date_cols = {}

    # 헤더행 찾기 (날짜가 있는 행) - 1개 이상이면 OK
    header_row_idx = 4
    for r_idx in range(1, 10):
        r = list(ws.iter_rows(min_row=r_idx, max_row=r_idx))[0]
        count = sum(1 for c in r if isinstance(c.value, datetime))
        if count >= 1:
            header_row_idx = r_idx
            for cell in r:
                if isinstance(cell.value, datetime):
                    date_cols[cell.column_letter] = cell.value.strftime('%Y-%m')
            break

    result["dates"] = date_cols

    items = {}
    current_label = None
    for row in ws.iter_rows(min_row=header_row_idx + 1, max_row=ws.max_row):
        for cell in row:
            try:
                col = cell.column_letter
            except AttributeError:
                continue
            val = cell.value
            if col == "A" and val:
                lbl = str(val).strip()
                current_label = lbl
                if lbl not in items:
                    items[lbl] = {}
            elif val is not None and current_label and col in date_cols:
                items[current_label][date_cols[col]] = safe_value(val)

    result["items"] = items
    return result


def parse_asset_quality(ws) -> Dict:
    """자산건전성 시트 파싱"""
    result = {"dates": {}, "items": {}}
    date_cols = {}
    data_start = 4

    # 헤더행 자동 탐색
    for r_idx in range(1, 10):
        r = list(ws.iter_rows(min_row=r_idx, max_row=r_idx))[0]
        count = sum(1 for c in r if isinstance(c.value, datetime))
        if count > 5:
            for cell in r:
                if isinstance(cell.value, datetime):
                    date_cols[cell.column_letter] = cell.value.strftime('%Y-%m')
            data_start = r_idx + 1
            break

    result["dates"] = date_cols

    items = {}
    for row in ws.iter_rows(min_row=data_start, max_row=min(30, ws.max_row)):
        row_label = None
        for cell in row:
            try:
                col = cell.column_letter
            except AttributeError:
                continue
            val = cell.value
            if col == "A" and val:
                row_label = str(val).strip()
                items[row_label] = {}
            elif val is not None and row_label and col in date_cols:
                items[row_label][date_cols[col]] = safe_value(val)

    result["items"] = items
    return result


def parse_business_status(ws) -> Dict:
    """영업현황 시트 파싱"""
    result = {"dates": {}, "items": {}}
    date_cols = {}
    data_start = 4

    # 헤더 행 찾기 (날짜가 많은 행)
    for header_row_idx in range(1, 10):
        header_row = list(ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx))[0]
        count = sum(1 for c in header_row if isinstance(c.value, datetime))
        if count >= 1:
            for cell in header_row:
                try:
                    if isinstance(cell.value, datetime):
                        date_cols[cell.column_letter] = cell.value.strftime('%Y-%m')
                except AttributeError:
                    continue
            data_start = header_row_idx + 1
            break

    result["dates"] = date_cols

    items = {}
    for row in ws.iter_rows(min_row=data_start, max_row=min(100, ws.max_row)):
        row_label = None
        for cell in row:
            try:
                col = cell.column_letter
            except AttributeError:
                continue
            val = cell.value
            if col == "A" and val:
                row_label = str(val).strip()
                items[row_label] = {}
            elif val is not None and row_label and col in date_cols:
                items[row_label][date_cols[col]] = safe_value(val)

    result["items"] = items
    return result


def parse_loan_asset(ws) -> Dict:
    """대출자산속성 시트 파싱"""
    result = {"dates": {}, "items": {}}
    date_cols = {}
    data_start = 4

    for header_row_idx in range(1, 10):
        header_row = list(ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx))[0]
        count = sum(1 for c in header_row if isinstance(c.value, datetime))
        if count >= 1:
            for cell in header_row:
                try:
                    if isinstance(cell.value, datetime):
                        date_cols[cell.column_letter] = cell.value.strftime('%Y-%m')
                except AttributeError:
                    continue
            data_start = header_row_idx + 1
            break

    result["dates"] = date_cols

    items = {}
    for row in ws.iter_rows(min_row=data_start, max_row=min(100, ws.max_row)):
        row_label = None
        for cell in row:
            try:
                col = cell.column_letter
            except AttributeError:
                continue
            val = cell.value
            if col == "A" and val:
                row_label = str(val).strip()
                items[row_label] = {}
            elif val is not None and row_label and col in date_cols:
                items[row_label][date_cols[col]] = safe_value(val)

    result["items"] = items
    return result


def parse_borrowing(ws) -> Dict:
    """차입현황 시트 파싱"""
    result = {"headers": [], "rows": [], "raw": {}}

    rows_data = []
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200)):
        row_vals = []
        has_data = False
        for cell in row:
            val = safe_value(cell.value)
            row_vals.append(val)
            if val is not None:
                has_data = True
        if has_data:
            rows_data.append(row_vals)

    result["raw"] = rows_data
    return result


def parse_settlement_asset_quality(file_path: str) -> Dict:
    """
    결산자료 엑셀 파싱 → 자산건전성 연체 버킷 집계
    H열=현재상품, J열=연체일수, L열=잔액
    반환: { total, by_product, by_group, products, groups }
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active  # 결산자료_20260531

    # 버킷 정의
    BUCKETS = ['무연체', '1~10', '11~30', '31~60', '61~90', '91~120', '121~150', '151~180', '181~', '연체합계', '융잔합계']

    def bucket_of(days: int) -> List[str]:
        """연체일수 → 해당 버킷 목록"""
        buckets = []
        if days == 0:
            buckets.append('무연체')
        elif 1 <= days <= 10:
            buckets.append('1~10')
        elif 11 <= days <= 30:
            buckets.append('11~30')
        elif 31 <= days <= 60:
            buckets.append('31~60')
        elif 61 <= days <= 90:
            buckets.append('61~90')
        elif 91 <= days <= 120:
            buckets.append('91~120')
        elif 121 <= days <= 150:
            buckets.append('121~150')
        elif 151 <= days <= 180:
            buckets.append('151~180')
        else:  # >= 181
            buckets.append('181~')
        if days >= 1:
            buckets.append('연체합계')
        buckets.append('융잔합계')
        return buckets

    def empty_buckets() -> Dict:
        return {b: {'count': 0, 'balance': 0.0} for b in BUCKETS}

    # 상품 → 그룹 매핑 (업종 분류)
    PRODUCT_GROUPS = {
        '담보론': '담보',
        '담보론(지분대출)': '담보',
        '전월세론': '담보',
        'N론': '신용',
        'N론(하이브리드)': '신용',
        'OP론': '신용',
        '오투N론': '신용',
        '오투론': '신용',
        '토마토N론': '신용',
        '기타N': '신용',
        '토마토론': '신용',
        '토마토토탈론': '토탈',
        '토마토토탈론플러스': '토탈',
        '토탈론': '토탈',
        '스타론': '스타',
        '스타스위치론': '스타',
        '우량론': '우량/기타',
        '기타': '우량/기타',
        '플러스론': '우량/기타',
        '프리론': '우량/기타',
        '프리미엄론': '우량/기타',
        '큐브론': '우량/기타',
        '테일론': '우량/기타',
        '레이디론': '우량/기타',
        '다이렉트론(A)': '다이렉트',
        '다이렉트론(W)': '다이렉트',
        'T플러스론': '우량/기타',
    }

    total = empty_buckets()
    by_product: Dict[str, Dict] = {}
    by_group: Dict[str, Dict] = {}

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        product = row[7]   # H열
        overdue = row[9]   # J열
        balance = row[11]  # L열

        if product is None or balance is None:
            continue

        try:
            bal = float(balance)
            ov = int(overdue) if overdue is not None else 0
        except (ValueError, TypeError):
            continue

        product = str(product).strip()
        group = PRODUCT_GROUPS.get(product, '기타')

        # 상품별 초기화
        if product not in by_product:
            by_product[product] = empty_buckets()
        if group not in by_group:
            by_group[group] = empty_buckets()

        hit_buckets = bucket_of(ov)
        for bk in hit_buckets:
            total[bk]['count'] += 1
            total[bk]['balance'] += bal
            by_product[product][bk]['count'] += 1
            by_product[product][bk]['balance'] += bal
            by_group[group][bk]['count'] += 1
            by_group[group][bk]['balance'] += bal

    # 연체율 계산 헬퍼
    def add_rate(d: Dict) -> Dict:
        result = {}
        for bk, v in d.items():
            total_bal = d['융잔합계']['balance']
            rate = (v['balance'] / total_bal * 100) if total_bal > 0 else 0.0
            result[bk] = {
                'count': v['count'],
                'balance': round(v['balance']),
                'rate': round(rate, 2)
            }
        return result

    return {
        'buckets': BUCKETS,
        'total': add_rate(total),
        'by_product': {p: add_rate(v) for p, v in sorted(by_product.items())},
        'by_group': {g: add_rate(v) for g, v in sorted(by_group.items())},
        'products': sorted(by_product.keys()),
        'groups': sorted(by_group.keys()),
    }


def parse_excel_file(file_path: str) -> Dict:
    """전체 엑셀 파일 파싱"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    result = {}

    sheet_map = {
        "기업정보": "company_info",
        "BSPL_요약": "bspl_summary",
        "1. BS": "bs",
        "2. IS": "is",
        "현금흐름": "cashflow",
        "자산건전성 (APL기준)": "asset_quality",
        "영업현황 (APL기준)": "business_status",
        "대출자산속성 (APL기준)": "loan_asset",
        "차입현황 26.05": "borrowing"
    }

    try:
        if "기업정보" in wb.sheetnames:
            result["company_info"] = parse_company_info(wb["기업정보"])
    except Exception as e:
        result["company_info"] = {"error": str(e)}

    try:
        result["bs_is"] = parse_bs_is(wb)
    except Exception as e:
        result["bs_is"] = {"error": str(e)}

    try:
        if "현금흐름" in wb.sheetnames:
            result["cashflow"] = parse_cashflow(wb["현금흐름"])
    except Exception as e:
        result["cashflow"] = {"error": str(e)}

    try:
        if "자산건전성 (APL기준)" in wb.sheetnames:
            result["asset_quality"] = parse_asset_quality(wb["자산건전성 (APL기준)"])
    except Exception as e:
        result["asset_quality"] = {"error": str(e)}

    try:
        if "영업현황 (APL기준)" in wb.sheetnames:
            result["business_status"] = parse_business_status(wb["영업현황 (APL기준)"])
    except Exception as e:
        result["business_status"] = {"error": str(e)}

    try:
        if "대출자산속성 (APL기준)" in wb.sheetnames:
            result["loan_asset"] = parse_loan_asset(wb["대출자산속성 (APL기준)"])
    except Exception as e:
        result["loan_asset"] = {"error": str(e)}

    try:
        if "차입현황 26.05" in wb.sheetnames:
            result["borrowing"] = parse_borrowing(wb["차입현황 26.05"])
    except Exception as e:
        result["borrowing"] = {"error": str(e)}

    return result
