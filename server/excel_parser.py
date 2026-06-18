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
    """기업정보 시트 파싱"""
    data = {
        "company": {},
        "shareholders": [],
        "executives": [],
        "ceo_history": [],
        "company_history": [],
        "guarantor": {}
    }

    rows = []
    for row in ws.iter_rows():
        row_data = {cell.coordinate: safe_value(cell.value) for cell in row if cell.value is not None}
        rows.append(row_data)

    # 기업 기본정보
    data["company"] = {
        "name": "",
        "representative": "",
        "actual_manager": "",
        "business_number": "",
        "industry": "",
        "company_type": "",
        "phone": "",
        "rep_phone": "",
        "listed": "",
        "employees": "",
        "address": ""
    }

    for row in ws.iter_rows(min_row=4, max_row=7):
        cells = {cell.column_letter: safe_value(cell.value) for cell in row if cell.value is not None}
        row_num = row[0].row
        if row_num == 4:
            data["company"]["name"] = cells.get("B", "")
            data["company"]["representative"] = cells.get("D", "")
            data["company"]["actual_manager"] = cells.get("F", "") if "F" in cells else cells.get("E", "")
        elif row_num == 5:
            data["company"]["business_number"] = cells.get("B", "")
            data["company"]["industry"] = cells.get("D", "")
            data["company"]["company_type"] = cells.get("F", "")
        elif row_num == 6:
            data["company"]["phone"] = cells.get("B", "")
            data["company"]["rep_phone"] = cells.get("D", "")
            data["company"]["listed"] = cells.get("F", "")
        elif row_num == 7:
            data["company"]["employees"] = cells.get("B", "")
            data["company"]["address"] = cells.get("D", "")

    # 주주현황 (8~14행)
    shareholders = []
    for row in ws.iter_rows(min_row=9, max_row=14):
        cells = {cell.column_letter: safe_value(cell.value) for cell in row if cell.value is not None}
        if cells.get("B"):
            shareholders.append({
                "name": cells.get("B", ""),
                "shares": cells.get("C", ""),
                "ratio": cells.get("D", ""),
                "amount": cells.get("E", ""),
                "relation": cells.get("F", "")
            })
    data["shareholders"] = shareholders

    # 등기임원
    for row in ws.iter_rows(min_row=15, max_row=15):
        for cell in row:
            if cell.column_letter != "A" and cell.value:
                data["executives"].append(str(cell.value))

    # 대표이사 이력 (16~24행)
    for row in ws.iter_rows(min_row=17, max_row=24):
        cells = {cell.column_letter: safe_value(cell.value) for cell in row if cell.value is not None}
        if cells.get("B") or cells.get("C"):
            data["ceo_history"].append({
                "date": str(cells.get("B", "")),
                "content": str(cells.get("C", ""))
            })

    # 연혁 (25~40행)
    for row in ws.iter_rows(min_row=26, max_row=40):
        cells = {cell.column_letter: safe_value(cell.value) for cell in row if cell.value is not None}
        if cells.get("B") or cells.get("C"):
            data["company_history"].append({
                "date": str(cells.get("B", "")),
                "content": str(cells.get("C", ""))
            })

    # 보증인 정보
    for row in ws.iter_rows(min_row=43, max_row=46):
        cells = {cell.column_letter: safe_value(cell.value) for cell in row if cell.value is not None}
        row_num = row[0].row
        if row_num == 43:
            data["guarantor"]["name"] = cells.get("B", "")
            data["guarantor"]["representative"] = cells.get("C", "")
        elif row_num == 44:
            data["guarantor"]["business_number"] = cells.get("B", "")
            data["guarantor"]["industry"] = cells.get("D", "")
        elif row_num == 45:
            data["guarantor"]["phone"] = cells.get("B", "")

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
