"""
Full contract list parser for daily close reports.

Columns used:
  A  : contract number
  C  : product name
  L  : advertising media / agent channel
  V  : loan balance, won
  AD : contract date
  AE : contract expiry
  AH : overdue days
  AL : rank
  CG : collateral region, optional
  CJ : document status
  CK : collateral kind, optional
  CW : collateral division, optional
  CX : relationship, optional
  DU : LTV, optional
  CE : receivable collateral company, optional
  EK : NICE score, optional
  EO : K score, optional
  FI : collateral provider, optional

The parsed shape intentionally mirrors the settlement asset-quality subset used
by the daily close screen: section1, section3, section4 and loan_loss_rows.
"""

from __future__ import annotations

from collections import defaultdict
import re
from datetime import date, datetime

import openpyxl
from openpyxl.utils import column_index_from_string


def _safe_amt(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except (TypeError, ValueError):
        return 0.0


def _safe_days(value) -> int:
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return max(0, int(value))
    match = re.search(r"-?\d+", str(value).replace(",", ""))
    if not match:
        return 0
    try:
        return max(0, int(match.group(0)))
    except ValueError:
        return 0


def _safe_int(value):
    if value is None:
        return None
    try:
        return int(float(str(value).strip().replace(",", "")))
    except (TypeError, ValueError):
        return None


def _contract_no(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _header_key(value) -> str:
    text = str(value or "").strip().lower()
    return re.sub(r"[\s\[\]\(\)\{\}:/_\-·.]+", "", text)


def _find_header_column(worksheet, aliases: list[str]) -> int | None:
    targets = {_header_key(alias) for alias in aliases if str(alias or "").strip()}
    if not targets:
        return None
    for row_idx in range(1, min(worksheet.max_row, 5) + 1):
        for col_idx in range(1, worksheet.max_column + 1):
            key = _header_key(worksheet.cell(row=row_idx, column=col_idx).value)
            if key and key in targets:
                return col_idx
    return None


def _looks_numeric_text(value) -> bool:
    text = str(value or "").strip().replace(",", "")
    return bool(text) and bool(re.fullmatch(r"-?\d+(\.\d+)?", text))


_INVALID_ENTITY_LABELS = {"중도수수료"}


def _clean_entity_text(value) -> str:
    text = str(value or "").strip()
    if not text or _looks_numeric_text(text) or text in _INVALID_ENTITY_LABELS:
        return ""
    return text


def _mostly_numeric_column(worksheet, col_idx: int | None, start_row: int = 2, sample_size: int = 40) -> bool:
    if not col_idx:
        return False
    checked = 0
    numeric = 0
    for row_idx in range(start_row, worksheet.max_row + 1):
        value = worksheet.cell(row=row_idx, column=col_idx).value
        if value in (None, ""):
            continue
        checked += 1
        if _looks_numeric_text(value):
            numeric += 1
        if checked >= sample_size:
            break
    return checked >= 5 and numeric / checked >= 0.8


def _resolve_text_column(worksheet, aliases: list[str], fallback_col: int | None) -> int | None:
    header_col = _find_header_column(worksheet, aliases)
    if header_col:
        return header_col
    if fallback_col and not _mostly_numeric_column(worksheet, fallback_col):
        return fallback_col
    return None


def _date_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    text = str(value or "").strip()
    match = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text)
    if match:
        return f"{match.group(1)}-{int(match.group(2)):02d}-{int(match.group(3)):02d}"
    match = re.search(r"(?<!\d)(\d{2})\D+(\d{1,2})\D+(\d{1,2})(?!\d)", text)
    if match:
        return f"20{match.group(1)}-{int(match.group(2)):02d}-{int(match.group(3)):02d}"
    return ""


def _normalize_period(value: str | None) -> str:
    text = str(value or "").strip()
    match = re.search(r"(\d{4})\D+(\d{1,2})", text)
    if match:
        return f"{match.group(1)}-{int(match.group(2)):02d}"
    match = re.search(r"(?<!\d)(\d{2})\D+(\d{1,2})(?!\d)", text)
    if match:
        return f"20{match.group(1)}-{int(match.group(2)):02d}"
    return text


def _group_products(group: dict) -> list[str]:
    items = group.get("products")
    if items is None:
        items = group.get("items")
    return [str(item).strip() for item in (items or []) if str(item).strip()]


def _product_group_map(product_groups: list[dict] | None) -> tuple[dict[str, str], list[str]]:
    product_to_group: dict[str, str] = {}
    group_names: list[str] = []
    for group in product_groups or []:
        name = str(group.get("name", "")).strip()
        if not name:
            continue
        group_names.append(name)
        for product in _group_products(group):
            product_to_group[product] = name
    return product_to_group, group_names


def _sheet_for_contracts(workbook):
    for sheet_name in workbook.sheetnames:
        if "계약" in sheet_name or "진행" in sheet_name:
            return workbook[sheet_name]
    return workbook.active


def _build_from_rows(rows: list[dict], period: str = "", group_names: list[str] | None = None) -> dict:
    total_won = 0.0
    product_totals: dict[str, float] = defaultdict(float)
    group_totals: dict[str, float] = defaultdict(float)
    products: set[str] = set()
    groups_seen: set[str] = set()

    for row in rows:
        product = str(row.get("product", "")).strip()
        group = str(row.get("group", "")).strip()
        balance = _safe_amt(row.get("balance"))
        total_won += balance
        if product:
            products.add(product)
            product_totals[product] += balance
        if group:
            groups_seen.add(group)
            group_totals[group] += balance

    ordered_groups = list(group_names or [])
    for group in sorted(groups_seen):
        if group not in ordered_groups:
            ordered_groups.append(group)

    section3 = {
        group: {"융잔합계": group_totals.get(group, 0.0) / 1_000_000}
        for group in ordered_groups
    }
    section4 = {
        product: {"융잔합계": product_totals[product] / 1_000_000}
        for product in sorted(products)
    }

    periods = [period] if period else sorted({str(row.get("period", "")).strip() for row in rows if row.get("period")})
    return {
        "period": period,
        "periods": periods,
        "products": sorted(products),
        "section1": {"융잔합계": total_won / 1_000_000},
        "section3": section3,
        "section4": section4,
        "loan_loss_rows": rows,
        "rows": len(rows),
        "source": "full_contract",
    }


def parse_full_contract_list(filepath: str, product_groups: list[dict] | None = None, period: str = "") -> dict:
    period_key = _normalize_period(period)
    workbook = openpyxl.load_workbook(filepath, data_only=True)
    worksheet = _sheet_for_contracts(workbook)
    product_to_group, group_names = _product_group_map(product_groups)
    fixed_cols = {
        "contract_no": column_index_from_string("A"),
        "product": column_index_from_string("C"),
        "channel": column_index_from_string("L"),
        "balance": column_index_from_string("V"),
        "contract_date": column_index_from_string("AD"),
        "contract_expiry": column_index_from_string("AE"),
        "overdue_days": column_index_from_string("AH"),
        "rank": column_index_from_string("AL"),
        "collateral_region": column_index_from_string("CG"),
        "document_status": column_index_from_string("CJ"),
        "collateral_kind": column_index_from_string("CK"),
        "collateral_division": column_index_from_string("CW"),
        "relationship": column_index_from_string("CX"),
        "appraisal_agency": column_index_from_string("DA"),
        "ltv": column_index_from_string("DU"),
        "receivable_company": column_index_from_string("CE"),
        "nice_score": column_index_from_string("EK"),
        "k_score": column_index_from_string("EO"),
        "collateral_provider": column_index_from_string("FI"),
    }
    cols = dict(fixed_cols)
    cols["receivable_company"] = _resolve_text_column(
        worksheet,
        ["채권담보업체", "채권담보 업체", "채권 담보 업체"],
        fixed_cols["receivable_company"],
    )
    cols["collateral_provider"] = _resolve_text_column(
        worksheet,
        ["담보제공업체", "담보제공 업체", "담보 제공 업체"],
        fixed_cols["collateral_provider"],
    )

    rows: list[dict] = []
    for row_idx in range(2, worksheet.max_row + 1):
        contract_no = _contract_no(worksheet.cell(row=row_idx, column=cols["contract_no"]).value)
        product = str(worksheet.cell(row=row_idx, column=cols["product"]).value or "").strip()
        balance = _safe_amt(worksheet.cell(row=row_idx, column=cols["balance"]).value)
        if not product and balance == 0:
            continue
        if not product:
            continue
        group = product_to_group.get(product, "")
        channel = str(worksheet.cell(row=row_idx, column=cols["channel"]).value or "").strip()
        contract_date = _date_str(worksheet.cell(row=row_idx, column=cols["contract_date"]).value)
        contract_expiry_cell = worksheet.cell(row=row_idx, column=cols["contract_expiry"]).value
        contract_expiry = _date_str(contract_expiry_cell)
        contract_expiry_raw = str(contract_expiry_cell or "").strip()
        overdue_days = _safe_days(worksheet.cell(row=row_idx, column=cols["overdue_days"]).value)
        rank = str(worksheet.cell(row=row_idx, column=cols["rank"]).value or "").strip()
        collateral_region = str(worksheet.cell(row=row_idx, column=cols["collateral_region"]).value or "").strip()
        document_status = str(worksheet.cell(row=row_idx, column=cols["document_status"]).value or "").strip()
        collateral_kind = str(worksheet.cell(row=row_idx, column=cols["collateral_kind"]).value or "").strip()
        collateral_division = str(worksheet.cell(row=row_idx, column=cols["collateral_division"]).value or "").strip()
        relationship = str(worksheet.cell(row=row_idx, column=cols["relationship"]).value or "").strip()
        appraisal_agency = str(worksheet.cell(row=row_idx, column=cols["appraisal_agency"]).value or "").strip()
        ltv = _safe_amt(worksheet.cell(row=row_idx, column=cols["ltv"]).value)
        receivable_company = _clean_entity_text(worksheet.cell(row=row_idx, column=cols["receivable_company"]).value) if cols.get("receivable_company") else ""
        nice_score = _safe_int(worksheet.cell(row=row_idx, column=cols["nice_score"]).value)
        k_score = _safe_int(worksheet.cell(row=row_idx, column=cols["k_score"]).value)
        collateral_provider = _clean_entity_text(worksheet.cell(row=row_idx, column=cols["collateral_provider"]).value) if cols.get("collateral_provider") else ""
        rows.append({
            "period": period_key,
            "contract_no": contract_no,
            "contract_date": contract_date,
            "contract_expiry": contract_expiry,
            "contract_expiry_raw": contract_expiry_raw,
            "contract_period": contract_date[:7] if contract_date else "",
            "product": product,
            "group": group,
            "channel": channel,
            "channel_raw": channel,
            "days": overdue_days,
            "rank": rank,
            "collateral_region": collateral_region,
            "document_status": document_status,
            "balance": balance,
            "balance_million": balance / 1_000_000,
            "nice_score": nice_score,
            "k_score": k_score,
            "collateral_kind": collateral_kind,
            "collateral_division": collateral_division,
            "relationship": relationship,
            "collateral_relation": relationship,
            "appraisal_agency": appraisal_agency,
            "ltv": ltv,
            "LTV": ltv,
            "DU": ltv,
            "ltv_value": ltv,
            "ltv_raw": ltv,
            "ltv_du": ltv,
            "raw_ltv": ltv,
            "receivable_company": receivable_company,
            "collateral_provider": collateral_provider,
        })

    parsed = _build_from_rows(rows, period_key, group_names)
    parsed["source_file"] = filepath
    parsed["column_map"] = cols
    return parsed


def build_full_contract_period(data: dict, period: str) -> dict:
    period_key = _normalize_period(period)
    rows = [
        dict(row)
        for row in (data.get("loan_loss_rows") or [])
        if not period_key or str(row.get("period", "")).strip() == period_key
    ]
    group_names = list((data.get("section3") or {}).keys())
    return _build_from_rows(rows, period_key, group_names)
