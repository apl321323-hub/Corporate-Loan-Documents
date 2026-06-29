from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
import tempfile
import os
import json
import io
from excel_parser import parse_excel_file, parse_settlement_asset_quality, parse_company_info, parse_bsis
from asset_quality_parser import parse_asset_quality_excel
from settlement_aq_parser import parse_settlement_asset_quality
from business_parser import parse_business_excel
from contract_parser import parse_contract_list
from payment_parser import parse_payment
from writeoff_parser import parse_writeoff
from sale_parser import parse_sale
from asset_parser import parse_asset
from loan_count_parser import parse_loan_count
from pdf_parser import parse_pdf_fs, apply_pdf_to_bsis
import openpyxl

from contextlib import asynccontextmanager

# ── 최근 업로드 PDF 경로 (raw값 복원용) ───────────────────────────
_UPLOADED_FILES_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "uploaded_files")

def _find_latest_pdf() -> str | None:
    """uploaded_files 디렉토리에서 가장 최근 PDF 반환"""
    d = os.path.abspath(_UPLOADED_FILES_DIR)
    if not os.path.exists(d):
        return None
    pdfs = sorted(
        [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith(".pdf")],
        key=os.path.getmtime, reverse=True
    )
    return pdfs[0] if pdfs else None

@asynccontextmanager
async def lifespan(app_):
    """서버 시작 시 최근 PDF → pdf_raw 복원 + settlement_aq maturity 재처리"""
    _data_dir = os.path.join(os.path.dirname(__file__), "data")
    _raw_path = os.path.join(_data_dir, "pdf_raw.json")
    # 이미 저장된 pdf_raw가 없으면 최신 PDF 파싱
    if not os.path.exists(_raw_path):
        pdf_path = _find_latest_pdf()
        if pdf_path:
            try:
                from pdf_parser import parse_pdf_fs
                parsed = parse_pdf_fs(pdf_path)
                raw = parsed.get("raw", {})
                os.makedirs(_data_dir, exist_ok=True)
                with open(_raw_path, "w", encoding="utf-8") as f:
                    json.dump(raw, f, ensure_ascii=False)
                print(f"[startup] pdf_raw 복원 완료: {len(raw)}개 항목")
            except Exception as ex:
                print(f"[startup] pdf_raw 복원 실패: {ex}")

    # ── settlement_aq maturity 재처리 ────────────────────────────────
    # 저장된 settlement_aq에 maturity가 비어있으면 uploaded_files에서 원본 파일 재파싱
    _saq_path = os.path.join(_data_dir, "settlement_aq.json")
    _asset_path = os.path.join(_data_dir, "asset_data.json")
    if os.path.exists(_saq_path) and os.path.exists(_asset_path):
        try:
            with open(_saq_path, encoding="utf-8") as f:
                saq_store = json.load(f)
            needs_rebuild = any(
                not pdata.get("maturity") or not any(v > 0 for v in pdata["maturity"].values())
                or not pdata.get("avg_rate")
                or not pdata.get("repayment")
                or not pdata.get("amount_band")
                or not pdata.get("channels")
                for pdata in saq_store.values()
            )
            if needs_rebuild:
                # uploaded_files 디렉토리에서 결산자료 파일 탐색
                _search_dirs = [
                    os.path.join(os.path.dirname(__file__), "..", "..", "uploaded_files"),
                    os.path.join(os.path.dirname(__file__), "..", "..", "uploaded_files2"),
                ]
                settlement_files = []
                for d in _search_dirs:
                    d = os.path.abspath(d)
                    if os.path.exists(d):
                        for fn in os.listdir(d):
                            if '결산' in fn and fn.lower().endswith(('.xlsx', '.xls')):
                                settlement_files.append(os.path.join(d, fn))
                settlement_files.sort(key=os.path.getmtime, reverse=True)

                if settlement_files:
                    from settlement_aq_parser import parse_settlement_asset_quality
                    with open(os.path.join(_data_dir, "product_groups.json"), encoding="utf-8") as f:
                        product_groups = json.load(f)
                    with open(_asset_path, encoding="utf-8") as f:
                        asset_data = json.load(f)

                    for sf in settlement_files:
                        try:
                            result = parse_settlement_asset_quality(sf, product_groups)
                            pkey = result.get("period", "")
                            maturity = result.get("maturity", {})
                            if not pkey or not maturity:
                                continue
                            # settlement_aq 업데이트
                            if pkey in saq_store:
                                saq_store[pkey]["maturity"] = maturity
                            # asset_data 대출만기 섹션 업데이트
                            sections = asset_data.get("sections", {})
                            sec_m = sections.get("대출만기", {})
                            for row_key, val in maturity.items():
                                if val is not None:
                                    if row_key not in sec_m:
                                        sec_m[row_key] = {}
                                    # 백만원 → 원 단위 변환 (기존 데이터와 단위 통일)
                                    sec_m[row_key][pkey] = float(val) * 1_000_000
                            sections["대출만기"] = sec_m
                            asset_periods = asset_data.get("periods", [])
                            if pkey not in asset_periods:
                                asset_periods.append(pkey)
                                asset_periods.sort()
                                asset_data["periods"] = asset_periods
                            asset_data["sections"] = sections

                            # ── 평균이자율 섹션 업데이트 ─────────────────────
                            avg_rate = result.get("avg_rate", {})
                            if avg_rate:
                                if pkey in saq_store:
                                    saq_store[pkey]["avg_rate"] = avg_rate
                                sections = asset_data.get("sections", {})
                                sec_rate  = sections.get("평균이자율",  {})
                                sec_rate2 = sections.get("평균이자율2", {})
                                for row_key, val in avg_rate.items():
                                    if val is None:
                                        continue
                                    if row_key == "대출자산평균":
                                        if row_key not in sec_rate:
                                            sec_rate[row_key] = {}
                                        sec_rate[row_key][pkey] = val
                                    else:
                                        if row_key not in sec_rate2:
                                            sec_rate2[row_key] = {}
                                        sec_rate2[row_key][pkey] = val
                                sections["평균이자율"]  = sec_rate
                                sections["평균이자율2"] = sec_rate2
                                asset_data["sections"] = sections

                            # ── 상환방식 섹션 업데이트 ───────────────────────
                            repayment = result.get("repayment", {})
                            if repayment:
                                if pkey in saq_store:
                                    saq_store[pkey]["repayment"] = repayment
                                sections = asset_data.get("sections", {})
                                sec_rep = sections.get("상환방식", {})
                                for row_key, val in repayment.items():
                                    if val is not None:
                                        if row_key not in sec_rep:
                                            sec_rep[row_key] = {}
                                        sec_rep[row_key][pkey] = float(val) * 1_000_000
                                sections["상환방식"] = sec_rep
                                asset_data["sections"] = sections

                            # ── 금액별 섹션 업데이트 ─────────────────────────
                            amount_band = result.get("amount_band", {})
                            if amount_band:
                                if pkey in saq_store:
                                    saq_store[pkey]["amount_band"] = amount_band
                                sections = asset_data.get("sections", {})
                                sec_amt = sections.get("금액별", {})
                                for row_key, val in amount_band.items():
                                    if val is not None:
                                        if row_key not in sec_amt:
                                            sec_amt[row_key] = {}
                                        sec_amt[row_key][pkey] = float(val) * 1_000_000
                                sections["금액별"] = sec_amt
                                asset_data["sections"] = sections

                            # ── channels(광고매체) 업데이트 ──────────────────
                            channels = result.get("channels", [])
                            if channels and pkey in saq_store:
                                saq_store[pkey]["channels"] = channels

                            print(f"[startup] settlement_aq 재처리 완료: {pkey} → maturity={maturity}, avg_rate={avg_rate}, repayment keys={list(repayment.keys()) if repayment else []}, amount_band keys={list(amount_band.keys()) if amount_band else []}")
                        except Exception as ex:
                            print(f"[startup] {sf} 재처리 실패: {ex}")

                    # 변경된 데이터 저장
                    with open(_saq_path, "w", encoding="utf-8") as f:
                        json.dump(saq_store, f, ensure_ascii=False, indent=2)
                    with open(_asset_path, "w", encoding="utf-8") as f:
                        json.dump(asset_data, f, ensure_ascii=False, indent=2)
                    print("[startup] maturity 재처리 저장 완료")
        except Exception as ex:
            print(f"[startup] maturity 재처리 오류: {ex}")

    yield

app = FastAPI(title="기업여신자료 분석 시스템", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── 파일 기반 영구 저장소 ──────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)

def _data_path(key: str) -> str:
    return os.path.join(DATA_DIR, f"{key}.json")

def save_data(key: str, value) -> None:
    """데이터를 JSON 파일로 저장"""
    with open(_data_path(key), "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)

def load_data(key: str):
    """JSON 파일에서 데이터 로드 (없으면 None)"""
    path = _data_path(key)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

class PersistentStore:
    """dict처럼 사용하되 get/update/__setitem__/__contains__ 시 파일 I/O"""
    def get(self, key, default=None):
        v = load_data(key)
        return v if v is not None else default

    def __setitem__(self, key, value):
        save_data(key, value)

    def __contains__(self, key):
        return os.path.exists(_data_path(key))

    def __delitem__(self, key):
        path = _data_path(key)
        if os.path.exists(path):
            os.remove(path)

    def pop(self, key, *args):
        """저장 파일 삭제 후 반환 (dict.pop 호환)"""
        if key in self:
            val = self.get(key)
            del self[key]
            return val
        if args:
            return args[0]   # default 반환
        raise KeyError(key)

    def update(self, d: dict):
        for k, v in d.items():
            save_data(k, v)

    def keys(self):
        names = [f[:-5] for f in os.listdir(DATA_DIR) if f.endswith(".json")]
        return names

uploaded_data = PersistentStore()

# 정적 파일 서빙
static_dir = os.path.join(os.path.dirname(__file__), "..", "public")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=os.path.join(static_dir, "static")), name="static")


@app.get("/")
async def root():
    index_path = os.path.join(os.path.dirname(__file__), "..", "public", "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {"message": "기업여신자료 분석 시스템 API"}


@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...), period: str = ""):
    """엑셀 파일 업로드 및 파싱 (period: YYYY.MM 업데이트 기준 년월)"""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_excel_file(tmp_path)
        uploaded_data.update(data)
        if period:
            uploaded_data["upload_period_main"] = period

        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"파일 '{file.filename}' 업로드 및 분석 완료",
            "sheets": list(data.keys()),
            "upload_period": period or "(미지정)",
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}")


@app.get("/api/data")
async def get_all_data():
    """전체 데이터 반환 — PersistentStore를 dict로 직렬화해서 반환"""
    # PersistentStore는 dict가 아니므로 FastAPI가 {}로 직렬화함
    # 모든 키를 명시적으로 읽어 dict로 변환
    _DATA_KEYS = [
        "bsis", "bs_is", "cashflow", "asset_quality",
        "business_status", "loan_asset", "borrowing",
        "company_info", "settlement_asset_quality",
        "asset_quality_detail",
    ]
    result = {}
    for k in _DATA_KEYS:
        v = uploaded_data.get(k)
        if v is not None:
            result[k] = v
    return JSONResponse(result)


@app.post("/api/upload/settlement")
async def upload_settlement(file: UploadFile = File(...), period: str = ""):
    """결산자료 엑셀 업로드 및 자산건전성 계산"""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_settlement_asset_quality(tmp_path)
        if period:
            data['upload_period'] = period
        uploaded_data['settlement_asset_quality'] = data
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"결산자료 '{file.filename}' 업로드 및 분석 완료",
            "total_products": len(data.get('products', [])),
            "total_groups": len(data.get('groups', [])),
            "upload_period": period or "(미지정)",
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}")


@app.get("/api/data/settlement_asset_quality")
async def get_settlement_asset_quality():
    """결산자료 기반 자산건전성 데이터"""
    return uploaded_data.get("settlement_asset_quality", {})


@app.get("/api/export/asset_quality")
async def export_asset_quality(view: str = "total", product: str = "", group: str = ""):
    """자산건전성 데이터를 엑셀로 다운로드"""
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

        data = uploaded_data.get("settlement_asset_quality", {})
        if not data:
            raise HTTPException(status_code=404, detail="결산자료가 업로드되지 않았습니다.")

        buckets = data.get('buckets', [])
        wb = openpyxl.Workbook()

        # 스타일 정의
        header_fill = PatternFill("solid", fgColor="1e3a5f")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        subheader_fill = PatternFill("solid", fgColor="2d5a8e")
        subheader_font = Font(color="FFFFFF", bold=True, size=10)
        total_fill = PatternFill("solid", fgColor="FFF3CD")
        overdue_fill = PatternFill("solid", fgColor="FFE0E0")
        center_align = Alignment(horizontal="center", vertical="center")
        right_align = Alignment(horizontal="right", vertical="center")
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

        def write_bucket_table(ws, title: str, bk_data: dict, start_row: int) -> int:
            """버킷 테이블 작성, 다음 시작행 반환"""
            # 타이틀
            ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=12, color="1e3a5f")
            start_row += 1

            # 헤더
            headers = ['구분', '건수', '잔액(원)', '연체율(%)']
            for ci, h in enumerate(headers, 1):
                c = ws.cell(row=start_row, column=ci, value=h)
                c.fill = header_fill; c.font = header_font
                c.alignment = center_align; c.border = border
            start_row += 1

            # 데이터 행
            for bk in buckets:
                v = bk_data.get(bk, {'count': 0, 'balance': 0, 'rate': 0.0})
                row_vals = [bk, v['count'], v['balance'], v['rate']]
                for ci, val in enumerate(row_vals, 1):
                    c = ws.cell(row=start_row, column=ci, value=val)
                    c.border = border
                    if bk in ('융잔합계',):
                        c.fill = total_fill; c.font = Font(bold=True)
                    elif bk in ('연체합계',):
                        c.fill = overdue_fill; c.font = Font(bold=True, color="CC0000")
                    if ci > 1:
                        c.alignment = right_align
                    if ci == 3 and isinstance(val, (int, float)):
                        c.number_format = '#,##0'
                    if ci == 4 and isinstance(val, (int, float)):
                        c.number_format = '0.00"%"'
                start_row += 1

            return start_row + 1  # 빈 행 하나 추가

        if view == "total":
            ws = wb.active
            ws.title = "전체"
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 12
            write_bucket_table(ws, "■ 전체 자산건전성 현황", data['total'], 1)
            filename = "자산건전성_전체.xlsx"

        elif view == "by_group":
            ws = wb.active
            ws.title = "상품그룹별"
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 12
            row = 1
            for g, gdata in data['by_group'].items():
                row = write_bucket_table(ws, f"■ {g}", gdata, row)
            filename = "자산건전성_상품그룹별.xlsx"

        elif view == "by_product":
            ws = wb.active
            ws.title = "상품별"
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 12
            row = 1
            target_product = product or None
            items = data['by_product']
            if target_product and target_product in items:
                row = write_bucket_table(ws, f"■ {target_product}", items[target_product], row)
            else:
                for p, pdata in items.items():
                    row = write_bucket_table(ws, f"■ {p}", pdata, row)
            filename = f"자산건전성_{target_product or '상품별'}.xlsx"

        else:
            # 전체 시트 묶음
            ws = wb.active
            ws.title = "전체"
            for col_width, col in zip([15, 10, 20, 12], ['A', 'B', 'C', 'D']):
                ws.column_dimensions[col].width = col_width
            write_bucket_table(ws, "■ 전체 자산건전성 현황", data['total'], 1)
            filename = "자산건전성_전체.xlsx"

        # 스트리밍 응답
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        from urllib.parse import quote
        encoded_filename = quote(filename)

        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"엑셀 생성 오류: {str(e)}")


@app.post("/api/upload/company_info")
async def upload_company_info(file: UploadFile = File(...), period: str = ""):
    """기업정보 엑셀 전용 업로드"""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        wb = openpyxl.load_workbook(tmp_path, data_only=True)
        ws = None
        for sname in wb.sheetnames:
            if '기업정보' in sname or '기업현황' in sname:
                ws = wb[sname]
                break
        if ws is None:
            ws = wb.active

        data = parse_company_info(ws)
        if period:
            data['upload_period'] = period
        uploaded_data['company_info'] = data
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"기업정보 '{file.filename}' 업로드 완료",
            "company_name": data.get('company', {}).get('name', ''),
            "upload_period": period or "(미지정)",
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}")


@app.get("/api/data/company_info")
async def get_company_info():
    """기업정보 데이터"""
    return uploaded_data.get("company_info", {})


@app.get("/api/data/bs_is")
async def get_bs_is():
    """BS/IS 데이터"""
    return uploaded_data.get("bs_is", {})


@app.get("/api/data/cashflow")
async def get_cashflow():
    """현금흐름 데이터"""
    return uploaded_data.get("cashflow", {})


@app.get("/api/data/asset_quality")
async def get_asset_quality():
    """자산건전성 데이터"""
    return uploaded_data.get("asset_quality", {})


@app.post("/api/upload/asset-quality")
async def upload_asset_quality(file: UploadFile = File(...), period: str = ""):
    """자산건전성 엑셀 업로드 (기업여신자료 형식)
    period: YYYY.MM 형식 업데이트 기준 년월 (선택)
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_asset_quality_excel(tmp_path)
        if period:
            data['upload_period'] = period
        uploaded_data['asset_quality_detail'] = data
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"자산건전성 '{file.filename}' 업로드 완료",
            "periods": len(data.get('periods', [])),
            "products": len(data.get('products', [])),
            "upload_period": period or "(미지정)",
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}")


@app.get("/api/data/asset-quality-detail")
async def get_asset_quality_detail():
    """자산건전성 상세 데이터 (기업여신자료 형식)"""
    d = uploaded_data.get("asset_quality_detail")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    return JSONResponse(d)


@app.post("/api/upload/settlement-aq")
async def upload_settlement_aq(file: UploadFile = File(...), period: str = ""):
    """결산자료 엑셀 업로드 → 자산건전성 데이터 계산
    H열(현재상품), J열(연체일수), L열(잔액) 기준으로
    섹션1(금액), 섹션3(담보구분별), 섹션4(상품별) 집계
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        # 상품 그룹 로드 (담보구분별 집계에 사용)
        product_groups = uploaded_data.get("product_groups", [])

        data = parse_settlement_asset_quality(tmp_path, product_groups)
        if period:
            data['upload_period'] = period
        os.unlink(tmp_path)

        # 기존 asset_quality_detail 에 결산자료 기반 데이터를 병합
        # period 기준으로 단일 기간 데이터로 저장
        existing = uploaded_data.get("asset_quality_detail") or {}
        # settlement_aq 키에 별도 저장 (기간별 여러 건 누적 가능)
        saq_store = uploaded_data.get("settlement_aq") or {}
        pkey = data["period"] or period or "unknown"
        saq_store[pkey] = data
        uploaded_data["settlement_aq"] = saq_store

        # ── 대출만기 → asset_data.sections['대출만기'] 병합 ─────────
        # 결산자료의 maturity 집계값을 대출자산속성 '대출만기' 섹션에 반영
        # pkey: 'YYYY-MM' 형태 (대출자산속성 기간 키와 동일)
        # ※ 대출자산속성 기존 데이터는 원 단위, maturity는 백만원 단위
        #   → asset_data에 저장 시 백만원 × 1,000,000 = 원 단위로 변환
        maturity = data.get("maturity", {})
        if maturity and pkey and pkey != "unknown":
            asset_data = uploaded_data.get("asset_data")
            if asset_data is not None:
                sections = asset_data.get("sections", {})
                sec_maturity = sections.get("대출만기", {})
                MATURITY_ROW_KEYS = [
                    "12개월 미만", "12~24개월 미만",
                    "24~36개월 미만", "36개월 이상", "전체융잔 합계"
                ]
                for row_key in MATURITY_ROW_KEYS:
                    val = maturity.get(row_key)
                    if val is not None:
                        if row_key not in sec_maturity:
                            sec_maturity[row_key] = {}
                        # 백만원 → 원 단위 변환 (기존 데이터와 단위 통일)
                        sec_maturity[row_key][pkey] = float(val) * 1_000_000
                sections["대출만기"] = sec_maturity

                # periods 통합
                asset_periods = asset_data.get("periods", [])
                if pkey not in asset_periods:
                    asset_periods.append(pkey)
                    asset_periods.sort()
                    asset_data["periods"] = asset_periods

                asset_data["sections"] = sections
                uploaded_data["asset_data"] = asset_data

        # ── 평균이자율 → asset_data.sections['평균이자율'/'평균이자율2'] 병합 ────
        # avg_rate 구조: {"대출자산평균": float, "신용": float, "담보": float, ...}
        # - "대출자산평균" → sections['평균이자율']['대출자산평균']
        # - 그룹명(신용/담보) → sections['평균이자율2'][그룹명]
        avg_rate = data.get("avg_rate", {})
        if avg_rate and pkey and pkey != "unknown":
            asset_data = uploaded_data.get("asset_data")
            if asset_data is not None:
                sections = asset_data.get("sections", {})
                sec_rate  = sections.get("평균이자율",  {})
                sec_rate2 = sections.get("평균이자율2", {})
                for row_key, val in avg_rate.items():
                    if val is None:
                        continue
                    if row_key == "대출자산평균":
                        if row_key not in sec_rate:
                            sec_rate[row_key] = {}
                        sec_rate[row_key][pkey] = val
                    else:
                        # 그룹별(신용/담보 등) → 평균이자율2
                        if row_key not in sec_rate2:
                            sec_rate2[row_key] = {}
                        sec_rate2[row_key][pkey] = val
                sections["평균이자율"]  = sec_rate
                sections["평균이자율2"] = sec_rate2
                asset_data["sections"] = sections
                uploaded_data["asset_data"] = asset_data

        # ── 상환방식 → asset_data.sections['상환방식'] 병합 ──────────
        # repayment 구조: {"신용 _원리금균등상환": int(백만원), ..., "전체융잔 합계": int}
        # asset_data의 상환방식 섹션은 {행키: {기간: 원값}} 형태
        # 단위 통일: repayment는 백만원 → 원 단위(×1_000_000)로 변환 후 저장
        repayment = data.get("repayment", {})
        if repayment and pkey and pkey != "unknown":
            asset_data = uploaded_data.get("asset_data")
            if asset_data is not None:
                sections = asset_data.get("sections", {})
                sec_rep = sections.get("상환방식", {})
                for row_key, val in repayment.items():
                    if val is not None:
                        if row_key not in sec_rep:
                            sec_rep[row_key] = {}
                        sec_rep[row_key][pkey] = float(val) * 1_000_000
                sections["상환방식"] = sec_rep
                asset_data["sections"] = sections
                uploaded_data["asset_data"] = asset_data

        # ── 금액별 → asset_data.sections['금액별'] 병합 ──────────────
        # amount_band 구조: {"300만원이하": int(백만원), ..., "전체융잔 합계": int}
        # 단위 통일: amount_band는 백만원 → 원 단위(×1_000_000)로 변환 후 저장
        amount_band = data.get("amount_band", {})
        if amount_band and pkey and pkey != "unknown":
            asset_data = uploaded_data.get("asset_data")
            if asset_data is not None:
                sections = asset_data.get("sections", {})
                sec_amt = sections.get("금액별", {})
                for row_key, val in amount_band.items():
                    if val is not None:
                        if row_key not in sec_amt:
                            sec_amt[row_key] = {}
                        sec_amt[row_key][pkey] = float(val) * 1_000_000
                sections["금액별"] = sec_amt
                asset_data["sections"] = sections
                uploaded_data["asset_data"] = asset_data

        return JSONResponse({
            "success": True,
            "message": f"결산자료 '{file.filename}' 업로드 완료",
            "period": data["period"],
            "products": len(data.get("products", [])),
            "section1_total": data["section1"].get("융잔합계", 0),
            "upload_period": period or "(미지정)",
            "maturity_updated":  bool(maturity   and pkey and pkey != "unknown" and uploaded_data.get("asset_data") is not None),
            "avg_rate_updated":  bool(avg_rate   and pkey and pkey != "unknown" and uploaded_data.get("asset_data") is not None),
            "repayment_updated":    bool(repayment   and pkey and pkey != "unknown" and uploaded_data.get("asset_data") is not None),
            "amount_band_updated": bool(amount_band and pkey and pkey != "unknown" and uploaded_data.get("asset_data") is not None),
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")


@app.get("/api/data/settlement-aq")
async def get_settlement_aq(period: str = ""):
    """결산자료 기반 자산건전성 데이터 반환
    period 지정 시 해당 기간, 미지정 시 최신 기간
    """
    store = uploaded_data.get("settlement_aq") or {}
    if not store:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    if period and period in store:
        return JSONResponse(store[period])
    # 최신 기간 반환
    latest_key = sorted(store.keys())[-1]
    return JSONResponse(store[latest_key])


@app.get("/api/data/settlement-aq/periods")
async def get_settlement_aq_periods():
    """업로드된 결산자료 기간 목록 반환"""
    store = uploaded_data.get("settlement_aq") or {}
    return JSONResponse(sorted(store.keys()))


# ── 상품 그룹 API ─────────────────────────────────────────
@app.get("/api/product-groups")
async def get_product_groups():
    """저장된 상품 그룹 목록 반환"""
    groups = uploaded_data.get("product_groups", [])
    return JSONResponse(groups)


@app.post("/api/product-groups")
async def save_product_groups(request: Request):
    """상품 그룹 전체 저장 (덮어쓰기)"""
    body = await request.json()
    # body: [{"name": "신용", "products": ["N론", "V론", ...]}, ...]
    if not isinstance(body, list):
        raise HTTPException(status_code=400, detail="배열 형태로 전송하세요")
    uploaded_data["product_groups"] = body
    return JSONResponse({"success": True, "count": len(body)})


@app.get("/api/product-groups/all-products")
async def get_all_products():
    """자산건전성 데이터에서 상품 목록 반환"""
    aq = uploaded_data.get("asset_quality_detail")
    if aq and aq.get("products"):
        return JSONResponse(aq["products"])
    # 기본 상품 목록 (엑셀 미업로드 시 fallback)
    default_products = [
        "N론","V론","담보론","담보론(지분대출)","레이디론","스타론","스타스위치론",
        "우량론","큐브론","테일론","토탈론","프리론","프리미엄론","전월세론","플러스론",
        "토마토론","토마토N론","토마토토탈론","T플러스론","OP론","충성론","오투론",
        "오투N론","토마토토탈론플러스","다이렉트론(A)","다이렉트론(W)","N론(하이브리드)",
        "기타","기타N"
    ]
    return JSONResponse(default_products)


# ── 접수경로 그룹 API ─────────────────────────────────────────────────────
@app.get("/api/channel-groups")
async def get_channel_groups():
    """저장된 접수경로 그룹 목록 반환"""
    groups = uploaded_data.get("channel_groups", [])
    # 파일 폴백
    if not groups:
        _path = os.path.join(os.path.dirname(__file__), "data", "channel_groups.json")
        if os.path.exists(_path):
            with open(_path, encoding="utf-8") as f:
                groups = json.load(f)
            uploaded_data["channel_groups"] = groups
    return JSONResponse(groups)


@app.post("/api/channel-groups")
async def save_channel_groups(request: Request):
    """접수경로 그룹 전체 저장 (덮어쓰기)"""
    body = await request.json()
    if not isinstance(body, list):
        raise HTTPException(status_code=400, detail="배열 형태로 전송하세요")
    uploaded_data["channel_groups"] = body
    # 파일에도 저장
    _path = os.path.join(os.path.dirname(__file__), "data", "channel_groups.json")
    os.makedirs(os.path.dirname(_path), exist_ok=True)
    with open(_path, "w", encoding="utf-8") as f:
        json.dump(body, f, ensure_ascii=False, indent=2)
    return JSONResponse({"success": True, "count": len(body)})


@app.get("/api/channel-groups/all-channels")
async def get_all_channels():
    """접수경로 전체 항목 반환
    우선순위: 결산자료 Q열(광고매체) 고유값 → asset_data 접수경로 섹션 키 → 기본값
    """
    # 1순위: 결산자료 파싱 결과의 channels 목록 (Q열 광고매체 고유값)
    saq = uploaded_data.get("settlement_aq") or {}
    if not saq:
        _saq_file = os.path.join(os.path.dirname(__file__), "data", "settlement_aq.json")
        if os.path.exists(_saq_file):
            with open(_saq_file, encoding="utf-8") as f:
                saq = json.load(f)
    if saq:
        # settlement_aq는 {pkey: {channels:[...], ...}} 구조
        for pkey, pdata in saq.items():
            channels = pdata.get("channels", [])
            if channels:
                return JSONResponse(sorted(channels))

    # 2순위: asset_data 접수경로 섹션 키
    asset_data = uploaded_data.get("asset_data")
    if asset_data:
        sec = asset_data.get("sections", {}).get("접수경로", {})
        channels = [k for k in sec.keys() if k != "전체융잔 합계"]
        if channels:
            return JSONResponse(channels)

    # 3순위: 기본값 (fallback)
    default_channels = ["에이전트", "직접대출", "홈페이지", "온라인플랫폼", "기타(고객추천)"]
    return JSONResponse(default_channels)


@app.get("/api/data/business_status")
async def get_business_status():
    """영업현황 데이터 (기존 형식)"""
    return uploaded_data.get("business_status", {})


@app.get("/api/data/business")
async def get_business_detail():
    """영업현황 3섹션 데이터 (엑셀 파싱 기반)"""
    d = uploaded_data.get("business_detail")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    return JSONResponse(d)


@app.post("/api/upload/business")
async def upload_business(file: UploadFile = File(...), period: str = ""):
    """영업현황 엑셀 업로드 (기업여신자료 - 영업현황.xlsx)"""
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_business_excel(tmp_path)
        if period:
            data['upload_period'] = period
        uploaded_data['business_detail'] = data
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"영업현황 '{file.filename}' 업로드 완료",
            "periods": len(data.get('periods', [])),
            "upload_period": period or "(미지정)",
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")


@app.get("/api/data/loan_asset")
async def get_loan_asset():
    """대출자산속성 데이터"""
    return uploaded_data.get("loan_asset", {})


@app.get("/api/data/borrowing")
async def get_borrowing():
    """차입현황 데이터"""
    return uploaded_data.get("borrowing", {})


@app.post("/api/data/company_info")
async def update_company_info(data: dict):
    """기업정보 수동 수정"""
    if "company_info" not in uploaded_data:
        uploaded_data["company_info"] = {}
    uploaded_data["company_info"].update(data)
    return {"success": True, "message": "기업정보가 업데이트되었습니다."}


# ─── BS/IS 엔드포인트 ────────────────────────────────────────────────────────

@app.post("/api/upload/bsis")
async def upload_bsis(file: UploadFile = File(...), period: str = ""):
    """BS/IS 엑셀 파일 업로드 (재무상태표 + 손익계산서 + BSPL_요약)"""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        wb = openpyxl.load_workbook(tmp_path, data_only=True)
        data = parse_bsis(wb)
        if period:
            data['upload_period'] = period
        uploaded_data['bsis'] = data
        os.unlink(tmp_path)

        sheet_info = []
        if 'bs' in data:
            sheet_info.append(f"BS {len(data['bs']['headers'])}개월")
        if 'is_' in data:
            sheet_info.append(f"IS {len(data['is_']['headers'])}개월")
        if 'summary' in data:
            sheet_info.append(f"요약 {len(data['summary']['headers'])}개년")

        return JSONResponse({
            "success": True,
            "message": f"BS/IS '{file.filename}' 업로드 완료",
            "sheets": sheet_info,
            "upload_period": period or "(미지정)",
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}")


@app.get("/api/data/bsis")
async def get_bsis():
    """BS/IS 데이터 반환"""
    return uploaded_data.get("bsis", {})


@app.post("/api/upload/pdf-fs")
async def upload_pdf_fs(
    file: UploadFile = File(...),
    period: str = "2026.05"
):
    """
    재무제표 PDF 업로드 → BS/IS/Summary 특정 기간 컬럼 업데이트
    period: "YYYY.MM" 형식 (예: "2026.05")
    """
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="PDF 파일만 업로드 가능합니다.")

    bsis_data = uploaded_data.get("bsis", {})
    if not bsis_data:
        raise HTTPException(status_code=400, detail="먼저 BS/IS 엑셀 파일(섹터 4)을 업로드하세요.")

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        # PDF 파싱 (저장된 사용자 매핑 우선 적용)
        user_mapping = uploaded_data.get("pdf_mapping", None)
        parsed = parse_pdf_fs(tmp_path, user_mapping=user_mapping)
        os.unlink(tmp_path)

        # raw 값 저장 (매핑 설정 UI의 '당기값' 표시용)
        if "raw" in parsed:
            uploaded_data["pdf_raw"] = parsed["raw"]

        # 기간이 bsis headers에 없으면 추가
        added_period = False
        for sheet_key in ["bs", "is_", "summary"]:
            sheet = bsis_data.get(sheet_key, {})
            headers = sheet.get("headers", [])
            if period not in headers:
                # 기간을 정렬 삽입 (YYYY.MM 기준)
                headers.append(period)
                headers.sort()
                sheet["headers"] = headers
                # 모든 row의 values에 None 추가 (삽입 위치에)
                insert_idx = headers.index(period)
                for row in sheet.get("rows", []):
                    row["values"].insert(insert_idx, None)
                added_period = True

        # PDF 값으로 bsis 업데이트
        updated_bsis, sheets_updated = apply_pdf_to_bsis(bsis_data, parsed, period)
        uploaded_data['bsis'] = updated_bsis

        return JSONResponse({
            "success": True,
            "message": f"재무제표 '{file.filename}' → {period} 업데이트 완료",
            "period": period,
            "period_added": added_period,
            "sheets_updated": sheets_updated,
            "parsed_counts": {
                "bs": len(parsed["bs"]),
                "is_": len(parsed["is_"]),
                "summary": len(parsed["summary"]),
            }
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")


@app.get("/api/bsis/periods")
async def get_bsis_periods():
    """BS/IS 데이터의 사용 가능한 기간 목록 반환"""
    bsis_data = uploaded_data.get("bsis", {})
    periods = {}
    for sheet_key in ["bs", "is_", "summary"]:
        sheet = bsis_data.get(sheet_key, {})
        periods[sheet_key] = sheet.get("headers", [])
    return JSONResponse(periods)


@app.get("/api/pdf-mapping")
async def get_pdf_mapping():
    """
    PDF 과목 → 엑셀 과목 매핑 설정 반환
    저장된 설정이 없으면 pdf_mapping.py 기본값으로 초기화
    """
    saved = uploaded_data.get("pdf_mapping", None)
    pdf_raw = uploaded_data.get("pdf_raw", {})

    # raw 값에서 당기값(cur) 추출 헬퍼
    def get_pdf_val(pdf_key):
        entry = pdf_raw.get(pdf_key)
        if entry and isinstance(entry, list) and len(entry) > 0 and entry[0] is not None:
            return entry[0]
        return None

    from pdf_mapping import PDF_BS_MAPPING, PDF_IS_MAPPING, PDF_SUMMARY_MAPPING, \
                            DEFAULT_SIGN, DEFAULT_SIGN_FALLBACK

    bsis_data = uploaded_data.get("bsis", {})
    excel_opts = {
        "bs":      [r["label"] for r in bsis_data.get("bs",      {}).get("rows", [])],
        "is_":     [r["label"] for r in bsis_data.get("is_",     {}).get("rows", [])],
        "summary": [r["label"] for r in bsis_data.get("summary", {}).get("rows", [])],
    }

    from pdf_parser import _normalize

    # ── pdf_map 키 기준 역방향 탐색으로 행 생성 ──────────────────────
    # 문제: raw 키 중 "대손충당금2,308,865,085..." 같은 가비지 키는
    #       _normalize 후에도 매핑 딕셔너리의 "대손충당금"과 불일치 →
    #       raw 순서 기반 탐색에서 해당 과목이 누락됨
    # 해결: pdf_map의 각 norm_key에 대해 raw 전체를 스캔하여
    #       _normalize(raw_k) == norm_key 인 첫 번째 raw 인덱스를 위치로 사용
    def build_rows_from_raw(pdf_map):
        raw_keys = list(pdf_raw.keys())

        # norm_key → 해당 norm_key와 매칭되는 가장 이른 raw 인덱스
        norm_to_first_raw_idx: dict = {}
        for i, raw_k in enumerate(raw_keys):
            nk = _normalize(raw_k)
            if nk not in norm_to_first_raw_idx:
                norm_to_first_raw_idx[nk] = i

        rows = []
        for norm_key, excel_label in pdf_map.items():
            idx = norm_to_first_raw_idx.get(norm_key, 99999)
            rows.append({
                "pdf_key":     norm_key,
                "pdf_val":     get_pdf_val(norm_key),
                "excel_label": excel_label,
                "enabled":     excel_label is not None,
                "sign":        DEFAULT_SIGN.get(norm_key, DEFAULT_SIGN_FALLBACK),
                "_sort_idx":   idx,
            })

        # raw 파싱 순서대로 정렬 후 _sort_idx 제거
        rows.sort(key=lambda r: r["_sort_idx"])
        for r in rows:
            r.pop("_sort_idx")
        return rows

    if saved:
        # 저장된 매핑 → raw 순서로 재정렬 + pdf_val 주입 + sign 없으면 기본값 보완
        # build_rows_from_raw 와 동일한 논리: norm_key별 첫 등장 raw 인덱스 사용
        norm_order: dict = {}
        for i, raw_k in enumerate(pdf_raw.keys()):
            nk = _normalize(raw_k)
            if nk not in norm_order:
                norm_order[nk] = i
        for sheet_key in ["bs", "is_", "summary"]:
            for row in saved.get(sheet_key, []):
                row["pdf_val"] = get_pdf_val(row.get("pdf_key", ""))
                # 구버전 저장 데이터에 sign 없으면 기본값 주입
                if "sign" not in row:
                    row["sign"] = DEFAULT_SIGN.get(row.get("pdf_key", ""), DEFAULT_SIGN_FALLBACK)
            saved[sheet_key] = sorted(
                saved.get(sheet_key, []),
                key=lambda r: norm_order.get(r["pdf_key"], 99999)
            )
        return JSONResponse(saved)

    config = {
        "bs":      build_rows_from_raw(PDF_BS_MAPPING),
        "is_":     build_rows_from_raw(PDF_IS_MAPPING),
        "summary": build_rows_from_raw(PDF_SUMMARY_MAPPING),
        "excel_options": excel_opts,
    }
    return JSONResponse(config)


@app.post("/api/pdf-mapping")
async def save_pdf_mapping(request: Request):
    """
    PDF 매핑 설정 저장
    Body: { bs: [{pdf_key, excel_label, enabled}], is_: [...], summary: [...] }
    { _reset: true } 전달 시 저장된 설정 삭제 → 기본값으로 복원
    """
    body = await request.json()

    # _reset: true 처리 — 저장된 매핑 삭제 (기본값으로 복원)
    if body.get("_reset"):
        if "pdf_mapping" in uploaded_data:
            del uploaded_data["pdf_mapping"]
        return JSONResponse({"success": True, "message": "매핑 설정이 기본값으로 복원되었습니다."})

    body.pop("excel_options", None)
    body.pop("_reset", None)
    # _custom_labels, _ordered_labels는 별도 키로 분리 저장
    custom_labels  = body.pop("_custom_labels", None)
    ordered_labels = body.pop("_ordered_labels", None)

    # 모든 시트가 빈 배열이면 저장하지 않음 (기본 매핑 덮어쓰기 방지)
    total_rows = sum(len(body.get(k, [])) for k in ["bs", "is_", "summary"])
    if total_rows > 0:
        uploaded_data["pdf_mapping"] = body
    # custom_labels / ordered_labels는 내용과 관계없이 항상 저장
    if custom_labels is not None:
        uploaded_data["custom_excel_labels"] = custom_labels
    if ordered_labels is not None:
        uploaded_data["ordered_excel_labels"] = ordered_labels
    return JSONResponse({"success": True, "message": "매핑 설정이 저장되었습니다."})


@app.get("/api/pdf-mapping/excel-options")
async def get_excel_options():
    """엑셀 BS/IS/Summary 과목 목록 반환 (매핑 드롭다운용)
    - bsis rows 전체 label 반환
    - 수기 추가 과목(custom_excel_labels) 별도 포함
    """
    bsis_data     = uploaded_data.get("bsis", {})
    custom_labels = uploaded_data.get("custom_excel_labels", {"bs": [], "is_": [], "summary": []})
    ordered_labels = uploaded_data.get("ordered_excel_labels", {"bs": None, "is_": None, "summary": None})
    return JSONResponse({
        "bs":      [r["label"] for r in bsis_data.get("bs",      {}).get("rows", [])],
        "is_":     [r["label"] for r in bsis_data.get("is_",     {}).get("rows", [])],
        "summary": [r["label"] for r in bsis_data.get("summary", {}).get("rows", [])],
        "_custom_labels":  custom_labels,
        "_ordered_labels": ordered_labels,
    })


@app.delete("/api/custom-excel-label")
async def delete_custom_excel_label(request: Request):
    """
    수기 추가 과목 삭제
    Body: { "sheet": "bs", "label": "테스트" }
    - custom_excel_labels.json에서 제거
    - ordered_excel_labels.json에서 제거
    - bsis.json 해당 시트 rows에서 제거
    """
    body = await request.json()
    sheet = body.get("sheet")
    label = body.get("label")
    if not sheet or not label:
        raise HTTPException(status_code=400, detail="sheet, label 필수")

    # 1. custom_excel_labels 업데이트
    custom = uploaded_data.get("custom_excel_labels", {"bs": [], "is_": [], "summary": []})
    if label in custom.get(sheet, []):
        custom[sheet] = [l for l in custom[sheet] if l != label]
        uploaded_data["custom_excel_labels"] = custom

    # 2. ordered_excel_labels 업데이트
    ordered = uploaded_data.get("ordered_excel_labels", {"bs": None, "is_": None, "summary": None})
    if ordered.get(sheet) and label in ordered[sheet]:
        ordered[sheet] = [l for l in ordered[sheet] if l != label]
        uploaded_data["ordered_excel_labels"] = ordered

    # 3. bsis rows에서 해당 label 행 제거
    bsis_data = uploaded_data.get("bsis", {})
    if bsis_data and sheet in bsis_data:
        rows = bsis_data[sheet].get("rows", [])
        new_rows = [r for r in rows if r.get("label") != label]
        if len(new_rows) != len(rows):
            bsis_data[sheet]["rows"] = new_rows
            uploaded_data["bsis"] = bsis_data

    return JSONResponse({"success": True, "message": f"'{label}' 과목이 삭제되었습니다."})


@app.post("/api/pdf-mapping/apply")
async def apply_pdf_mapping(request: Request):
    """
    저장된 pdf_raw + pdf_mapping으로 bsis를 재반영 (PDF 재업로드 없이 즉시 적용)
    Body: { "period": "2026.05" }  — 생략 시 _pdf_updated에 기록된 최근 기간 사용
    """
    body = await request.json()
    period = body.get("period", None)

    # 1. 저장된 raw 확인
    pdf_raw = uploaded_data.get("pdf_raw", {})
    if not pdf_raw:
        raise HTTPException(status_code=400, detail="저장된 PDF raw 데이터가 없습니다. PDF를 먼저 업로드하세요.")

    # 2. 매핑 확인
    user_mapping = uploaded_data.get("pdf_mapping", None)
    if not user_mapping:
        raise HTTPException(status_code=400, detail="저장된 매핑 설정이 없습니다.")

    # 3. bsis 확인
    bsis_data = uploaded_data.get("bsis", {})
    if not bsis_data:
        raise HTTPException(status_code=400, detail="BS/IS 엑셀 파일을 먼저 업로드하세요.")

    # 4. period 자동 결정 (지정 안 하면 _pdf_updated에서 가장 최근 기간)
    if not period:
        pdf_updated = bsis_data.get("_pdf_updated", {})
        periods = sorted([p for p, v in pdf_updated.items() if v], reverse=True)
        if periods:
            period = periods[0]
        else:
            # bsis headers에서 가장 최근 period 사용
            headers = bsis_data.get("bs", {}).get("headers", [])
            period = headers[-1] if headers else None

    if not period:
        raise HTTPException(status_code=400, detail="적용할 기간을 지정하세요 (예: {\"period\": \"2026.05\"})")

    # 5. 수기 추가 과목을 bsis rows에 없으면 삽입
    custom_labels  = uploaded_data.get("custom_excel_labels", {})
    ordered_labels = uploaded_data.get("ordered_excel_labels", {})

    sheet_map = {"bs": "bs", "is_": "is_", "summary": "summary"}
    for sh_key, sh_name in sheet_map.items():
        sheet = bsis_data.get(sh_key)
        if not sheet:
            continue
        existing_labels = {r["label"] for r in sheet.get("rows", [])}
        custom_list = custom_labels.get(sh_key, [])
        if not custom_list:
            continue

        # ordered_labels가 있으면 그 순서대로, 없으면 기존 rows 뒤에 append
        ordered = ordered_labels.get(sh_key)  # None 이면 순서 정보 없음
        headers = sheet.get("headers", [])
        n_cols  = len(headers)

        for lbl in custom_list:
            if lbl not in existing_labels:
                # 새 row 생성 (모든 기간 None으로 초기화)
                sheet["rows"].append({
                    "label":  lbl,
                    "values": [None] * n_cols,
                    "indent": 0,
                    "bold":   False,
                })
                existing_labels.add(lbl)

        # ordered 순서가 있으면 rows를 재정렬
        if ordered:
            # ordered 에 없는 label은 뒤에 붙임
            order_idx = {lbl: i for i, lbl in enumerate(ordered)}
            sheet["rows"].sort(
                key=lambda r: order_idx.get(r["label"], len(ordered))
            )

    # raw는 { norm_key: [cur, prev] } 형태 — apply_user_mapping과 동일 형태
    from pdf_parser import apply_user_mapping, apply_pdf_to_bsis
    mapped = apply_user_mapping(pdf_raw, user_mapping)
    updated_bsis, sheets_updated = apply_pdf_to_bsis(bsis_data, mapped, period)
    uploaded_data["bsis"] = updated_bsis

    return JSONResponse({
        "success": True,
        "period": period,
        "sheets_updated": sheets_updated,
        "applied_counts": {
            "bs":      len(mapped["bs"]),
            "is_":     len(mapped["is_"]),
            "summary": len(mapped["summary"]),
        },
    })


@app.post("/api/upload/contract")
async def upload_contract(file: UploadFile = File(...), period: str = ""):
    """
    계약리스트 엑셀 업로드
    C열(상품명), Q열(계약구분), AD열(계약일), T열(대출액) 기준으로
    section1(전체 취급액), section3(담보구분별) 집계
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        product_groups = uploaded_data.get("product_groups", [])
        data = parse_contract_list(tmp_path, product_groups)
        if period:
            data['upload_period'] = period
        uploaded_data['contract_data'] = data
        os.unlink(tmp_path)

        # ── asset_data['sections']['대출채권잔액'] 자동 업데이트 ──
        section_balance = data.get("section_balance", {})
        if section_balance:
            asset_data = uploaded_data.get("asset_data")
            if asset_data is not None:
                sections = asset_data.get("sections", {})
                # 기존 대출채권잔액 섹션에 기간별 값 병합 (덮어쓰기)
                existing_sec = sections.get("대출채권잔액", {})
                for row_label, ym_vals in section_balance.items():
                    if row_label not in existing_sec:
                        existing_sec[row_label] = {}
                    existing_sec[row_label].update(ym_vals)
                sections["대출채권잔액"] = existing_sec

                # ── asset_data['sections']['대출취급액'] 자동 업데이트 ──
                section_deal = data.get("section_deal", {})
                if section_deal:
                    existing_deal = sections.get("대출취급액", {})
                    for row_label, ym_vals in section_deal.items():
                        if row_label not in existing_deal:
                            existing_deal[row_label] = {}
                        existing_deal[row_label].update(ym_vals)
                    sections["대출취급액"] = existing_deal

                asset_data["sections"] = sections
                # periods 병합
                asset_periods = asset_data.get("periods", [])
                for ym in data.get("periods", []):
                    if ym not in asset_periods:
                        asset_periods.append(ym)
                asset_periods.sort()
                asset_data["periods"] = asset_periods
                uploaded_data["asset_data"] = asset_data

        # ── 평균이자율(취급대출평균) → asset_data.sections['평균이자율'] 병합 ──
        # 계약리스트의 avg_rate_deal {ym: float} → '취급대출평균' 행으로 반영
        avg_rate_deal = data.get("avg_rate_deal", {})
        if avg_rate_deal:
            asset_data = uploaded_data.get("asset_data")
            if asset_data is not None:
                sections = asset_data.get("sections", {})
                sec_rate = sections.get("평균이자율", {})
                if "취급대출평균" not in sec_rate:
                    sec_rate["취급대출평균"] = {}
                for ym, val in avg_rate_deal.items():
                    if val is not None:
                        sec_rate["취급대출평균"][ym] = val
                sections["평균이자율"] = sec_rate
                asset_data["sections"] = sections
                uploaded_data["asset_data"] = asset_data

        return JSONResponse({
            "success": True,
            "message": f"계약리스트 '{file.filename}' 업로드 완료",
            "periods": data.get('periods', []),
            "products": len(data.get('products', [])),
            "balance_rows": list(section_balance.keys()),
            "deal_rows": list(data.get("section_deal", {}).keys()),
            "avg_rate_deal": {k: v for k, v in avg_rate_deal.items() if v is not None},
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")


@app.get("/api/data/contract")
async def get_contract_data():
    """계약리스트 집계 데이터 반환 (section1 / section3)"""
    d = uploaded_data.get("contract_data")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    return JSONResponse(d)


@app.post("/api/upload/payment")
async def upload_payment(file: UploadFile = File(...), period: str = ""):
    """
    입금명세 엑셀 업로드
    K열(회계)='+', AY열(계약구분)=신규/추가대출/재대출 필터 후
    N열(원금입금) → 원금회수액, Q열(이자입금) → 이자회수액 월별 집계
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_payment(tmp_path)
        if period:
            data['upload_period'] = period
        uploaded_data['payment_data'] = data
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"입금명세 '{file.filename}' 업로드 완료",
            "periods": data.get('periods', []),
            "rows": {
                "원금회수액": list(data.get('section1', {}).get('원금회수액', {}).values()),
                "이자회수액": list(data.get('section1', {}).get('이자회수액', {}).values()),
            },
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")


@app.get("/api/data/payment")
async def get_payment_data():
    """입금명세 집계 데이터 반환 (원금회수액 / 이자회수액)"""
    d = uploaded_data.get("payment_data")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    return JSONResponse(d)


@app.post("/api/upload/writeoff")
async def upload_writeoff(file: UploadFile = File(...), period: str = ""):
    """
    대손리스트 엑셀 업로드
    AZ열(대손일) 기준 월별, V열(대출잔액) 합산 → 월중상각액
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_writeoff(tmp_path)
        if period:
            data['upload_period'] = period
        uploaded_data['writeoff_data'] = data
        os.unlink(tmp_path)

        s1 = data.get('section1', {})
        series = s1.get('실상각액_월중상각액', {})
        total_vals = list(series.values())

        return JSONResponse({
            "success": True,
            "message": f"대손리스트 '{file.filename}' 업로드 완료",
            "periods": data.get('periods', []),
            "total": sum(total_vals),
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")


@app.get("/api/data/writeoff")
async def get_writeoff_data():
    """대손리스트 집계 데이터 반환 (월중상각액)"""
    d = uploaded_data.get("writeoff_data")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    return JSONResponse(d)


@app.post("/api/upload/sale")
async def upload_sale(file: UploadFile = File(...), period: str = ""):
    """매각리스트 xlsx 업로드 → 실상각액_월중매각액 월별 집계"""
    try:
        content = await file.read()
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_sale(tmp_path)
        if period:
            data['upload_period'] = period
        uploaded_data['sale_data'] = data
        os.unlink(tmp_path)

        s1 = data.get('section1', {})
        series = s1.get('실상각액_월중매각액', {})
        total_vals = list(series.values())

        return JSONResponse({
            "success": True,
            "message": f"매각리스트 '{file.filename}' 업로드 완료",
            "periods": data.get('periods', []),
            "total": sum(total_vals),
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")


@app.get("/api/data/sale")
async def get_sale_data():
    """매각리스트 집계 데이터 반환 (월중매각액)"""
    d = uploaded_data.get("sale_data")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    return JSONResponse(d)


@app.post("/api/upload/asset")
async def upload_asset(file: UploadFile = File(...)):
    import tempfile, os
    suffix = os.path.splitext(file.filename)[1] if file.filename else ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    try:
        data = parse_asset(tmp_path)
        uploaded_data['asset_data'] = data
        return JSONResponse({
            "status": "ok",
            "periods": len(data["periods"]),
            "sections": list(data["sections"].keys()),
            "cb_products": list(data["cb_products"].keys()),
        })
    except Exception as e:
        import traceback
        return JSONResponse({"error": str(e), "detail": traceback.format_exc()}, status_code=500)
    finally:
        os.unlink(tmp_path)


@app.get("/api/data/asset")
async def get_asset_data():
    d = uploaded_data.get("asset_data")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)

    import copy
    result = copy.deepcopy(d)

    # ── settlement_aq section4 → sections['상품별'] 동적 구성 ──────────
    # H열(현재상품)별 L열(잔액) 합산값을 기간별로 집계하여 상품별 섹션을 대체
    saq = uploaded_data.get("settlement_aq") or {}
    if saq:
        prod_sec: dict[str, dict[str, float]] = {}
        for pkey, pdata in saq.items():
            # pkey 형식: "2025-06" 또는 "2025-06_suffix"
            ym = pkey[:7] if len(pkey) >= 7 and pkey[4] == '-' else None
            if not ym:
                continue
            s4 = pdata.get("section4") or {}
            for prod, buckets in s4.items():
                # 융잔합계 키 (백만원 단위로 이미 변환됨)
                val = buckets.get("융잔합계")
                if val is None:
                    continue
                if prod not in prod_sec:
                    prod_sec[prod] = {}
                prod_sec[prod][ym] = float(val) * 1_000_000  # 원 단위로 복원

        if prod_sec:
            # 전체융잔 합계: 기간별 모든 상품 합산
            total_by_ym: dict[str, float] = {}
            for prod, ymap in prod_sec.items():
                for ym, val in ymap.items():
                    total_by_ym[ym] = total_by_ym.get(ym, 0.0) + val
            prod_sec["전체융잔 합계"] = total_by_ym

            sections = result.get("sections", {})
            sections["상품별"] = prod_sec
            result["sections"] = sections

    return JSONResponse(result)


@app.post("/api/upload/loan-count")
async def upload_loan_count(file: UploadFile = File(...), period: str = Form(default="")):
    import tempfile, os
    suffix = os.path.splitext(file.filename)[1] if file.filename else ".xls"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    try:
        data = parse_loan_count(tmp_path, period)
        # 기존 저장 데이터에 해당 기간 값을 병합
        existing = uploaded_data.get("loan_count_data", {"periods": [], "rows": {}})
        ym = data["period"]
        if ym and ym not in existing["periods"]:
            existing["periods"].append(ym)
            existing["periods"].sort()
        for key, val in data["rows"].items():
            if key not in existing["rows"]:
                existing["rows"][key] = {}
            if ym:
                existing["rows"][key][ym] = val
        uploaded_data["loan_count_data"] = existing

        # asset_data의 대출취급건수 섹션도 업데이트 (연동)
        if ym:
            asset_data = uploaded_data.get("asset_data")
            if asset_data is not None:
                sections = asset_data.get("sections", {})
                sec = sections.get("대출취급건수", {})
                for key, val in data["rows"].items():
                    if key not in sec:
                        sec[key] = {}
                    sec[key][ym] = float(val)
                sections["대출취급건수"] = sec
                asset_data["sections"] = sections
                # periods에도 추가
                asset_periods = asset_data.get("periods", [])
                if ym not in asset_periods:
                    asset_periods.append(ym)
                    asset_periods.sort()
                    asset_data["periods"] = asset_periods
                uploaded_data["asset_data"] = asset_data

        return JSONResponse({
            "status": "ok",
            "period": ym,
            "rows": data["rows"],
            "raw": data["raw"],
        })
    except Exception as e:
        import traceback
        return JSONResponse({"error": str(e), "detail": traceback.format_exc()}, status_code=500)
    finally:
        os.unlink(tmp_path)


@app.get("/api/data/loan-count")
async def get_loan_count_data():
    d = uploaded_data.get("loan_count_data")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    return JSONResponse(d)


@app.post("/api/settlement-aq/rebuild-maturity")
async def rebuild_maturity():
    """
    저장된 settlement_aq 데이터에서 maturity 정보가 비어있는 경우
    asset_data의 대출만기 섹션을 재빌드합니다.
    (기존 업로드된 결산자료 파일이 있으면 재파싱, 없으면 settlement_aq 데이터에서 재계산 불가)
    """
    import traceback
    from datetime import date
    import calendar

    saq_store = uploaded_data.get("settlement_aq") or {}
    if not saq_store:
        # 메모리에 없으면 파일에서 로드
        _saq_file = os.path.join(os.path.dirname(__file__), "data", "settlement_aq.json")
        if os.path.exists(_saq_file):
            with open(_saq_file, encoding="utf-8") as _f:
                saq_store = json.load(_f)
            uploaded_data["settlement_aq"] = saq_store  # 메모리에도 캐시
    if not saq_store:
        return JSONResponse({"error": "결산자료 데이터가 없습니다. 결산자료를 먼저 업로드하세요."}, status_code=404)

    # asset_data: 메모리에 없으면 파일에서 로드
    asset_data = uploaded_data.get("asset_data")
    if asset_data is None:
        _asset_path = os.path.join(os.path.dirname(__file__), "data", "asset_data.json")
        if os.path.exists(_asset_path):
            with open(_asset_path, encoding="utf-8") as _f:
                asset_data = json.load(_f)
            uploaded_data["asset_data"] = asset_data  # 메모리에도 캐시

    results = []
    for pkey, pdata in saq_store.items():
        maturity = pdata.get("maturity", {})
        # maturity가 비어있거나 전체융잔 합계가 0이면 재처리 시도
        if not maturity or not any(v > 0 for v in maturity.values()):
            results.append({"period": pkey, "status": "maturity_empty", "note": "결산자료 파일을 다시 업로드하세요."})
            continue

        # ── amount_band가 없으면 원본 결산자료 파일 재파싱 ──────────
        if not pdata.get("amount_band"):
            _search_dirs = [
                os.path.join(os.path.dirname(__file__), "..", "..", "uploaded_files"),
                os.path.join(os.path.dirname(__file__), "..", "..", "uploaded_files2"),
            ]
            _reparsed = None
            for _d in _search_dirs:
                _d = os.path.abspath(_d)
                if not os.path.exists(_d):
                    continue
                for _fn in sorted(os.listdir(_d), key=lambda n: os.path.getmtime(os.path.join(_d, n)), reverse=True):
                    if "결산" in _fn and _fn.lower().endswith((".xlsx", ".xls")):
                        try:
                            from settlement_aq_parser import parse_settlement_asset_quality
                            _pg_path = os.path.join(os.path.dirname(__file__), "data", "product_groups.json")
                            with open(_pg_path, encoding="utf-8") as _pf:
                                _pg = json.load(_pf)
                            _res = parse_settlement_asset_quality(os.path.join(_d, _fn), _pg)
                            if _res.get("period") == pkey:
                                _reparsed = _res
                                break
                        except Exception:
                            continue
                if _reparsed:
                    break
            if _reparsed:
                pdata["amount_band"] = _reparsed.get("amount_band", {})
                pdata["repayment"]   = _reparsed.get("repayment", pdata.get("repayment", {}))
                pdata["maturity"]    = _reparsed.get("maturity", maturity)
                pdata["channels"]    = _reparsed.get("channels", pdata.get("channels", []))
                maturity = pdata["maturity"]
                saq_store[pkey] = pdata
                # saq_store 파일에도 저장
                _saq_save = os.path.join(os.path.dirname(__file__), "data", "settlement_aq.json")
                with open(_saq_save, "w", encoding="utf-8") as _sf:
                    json.dump(saq_store, _sf, ensure_ascii=False, indent=2)

        # maturity 값이 있는 경우 → asset_data에 반영
        if asset_data is None:
            results.append({"period": pkey, "status": "no_asset_data", "note": "대출자산속성 파일 없음 — 반영 불가"})
            continue

        sections = asset_data.get("sections", {})
        sec_maturity = sections.get("대출만기", {})
        MATURITY_ROW_KEYS = ["12개월 미만", "12~24개월 미만", "24~36개월 미만", "36개월 이상", "전체융잔 합계"]
        for row_key in MATURITY_ROW_KEYS:
            val = maturity.get(row_key)
            if val is not None:
                if row_key not in sec_maturity:
                    sec_maturity[row_key] = {}
                # 백만원 → 원 단위 변환 (기존 데이터와 단위 통일)
                sec_maturity[row_key][pkey] = float(val) * 1_000_000
        sections["대출만기"] = sec_maturity

        asset_periods = asset_data.get("periods", [])
        if pkey not in asset_periods:
            asset_periods.append(pkey)
            asset_periods.sort()
            asset_data["periods"] = asset_periods

        asset_data["sections"] = sections

        # ── 평균이자율 섹션도 함께 재적용 ────────────────────────
        avg_rate = pdata.get("avg_rate", {})
        if avg_rate:
            sections = asset_data.get("sections", {})
            sec_rate  = sections.get("평균이자율",  {})
            sec_rate2 = sections.get("평균이자율2", {})
            for row_key, val in avg_rate.items():
                if val is None:
                    continue
                if row_key == "대출자산평균":
                    if row_key not in sec_rate:
                        sec_rate[row_key] = {}
                    sec_rate[row_key][pkey] = val
                else:
                    if row_key not in sec_rate2:
                        sec_rate2[row_key] = {}
                    sec_rate2[row_key][pkey] = val
            sections["평균이자율"]  = sec_rate
            sections["평균이자율2"] = sec_rate2
            asset_data["sections"] = sections

        # ── 상환방식 섹션도 함께 재적용 ──────────────────────────
        repayment = pdata.get("repayment", {})
        if repayment:
            sections = asset_data.get("sections", {})
            sec_rep = sections.get("상환방식", {})
            for row_key, val in repayment.items():
                if val is not None:
                    if row_key not in sec_rep:
                        sec_rep[row_key] = {}
                    sec_rep[row_key][pkey] = float(val) * 1_000_000
            sections["상환방식"] = sec_rep
            asset_data["sections"] = sections

        # ── 금액별 섹션도 함께 재적용 ────────────────────────────
        amount_band = pdata.get("amount_band", {})
        if amount_band:
            sections = asset_data.get("sections", {})
            sec_amt = sections.get("금액별", {})
            for row_key, val in amount_band.items():
                if val is not None:
                    if row_key not in sec_amt:
                        sec_amt[row_key] = {}
                    sec_amt[row_key][pkey] = float(val) * 1_000_000
            sections["금액별"] = sec_amt
            asset_data["sections"] = sections

        uploaded_data["asset_data"] = asset_data
        results.append({"period": pkey, "status": "merged", "maturity": maturity, "avg_rate": avg_rate, "repayment": repayment, "amount_band": amount_band})

    # ── rebuild 결과를 파일에 저장 ──────────────────────────────
    if any(r.get("status") == "merged" for r in results) and asset_data is not None:
        _asset_save_path = os.path.join(os.path.dirname(__file__), "data", "asset_data.json")
        try:
            with open(_asset_save_path, "w", encoding="utf-8") as _f:
                json.dump(asset_data, _f, ensure_ascii=False, indent=2)
        except Exception as _e:
            pass  # 저장 실패는 무시 (메모리 반영은 이미 완료)

    return JSONResponse({"success": True, "results": results})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=3000)
