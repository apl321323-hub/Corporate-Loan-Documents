from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
import tempfile
import os
import json
import io
from excel_parser import parse_excel_file, parse_settlement_asset_quality, parse_company_info
import openpyxl

app = FastAPI(title="기업여신자료 분석 시스템")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 업로드된 데이터 저장소 (메모리)
uploaded_data = {}

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
async def upload_excel(file: UploadFile = File(...)):
    """엑셀 파일 업로드 및 파싱"""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")

    try:
        # 임시 파일로 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        # 파싱
        data = parse_excel_file(tmp_path)
        uploaded_data.update(data)

        # 임시 파일 삭제
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"파일 '{file.filename}' 업로드 및 분석 완료",
            "sheets": list(data.keys())
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}")


@app.get("/api/data")
async def get_all_data():
    """전체 데이터 반환"""
    return uploaded_data


@app.post("/api/upload/settlement")
async def upload_settlement(file: UploadFile = File(...)):
    """결산자료 엑셀 업로드 및 자산건전성 계산"""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        data = parse_settlement_asset_quality(tmp_path)
        uploaded_data['settlement_asset_quality'] = data
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"결산자료 '{file.filename}' 업로드 및 분석 완료",
            "total_products": len(data.get('products', [])),
            "total_groups": len(data.get('groups', []))
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
async def upload_company_info(file: UploadFile = File(...)):
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
        uploaded_data['company_info'] = data
        os.unlink(tmp_path)

        return JSONResponse({
            "success": True,
            "message": f"기업정보 '{file.filename}' 업로드 완료",
            "company_name": data.get('company', {}).get('name', '')
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


@app.get("/api/data/business_status")
async def get_business_status():
    """영업현황 데이터"""
    return uploaded_data.get("business_status", {})


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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=3000)
