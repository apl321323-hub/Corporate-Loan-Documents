from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
import tempfile
import os
import json
from excel_parser import parse_excel_file

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
