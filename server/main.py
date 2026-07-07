from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request



from fastapi.middleware.cors import CORSMiddleware



from fastapi.staticfiles import StaticFiles



from fastapi.responses import FileResponse, JSONResponse, StreamingResponse



import tempfile



import os



import json



import io



import re



import copy



from datetime import datetime



from excel_parser import parse_excel_file, parse_settlement_asset_quality, parse_company_info, parse_bsis



from asset_quality_parser import parse_asset_quality_excel, SECTION1_ROWS, SECTION2_ROWS, SECTION3_GROUPS, SECTION4_PRODUCTS, PRODUCT_ROW_OFFSETS



from settlement_aq_parser import parse_settlement_asset_quality



from business_parser import parse_business_excel



from contract_parser import parse_contract_list, normalize_period_key



from full_contract_parser import parse_full_contract_list, build_full_contract_period

from payment_parser import parse_payment

from supabase_store import SupabaseStore



from writeoff_parser import parse_writeoff



from sale_parser import parse_sale



from asset_parser import parse_asset



from loan_count_parser import parse_loan_count



from pdf_parser import parse_pdf_fs, apply_pdf_to_bsis



import openpyxl







from contextlib import asynccontextmanager

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _load_env_file() -> None:
    """Load simple KEY=VALUE entries from .env without adding a dependency."""
    for path in (os.path.join(ROOT_DIR, ".env"), os.path.join(os.path.dirname(__file__), ".env")):
        if not os.path.exists(path):
            continue
        try:
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    text = line.strip()
                    if not text or text.startswith("#") or "=" not in text:
                        continue
                    key, value = text.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if key and key not in os.environ:
                        os.environ[key] = value
        except Exception as ex:
            print(f"[env] .env 로드 실패({path}): {ex}")


_load_env_file()







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



    _data_dir = DATA_DIR



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



                or not pdata.get("channel_amount")



                or not pdata.get("channels")



                or not pdata.get("hwahae_bw")



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




                            channel_amount = result.get("channel_amount", {})



                            if channel_amount:



                                if pkey in saq_store:



                                    saq_store[pkey]["channel_amount"] = channel_amount



                                sections = asset_data.get("sections", {})



                                sec_channel = sections.get("접수경로", {})



                                for row_key, val in channel_amount.items():



                                    if val is not None:



                                        if row_key not in sec_channel:



                                            sec_channel[row_key] = {}



                                        sec_channel[row_key][pkey] = float(val) * 1_000_000



                                sections["접수경로"] = sec_channel



                                asset_data["sections"] = sections







                            # ── BW/BX 화해채권 업데이트 ──────────────────────



                            hwahae_rank = result.get("hwahae_rank", [])



                            hwahae_bw = result.get("hwahae_bw", [])



                            hwahae_bx = result.get("hwahae_bx", [])



                            section5_rank = result.get("section5_rank", {})



                            section5_bw = result.get("section5_bw", {})



                            section5_bx = result.get("section5_bx", {})



                            section5_rows = result.get("section5_rows", [])



                            if pkey in saq_store:



                                saq_store[pkey]["hwahae_basis"] = result.get("hwahae_basis", "rank_bw_bx")



                                if hwahae_rank:



                                    saq_store[pkey]["hwahae_rank"] = hwahae_rank



                                if hwahae_bw:



                                    saq_store[pkey]["hwahae_bw"] = hwahae_bw



                                if hwahae_bx:



                                    saq_store[pkey]["hwahae_bx"] = hwahae_bx



                                if section5_rank:



                                    saq_store[pkey]["section5_rank"] = section5_rank



                                if section5_bw:



                                    saq_store[pkey]["section5_bw"] = section5_bw



                                if section5_bx:



                                    saq_store[pkey]["section5_bx"] = section5_bx



                                if section5_rows:



                                    saq_store[pkey]["section5_rows"] = section5_rows



                                gender_r = result.get("gender", {})



                                age_r    = result.get("age_band", {})



                                job_r    = result.get("job_raw", {})



                                region_r = result.get("region", {})



                                cb_raw_r = result.get("cb_raw", {})



                                if gender_r:



                                    saq_store[pkey]["gender"]   = gender_r



                                if age_r:



                                    saq_store[pkey]["age_band"] = age_r



                                if job_r:



                                    saq_store[pkey]["job_raw"]  = job_r



                                if region_r:



                                    saq_store[pkey]["region"]   = region_r



                                if cb_raw_r:



                                    saq_store[pkey]["cb_raw"]   = {str(k): v for k, v in cb_raw_r.items()}







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



# 프로젝트 폴더 안의 server/data는 코드 갱신/재클론 시 함께 바뀔 수 있어,



# 기본 저장 위치는 사용자 문서의 별도 폴더로 둔다. 필요하면 APL_DATA_DIR로 덮어쓴다.



LEGACY_DATA_DIR = os.path.join(os.path.dirname(__file__), "data")



DEFAULT_DATA_DIR = os.path.join(os.path.expanduser("~"), "Documents", "기업여신자료_업로드데이터")



DATA_DIR = os.environ.get("APL_DATA_DIR") or DEFAULT_DATA_DIR







try:



    os.makedirs(DATA_DIR, exist_ok=True)



except Exception as ex:



    print(f"[data-store] 외부 저장소 생성 실패, 프로젝트 내부 저장소 사용: {ex}")



    DATA_DIR = LEGACY_DATA_DIR



    os.makedirs(DATA_DIR, exist_ok=True)







FILES_DIR = os.path.join(DATA_DIR, "files")



LEGACY_FILES_DIR = os.path.join(LEGACY_DATA_DIR, "files")



os.makedirs(FILES_DIR, exist_ok=True)







def _data_path(key: str) -> str:



    return os.path.join(DATA_DIR, f"{key}.json")







def _legacy_data_path(key: str) -> str:



    return os.path.join(LEGACY_DATA_DIR, f"{key}.json")







def _file_path(filename: str) -> str:



    return os.path.join(FILES_DIR, filename)







def _stored_file_path(filename: str) -> str | None:



    path = os.path.join(FILES_DIR, filename)



    if os.path.exists(path):



        return path



    legacy_path = os.path.join(LEGACY_FILES_DIR, filename)



    if os.path.exists(legacy_path):



        return legacy_path



    return None







def migrate_legacy_data() -> None:



    """기존 server/data/*.json을 새 저장소로 1회 이관"""



    if os.path.abspath(DATA_DIR) == os.path.abspath(LEGACY_DATA_DIR):



        return



    if not os.path.exists(LEGACY_DATA_DIR):



        return



    for name in os.listdir(LEGACY_DATA_DIR):



        if not name.endswith(".json"):



            continue



        src = os.path.join(LEGACY_DATA_DIR, name)



        dst = os.path.join(DATA_DIR, name)



        if os.path.exists(dst):



            continue



        try:



            with open(src, "r", encoding="utf-8") as f:



                data = json.load(f)



            with open(dst, "w", encoding="utf-8") as f:



                json.dump(data, f, ensure_ascii=False, indent=2)



            print(f"[data-store] 기존 데이터 이관: {name}")



        except Exception as ex:



            print(f"[data-store] 기존 데이터 이관 실패({name}): {ex}")







migrate_legacy_data()

SUPABASE_STORE = SupabaseStore()
SUPABASE_READ_FIRST = str(os.environ.get("SUPABASE_READ_FIRST", "1")).strip().lower() not in {"0", "false", "no", "off"}


def _save_local_data(key: str, value) -> None:
    with open(_data_path(key), "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)


def _load_local_data(key: str):
    path = _data_path(key)
    if not os.path.exists(path):
        legacy_path = _legacy_data_path(key)
        if os.path.abspath(path) == os.path.abspath(legacy_path) or not os.path.exists(legacy_path):
            return None
        path = legacy_path
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None







def save_data(key: str, value) -> None:



    """데이터를 로컬 JSON에 저장하고, 설정된 경우 Supabase에도 동기화"""



    _save_local_data(key, value)
    if SUPABASE_STORE.enabled:
        try:
            SUPABASE_STORE.save(key, value)
        except Exception as ex:
            print(f"[supabase] 저장 실패({key}): {ex}")







def load_data(key: str):



    """Supabase/로컬 JSON에서 데이터 로드 (없으면 None)"""



    if SUPABASE_STORE.enabled and SUPABASE_READ_FIRST:
        try:
            remote = SUPABASE_STORE.load(key)
            if remote is not None:
                _save_local_data(key, remote)
                return remote
        except Exception as ex:
            print(f"[supabase] 로드 실패({key}): {ex}")

    local = _load_local_data(key)
    if local is not None:
        return local

    if SUPABASE_STORE.enabled and not SUPABASE_READ_FIRST:
        try:
            return SUPABASE_STORE.load(key)
        except Exception as ex:
            print(f"[supabase] 로드 실패({key}): {ex}")
    return None







class PersistentStore:



    """dict처럼 사용하되 get/update/__setitem__/__contains__ 시 파일 I/O"""



    def get(self, key, default=None):



        v = load_data(key)



        return v if v is not None else default







    def __getitem__(self, key):



        v = load_data(key)



        if v is None:



            raise KeyError(key)



        return v







    def __setitem__(self, key, value):



        save_data(key, value)







    def __contains__(self, key):



        if os.path.exists(_data_path(key)) or os.path.exists(_legacy_data_path(key)):
            return True
        return load_data(key) is not None







    def __delitem__(self, key):



        for path in {_data_path(key), _legacy_data_path(key)}:



            if os.path.exists(path):



                os.remove(path)
        if SUPABASE_STORE.enabled:
            try:
                SUPABASE_STORE.delete(key)
            except Exception as ex:
                print(f"[supabase] 삭제 실패({key}): {ex}")







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



        if SUPABASE_STORE.enabled:
            try:
                names = sorted(set(names) | set(SUPABASE_STORE.keys()))
            except Exception as ex:
                print(f"[supabase] 키 목록 로드 실패: {ex}")
        return names







uploaded_data = PersistentStore()


@app.get("/api/storage/status")
async def storage_status():
    """현재 저장소 연결 상태 반환."""
    return JSONResponse({
        "local_data_dir": DATA_DIR,
        "supabase": SUPABASE_STORE.status(),
        "supabase_read_first": SUPABASE_READ_FIRST,
    })


@app.post("/api/storage/sync-local-to-supabase")
async def sync_local_to_supabase():
    """로컬 JSON 저장 데이터를 Supabase로 1회 동기화."""
    if not SUPABASE_STORE.enabled:
        raise HTTPException(status_code=400, detail="Supabase 환경변수가 설정되지 않았습니다.")
    synced = []
    failed = []
    for name in os.listdir(DATA_DIR):
        if not name.endswith(".json"):
            continue
        key = name[:-5]
        value = _load_local_data(key)
        if value is None:
            continue
        try:
            SUPABASE_STORE.save(key, value)
            synced.append(key)
        except Exception as ex:
            failed.append({"key": key, "error": str(ex)})
    return JSONResponse({"success": not failed, "synced": synced, "failed": failed})







# 정적 파일 서빙



static_dir = os.path.join(os.path.dirname(__file__), "..", "public")



if os.path.exists(static_dir):



    app.mount("/static", StaticFiles(directory=os.path.join(static_dir, "static")), name="static")











def _index_file_response():



    index_path = os.path.join(os.path.dirname(__file__), "..", "public", "index.html")



    if os.path.exists(index_path):



        return FileResponse(



            index_path,



            headers={



                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",



                "Pragma": "no-cache",



            },



        )



    return {"message": "\uae30\uc5c5\uc5ec\uc2e0\uc790\ub8cc \ubd84\uc11d API"}











@app.get("/")



async def root():



    return _index_file_response()











@app.get("/index.html")



async def index_html():



    return _index_file_response()











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















def _ci_blank(value):



    return value is None or (isinstance(value, str) and value.strip() == "")







def _ci_text(value):



    if _ci_blank(value):



        return None



    return str(value)







def _ci_number(value):



    if _ci_blank(value):



        return None



    if isinstance(value, (int, float)) and not isinstance(value, bool):



        return value



    s = str(value).strip().replace(',', '').replace('\uc8fc', '').replace('\uc6d0', '')



    if not s:



        return None



    try:



        n = float(s)



        return int(n) if n.is_integer() else n



    except ValueError:



        return str(value)







def _ci_ratio(value):



    if _ci_blank(value):



        return None



    if isinstance(value, (int, float)) and not isinstance(value, bool):



        n = float(value)



        return n / 100 if n > 1 else n



    s = str(value).strip().replace(',', '')



    if s.endswith('%'):



        s = s[:-1].strip()



    try:



        n = float(s)



        return n / 100 if n > 1 else n



    except ValueError:



        return str(value)







def _ci_set(ws, coord: str, value, kind: str = "text") -> None:



    if kind == "number":



        ws[coord] = _ci_number(value)



    elif kind == "ratio":



        ws[coord] = _ci_ratio(value)



    else:



        ws[coord] = _ci_text(value)







def _ci_clear(ws, coords: list[str]) -> None:



    for coord in coords:



        ws[coord] = None







def _write_company_info_to_sheet(ws, data: dict) -> None:



    company = data.get("company") or {}



    _ci_set(ws, "B4", company.get("name"))



    _ci_set(ws, "D4", company.get("representative"))



    _ci_set(ws, "F4", company.get("actual_manager"))



    _ci_set(ws, "B5", company.get("business_number"))



    _ci_set(ws, "D5", company.get("industry"))



    _ci_set(ws, "F5", company.get("company_type"))



    _ci_set(ws, "B6", company.get("phone"))



    _ci_set(ws, "D6", company.get("rep_phone"))



    _ci_set(ws, "F6", company.get("listed"))



    _ci_set(ws, "B7", company.get("employees"), "number")



    _ci_set(ws, "D7", company.get("address"))







    shareholders = data.get("shareholders") or []



    for idx in range(6):



        row = 9 + idx



        _ci_clear(ws, [f"{col}{row}" for col in "BCDEF"])



        if idx >= len(shareholders):



            continue



        item = shareholders[idx] or {}



        _ci_set(ws, f"B{row}", item.get("name"))



        _ci_set(ws, f"C{row}", item.get("shares"), "number")



        _ci_set(ws, f"D{row}", item.get("ratio"), "ratio")



        _ci_set(ws, f"E{row}", item.get("amount"), "number")



        _ci_set(ws, f"F{row}", item.get("relation"))







    _ci_clear(ws, [f"{col}15" for col in "BCDEFGHIJK"])



    for idx, name in enumerate(data.get("executives") or []):



        if idx >= 10:



            break



        _ci_set(ws, f"{chr(ord('B') + idx)}15", name)







    ceo_history = data.get("ceo_history") or []



    for idx in range(8):



        row = 17 + idx



        _ci_clear(ws, [f"B{row}", f"C{row}"])



        if idx >= len(ceo_history):



            continue



        item = ceo_history[idx] or {}



        _ci_set(ws, f"B{row}", item.get("date"))



        _ci_set(ws, f"C{row}", item.get("content"))







    company_history = data.get("company_history") or []



    for idx in range(15):



        row = 26 + idx



        _ci_clear(ws, [f"B{row}", f"C{row}"])



        if idx >= len(company_history):



            continue



        item = company_history[idx] or {}



        _ci_set(ws, f"B{row}", item.get("date"))



        _ci_set(ws, f"C{row}", item.get("content"))







    guarantor = data.get("guarantor") or {}



    _ci_set(ws, "B43", guarantor.get("name"))



    _ci_set(ws, "C43", guarantor.get("representative"))



    _ci_set(ws, "F43", guarantor.get("actual_manager"))



    _ci_set(ws, "B44", guarantor.get("business_number"))



    _ci_set(ws, "D44", guarantor.get("industry"))



    _ci_set(ws, "F44", guarantor.get("company_type"))



    _ci_set(ws, "B45", guarantor.get("phone"))



    _ci_set(ws, "D45", guarantor.get("rep_phone"))



    _ci_set(ws, "F45", guarantor.get("listed"))



    _ci_set(ws, "B46", guarantor.get("employees"), "number")



    _ci_set(ws, "D46", guarantor.get("address"))







    guarantor_shareholders = data.get("guarantor_shareholders") or []



    for idx in range(6):



        row = 48 + idx



        _ci_clear(ws, [f"{col}{row}" for col in "BCDEF"])



        if idx >= len(guarantor_shareholders):



            continue



        item = guarantor_shareholders[idx] or {}



        _ci_set(ws, f"B{row}", item.get("name"))



        _ci_set(ws, f"C{row}", item.get("shares"), "number")



        _ci_set(ws, f"D{row}", item.get("ratio"), "ratio")



        _ci_set(ws, f"E{row}", item.get("amount"), "number")



        _ci_set(ws, f"F{row}", item.get("relation"))







    _ci_clear(ws, [f"{col}54" for col in "BCDEFGHIJK"])



    for idx, name in enumerate(data.get("guarantor_executives") or []):



        if idx >= 10:



            break



        _ci_set(ws, f"{chr(ord('B') + idx)}54", name)







    guarantor_history = data.get("guarantor_history") or []



    for idx in range(8):



        row = 56 + idx



        _ci_clear(ws, [f"B{row}", f"C{row}"])



        if idx >= len(guarantor_history):



            continue



        item = guarantor_history[idx] or {}



        _ci_set(ws, f"B{row}", item.get("date"))



        _ci_set(ws, f"C{row}", item.get("content"))







    _ci_set(ws, "B67", data.get("business_direction"))



    _ci_set(ws, "B68", data.get("affiliates"))







def _company_info_sheet(wb, preferred_name: str | None = None):



    company_info_name = "\uae30\uc5c5\uc815\ubcf4"



    company_status_name = "\uae30\uc5c5\ud604\ud669"



    if preferred_name and preferred_name in wb.sheetnames:



        return wb[preferred_name]



    for sname in wb.sheetnames:



        if company_info_name in sname or company_status_name in sname:



            return wb[sname]



    return wb.active







@app.post("/api/upload/company_info")



async def upload_company_info(file: UploadFile = File(...), period: str = ""):



    """기업정보 엑셀 전용 업로드"""



    if not file.filename or not file.filename.lower().endswith(('.xlsx', '.xls')):



        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")



    tmp_path = None



    try:



        suffix = os.path.splitext(file.filename)[1] or '.xlsx'



        content = await file.read()



        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:



            tmp.write(content)



            tmp_path = tmp.name







        wb = openpyxl.load_workbook(tmp_path, data_only=True)



        ws = _company_info_sheet(wb)







        data = parse_company_info(ws)



        if period:



            data['upload_period'] = period



        uploaded_data['company_info'] = data







        template_path = _file_path("company_info_template.xlsx")



        with open(template_path, "wb") as template_file:



            template_file.write(content)



        uploaded_data["company_info_template"] = {



            "filename": file.filename,



            "sheet_name": ws.title,



        }







        return JSONResponse({



            "success": True,



            "message": f"기업정보 '{file.filename}' 업로드 완료",



            "company_name": data.get('company', {}).get('name', ''),



            "upload_period": period or "(미지정)",



        })



    except Exception as e:



        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}")



    finally:



        if tmp_path and os.path.exists(tmp_path):



            os.unlink(tmp_path)











@app.get("/api/data/company_info")



async def get_company_info():



    """기업정보 데이터"""



    return uploaded_data.get("company_info", {})















@app.get("/api/export/company_info")



async def export_company_info():



    """업로드한 기업현황 엑셀 서식을 보존하여 현재 기업정보를 다운로드"""



    data = uploaded_data.get("company_info")



    if not data:



        raise HTTPException(status_code=404, detail="기업정보 데이터가 없습니다. 기업현황 엑셀을 먼저 업로드하세요.")







    template_path = _stored_file_path("company_info_template.xlsx")



    if not template_path:



        raise HTTPException(status_code=404, detail="기업현황 원본 엑셀 서식이 없습니다. 기업현황 엑셀을 다시 업로드하세요.")







    try:



        meta = uploaded_data.get("company_info_template", {}) or {}



        wb = openpyxl.load_workbook(template_path)



        ws = _company_info_sheet(wb, meta.get("sheet_name"))



        _write_company_info_to_sheet(ws, data)







        buf = io.BytesIO()



        wb.save(buf)



        buf.seek(0)







        from urllib.parse import quote



        company_name = (data.get("company") or {}).get("name") or "기업정보"



        safe_company = re.sub(r'[\\/:*?"<>|]+', '_', str(company_name)).strip() or "기업정보"



        filename = f"기업여신자료_기업현황_{safe_company}.xlsx"



        encoded_filename = quote(filename)







        return StreamingResponse(



            buf,



            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",



            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"},



        )



    except HTTPException:



        raise



    except Exception as e:



        raise HTTPException(status_code=500, detail=f"기업정보 엑셀 생성 오류: {str(e)}")















def _aq_norm_period(value):



    if value is None:



        return ""



    if hasattr(value, "strftime") and not isinstance(value, str):



        try:



            return value.strftime("%Y.%m")



        except Exception:



            pass



    s = str(value).strip()



    m = re.search(r'(\d{4})[-.](\d{1,2})', s)



    if m:



        return f"{m.group(1)}.{int(m.group(2)):02d}"



    return s











def _aq_period_sort_key(period):



    period = _aq_norm_period(period)



    m = re.search(r'(\d{4})[-.](\d{1,2})', period)



    if not m:



        return (9999, 99, period)



    return (int(m.group(1)), int(m.group(2)), period)











def _aq_insert_blank_period(data, index):



    def insert_list(values):



        if isinstance(values, list):



            values.insert(index, None)







    for values in (data.get("section1") or {}).values():



        insert_list(values)



    for values in (data.get("section2") or {}).values():



        insert_list(values)



    for group_data in (data.get("section3") or {}).values():



        if isinstance(group_data, dict):



            for values in group_data.values():



                insert_list(values)



    for product_data in (data.get("section4") or {}).values():



        if isinstance(product_data, dict):



            for values in product_data.values():



                insert_list(values)











def _aq_ensure_period(data, period):



    period = _aq_norm_period(period)



    periods = [_aq_norm_period(p) for p in (data.get("periods") or [])]



    data["periods"] = periods



    if period in periods:



        return periods.index(period)







    periods.append(period)



    sorted_periods = sorted(periods, key=_aq_period_sort_key)



    index = sorted_periods.index(period)



    data["periods"] = sorted_periods



    _aq_insert_blank_period(data, index)



    return index











def _aq_set_series_value(container, label, index, value, period_count):



    if not isinstance(container, dict):



        return



    values = container.get(label)



    if not isinstance(values, list):



        values = [None] * period_count



        container[label] = values



    while len(values) < period_count:



        values.append(None)



    values[index] = value











def _aq_settlement_amount_to_won(value):



    if isinstance(value, (int, float)):



        return int(round(float(value) * 1_000_000))



    return value











def _aq_loan_loss_allowance_by_period(settlement_store):



    """Return total loan-loss allowance by period in won for asset-quality section2."""



    result = {}



    if not isinstance(settlement_store, dict):



        return result



    config = _loss_config_with_products()



    for key, period_data in settlement_store.items():



        if not isinstance(period_data, dict):



            continue



        period = _aq_norm_period(period_data.get("period") or period_data.get("upload_period") or key)



        if not period or _aq_period_sort_key(period) < _aq_period_sort_key("2024.09"):



            continue



        try:



            summary = _loss_calculate_period(period, period_data, config).get("summary") or {}



            allowance = summary.get("total_allowance")



            if isinstance(allowance, (int, float)) and not isinstance(allowance, bool):



                result[period] = int(round(float(allowance)))



        except Exception:



            continue



    return result











def _aq_apply_loan_loss_allowance(merged, settlement_store):



    allowance_by_period = _aq_loan_loss_allowance_by_period(settlement_store)



    if not allowance_by_period:



        return False



    section2 = merged.setdefault("section2", {})



    applied = False



    for period, allowance in allowance_by_period.items():



        index = _aq_ensure_period(merged, period)



        period_count = len(merged.get("periods") or [])



        _aq_set_series_value(section2, "\ub300\uc190\ucda9\ub2f9\uae08", index, allowance, period_count)



        applied = True



    return applied











def _merge_settlement_into_asset_quality_detail(aq_data, settlement_store):



    if not aq_data:



        return aq_data



    merged = copy.deepcopy(aq_data)



    merged["periods"] = [_aq_norm_period(p) for p in (merged.get("periods") or [])]



    if not isinstance(settlement_store, dict):



        return merged







    applied = False



    for key, period_data in settlement_store.items():



        if not isinstance(period_data, dict):



            continue



        period = _aq_norm_period(period_data.get("period") or period_data.get("upload_period") or key)



        if not period:



            continue







        index = _aq_ensure_period(merged, period)



        period_count = len(merged.get("periods") or [])







        section1 = merged.setdefault("section1", {})



        for label, value in (period_data.get("section1") or {}).items():



            _aq_set_series_value(section1, label, index, _aq_settlement_amount_to_won(value), period_count)







        section3 = merged.setdefault("section3", {})



        for group, group_values in (period_data.get("section3") or {}).items():



            target_group = section3.setdefault(group, {})



            if isinstance(group_values, dict):



                for label, value in group_values.items():



                    normalized_value = value if label == "\uc5f0\uccb4\uc728(%)" else _aq_settlement_amount_to_won(value)



                    _aq_set_series_value(target_group, label, index, normalized_value, period_count)







        section4 = merged.setdefault("section4", {})



        products = merged.setdefault("products", [])



        for product, product_values in (period_data.get("section4") or {}).items():



            target_product = section4.setdefault(product, {})



            if product not in products:



                products.append(product)



            if isinstance(product_values, dict):



                for label, value in product_values.items():



                    _aq_set_series_value(target_product, label, index, _aq_settlement_amount_to_won(value), period_count)







        applied = True







    if _aq_apply_loan_loss_allowance(merged, settlement_store):



        applied = True







    if applied:



        merged["settlement_applied"] = True



    return merged











def _asset_quality_sheet(wb, preferred_name=None):



    if preferred_name and preferred_name in wb.sheetnames:



        return wb[preferred_name]



    for sheet_name in wb.sheetnames:



        if "\uc790\uc0b0\uac74\uc804\uc131" in sheet_name:



            return wb[sheet_name]



    return wb.active











def _aq_period_column_map(ws):



    mapping = {}



    for col in range(2, ws.max_column + 1):



        period = _aq_norm_period(ws.cell(row=3, column=col).value)



        if period:



            mapping[period] = col



    return mapping















AQ_UNIT_OPTIONS = {



    "won": {"label": "\uc6d0", "factor": 1},



    "million": {"label": "\ubc31\ub9cc\uc6d0", "factor": 1 / 1_000_000},



    "ten_million": {"label": "\ucc9c\ub9cc\uc6d0", "factor": 1 / 10_000_000},



}



AQ_UNIT_ALIASES = {



    "\uc6d0": "won",



    "\ubc31\ub9cc": "million",



    "\ubc31\ub9cc\uc6d0": "million",



    "\ucc9c\ub9cc": "ten_million",



    "\ucc9c\ub9cc\uc6d0": "ten_million",



}



AQ_SECTION2_AMOUNT_ROWS = {31, 34, 35}



AQ_SECTION2_EXPORT_ROWS = [



    {"row": 19, "label": "무연체", "key": "무연체", "kind": "pct"},



    {"row": 20, "label": "1~30", "key": "1~30", "kind": "pct"},



    {"row": 21, "label": "1~10", "key": "1~10", "kind": "pct"},



    {"row": 22, "label": "11~30", "key": "11~30", "kind": "pct"},



    {"row": 23, "label": "31~60", "key": "31~60", "kind": "pct"},



    {"row": 24, "label": "61~90", "key": "61~90", "kind": "pct"},



    {"row": 25, "label": "91~120", "key": "91~120", "kind": "pct"},



    {"row": 26, "label": "121~150", "key": "121~150", "kind": "pct"},



    {"row": 27, "label": "151~180", "key": "151~180", "kind": "pct"},



    {"row": 28, "label": "181~", "key": "181~", "kind": "pct"},



    {"row": 29, "label": "합계(%)", "key": "__total_pct__", "kind": "pct"},



    {"row": 30, "label": "1일이상 연체율", "key": "연체율(%)", "kind": "pct"},



    {"row": 31, "label": "연체합계", "key": "연체합계", "kind": "amount"},



    {"row": 32, "label": "", "key": None, "kind": "blank"},



    {"row": 33, "label": "31일이상 연체율", "key": "30일이상연체율", "kind": "pct"},



    {"row": 34, "label": "31일이상 연체합계", "key": "31일이상연체합계", "kind": "amount"},



    {"row": 35, "label": "대손충당금", "key": "대손충당금", "kind": "amount"},



    {"row": 36, "label": "연체대비 충당비율", "key": "__allowance_to_31plus_ratio__", "kind": "pct"},



]



AQ_AMOUNT_NUMBER_FORMAT = r'#,##0;[Red]\(#,##0\)'



AQ_PERCENT_NUMBER_FORMAT = '0.00%'







def _aq_unit_key(unit):



    key = str(unit or "million").strip()



    key = AQ_UNIT_ALIASES.get(key, key)



    if key not in AQ_UNIT_OPTIONS:



        key = "million"



    return key











def _aq_unit_label(unit):



    return AQ_UNIT_OPTIONS[_aq_unit_key(unit)]["label"]











def _aq_update_unit_labels(ws, unit):



    label = _aq_unit_label(unit)



    unit_pattern = re.compile(r'(\([^)]*:\s*)([^)]*)(\))')



    for row in range(1, min(ws.max_row, 120) + 1):



        for col in range(1, min(ws.max_column, 3) + 1):



            cell = ws.cell(row=row, column=col)



            if isinstance(cell.value, str) and "\ub2e8\uc704" in cell.value:



                cell.value = unit_pattern.sub(lambda match: f"{match.group(1)}{label}{match.group(3)}", cell.value)











def _aq_scale_amount(value, unit):



    if value is None:



        return None



    if not isinstance(value, (int, float)):



        return value



    factor = AQ_UNIT_OPTIONS[_aq_unit_key(unit)]["factor"]



    return value * factor







def _aq_section2_total_pct_values(data):



    section1 = data.get("section1") or {}



    periods = data.get("periods") or []



    total_values = section1.get("\uc735\uc794\ud569\uacc4") or []



    keys = ["\ubb34\uc5f0\uccb4", "1~30", "31~60", "61~90", "91~120", "121~150", "151~180", "181~"]



    result = []



    for idx, _ in enumerate(periods):



        total = total_values[idx] if idx < len(total_values) else None



        if not isinstance(total, (int, float)) or total == 0:



            result.append(None)



            continue



        amount_sum = 0



        has_value = False



        for key in keys:



            values = section1.get(key) or []



            value = values[idx] if idx < len(values) else None



            if isinstance(value, (int, float)):



                amount_sum += value



                has_value = True



        result.append(amount_sum / total if has_value else None)



    return result



















def _aq_section2_allowance_to_31plus_ratio_values(data):



    section2 = data.get("section2") or {}



    periods = data.get("periods") or []



    allowances = section2.get("\ub300\uc190\ucda9\ub2f9\uae08") or []



    overdue31_values = section2.get("31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4") or []



    result = []



    for idx, _ in enumerate(periods):



        allowance = allowances[idx] if idx < len(allowances) else None



        overdue31 = overdue31_values[idx] if idx < len(overdue31_values) else None



        if not isinstance(allowance, (int, float)) or not isinstance(overdue31, (int, float)) or overdue31 == 0:



            result.append(None)



            continue



        result.append(allowance / overdue31)



    return result







def _aq_write_series_row(ws, row, periods, period_columns, values, unit="million", scale_amount=True, number_format=None):



    if not isinstance(values, list):



        return



    for idx, value in enumerate(values):



        if idx >= len(periods):



            break



        col = period_columns.get(_aq_norm_period(periods[idx]))



        if col and value is not None:



            cell = ws.cell(row=row, column=col)



            cell.value = _aq_scale_amount(value, unit) if scale_amount else value



            if number_format:



                cell.number_format = number_format



















def _aq_copy_row_style(ws, source_row, target_row):



    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height



    for col in range(1, ws.max_column + 1):



        source = ws.cell(row=source_row, column=col)



        target = ws.cell(row=target_row, column=col)



        if source.has_style:



            target._style = copy.copy(source._style)



        if source.number_format:



            target.number_format = source.number_format



        if source.alignment:



            target.alignment = copy.copy(source.alignment)



        if source.protection:



            target.protection = copy.copy(source.protection)



        if source.comment:



            target.comment = copy.copy(source.comment)



        target.value = None











def _aq_append_product_block(ws, source_base_row, target_base_row, product_name):



    for offset in range(0, 10):



        _aq_copy_row_style(ws, source_base_row + offset, target_base_row + offset)



    ws.cell(row=target_base_row, column=1).value = product_name



    for offset, label in PRODUCT_ROW_OFFSETS.items():



        ws.cell(row=target_base_row + offset, column=1).value = label











def _aq_product_row_map(ws, data):



    product_rows = {product: base_row for product, base_row in SECTION4_PRODUCTS}



    section4 = data.get("section4") or {}



    product_names = list(data.get("products") or [])



    for product in section4.keys():



        if product not in product_names:



            product_names.append(product)







    source_base_row = SECTION4_PRODUCTS[-1][1]



    block_step = 11



    next_base_row = source_base_row + block_step



    for product in product_names:



        if product in product_rows or product not in section4:



            continue



        _aq_append_product_block(ws, source_base_row, next_base_row, product)



        product_rows[product] = next_base_row



        next_base_row += block_step



    return product_rows, product_names











def _aq_amount_rows(product_rows):



    rows = set(SECTION1_ROWS.keys()) | set(AQ_SECTION2_AMOUNT_ROWS) | {17}



    for group_info in SECTION3_GROUPS.values():



        group_rows = group_info.get("rows", group_info) if isinstance(group_info, dict) else {}



        for row, label in group_rows.items():



            if label != "\uc5f0\uccb4\uc728(%)":



                rows.add(row)



    for base_row in product_rows.values():



        for offset in PRODUCT_ROW_OFFSETS.keys():



            rows.add(base_row + offset)



    return rows











def _aq_normalize_leftover_amount_formats(ws, unit, period_columns, product_rows):



    factor = AQ_UNIT_OPTIONS[_aq_unit_key(unit)]["factor"]



    amount_rows = _aq_amount_rows(product_rows)



    period_cols = set(period_columns.values())



    for row_cells in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=ws.max_column):



        if not row_cells or row_cells[0].row not in amount_rows:



            continue



        for cell in row_cells:



            if cell.column not in period_cols or ",," not in str(cell.number_format):



                continue



            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):



                cell.value = cell.value * factor



            cell.number_format = AQ_AMOUNT_NUMBER_FORMAT











def _write_asset_quality_to_sheet(ws, data, unit="million"):



    unit = _aq_unit_key(unit)



    periods = [_aq_norm_period(p) for p in (data.get("periods") or [])]



    period_columns = _aq_period_column_map(ws)



    ws.cell(row=2, column=1).value = f"(\ub2e8\uc704: {_aq_unit_label(unit)})"



    _aq_update_unit_labels(ws, unit)







    section1 = data.get("section1") or {}



    for row, label in SECTION1_ROWS.items():



        _aq_write_series_row(ws, row, periods, period_columns, section1.get(label), unit=unit, number_format=AQ_AMOUNT_NUMBER_FORMAT)







    section2 = data.get("section2") or {}



    section2_total_pct = _aq_section2_total_pct_values(data)



    section2_allowance_to_31plus_ratio = _aq_section2_allowance_to_31plus_ratio_values(data)



    period_cols = list(period_columns.values())



    for item in AQ_SECTION2_EXPORT_ROWS:



        row = item["row"]



        kind = item.get("kind")



        ws.cell(row=row, column=1).value = item.get("label") or None



        for col in period_cols:



            ws.cell(row=row, column=col).value = None



        if kind == "blank":



            continue



        key = item.get("key")



        if key == "__total_pct__":



            values = section2_total_pct



        elif key == "__allowance_to_31plus_ratio__":



            values = section2_allowance_to_31plus_ratio



        else:



            values = section2.get(key)



        is_amount = kind == "amount"



        number_format = AQ_AMOUNT_NUMBER_FORMAT if is_amount else AQ_PERCENT_NUMBER_FORMAT if kind == "pct" else None



        _aq_write_series_row(



            ws, row, periods, period_columns, values,



            unit=unit,



            scale_amount=is_amount,



            number_format=number_format,



        )



    _aq_write_series_row(



        ws, 17, periods, period_columns, section2.get("\u0033\u0031\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4"),



        unit=unit,



        number_format=AQ_AMOUNT_NUMBER_FORMAT,



    )







    section3 = data.get("section3") or {}



    for group, group_info in SECTION3_GROUPS.items():



        group_data = section3.get(group) or {}



        rows = group_info.get("rows", group_info) if isinstance(group_info, dict) else {}



        for row, label in rows.items():



            is_amount = label != "\uc5f0\uccb4\uc728(%)"



            _aq_write_series_row(



                ws, row, periods, period_columns, group_data.get(label),



                unit=unit,



                scale_amount=is_amount,



                number_format=AQ_AMOUNT_NUMBER_FORMAT if is_amount else None,



            )







    section4 = data.get("section4") or {}



    product_rows, product_names = _aq_product_row_map(ws, data)



    for product in product_names:



        base_row = product_rows.get(product)



        product_data = section4.get(product) or {}



        if not base_row or not product_data:



            continue



        for offset, label in PRODUCT_ROW_OFFSETS.items():



            _aq_write_series_row(ws, base_row + offset, periods, period_columns, product_data.get(label), unit=unit, number_format=AQ_AMOUNT_NUMBER_FORMAT)







    _aq_normalize_leftover_amount_formats(ws, unit, period_columns, product_rows)







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



    """Upload asset-quality Excel and keep the original workbook template."""



    if not file.filename.lower().endswith(('.xlsx', '.xls')):



        raise HTTPException(status_code=400, detail="\uc5d1\uc140 \ud30c\uc77c(.xlsx, .xls)\ub9cc \uc5c5\ub85c\ub4dc \uac00\ub2a5\ud569\ub2c8\ub2e4.")







    tmp_path = None



    try:



        content = await file.read()



        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:



            tmp.write(content)



            tmp_path = tmp.name







        data = parse_asset_quality_excel(tmp_path)



        if period:



            data['upload_period'] = period



        uploaded_data['asset_quality_detail'] = data







        sheet_name = None



        wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=False)



        try:



            sheet_name = _asset_quality_sheet(wb).title



        finally:



            wb.close()







        template_path = _file_path("asset_quality_template.xlsx")



        with open(template_path, "wb") as template_file:



            template_file.write(content)



        uploaded_data["asset_quality_template"] = {



            "filename": file.filename,



            "sheet_name": sheet_name,



        }







        return JSONResponse({



            "success": True,



            "message": f"\uc790\uc0b0\uac74\uc804\uc131 '{file.filename}' \uc5c5\ub85c\ub4dc \uc644\ub8cc",



            "periods": len(data.get('periods', [])),



            "products": len(data.get('products', [])),



            "upload_period": period or "(\ubbf8\uc9c0\uc815)",



        })



    except Exception as e:



        raise HTTPException(status_code=500, detail=f"\ud30c\uc77c \ucc98\ub9ac \uc624\ub958: {str(e)}")



    finally:



        if tmp_path and os.path.exists(tmp_path):



            os.unlink(tmp_path)











@app.get("/api/data/asset-quality-detail")



async def get_asset_quality_detail():



    """Return asset-quality detail data with settlement months applied."""



    d = uploaded_data.get("asset_quality_detail")



    if not d:



        return JSONResponse({"error": "\ub370\uc774\ud130 \uc5c6\uc74c"}, status_code=404)



    merged = _merge_settlement_into_asset_quality_detail(d, uploaded_data.get("settlement_aq") or {})



    return JSONResponse(merged)











@app.get("/api/export/asset-quality-detail")



async def export_asset_quality_detail(unit: str = "million"):



    """Export current asset-quality data using the uploaded workbook template."""



    data = uploaded_data.get("asset_quality_detail")



    if not data:



        raise HTTPException(status_code=404, detail="\uc790\uc0b0\uac74\uc804\uc131 \ub370\uc774\ud130\uac00 \uc5c6\uc2b5\ub2c8\ub2e4. \uc790\uc0b0\uac74\uc804\uc131 \uc5d1\uc140\uc744 \uba3c\uc800 \uc5c5\ub85c\ub4dc\ud558\uc138\uc694.")







    template_path = _stored_file_path("asset_quality_template.xlsx")



    if not template_path:



        raise HTTPException(status_code=404, detail="\uc790\uc0b0\uac74\uc804\uc131 \uc6d0\ubcf8 \uc5d1\uc140 \uc11c\uc2dd\uc774 \uc5c6\uc2b5\ub2c8\ub2e4. \uc790\uc0b0\uac74\uc804\uc131 \uc5d1\uc140\uc744 \ub2e4\uc2dc \uc5c5\ub85c\ub4dc\ud558\uc138\uc694.")







    try:



        meta = uploaded_data.get("asset_quality_template", {}) or {}



        merged = _merge_settlement_into_asset_quality_detail(data, uploaded_data.get("settlement_aq") or {})



        wb = openpyxl.load_workbook(template_path, data_only=True, keep_links=False)



        ws = _asset_quality_sheet(wb, meta.get("sheet_name"))



        unit = _aq_unit_key(unit)



        _write_asset_quality_to_sheet(ws, merged, unit=unit)







        buf = io.BytesIO()



        wb.save(buf)



        buf.seek(0)







        from urllib.parse import quote



        filename = f"\uae30\uc5c5\uc5ec\uc2e0\uc790\ub8cc_\uc790\uc0b0\uac74\uc804\uc131_{_aq_unit_label(unit)}.xlsx"



        encoded_filename = quote(filename)



        return StreamingResponse(



            buf,



            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",



            headers={



                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",



                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",



                "Pragma": "no-cache",



                "X-Asset-Quality-Unit": unit,



            },



        )



    except HTTPException:



        raise



    except Exception as e:



        raise HTTPException(status_code=500, detail=f"\uc790\uc0b0\uac74\uc804\uc131 \uc5d1\uc140 \uc0dd\uc131 \uc624\ub958: {str(e)}")











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







        # ── 성별 → saq_store 저장 (asset 섹션은 /api/data/asset에서 동적 구성) ──



        # channel amount -> asset_data.sections

        channel_amount = data.get("channel_amount", {})

        if channel_amount and pkey and pkey != "unknown":

            asset_data = uploaded_data.get("asset_data")

            if asset_data is not None:

                sections = asset_data.get("sections", {})

                sec_channel = sections.get("접수경로", {})

                for row_key, val in channel_amount.items():

                    if val is not None:

                        if row_key not in sec_channel:

                            sec_channel[row_key] = {}

                        sec_channel[row_key][pkey] = float(val) * 1_000_000

                sections["접수경로"] = sec_channel

                asset_data["sections"] = sections

                uploaded_data["asset_data"] = asset_data

            if pkey in saq_store:

                saq_store[pkey]["channel_amount"] = channel_amount

                uploaded_data["settlement_aq"] = saq_store



        gender = data.get("gender", {})



        if gender and pkey in saq_store:



            saq_store[pkey]["gender"] = gender



            uploaded_data["settlement_aq"] = saq_store







        # ── 연령별 → saq_store 저장 ──────────────────────────────────



        age_band_data = data.get("age_band", {})



        if age_band_data and pkey in saq_store:



            saq_store[pkey]["age_band"] = age_band_data



            uploaded_data["settlement_aq"] = saq_store







        # ── 직업분류 → saq_store 저장 ────────────────────────────────



        job_raw_data = data.get("job_raw", {})



        if job_raw_data and pkey in saq_store:



            saq_store[pkey]["job_raw"] = job_raw_data



            uploaded_data["settlement_aq"] = saq_store







        # ── 지역별 → saq_store 저장 ──────────────────────────────────



        region_data = data.get("region", {})



        if region_data and pkey in saq_store:



            saq_store[pkey]["region"] = region_data



            uploaded_data["settlement_aq"] = saq_store







        # ── CB등급(NICE스코어) → saq_store 저장 ──────────────────────



        # cb_raw: {점수(int): 백만원} 형태 → JSON 직렬화 위해 str 키로 저장



        cb_raw_data = data.get("cb_raw", {})



        if cb_raw_data and pkey in saq_store:



            # 점수 키를 문자열로 변환하여 저장



            saq_store[pkey]["cb_raw"] = {str(k): v for k, v in cb_raw_data.items()}



            uploaded_data["settlement_aq"] = saq_store







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


            "channel_amount_updated": bool(channel_amount and pkey and pkey != "unknown" and uploaded_data.get("asset_data") is not None),



            "gender_updated":    bool(gender      and pkey and pkey != "unknown"),



            "age_band_updated2": bool(age_band_data and pkey and pkey != "unknown"),



            "job_updated":       bool(job_raw_data and pkey and pkey != "unknown"),



            "region_updated":    bool(region_data  and pkey and pkey != "unknown"),



            "cb_raw_updated":    bool(cb_raw_data  and pkey and pkey != "unknown"),



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



    """Return product names. Prefer settlement upload products, then AQ detail products."""



    saq = uploaded_data.get("settlement_aq") or {}



    if not saq:



        _saq_file = _data_path("settlement_aq")



        if os.path.exists(_saq_file):



            with open(_saq_file, encoding="utf-8") as f:



                saq = json.load(f)



    if saq:



        prod_set = set()



        for _pkey, pdata in saq.items():



            for prod in (pdata.get("products") or []):



                if prod:



                    prod_set.add(str(prod).strip())



        if prod_set:



            return JSONResponse(sorted(prod_set))



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











# -- Loan loss allowance / reserve calculation ------------------------------



LOSS_TARGET_GROUPS = ["\ub2f4\ubcf4", "\uc2e0\uc6a9"]



LOSS_BUCKET_TEMPLATE = [



    {"key": "1~10", "label": "\uc5f0\uccb4(1~10)", "lo": 1, "hi": 10},



    {"key": "11~30", "label": "\uc5f0\uccb4(11~30)", "lo": 11, "hi": 30},



    {"key": "31~60", "label": "\uc5f0\uccb4(31~60)", "lo": 31, "hi": 60},



    {"key": "61~90", "label": "\uc5f0\uccb4(61~90)", "lo": 61, "hi": 90},



    {"key": "91~180", "label": "\uc5f0\uccb4(91~180)", "lo": 91, "hi": 180},



    {"key": "181~", "label": "\uc5f0\uccb4(181~)", "lo": 181, "hi": None},



]



LOSS_FIXED_BUCKETS = [



    {"key": "1~10", "lo": 1, "hi": 10},



    {"key": "11~30", "lo": 11, "hi": 30},



    {"key": "31~60", "lo": 31, "hi": 60},



    {"key": "61~90", "lo": 61, "hi": 90},



    {"key": "91~120", "lo": 91, "hi": 120},



    {"key": "121~150", "lo": 121, "hi": 150},



    {"key": "151~180", "lo": 151, "hi": 180},



    {"key": "181~", "lo": 181, "hi": None},



]











def _loss_default_rate(group_name: str, bucket_key: str) -> float:



    if "\ub2f4\ubcf4" in str(group_name):



        return {"61~90": 50, "91~180": 80, "181~": 100}.get(bucket_key, 0)



    return {"11~30": 20, "31~60": 50, "61~90": 50, "91~180": 80, "181~": 100}.get(bucket_key, 0)











def _loss_default_buckets(group_name: str) -> list[dict]:



    return [{**bucket, "rate": _loss_default_rate(group_name, bucket["key"])} for bucket in LOSS_BUCKET_TEMPLATE]











def _loss_product_group_map() -> dict[str, list[str]]:



    groups = uploaded_data.get("product_groups", []) or []



    result: dict[str, list[str]] = {}



    for group in groups:



        name = str(group.get("name") or "").strip()



        if name:



            result[name] = [str(p).strip() for p in (group.get("products") or []) if str(p).strip()]



    return result











def _loss_normalize_bucket(bucket: dict, fallback: dict) -> dict:



    def _to_int(value, default=None):



        if value in (None, ""):



            return default



        try:



            return int(float(value))



        except (TypeError, ValueError):



            return default







    def _to_float(value, default=0.0):



        if value in (None, ""):



            return default



        try:



            return float(value)



        except (TypeError, ValueError):



            return default







    key = str(bucket.get("key") or fallback.get("key") or "").strip()



    label = str(bucket.get("label") or fallback.get("label") or key).strip()



    return {



        "key": key,



        "label": label,



        "lo": _to_int(bucket.get("lo"), fallback.get("lo")),



        "hi": _to_int(bucket.get("hi"), fallback.get("hi")),



        "rate": max(0.0, _to_float(bucket.get("rate"), fallback.get("rate", 0))),



    }











def _loss_config_with_products() -> dict:



    product_map = _loss_product_group_map()



    saved = uploaded_data.get("loss_rate_config", {}) or {}



    saved_groups = saved.get("groups") if isinstance(saved, dict) else []



    saved_map = {



        str(group.get("name") or "").strip(): group



        for group in (saved_groups or [])



        if isinstance(group, dict) and str(group.get("name") or "").strip()



    }



    names = [name for name in LOSS_TARGET_GROUPS if name in product_map]



    for name in LOSS_TARGET_GROUPS:



        if name not in names:



            names.append(name)



    groups = []



    for name in names:



        defaults = _loss_default_buckets(name)



        saved_buckets = (saved_map.get(name) or {}).get("buckets") or []



        buckets = []



        for idx2, fallback in enumerate(defaults):



            source = saved_buckets[idx2] if idx2 < len(saved_buckets) and isinstance(saved_buckets[idx2], dict) else fallback



            buckets.append(_loss_normalize_bucket(source, fallback))



        for extra in saved_buckets[len(defaults):]:



            if isinstance(extra, dict):



                buckets.append(_loss_normalize_bucket(extra, {"key": "custom", "label": "\uc0ac\uc6a9\uc790\uad6c\uac04", "lo": 1, "hi": None, "rate": 0}))



        products = product_map.get(name)



        if products is None:



            products = (saved_map.get(name) or {}).get("products") or []



        groups.append({"name": name, "products": products, "buckets": buckets})



    return {"groups": groups}











def _loss_period_label(period: str) -> str:



    m = re.search(r'(\d{4})[-.](\d{1,2})', str(period or ""))



    if not m:



        return str(period or "")



    return f"{m.group(1)}.{int(m.group(2))}\uc6d4"











def _loss_rate_to_allowance(amount: float, rate: float) -> int:



    return int(round((amount or 0) * (rate or 0) / 100))











def _loss_bucket_contains(config_bucket: dict, fixed_bucket: dict) -> bool:



    lo = config_bucket.get("lo")



    hi = config_bucket.get("hi")



    flo = fixed_bucket.get("lo")



    fhi = fixed_bucket.get("hi")



    if lo is not None and flo < lo:



        return False



    if hi is None:



        return True



    if fhi is None:



        return False



    return fhi <= hi











def _loss_amount_from_section4(section4: dict, products: list[str], bucket_key: str) -> float:



    total = 0.0



    for product in products:



        value = (section4.get(product) or {}).get(bucket_key)



        if isinstance(value, (int, float)):



            total += float(value) * 1_000_000



    return total











def _loss_aggregate_from_section4(period_data: dict, group_config: dict) -> dict:



    section4 = period_data.get("section4") or {}



    products = group_config.get("products") or []



    normal = _loss_amount_from_section4(section4, products, "\ubb34\uc5f0\uccb4")



    overdue_total = _loss_amount_from_section4(section4, products, "\uc5f0\uccb4\ud569\uacc4")



    bucket_amounts = []



    for bucket in group_config.get("buckets") or []:



        amount = 0.0



        for fixed in LOSS_FIXED_BUCKETS:



            if _loss_bucket_contains(bucket, fixed):



                amount += _loss_amount_from_section4(section4, products, fixed["key"])



        bucket_amounts.append({**bucket, "amount": amount})



    if not overdue_total:



        overdue_total = sum(item["amount"] for item in bucket_amounts)



    return {"normal": normal, "overdue_total": overdue_total, "buckets": bucket_amounts}











def _loss_aggregate_from_rows(period_data: dict, group_config: dict) -> dict | None:



    rows = period_data.get("loan_loss_rows")



    if not isinstance(rows, list) or not rows:



        return None



    products = set(group_config.get("products") or [])



    bucket_amounts = [{**bucket, "amount": 0.0} for bucket in (group_config.get("buckets") or [])]



    normal = 0.0



    overdue_total = 0.0



    for row in rows:



        product = str(row.get("product") or "").strip()



        if product not in products:



            continue



        try:



            days = int(float(row.get("days") or 0))



            balance = float(row.get("balance") or 0)



        except (TypeError, ValueError):



            continue



        if days <= 0:



            normal += balance



            continue



        overdue_total += balance



        for bucket in bucket_amounts:



            lo = bucket.get("lo")



            hi = bucket.get("hi")



            if lo is not None and days < lo:



                continue



            if hi is not None and days > hi:



                continue



            bucket["amount"] += balance



            break



    return {"normal": normal, "overdue_total": overdue_total, "buckets": bucket_amounts}











def _loss_build_group_block(group_config: dict, period_data: dict) -> dict:



    agg = _loss_aggregate_from_rows(period_data, group_config) or _loss_aggregate_from_section4(period_data, group_config)



    rows = [



        {"kind": "normal", "label": "\uc815 \uc0c1", "amount": int(round(agg["normal"])), "rate": None, "allowance": None},



        {"kind": "overdue_total", "label": "\ucd1d \uc5f0\uccb4", "amount": int(round(agg["overdue_total"])), "rate": None, "allowance": None},



    ]



    total_allowance = 0



    for bucket in agg["buckets"]:



        amount = int(round(bucket.get("amount") or 0))



        rate = float(bucket.get("rate") or 0)



        allowance = _loss_rate_to_allowance(amount, rate)



        total_allowance += allowance



        rows.append({"kind": "bucket", "key": bucket.get("key"), "label": bucket.get("label"), "lo": bucket.get("lo"), "hi": bucket.get("hi"), "amount": amount, "rate": rate, "allowance": allowance})



    rows.append({"kind": "total", "label": "\uc804 \uccb4", "amount": int(round(agg["normal"] + agg["overdue_total"])), "rate": None, "allowance": total_allowance})



    return {"name": group_config.get("name"), "products": group_config.get("products") or [], "rows": rows, "allowance": total_allowance}











def _loss_build_total_block(group_blocks: list[dict]) -> dict:



    row_map: dict[str, dict] = {}



    row_order: list[str] = []



    for block in group_blocks:



        for row in block.get("rows") or []:



            key = row.get("key") or row.get("kind") or row.get("label")



            if key not in row_map:



                row_map[key] = {**row, "amount": 0, "allowance": 0 if row.get("allowance") is not None else None}



                row_order.append(key)



            row_map[key]["amount"] += row.get("amount") or 0



            if row.get("allowance") is not None:



                row_map[key]["allowance"] = (row_map[key].get("allowance") or 0) + row.get("allowance")



            if row.get("rate") is not None:



                row_map[key]["rate"] = max(row_map[key].get("rate") or 0, row.get("rate") or 0)



    return {"name": "\ud569\uacc4", "products": [], "rows": [row_map[key] for key in row_order], "allowance": sum(block.get("allowance") or 0 for block in group_blocks)}











def _loss_calculate_period(period: str, period_data: dict, config: dict) -> dict:



    precomputed_blocks = period_data.get("loan_loss_blocks") if isinstance(period_data, dict) else None



    if isinstance(precomputed_blocks, list) and precomputed_blocks:



        total_block = next((b for b in precomputed_blocks if b.get("name") == "\ud569\uacc4"), precomputed_blocks[-1])



        total_row = next((r for r in (total_block.get("rows") or []) if r.get("kind") == "total"), {})



        total_balance = total_row.get("amount") or 0



        total_allowance = total_row.get("allowance")



        if total_allowance in (None, ""):



            total_allowance = total_block.get("allowance") or 0



        return {



            "period": period,



            "period_label": period_data.get("period_label") or _loss_period_label(period),



            "blocks": precomputed_blocks,



            "summary": {



                "total_balance": total_balance,



                "total_allowance": total_allowance,



                "allowance_rate": round(float(total_allowance or 0) / float(total_balance or 0) * 100, 2) if total_balance else 0,



            },



        }







    group_blocks = [_loss_build_group_block(group, period_data) for group in config.get("groups") or []]



    total_block = _loss_build_total_block(group_blocks)



    total_balance = 0



    for row in total_block.get("rows") or []:



        if row.get("kind") == "total":



            total_balance = row.get("amount") or 0



            break



    total_allowance = total_block.get("allowance") or 0



    return {"period": period, "period_label": _loss_period_label(period), "blocks": group_blocks + [total_block], "summary": {"total_balance": total_balance, "total_allowance": total_allowance, "allowance_rate": round(total_allowance / total_balance * 100, 2) if total_balance else 0}}











@app.get("/api/loss-rate-config")



async def get_loss_rate_config():



    return JSONResponse(_loss_config_with_products())











@app.post("/api/loss-rate-config")



async def save_loss_rate_config(request: Request):



    body = await request.json()



    if not isinstance(body, dict):



        raise HTTPException(status_code=400, detail="\uac1d\uccb4 \ud615\ud0dc\ub85c \uc804\uc1a1\ud558\uc138\uc694")



    uploaded_data["loss_rate_config"] = {"groups": body.get("groups") or []}



    return JSONResponse({"success": True})











@app.get("/api/data/loan-loss-allowance")



async def get_loan_loss_allowance(period: str = ""):



    store = uploaded_data.get("settlement_aq") or {}



    if not store:



        return JSONResponse({"error": "\ub370\uc774\ud130 \uc5c6\uc74c"}, status_code=404)



    config = _loss_config_with_products()



    periods = sorted(store.keys())



    selected_periods = [period] if period and period in store else periods



    by_period = {p: _loss_calculate_period(p, store[p], config) for p in selected_periods}



    latest = selected_periods[-1] if selected_periods else None



    return JSONResponse({"periods": periods, "latest_period": latest, "unit": "\uc6d0", "config": config, "by_period": by_period})











@app.get("/api/export/loan-loss-allowance")



async def export_loan_loss_allowance(period: str = ""):



    """Export loan-loss allowance using the current screen layout."""



    store = uploaded_data.get("settlement_aq") or {}



    if not store:



        raise HTTPException(status_code=404, detail="\uacb0\uc0b0\uc790\ub8cc\uac00 \uc5c6\uc2b5\ub2c8\ub2e4. \uba3c\uc800 \uacb0\uc0b0\uc790\ub8cc\ub97c \uc5c5\ub85c\ub4dc\ud558\uc138\uc694.")







    periods = sorted(store.keys())



    selected = period if period in store else periods[-1]



    data = _loss_calculate_period(selected, store[selected], _loss_config_with_products())







    try:



        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side



        from urllib.parse import quote







        wb = openpyxl.Workbook()



        ws = wb.active



        ws.title = "\ub300\uc190"



        ws.sheet_view.showGridLines = False







        # Original-style compact columns: B~F are the report body.



        widths = {"A": 2, "B": 10, "C": 18, "D": 18, "E": 12, "F": 18}



        for col, width in widths.items():



            ws.column_dimensions[col].width = width







        navy_fill = PatternFill("solid", fgColor="1E3A5F")



        header_fill = PatternFill("solid", fgColor="D9EAF7")



        group_fill = PatternFill("solid", fgColor="EAF7EF")



        normal_fill = PatternFill("solid", fgColor="EAF7EF")



        overdue_fill = PatternFill("solid", fgColor="FFF7ED")



        total_fill = PatternFill("solid", fgColor="FFF7BF")



        grand_fill = PatternFill("solid", fgColor="E0F2FE")



        white_font = Font(color="FFFFFF", bold=True, size=11)



        title_font = Font(color="0B1F3A", bold=True, size=14)



        header_font = Font(color="0B1F3A", bold=True, size=10)



        total_font = Font(color="111827", bold=True, size=10)



        red_font = Font(color="DC2626", bold=True, size=10)



        base_font = Font(color="0B1F3A", size=10)



        center = Alignment(horizontal="center", vertical="center")



        left = Alignment(horizontal="left", vertical="center")



        right = Alignment(horizontal="right", vertical="center")



        thin_side = Side(style="thin", color="C9D6E6")



        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)







        ws.merge_cells("B2:C2")



        ws.merge_cells("D2:F2")



        ws["B2"] = "\ucc44 \uad8c \ud604 \ud669"



        ws["D2"] = data.get("period_label") or _loss_period_label(selected)



        for cell in (ws["B2"], ws["D2"]):



            cell.fill = navy_fill



            cell.font = white_font



            cell.alignment = center



            cell.border = border



        for row in ws["B2:F2"]:



            for cell in row:



                cell.border = border



                if cell.value is None:



                    cell.fill = navy_fill







        headers = ["\uad6c\ubd84", "\ucc44\uad8c \uc0c1\ud0dc", "\uae08\uc561", "\ub300\uc190\uc728", "\ub300\uc190\ucda9\ub2f9\uae08"]



        row_idx = 4



        for col_idx, label in enumerate(headers, 2):



            c = ws.cell(row=row_idx, column=col_idx, value=label)



            c.fill = header_fill



            c.font = header_font



            c.alignment = center



            c.border = border



        row_idx += 1







        def _write_block(block: dict, start_row: int) -> int:



            rows = block.get("rows") or []



            if not rows:



                return start_row



            is_grand = block.get("name") == "\ud569\uacc4"



            end_row = start_row + len(rows) - 1



            if end_row > start_row:



                ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)



            group_cell = ws.cell(row=start_row, column=2, value=block.get("name") or "")



            group_cell.fill = grand_fill if is_grand else group_fill



            group_cell.font = total_font



            group_cell.alignment = center



            group_cell.border = border



            for rr in range(start_row, end_row + 1):



                ws.cell(row=rr, column=2).border = border



                ws.cell(row=rr, column=2).fill = grand_fill if is_grand else group_fill







            for offset, item in enumerate(rows):



                r = start_row + offset



                kind = item.get("kind")



                fill = None



                font = base_font



                if kind == "overdue_total":



                    fill = overdue_fill



                    font = total_font



                elif kind == "total":



                    fill = grand_fill if is_grand else total_fill



                    font = total_font



                elif kind == "normal":



                    fill = normal_fill



                if is_grand and kind != "total":



                    fill = grand_fill if kind in ("normal", "overdue_total") else None







                values = [



                    item.get("label") or "",



                    item.get("amount") or 0,



                    None if item.get("rate") in (None, "") else float(item.get("rate") or 0) / 100,



                    item.get("allowance"),



                ]



                for col_idx, value in enumerate(values, 3):



                    c = ws.cell(row=r, column=col_idx, value=value)



                    c.font = red_font if (kind == "total" and is_grand and col_idx in (4, 6)) else font



                    c.border = border



                    if fill:



                        c.fill = fill



                    if col_idx == 3:



                        c.alignment = left



                    else:



                        c.alignment = right



                    if col_idx in (4, 6):



                        c.number_format = '#,##0'



                    elif col_idx == 5:



                        c.number_format = '0%'



                    if col_idx == 6 and value is None:



                        c.value = None



                ws.row_dimensions[r].height = 20



            return end_row + 2







        for block in data.get("blocks") or []:



            row_idx = _write_block(block, row_idx)







        ws.freeze_panes = "B5"



        ws.page_margins.left = 0.25



        ws.page_margins.right = 0.25



        ws.page_margins.top = 0.5



        ws.page_margins.bottom = 0.5



        ws.page_setup.orientation = "portrait"



        ws.page_setup.fitToWidth = 1



        ws.page_setup.fitToHeight = 0







        buf = io.BytesIO()



        wb.save(buf)



        buf.seek(0)



        filename = f"\uae30\uc5c5\uc5ec\uc2e0\uc790\ub8cc_\ub300\uc190\ucda9\ub2f9\uae08_{data.get('period_label') or selected}.xlsx"



        encoded_filename = quote(filename)



        return StreamingResponse(



            buf,



            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",



            headers={



                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",



                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",



                "Pragma": "no-cache",



            },



        )



    except HTTPException:



        raise



    except Exception as e:



        raise HTTPException(status_code=500, detail=f"\ub300\uc190\ucda9\ub2f9\uae08 \uc5d1\uc140 \uc0dd\uc131 \uc624\ub958: {str(e)}")











@app.get("/api/channel-groups")



async def get_channel_groups():



    """저장된 접수경로 그룹 목록 반환"""



    groups = uploaded_data.get("channel_groups", [])



    # 파일 폴백



    if not groups:



        _path = _data_path("channel_groups")



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



    _path = _data_path("channel_groups")



    os.makedirs(os.path.dirname(_path), exist_ok=True)



    with open(_path, "w", encoding="utf-8") as f:



        json.dump(body, f, ensure_ascii=False, indent=2)



    return JSONResponse({"success": True, "count": len(body)})











# ── 일 마감보고 설정 API ──────────────────────────────────────────

DAILY_CLOSE_CONFIG_KEY = "daily_close_report_config"
DAILY_CLOSE_EXCELLENT_TARGET_MILLION = 12100
DAILY_CLOSE_EXCELLENT_RECORDS_KEY = "daily_close_excellent_records"


def _sanitize_daily_close_groups(value):
    groups = []
    if not isinstance(value, list):
        return groups
    for group in value:
        if not isinstance(group, dict):
            continue
        name = str(group.get("name") or "").strip()
        if not name:
            continue
        items = group.get("items") or []
        if not isinstance(items, list):
            items = []
        clean_items = []
        seen = set()
        for item in items:
            text = str(item or "").strip()
            if text and text not in seen:
                clean_items.append(text)
                seen.add(text)
        groups.append({"name": name, "items": clean_items})
    return groups


def _sanitize_daily_close_excellent(value):
    if not isinstance(value, dict):
        value = {}
    category_groups = value.get("categoryGroups") or []
    if not isinstance(category_groups, list):
        category_groups = []
    category_products = value.get("categoryProducts") or []
    if not isinstance(category_products, list):
        category_products = []
    force = value.get("forceContracts") or []
    if not isinstance(force, list):
        force = []
    collateral = value.get("collateralKinds") or []
    if not isinstance(collateral, list):
        collateral = []
    def _contract_no(v):
        text = str(v or "").strip()
        return text[:-2] if text.endswith(".0") else text
    return {
        "niceMin": str(value.get("niceMin") or "").strip(),
        "niceMax": str(value.get("niceMax") or "").strip(),
        "kMin": str(value.get("kMin") or "").strip(),
        "kMax": str(value.get("kMax") or "").strip(),
        "categoryGroups": list(dict.fromkeys([str(v or "").strip() for v in category_groups if str(v or "").strip()])),
        "categoryProducts": list(dict.fromkeys([str(v or "").strip() for v in category_products if str(v or "").strip()])),
        "collateralKinds": [str(v or "").strip() for v in collateral],
        "forceContracts": list(dict.fromkeys([_contract_no(v) for v in force if _contract_no(v)])),
    }


def _sanitize_daily_close_config(value):
    if not isinstance(value, dict):
        value = {}
    return {
        "product": _sanitize_daily_close_groups(value.get("product")),
        "category": _sanitize_daily_close_groups(value.get("category")),
        "agent": _sanitize_daily_close_groups(value.get("agent")),
        "excellent": _sanitize_daily_close_excellent(value.get("excellent")),
    }


def _load_daily_close_config():
    config = uploaded_data.get(DAILY_CLOSE_CONFIG_KEY)
    if config is None:
        config = load_data(DAILY_CLOSE_CONFIG_KEY) or {}
        uploaded_data[DAILY_CLOSE_CONFIG_KEY] = config
    return _sanitize_daily_close_config(config)


@app.get("/api/daily-close-config")
async def get_daily_close_config():
    """일 마감보고 설정 반환: 상품구성/상품구분/에이전트/우수대부업."""
    return JSONResponse(_load_daily_close_config())


@app.post("/api/daily-close-config")
async def save_daily_close_config(request: Request):
    """일 마감보고 설정 전체 저장."""
    body = await request.json()
    if not isinstance(body, dict):
        raise HTTPException(status_code=400, detail="객체 형태로 전송하세요")
    config = _sanitize_daily_close_config(body)
    uploaded_data[DAILY_CLOSE_CONFIG_KEY] = config
    save_data(DAILY_CLOSE_CONFIG_KEY, config)
    return JSONResponse({"success": True, "config": config})


def _load_daily_close_excellent_records():
    records = uploaded_data.get(DAILY_CLOSE_EXCELLENT_RECORDS_KEY)
    if records is None:
        records = load_data(DAILY_CLOSE_EXCELLENT_RECORDS_KEY) or {}
        uploaded_data[DAILY_CLOSE_EXCELLENT_RECORDS_KEY] = records
    if not isinstance(records, dict):
        records = {}
    daily = records.get("daily") if isinstance(records.get("daily"), dict) else {}
    monthly = records.get("monthly") if isinstance(records.get("monthly"), dict) else {}
    return {"daily": daily, "monthly": monthly}


def _clean_daily_close_number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _sanitize_daily_close_excellent_record(value):
    if not isinstance(value, dict):
        raise HTTPException(status_code=400, detail="record 객체가 필요합니다.")
    date = str(value.get("date") or "").strip()[:10]
    period = normalize_period_key(value.get("period")) or (date[:7] if date else "")
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", date):
        raise HTTPException(status_code=400, detail="date는 YYYY-MM-DD 형식이어야 합니다.")
    if not re.match(r"^\d{4}-\d{2}$", period):
        raise HTTPException(status_code=400, detail="period는 YYYY-MM 형식이어야 합니다.")

    clean = {
        "date": date,
        "period": period,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    for key in (
        "balance", "gap", "change", "previous_balance", "previous_gap",
        "today_execution", "month_execution", "today_excellent_execution", "month_excellent_execution",
    ):
        number = _clean_daily_close_number(value.get(key))
        if number is not None:
            clean[key] = number
    product_filter = value.get("product_filter") or []
    if isinstance(product_filter, list):
        clean["product_filter"] = list(dict.fromkeys([str(v or "").strip() for v in product_filter if str(v or "").strip()]))
    else:
        clean["product_filter"] = []
    return clean


@app.get("/api/daily-close-excellent-records")
async def get_daily_close_excellent_records():
    """우수대부업 일보고/월별 추이 저장 기록 반환."""
    return JSONResponse(_load_daily_close_excellent_records())


@app.post("/api/daily-close-excellent-records")
async def save_daily_close_excellent_record(request: Request):
    """우수대부업 일보고 현재 스냅샷 저장."""
    body = await request.json()
    record = _sanitize_daily_close_excellent_record(body.get("record") if isinstance(body, dict) and "record" in body else body)
    records = _load_daily_close_excellent_records()
    records["daily"][record["date"]] = record
    records["monthly"][record["period"]] = {
        "period": record["period"],
        "date": record["date"],
        "balance": record.get("balance", 0),
        "gap": record.get("gap", 0),
        "saved_at": record["saved_at"],
    }
    uploaded_data[DAILY_CLOSE_EXCELLENT_RECORDS_KEY] = records
    save_data(DAILY_CLOSE_EXCELLENT_RECORDS_KEY, records)
    return JSONResponse({"success": True, "records": records, "record": record})


def _daily_close_excellent_products(config: dict) -> set[str]:
    excellent = config.get("excellent") or {}
    products = {str(v or "").strip() for v in (excellent.get("categoryProducts") or []) if str(v or "").strip()}
    if products:
        return products
    selected_groups = {str(v or "").strip() for v in (excellent.get("categoryGroups") or []) if str(v or "").strip()}
    if not selected_groups:
        return set()
    for group in config.get("category") or []:
        name = str(group.get("name") or "").strip()
        if name not in selected_groups:
            continue
        for item in group.get("items") or []:
            text = str(item or "").strip()
            if text:
                products.add(text)
    return products


def _daily_close_contract_no(value) -> str:
    text = str(value or "").strip()
    return text[:-2] if text.endswith(".0") else text


def _daily_close_score_in_range(value, min_value, max_value) -> bool:
    try:
        score = float(value)
    except (TypeError, ValueError):
        return False
    has_min = str(min_value or "").strip() != ""
    has_max = str(max_value or "").strip() != ""
    if not has_min and not has_max:
        return False
    try:
        lo = float(min_value) if has_min else float("-inf")
        hi = float(max_value) if has_max else float("inf")
    except (TypeError, ValueError):
        return False
    return lo <= score <= hi


def _daily_close_excellent_match(row: dict, excellent: dict, products: set[str]) -> bool:
    contract_no = _daily_close_contract_no(row.get("contract_no") or row.get("contractNo"))
    forced = {_daily_close_contract_no(v) for v in (excellent.get("forceContracts") or []) if _daily_close_contract_no(v)}
    if contract_no and contract_no in forced:
        return True

    product = str(row.get("product") or row.get("product_name") or "").strip()
    if products and product not in products:
        return False

    score_configured = any(str(excellent.get(key) or "").strip() for key in ("niceMin", "niceMax", "kMin", "kMax"))
    collateral_configured = bool(excellent.get("collateralKinds") or [])
    if not score_configured and not collateral_configured:
        return True

    nice_hit = _daily_close_score_in_range(row.get("nice_score"), excellent.get("niceMin"), excellent.get("niceMax"))
    k_hit = _daily_close_score_in_range(row.get("k_score"), excellent.get("kMin"), excellent.get("kMax"))
    collateral = str(row.get("collateral_kind") or "").strip()
    score_hit = (not score_configured) or nice_hit or k_hit
    collateral_hit = (not collateral_configured) or any(str(v or "").strip() == collateral for v in (excellent.get("collateralKinds") or []))
    return score_hit and collateral_hit


def _daily_close_row_date(row: dict) -> str:
    return str(row.get("contract_date") or row.get("date") or "").strip()[:10]


def _daily_close_balance_million(row: dict) -> float:
    value = row.get("balance_million")
    try:
        if value not in (None, ""):
            return float(value)
    except (TypeError, ValueError):
        pass
    try:
        return float(row.get("balance") or 0) / 1_000_000
    except (TypeError, ValueError):
        return 0.0


def _daily_close_excel_filename(name: str) -> str:
    from urllib.parse import quote
    encoded = quote(name)
    return f"attachment; filename*=UTF-8''{encoded}"


@app.get("/api/export/daily-close-excellent-list")
async def export_daily_close_excellent_list(period: str = "", date: str = ""):
    """일 마감보고 우수대부업 기준에 포함된 전체 진행 계약리스트 행을 엑셀로 다운로드."""
    source = uploaded_data.get("full_contract_data") or {}
    if not source:
        raise HTTPException(status_code=404, detail="전체 진행 계약리스트 데이터가 없습니다.")

    selected_date = str(date or "").strip()[:10]
    selected_period = normalize_period_key(period) or (selected_date[:7] if selected_date else "")
    if not selected_period:
        periods = source.get("periods") or []
        selected_period = sorted([str(p) for p in periods if p])[-1] if periods else ""
    if not selected_period:
        raise HTTPException(status_code=400, detail="기간을 확인할 수 없습니다.")

    period_data = build_full_contract_period(source, selected_period)
    rows = [row for row in (period_data.get("loan_loss_rows") or []) if isinstance(row, dict)]
    config = _load_daily_close_config()
    excellent = config.get("excellent") or {}
    products = _daily_close_excellent_products(config)
    matched = [row for row in rows if _daily_close_excellent_match(row, excellent, products)]

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "우수대부업 리스트"
    summary = wb.create_sheet("집계요약", 0)

    month_start = f"{selected_period}-01"
    total_balance = sum(_daily_close_balance_million(row) for row in matched)
    shortage_balance = abs(DAILY_CLOSE_EXCELLENT_TARGET_MILLION - total_balance)
    today_balance = sum(_daily_close_balance_million(row) for row in matched if selected_date and _daily_close_row_date(row) == selected_date)
    month_balance = sum(
        _daily_close_balance_million(row)
        for row in matched
        if month_start <= _daily_close_row_date(row) <= (selected_date or f"{selected_period}-31")
    )

    condition_rows = [
        ("기준월", selected_period),
        ("마감보고 일자", selected_date or "-"),
        ("대상 건수", len(matched)),
        ("우수대부업 기준 잔액(백만원)", DAILY_CLOSE_EXCELLENT_TARGET_MILLION),
        ("융잔 합계(백만원)", total_balance),
        ("부족금/초과금(백만원)", shortage_balance),
        ("당일 금액(백만원)", today_balance),
        ("당월 금액(백만원)", month_balance),
        ("상품구분 그룹", ", ".join(excellent.get("categoryGroups") or []) or "전체"),
        ("상품 수", len(products)),
        ("NICE 구간", f"{excellent.get('niceMin') or ''}~{excellent.get('niceMax') or ''}".strip("~") or "-"),
        ("K 구간", f"{excellent.get('kMin') or ''}~{excellent.get('kMax') or ''}".strip("~") or "-"),
        ("담보종류", ", ".join([str(v or "빈칸") for v in (excellent.get("collateralKinds") or [])]) or "-"),
        ("강제적용 계약번호 수", len(excellent.get("forceContracts") or [])),
    ]
    summary.append(["항목", "값"])
    for key, value in condition_rows:
        summary.append([key, value])

    headers = [
        "순번", "기준월", "계약일", "계약번호", "상품명", "상품구성", "우수대부 상품구분",
        "대출잔액(원)", "대출잔액(백만원)", "연체일", "NICE스코어", "K스코어", "담보종류",
        "강제적용", "당일대상", "당월대상"
    ]
    ws.append(headers)
    product_to_category = {}
    for group in config.get("category") or []:
        name = str(group.get("name") or "").strip()
        for item in group.get("items") or []:
            text = str(item or "").strip()
            if text:
                product_to_category.setdefault(text, []).append(name)
    forced_set = {_daily_close_contract_no(v) for v in (excellent.get("forceContracts") or []) if _daily_close_contract_no(v)}
    for idx, row in enumerate(matched, start=1):
        row_date = _daily_close_row_date(row)
        contract_no = _daily_close_contract_no(row.get("contract_no") or row.get("contractNo"))
        balance_m = _daily_close_balance_million(row)
        ws.append([
            idx,
            selected_period,
            row_date,
            contract_no,
            row.get("product") or "",
            row.get("group") or "",
            ", ".join(product_to_category.get(str(row.get("product") or "").strip(), [])),
            row.get("balance") or 0,
            balance_m,
            row.get("days") if row.get("days") is not None else "",
            row.get("nice_score") if row.get("nice_score") is not None else "",
            row.get("k_score") if row.get("k_score") is not None else "",
            row.get("collateral_kind") or "",
            "Y" if contract_no and contract_no in forced_set else "",
            "Y" if selected_date and row_date == selected_date else "",
            "Y" if row_date and row_date.startswith(selected_period) else "",
        ])

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="B7C9D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for sheet in (summary, ws):
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center")
        for col in sheet.columns:
            letter = col[0].column_letter
            width = min(max(len(str(cell.value or "")) for cell in col) + 4, 38)
            sheet.column_dimensions[letter].width = width
    for row in ws.iter_rows(min_row=2, min_col=8, max_col=9):
        for cell in row:
            cell.number_format = '#,##0.000'
    for cell in summary["B"]:
        if isinstance(cell.value, (int, float)):
            cell.number_format = '#,##0.000'

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    filename = f"일마감보고_우수대부업_대상리스트_{selected_period}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _daily_close_excel_filename(filename)},
    )

# ── 상품구분 그룹 설정 API ──────────────────────────────────────────



# body: {"상품구분": [...], "상품구분상세": [...], "상품구분화해채권": [...]}



# 각 배열 원소: {"name": str, "items": [str, ...]}







@app.get("/api/product-category-groups")



async def get_product_category_groups():



    """저장된 상품구분 그룹 설정 반환"""



    groups = uploaded_data.get("product_category_groups", {})



    if not groups:



        _path = _data_path("product_category_groups")



        if os.path.exists(_path):



            with open(_path, encoding="utf-8") as f:



                groups = json.load(f)



            uploaded_data["product_category_groups"] = groups



    return JSONResponse(groups)











@app.post("/api/product-category-groups")



async def save_product_category_groups(request: Request):



    """상품구분 그룹 전체 저장 (덮어쓰기)"""



    body = await request.json()



    if not isinstance(body, dict):



        raise HTTPException(status_code=400, detail="객체 형태로 전송하세요")



    uploaded_data["product_category_groups"] = body



    _path = _data_path("product_category_groups")



    os.makedirs(os.path.dirname(_path), exist_ok=True)



    with open(_path, "w", encoding="utf-8") as f:



        json.dump(body, f, ensure_ascii=False, indent=2)



    return JSONResponse({"success": True})















@app.get("/api/product-category-groups/hwahae-items")



async def get_hwahae_items():



    """BW열(회생상태) + BX열(신복상태) 고유값 목록 반환



    구분설정 상품구분(화해채권) 탭의 항목 선택에 사용.



    반환: {"bw": [...], "bx": [...]}



    """



    saq = uploaded_data.get("settlement_aq") or {}



    if not saq:



        _saq_file = _data_path("settlement_aq")



        if os.path.exists(_saq_file):



            with open(_saq_file, encoding="utf-8") as f:



                saq = json.load(f)







    all_rank: set[str] = set()



    all_bw: set[str] = set()



    all_bx: set[str] = set()



    for pdata in saq.values():



        all_rank.update(pdata.get("hwahae_rank") or [])



        all_bw.update(pdata.get("hwahae_bw") or [])



        all_bx.update(pdata.get("hwahae_bx") or [])







    all_rank.add("\uD654\uD574")



    exclude_subranks = {"\u2605\uBCC4\uC81C\uAD8C\uBD80", "\u2605\uBCC4\uC81C\uAD8C\uBD80 \uB3D9\uC758"}



    all_bw.difference_update(exclude_subranks)



    all_bx.difference_update(exclude_subranks)







    return JSONResponse({



        "rank": sorted(all_rank),



        "bw": sorted(all_bw),



        "bx": sorted(all_bx),



    })












def _collect_job_group_raw_keys() -> list[str]:

    saq = uploaded_data.get("settlement_aq") or {}

    if not saq:

        _saq_file = _data_path("settlement_aq")

        if os.path.exists(_saq_file):

            with open(_saq_file, encoding="utf-8") as f:

                saq = json.load(f)

            uploaded_data["settlement_aq"] = saq

    all_jobs: set[str] = set()

    total_key = "\uc804\uccb4\uc735\uc794 \ud569\uacc4"

    for pdata in saq.values():

        if not isinstance(pdata, dict):

            continue

        job_raw = pdata.get("job_raw") or {}

        for key in job_raw.keys():

            if key and key != total_key:

                all_jobs.add(str(key))

    return sorted(all_jobs)



def _default_job_groups_from_raw(job_keys: list[str]) -> list[dict]:

    key_set = set(job_keys or [])

    defs = [

        ("\uc9c1\uc7a5\uc778", ["\uae09\uc5ec", "\uae09\uc5ec-\uacf5\ubb34\uc6d0"]),

        ("\uc790\uc601\uc5c5", ["\uc790\uc601\uc5c5\uc790", "\ubc95\uc778\uc0ac\uc5c5\uc790"]),

        ("\uc8fc\ubd80", ["\uc8fc\ubd80-\ubc30\uc6b0\uc790\uae09\uc5ec", "\uc8fc\ubd80-\ubc30\uc6b0\uc790\uc790\uc601\uc5c5"]),

        ("\uae30\ud0c0", ["\uae30\ud0c0", "\ubb34\uc9c1"]),

    ]

    groups: list[dict] = []

    assigned: set[str] = set()

    for name, items in defs:

        picked = [item for item in items if item in key_set]

        assigned.update(picked)

        if picked:

            groups.append({"name": name, "items": picked})

    extras = sorted(key_set - assigned)

    if extras:

        etc_name = "\uae30\ud0c0"

        for group in groups:

            if group.get("name") == etc_name:

                group["items"] = list(group.get("items") or []) + extras

                break

        else:

            groups.append({"name": etc_name, "items": extras})

    return groups



@app.get("/api/job-groups")



async def get_job_groups():



    """??? ???? ?? ?? ??"""



    groups = uploaded_data.get("job_groups", None)



    _path = _data_path("job_groups")



    if groups is None and os.path.exists(_path):



        with open(_path, encoding="utf-8") as f:



            groups = json.load(f)



        uploaded_data["job_groups"] = groups



    if groups is None:



        groups = _default_job_groups_from_raw(_collect_job_group_raw_keys())



        uploaded_data["job_groups"] = groups



    return JSONResponse(groups)







@app.post("/api/job-groups")



async def save_job_groups(request: Request):



    """직업분류 그룹 전체 저장 (덮어쓰기)"""



    body = await request.json()



    if not isinstance(body, list):



        raise HTTPException(status_code=400, detail="배열 형태로 전송하세요")



    uploaded_data["job_groups"] = body



    _path = _data_path("job_groups")



    os.makedirs(os.path.dirname(_path), exist_ok=True)



    with open(_path, "w", encoding="utf-8") as f:



        json.dump(body, f, ensure_ascii=False, indent=2)



    return JSONResponse({"success": True})











@app.get("/api/job-groups/all-jobs")



async def get_all_jobs():



    """???? AC?(????) ??? ?? ??"""



    return JSONResponse(_collect_job_group_raw_keys())







@app.get("/api/cb-grade-config")



async def get_cb_grade_config():



    """CB등급별 NICE스코어 구간 설정 반환 (1~10등급 고정 그룹)"""



    cfg = uploaded_data.get("cb_grade_config")



    if not cfg:



        _path = _data_path("cb_grade_config")



        if os.path.exists(_path):



            with open(_path, encoding="utf-8") as f:



                cfg = json.load(f)



            uploaded_data["cb_grade_config"] = cfg



        else:



            # 기본값: 1~10등급 NICE 점수 구간



            cfg = [



                {"grade": "1분위", "lo": 900, "hi": 1000},



                {"grade": "2분위", "lo": 870, "hi": 899},



                {"grade": "3분위", "lo": 840, "hi": 869},



                {"grade": "4분위", "lo": 805, "hi": 839},



                {"grade": "5분위", "lo": 750, "hi": 804},



                {"grade": "6분위", "lo": 665, "hi": 749},



                {"grade": "7분위", "lo": 600, "hi": 664},



                {"grade": "8분위", "lo": 515, "hi": 599},



                {"grade": "9분위", "lo": 445, "hi": 514},



                {"grade": "10분위", "lo": 0,   "hi": 444},



            ]



            uploaded_data["cb_grade_config"] = cfg



    return JSONResponse(cfg)











@app.post("/api/cb-grade-config")



async def save_cb_grade_config(request: Request):



    """CB등급별 NICE스코어 구간 설정 저장 (1~10등급 고정)"""



    body = await request.json()



    if not isinstance(body, list):



        raise HTTPException(status_code=400, detail="배열 형태로 전송하세요")



    # 필수 필드 검증



    for item in body:



        if not all(k in item for k in ("grade", "lo", "hi")):



            raise HTTPException(status_code=400, detail="grade, lo, hi 필드가 필요합니다")



    uploaded_data["cb_grade_config"] = body



    _path = _data_path("cb_grade_config")



    os.makedirs(os.path.dirname(_path), exist_ok=True)



    with open(_path, "w", encoding="utf-8") as f:



        json.dump(body, f, ensure_ascii=False, indent=2)



    return JSONResponse({"success": True})











@app.get("/api/channel-groups/all-channels")



async def get_all_channels():



    """접수경로 전체 항목 반환



    우선순위: 전체 진행 계약리스트 L열/결산자료 Q열(광고매체) 고유값 → asset_data 접수경로 섹션 키 → 기본값



    """



    full_contract = uploaded_data.get("full_contract_data") or {}
    channel_values = set()
    for row in full_contract.get("loan_loss_rows") or []:
        if not isinstance(row, dict):
            continue
        channel = str(row.get("channel") or row.get("channel_raw") or "").strip()
        if channel:
            channel_values.add(channel)

    # 1순위: 결산자료 파싱 결과의 channels 목록 (Q열 광고매체 고유값)



    saq = uploaded_data.get("settlement_aq") or {}



    if not saq:



        _saq_file = _data_path("settlement_aq")



        if os.path.exists(_saq_file):



            with open(_saq_file, encoding="utf-8") as f:



                saq = json.load(f)



    if saq:



        # settlement_aq는 {pkey: {channels:[...], ...}} 구조



        for pkey, pdata in saq.items():



            channels = pdata.get("channels", [])



            if channels:



                channel_values.update(str(channel or "").strip() for channel in channels if str(channel or "").strip())



    if channel_values:



        return JSONResponse(sorted(channel_values))







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















BUSINESS_AMOUNT_NUMBER_FORMAT = r'#,##0;[Red]\(#,##0\)'



BUSINESS_AMOUNT_DECIMAL_NUMBER_FORMAT = r'#,##0.0;[Red]\(#,##0.0\)'



BUSINESS_UNIT_OPTIONS = {



    "won": {"label": "\uc6d0", "factor": 1_000_000, "number_format": BUSINESS_AMOUNT_NUMBER_FORMAT},



    "million": {"label": "\ubc31\ub9cc\uc6d0", "factor": 1, "number_format": BUSINESS_AMOUNT_NUMBER_FORMAT},



    "ten_million": {"label": "\ucc9c\ub9cc\uc6d0", "factor": 0.1, "number_format": BUSINESS_AMOUNT_DECIMAL_NUMBER_FORMAT},



}



BUSINESS_UNIT_ALIASES = {



    "\uc6d0": "won",



    "\ubc31\ub9cc": "million",



    "\ubc31\ub9cc\uc6d0": "million",



    "\ucc9c\ub9cc": "ten_million",



    "\ucc9c\ub9cc\uc6d0": "ten_million",



}



BUSINESS_S1_ROWS = {



    4: "\uae30\ucd08\ub300\ucd9c\uc790\uc0b0",



    5: "\ub300\ucd9c_\uc2e0\uaddc",



    6: "\ub300\ucd9c_\ucd94\uac00",



    7: "\ub300\ucd9c_\uc7ac\ub300\ucd9c",



    8: "\ub300\ucd9c_\ucde8\uae09\uc561(\ub300\ucd9c)",



    9: "\ub9e4\uc785",



    10: "\uc6d0\uae08\ud68c\uc218\uc561",



    11: "\uc774\uc790\ud68c\uc218\uc561",



    12: "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\ub9e4\uac01\uc561",



    13: "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\uc0c1\uac01\uc561",



    14: "\uae30\ub9d0\uc790\uc0b0",



    15: "\ub300\uc190\ucda9\ub2f9_\uae30\ucd08 \ucda9\ub2f9\uae08",



    16: "\ub300\uc190\ucda9\ub2f9_\uc124\uc815\ubd84",



    17: "\ub300\uc190\ucda9\ub2f9_\uc0ac\uc6a9\ubd84(\ud658\uc785)",



    18: "\uae30\ub9d0 \ucda9\ub2f9\uae08",



}



BUSINESS_S1_CONTRACT_KEY_MAP = {



    "\ub300\ucd9c_\uc2e0\uaddc": "\uc2e0\uaddc",



    "\ub300\ucd9c_\ucd94\uac00": "\ucd94\uac00\ub300\ucd9c",



    "\ub300\ucd9c_\uc7ac\ub300\ucd9c": "\uc7ac\ub300\ucd9c",



    "\ub300\ucd9c_\ucde8\uae09\uc561(\ub300\ucd9c)": "\ucde8\uae09\uc561(\ub300\ucd9c)",



}



BUSINESS_S1_PAYMENT_KEY_MAP = {



    "\uc6d0\uae08\ud68c\uc218\uc561": "\uc6d0\uae08\ud68c\uc218\uc561",



    "\uc774\uc790\ud68c\uc218\uc561": "\uc774\uc790\ud68c\uc218\uc561",



}



BUSINESS_S1_WRITEOFF_KEY_MAP = {



    "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\uc0c1\uac01\uc561": "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\uc0c1\uac01\uc561",



    "\ub300\uc190\ucda9\ub2f9_\uc0ac\uc6a9\ubd84(\ud658\uc785)": "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\uc0c1\uac01\uc561",



}



BUSINESS_S1_SALE_KEY_MAP = {"\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\ub9e4\uac01\uc561": "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\ub9e4\uac01\uc561"}



BUSINESS_COLLATERAL_SUB_ROWS = [



    {"key": "\uc2e0\uaddc", "fallback_sub": "\uc2e0\uaddc", "label": "\uc2e0\uaddc"},



    {"key": "\ucd94\uac00\ub300\ucd9c", "fallback_sub": "\ucd94\uac00", "label": "\ucd94\uac00"},



    {"key": "\uc7ac\ub300\ucd9c", "fallback_sub": "\uc7ac\ub300\ucd9c", "label": "\uc7ac\ub300\ucd9c"},



    {"key": "\ucde8\uae09\uc561(\ub300\ucd9c)", "fallback_sub": "\ucde8\uae09\uc561(\ub300\ucd9c)", "label": "\ucde8\uae09\uc561(\ub300\ucd9c)"},



]











def _business_unit_key(unit):



    key = str(unit or "million").strip()



    key = BUSINESS_UNIT_ALIASES.get(key, key)



    if key not in BUSINESS_UNIT_OPTIONS:



        key = "million"



    return key











def _business_unit_label(unit):



    return BUSINESS_UNIT_OPTIONS[_business_unit_key(unit)]["label"]











def _business_scale_amount(value, unit):



    if value is None:



        return None



    if not isinstance(value, (int, float)) or isinstance(value, bool):



        return value



    return value * BUSINESS_UNIT_OPTIONS[_business_unit_key(unit)]["factor"]











def _business_norm_period(value):



    period = _aq_norm_period(value)



    return period.replace(".", "-") if period else ""











def _business_period_sort_key(period):



    period = _business_norm_period(period)



    m = re.search(r'(\d{4})[-.](\d{1,2})', period)



    if not m:



        return (9999, 99, period)



    return (int(m.group(1)), int(m.group(2)), period)











def _business_period_datetime(period):



    period = _business_norm_period(period)



    m = re.search(r'(\d{4})[-.](\d{1,2})', period)



    if not m:



        return period



    return datetime(int(m.group(1)), int(m.group(2)), 1)











def _business_sheet(wb, preferred_name=None):



    if preferred_name and preferred_name in wb.sheetnames:



        return wb[preferred_name]



    for sheet_name in wb.sheetnames:



        if "\uc601\uc5c5\ud604\ud669" in sheet_name:



            return wb[sheet_name]



    return wb.active











def _business_period_column_map(ws):



    mapping = {}



    for header_row in (3, 20):



        if header_row > ws.max_row:



            continue



        for col in range(3, ws.max_column + 1):



            period = _business_norm_period(ws.cell(row=header_row, column=col).value)



            if re.fullmatch(r'\d{4}-\d{2}', period or ''):



                mapping.setdefault(period, col)



    return mapping











def _business_copy_column_style(ws, source_col, target_col):



    from openpyxl.utils import get_column_letter



    source_letter = get_column_letter(source_col)



    target_letter = get_column_letter(target_col)



    ws.column_dimensions[target_letter].width = ws.column_dimensions[source_letter].width



    for row in range(1, ws.max_row + 1):



        source = ws.cell(row=row, column=source_col)



        target = ws.cell(row=row, column=target_col)



        if source.has_style:



            target._style = copy.copy(source._style)



        if source.number_format:



            target.number_format = source.number_format



        if source.alignment:



            target.alignment = copy.copy(source.alignment)



        if source.protection:



            target.protection = copy.copy(source.protection)



        target.value = None











def _business_ensure_period_columns(ws, periods):



    mapping = _business_period_column_map(ws)



    last_col = max(mapping.values(), default=2)



    for period in sorted({_business_norm_period(p) for p in periods if _business_norm_period(p)}, key=_business_period_sort_key):



        if period in mapping:



            continue



        last_col += 1



        _business_copy_column_style(ws, last_col - 1, last_col)



        for header_row in (3, 20):



            if header_row <= ws.max_row:



                cell = ws.cell(row=header_row, column=last_col)



                cell.value = _business_period_datetime(period)



                cell.number_format = ws.cell(row=header_row, column=last_col - 1).number_format



        mapping[period] = last_col



    return mapping











def _business_update_unit_labels(ws, unit):



    label = _business_unit_label(unit)



    unit_pattern = re.compile(r'(\(?\s*\ub2e8\uc704\s*:\s*)([^)]*)(\)?)')



    for row in range(1, min(ws.max_row, 40) + 1):



        for col in range(1, min(ws.max_column, 4) + 1):



            cell = ws.cell(row=row, column=col)



            if isinstance(cell.value, str) and "\ub2e8\uc704" in cell.value:



                cell.value = unit_pattern.sub(lambda match: f"{match.group(1)}{label}{match.group(3)}", cell.value)



    if ws.max_row >= 2:



        ws.cell(row=2, column=1).value = f"(\ub2e8\uc704:{label})"











def _business_value_at(series, periods, idx=None, ym=None):



    if series is None:



        return None



    if isinstance(series, dict):



        if ym is None:



            return None



        return series.get(_business_norm_period(ym))



    if isinstance(series, list):



        if ym is not None and isinstance(periods, list):



            try:



                found = [_business_norm_period(p) for p in periods].index(_business_norm_period(ym))



            except ValueError:



                found = -1



            if found >= 0:



                return series[found] if found < len(series) else None



        if idx is not None:



            return series[idx] if idx < len(series) else None



    return None











def _business_section_value(data, section_name, key, idx=None, ym=None):



    if not isinstance(data, dict) or not key:



        return None



    return _business_value_at((data.get(section_name) or {}).get(key), data.get("periods") or [], idx, ym)











def _business_contract_maturity_extension_amount(ym):



    contract = uploaded_data.get("contract_data") or {}



    value = _business_value_at(contract.get("maturity_extension_amount"), contract.get("periods") or [], None, ym)



    if isinstance(value, (int, float)) and not isinstance(value, bool):



        return float(value)



    return 0.0











def _business_s1_raw_value(business, key, idx, ym):



    base = _business_section_value(business, "section1", key, idx, ym)



    if key == "\uc6d0\uae08\ud68c\uc218\uc561":



        payment_value = _business_section_value(



            uploaded_data.get("payment_data") or {},



            "section1",



            BUSINESS_S1_PAYMENT_KEY_MAP.get(key),



            idx,



            ym,



        )



        if payment_value is not None:



            return payment_value - _business_contract_maturity_extension_amount(ym)



    sources = [



        (uploaded_data.get("sale_data") or {}, BUSINESS_S1_SALE_KEY_MAP.get(key)),



        (uploaded_data.get("writeoff_data") or {}, BUSINESS_S1_WRITEOFF_KEY_MAP.get(key)),



        (uploaded_data.get("payment_data") or {}, BUSINESS_S1_PAYMENT_KEY_MAP.get(key)),



        (uploaded_data.get("contract_data") or {}, BUSINESS_S1_CONTRACT_KEY_MAP.get(key)),



    ]



    for source, source_key in sources:



        value = _business_section_value(source, "section1", source_key, idx, ym)



        if value is not None:



            return value



    return base











def _business_period_index(business, period):



    period = _business_norm_period(period)



    periods = [_business_norm_period(p) for p in (business.get("periods") or [])]



    try:



        return periods.index(period)



    except ValueError:



        return None











def _business_previous_period(period, periods):



    period = _business_norm_period(period)



    normalized = [_business_norm_period(p) for p in (periods or [])]



    try:



        index = normalized.index(period)



    except ValueError:



        return None



    if index <= 0:



        return None



    return normalized[index - 1]











def _business_loan_loss_allowance_million(period):



    period = _business_norm_period(period)



    store = uploaded_data.get("settlement_aq") or {}



    if not isinstance(store, dict):



        return None



    matched_key = None



    for key in store.keys():



        if _business_norm_period(key) == period:



            matched_key = key



            break



    if not matched_key:



        return None



    try:



        summary = _loss_calculate_period(matched_key, store[matched_key], _loss_config_with_products()).get("summary") or {}



        allowance = summary.get("total_allowance")



        if isinstance(allowance, (int, float)) and not isinstance(allowance, bool):



            return float(allowance) / 1_000_000



    except Exception:



        return None



    return None











def _business_s1_value(business, key, idx, ym, periods=None, cache=None, stack=None):



    period = _business_norm_period(ym)



    periods = periods or _business_collect_periods(business)



    cache = cache if cache is not None else {}



    stack = stack if stack is not None else set()



    cache_key = (key, period)



    if cache_key in cache:



        return cache[cache_key]



    if cache_key in stack:



        return _business_s1_raw_value(business, key, idx, period)



    stack.add(cache_key)







    if key == "\uae30\ucd08\ub300\ucd9c\uc790\uc0b0":



        prev = _business_previous_period(period, periods)



        prev_idx = _business_period_index(business, prev) if prev else None



        prev_ending = _business_s1_value(business, "\uae30\ub9d0\uc790\uc0b0", prev_idx, prev, periods, cache, stack) if prev else None



        value = prev_ending if prev_ending is not None else _business_s1_raw_value(business, key, idx, period)



    elif key == "\uae30\ub9d0\uc790\uc0b0":



        raw_ending = _business_s1_raw_value(business, key, idx, period)



        if _business_period_sort_key(period) < _business_period_sort_key("2026-05"):



            value = raw_ending



        else:



            beginning = _business_s1_value(business, "\uae30\ucd08\ub300\ucd9c\uc790\uc0b0", idx, period, periods, cache, stack)



            deal = _business_s1_value(business, "\ub300\ucd9c_\ucde8\uae09\uc561(\ub300\ucd9c)", idx, period, periods, cache, stack)



            repayment = _business_s1_value(business, "\uc6d0\uae08\ud68c\uc218\uc561", idx, period, periods, cache, stack)



            sale = _business_s1_value(business, "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\ub9e4\uac01\uc561", idx, period, periods, cache, stack)



            writeoff = _business_s1_value(business, "\uc2e4\uc0c1\uac01\uc561_\uc6d4\uc911\uc0c1\uac01\uc561", idx, period, periods, cache, stack)



            if beginning is not None and deal is not None and repayment is not None:



                value = beginning + deal - repayment - (sale or 0) - (writeoff or 0)



            else:



                value = raw_ending



    elif key == "\ub300\uc190\ucda9\ub2f9_\uae30\ucd08 \ucda9\ub2f9\uae08":



        prev = _business_previous_period(period, periods)



        prev_idx = _business_period_index(business, prev) if prev else None



        prev_ending = _business_s1_value(business, "\uae30\ub9d0 \ucda9\ub2f9\uae08", prev_idx, prev, periods, cache, stack) if prev else None



        value = prev_ending if prev_ending is not None else _business_s1_raw_value(business, key, idx, period)



    elif key == "\uae30\ub9d0 \ucda9\ub2f9\uae08":



        allowance = _business_loan_loss_allowance_million(period)



        value = allowance if allowance is not None else _business_s1_raw_value(business, key, idx, period)



    elif key == "\ub300\uc190\ucda9\ub2f9_\uc0ac\uc6a9\ubd84(\ud658\uc785)":



        value = _business_s1_raw_value(business, key, idx, period)



    elif key == "\ub300\uc190\ucda9\ub2f9_\uc124\uc815\ubd84":



        ending = _business_s1_value(business, "\uae30\ub9d0 \ucda9\ub2f9\uae08", idx, period, periods, cache, stack)



        beginning = _business_s1_value(business, "\ub300\uc190\ucda9\ub2f9_\uae30\ucd08 \ucda9\ub2f9\uae08", idx, period, periods, cache, stack)



        used = _business_s1_value(business, "\ub300\uc190\ucda9\ub2f9_\uc0ac\uc6a9\ubd84(\ud658\uc785)", idx, period, periods, cache, stack)



        value = ending - beginning + used if ending is not None and beginning is not None and used is not None else _business_s1_raw_value(business, key, idx, period)



    else:



        value = _business_s1_raw_value(business, key, idx, period)







    stack.discard(cache_key)



    cache[cache_key] = value



    return value











def _business_normalize_group_name(name):



    return re.sub(r'\s+', '', str(name or '')).replace("\ub300\ucd9c", "")











def _business_configured_collateral_group_names():



    names = []



    groups = uploaded_data.get("product_groups") or []



    if isinstance(groups, list):



        for group in groups:



            if isinstance(group, dict):



                name = str(group.get("name") or "").strip()



                if name:



                    names.append(name)



    contract_section3 = (uploaded_data.get("contract_data") or {}).get("section3") or {}



    if not names and isinstance(contract_section3, dict):



        names.extend(str(name).strip() for name in contract_section3.keys() if str(name).strip())



    if not names:



        names = ["\uc2e0\uc6a9", "\ub2f4\ubcf4"]



    unique = []



    for name in names:



        if name not in unique:



            unique.append(name)



    return unique











def _business_contract_group_data(group_name):



    section3 = (uploaded_data.get("contract_data") or {}).get("section3") or {}



    if group_name in section3:



        return section3[group_name]



    target = _business_normalize_group_name(group_name)



    for key, value in section3.items():



        if _business_normalize_group_name(key) == target:



            return value



    return None











def _business_fallback_section3_prefix(group_name):



    normalized = _business_normalize_group_name(group_name)



    if "\ub2f4\ubcf4" in normalized:



        return "\ub2f4\ubcf4\ub300\ucd9c"



    if "\uc2e0\uc6a9" in normalized:



        return "\uc2e0\uc6a9\ub300\ucd9c"



    return f"{group_name}\ub300\ucd9c"











def _business_group_display_label(group_name):



    text = str(group_name or "").strip()



    return text if text.endswith("\ub300\ucd9c") else f"{text}\ub300\ucd9c"











def _business_collateral_value(business, group_name, sub_key, fallback_sub, idx, ym):



    contract_group = _business_contract_group_data(group_name)



    contract_value = _business_value_at(contract_group.get(sub_key) if isinstance(contract_group, dict) else None, (uploaded_data.get("contract_data") or {}).get("periods") or [], None, ym)



    if contract_value is not None:



        return contract_value







    prefix = _business_fallback_section3_prefix(group_name)



    key = f"{prefix}_{fallback_sub}"



    for section_name in ("section3", "section2"):



        value = _business_section_value(business, section_name, key, idx, ym)



        if value is not None:



            return value



    if prefix == "\ub2f4\ubcf4\ub300\ucd9c":



        legacy_key = f"\uae30\uc5c5\ub300\ucd9c_{fallback_sub}"



        value = _business_section_value(business, "section2", legacy_key, idx, ym)



        if value is not None:



            return value



    return None











def _business_collect_periods(business):



    periods = set(_business_norm_period(p) for p in (business.get("periods") or []) if _business_norm_period(p))



    for key in ("contract_data", "payment_data", "writeoff_data", "sale_data"):



        source = uploaded_data.get(key) or {}



        periods.update(_business_norm_period(p) for p in (source.get("periods") or []) if _business_norm_period(p))



    settlement_store = uploaded_data.get("settlement_aq") or {}



    if isinstance(settlement_store, dict):



        periods.update(_business_norm_period(p) for p in settlement_store.keys() if _business_norm_period(p))



    return sorted(periods, key=_business_period_sort_key)











def _business_write_amount_cell(ws, row, col, value, unit):



    cell = ws.cell(row=row, column=col)



    if value is None:



        cell.value = None



    else:



        cell.value = _business_scale_amount(value, unit)



    cell.number_format = BUSINESS_UNIT_OPTIONS[_business_unit_key(unit)]["number_format"]











def _business_collateral_base_rows(ws):



    if ws.max_row >= 38 and ws.cell(row=31, column=1).value:



        return [31, 35]



    return [21, 25]











def _write_business_to_sheet(ws, business, unit="million"):



    unit = _business_unit_key(unit)



    periods = _business_collect_periods(business)



    period_columns = _business_ensure_period_columns(ws, periods)



    _business_update_unit_labels(ws, unit)







    business_periods = [_business_norm_period(p) for p in (business.get("periods") or [])]



    s1_cache = {}



    for row, key in BUSINESS_S1_ROWS.items():



        for period in periods:



            col = period_columns.get(period)



            if not col:



                continue



            idx = business_periods.index(period) if period in business_periods else None



            value = _business_s1_value(business, key, idx, period, periods, s1_cache)



            _business_write_amount_cell(ws, row, col, value, unit)







    base_rows = _business_collateral_base_rows(ws)



    group_names = _business_configured_collateral_group_names()



    for group_index, base_row in enumerate(base_rows):



        if group_index >= len(group_names):



            break



        group_name = group_names[group_index]



        ws.cell(row=base_row, column=1).value = _business_group_display_label(group_name)



        for offset, sub in enumerate(BUSINESS_COLLATERAL_SUB_ROWS):



            row = base_row + offset



            ws.cell(row=row, column=2).value = sub["label"]



            for period in periods:



                col = period_columns.get(period)



                if not col:



                    continue



                idx = business_periods.index(period) if period in business_periods else None



                value = _business_collateral_value(business, group_name, sub["key"], sub["fallback_sub"], idx, period)



                _business_write_amount_cell(ws, row, col, value, unit)







@app.get("/api/data/business_status")



async def get_business_status():



    """영업현황 데이터 (기존 형식)"""



    return uploaded_data.get("business_status", {})











@app.get("/api/data/business")



async def get_business_detail():



    """\uc601\uc5c5\ud604\ud669 3\uc139\uc158 \ub370\uc774\ud130 (\uc5d1\uc140 \ud30c\uc2f1 \uae30\ubc18)"""



    d = uploaded_data.get("business_detail")



    if not d:



        return JSONResponse({"error": "\ub370\uc774\ud130 \uc5c6\uc74c"}, status_code=404)



    return JSONResponse(d)











@app.post("/api/upload/business")



async def upload_business(file: UploadFile = File(...), period: str = ""):



    """\uc601\uc5c5\ud604\ud669 \uc5d1\uc140 \uc5c5\ub85c\ub4dc (\uae30\uc5c5\uc5ec\uc2e0\uc790\ub8cc - \uc601\uc5c5\ud604\ud669.xlsx)"""



    filename = file.filename or "business.xlsx"



    if not filename.lower().endswith(('.xlsx', '.xls')):



        raise HTTPException(status_code=400, detail="\uc5d1\uc140 \ud30c\uc77c\ub9cc \uc5c5\ub85c\ub4dc \uac00\ub2a5\ud569\ub2c8\ub2e4.")







    tmp_path = None



    try:



        content = await file.read()



        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:



            tmp.write(content)



            tmp_path = tmp.name







        data = parse_business_excel(tmp_path)



        if period:



            data['upload_period'] = period



        uploaded_data['business_detail'] = data







        sheet_name = None



        wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=False)



        try:



            sheet_name = _business_sheet(wb).title



        finally:



            wb.close()







        template_path = _file_path("business_template.xlsx")



        with open(template_path, "wb") as template_file:



            template_file.write(content)



        uploaded_data["business_template"] = {



            "filename": filename,



            "sheet_name": sheet_name,



        }







        return JSONResponse({



            "success": True,



            "message": f"\uc601\uc5c5\ud604\ud669 '{filename}' \uc5c5\ub85c\ub4dc \uc644\ub8cc",



            "periods": len(data.get('periods', [])),



            "upload_period": period or "(\ubbf8\uc9c0\uc815)",



        })



    except Exception as e:



        import traceback



        raise HTTPException(status_code=500, detail=f"\ud30c\uc77c \ucc98\ub9ac \uc624\ub958: {str(e)}\n{traceback.format_exc()}")



    finally:



        if tmp_path and os.path.exists(tmp_path):



            os.unlink(tmp_path)











@app.get("/api/export/business-detail")



async def export_business_detail(unit: str = "million"):



    """Export current business-status data using the uploaded workbook template."""



    data = uploaded_data.get("business_detail")



    if not data:



        raise HTTPException(status_code=404, detail="\uc601\uc5c5\ud604\ud669 \ub370\uc774\ud130\uac00 \uc5c6\uc2b5\ub2c8\ub2e4. \uc601\uc5c5\ud604\ud669 \uc5d1\uc140\uc744 \uba3c\uc800 \uc5c5\ub85c\ub4dc\ud558\uc138\uc694.")







    template_path = _stored_file_path("business_template.xlsx")



    if not template_path:



        raise HTTPException(status_code=404, detail="\uc601\uc5c5\ud604\ud669 \uc6d0\ubcf8 \uc5d1\uc140 \uc11c\uc2dd\uc774 \uc5c6\uc2b5\ub2c8\ub2e4. \uc601\uc5c5\ud604\ud669 \uc5d1\uc140\uc744 \ub2e4\uc2dc \uc5c5\ub85c\ub4dc\ud558\uc138\uc694.")







    try:



        meta = uploaded_data.get("business_template", {}) or {}



        wb = openpyxl.load_workbook(template_path, data_only=False, keep_links=False)



        ws = _business_sheet(wb, meta.get("sheet_name"))



        unit = _business_unit_key(unit)



        _write_business_to_sheet(ws, data, unit=unit)







        buf = io.BytesIO()



        wb.save(buf)



        wb.close()



        buf.seek(0)







        from urllib.parse import quote



        filename = f"\uae30\uc5c5\uc5ec\uc2e0\uc790\ub8cc_\uc601\uc5c5\ud604\ud669_{_business_unit_label(unit)}.xlsx"



        encoded_filename = quote(filename)



        return StreamingResponse(



            buf,



            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",



            headers={



                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",



                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",



                "Pragma": "no-cache",



                "X-Business-Unit": unit,



            },



        )



    except HTTPException:



        raise



    except Exception as e:



        raise HTTPException(status_code=500, detail=f"\uc601\uc5c5\ud604\ud669 \uc5d1\uc140 \uc0dd\uc131 \uc624\ub958: {str(e)}")











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



    company_info = uploaded_data.get("company_info", {})



    company_info.update(data)



    uploaded_data["company_info"] = company_info



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







    from pdf_mapping import (
        PDF_BS_MAPPING, PDF_IS_MAPPING, PDF_SUMMARY_MAPPING,
        DEFAULT_SIGN, DEFAULT_SIGN_FALLBACK,
    )







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



















def _merge_period_values(existing, incoming):



    if isinstance(existing, dict) and isinstance(incoming, dict):



        merged = copy.deepcopy(existing)



        for key, value in incoming.items():



            if key == "periods":



                continue



            if key in merged:



                merged[key] = _merge_period_values(merged[key], value)



            else:



                merged[key] = copy.deepcopy(value)



        periods = _merge_period_list(



            existing.get("periods") if isinstance(existing, dict) else [],



            incoming.get("periods") if isinstance(incoming, dict) else [],



        )



        if periods:



            merged["periods"] = periods



        return merged



    if isinstance(existing, list) and isinstance(incoming, list):



        merged = list(existing)



        for item in incoming:



            if item not in merged:



                merged.append(item)



        return merged



    return copy.deepcopy(incoming)











def _merge_period_list(*period_lists):



    periods = []



    for period_list in period_lists:



        if not isinstance(period_list, list):



            continue



        for period in period_list:



            norm = _business_norm_period(period)



            if norm and norm not in periods:



                periods.append(norm)



    return sorted(periods, key=_business_period_sort_key)











def _merge_uploaded_periodic_data(key, incoming):



    existing = uploaded_data.get(key) or {}



    if not existing:



        return incoming



    return _merge_period_values(existing, incoming)







@app.post("/api/upload/contract")



async def upload_contract(file: UploadFile = File(...), period: str = Form(default=""), confirm_period_mismatch: str = Form(default="")):



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



        selected_period = normalize_period_key(period)



        parsed_periods = data.get("periods", []) or []



        confirmed_mismatch = str(confirm_period_mismatch).strip().lower() in {"1", "true", "yes", "y"}



        period_mismatch = bool(



            selected_period



            and parsed_periods



            and (selected_period not in parsed_periods or len(parsed_periods) != 1)



        )



        if selected_period:



            data["selected_period"] = selected_period



        if period_mismatch and not confirmed_mismatch:



            os.unlink(tmp_path)



            return JSONResponse({



                "success": False,



                "code": "PERIOD_MISMATCH",



                "selected_period": selected_period,



                "parsed_periods": parsed_periods,



                "message": "Selected upload period differs from contract date periods in the file.",



            }, status_code=409)



        if period_mismatch:



            data["period_mismatch_confirmed"] = True







        existing_contract_data = uploaded_data.get('contract_data') or {}
        incoming_periods = set(data.get('periods') or [])
        existing_raw_rows = existing_contract_data.get('raw_rows') or []
        preserved_raw_rows = [
            row for row in existing_raw_rows
            if not isinstance(row, dict) or row.get('period') not in incoming_periods
        ]
        merged_contract_data = _merge_uploaded_periodic_data('contract_data', data)
        if 'raw_rows' in data:
            merged_contract_data['raw_rows'] = preserved_raw_rows + (data.get('raw_rows') or [])
        uploaded_data['contract_data'] = merged_contract_data



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



            "maturity_extension_amount": data.get("maturity_extension_amount", {}),



            "avg_rate_deal": {k: v for k, v in avg_rate_deal.items() if v is not None},

            "raw_rows": len(data.get("raw_rows") or []),

            "channel_column": data.get("channel_column"),



        })



    except Exception as e:



        import traceback



        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")











@app.post("/api/upload/full-contract")
async def upload_full_contract(file: UploadFile = File(...), period: str = Form(default="")):
    """
    전체 진행 계약리스트 업로드

    C열(상품명), L열(광고매체), V열(대출잔액), AH열(연체일) 기준으로
    일 마감보고의 당월 융자잔고/연체/에이전트 집계에 사용한다.
    """
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        selected_period = normalize_period_key(period) or str(period or '').strip()
        product_groups = uploaded_data.get("product_groups", []) or []
        data = parse_full_contract_list(tmp_path, product_groups, selected_period)

        existing_data = uploaded_data.get('full_contract_data') or {}
        incoming_periods = set(data.get('periods') or [])
        existing_rows = existing_data.get('loan_loss_rows') or []
        preserved_rows = [
            row for row in existing_rows
            if not isinstance(row, dict) or row.get('period') not in incoming_periods
        ]
        merged_data = _merge_uploaded_periodic_data('full_contract_data', data)
        merged_data['loan_loss_rows'] = preserved_rows + (data.get('loan_loss_rows') or [])
        merged_data['source_file'] = file.filename
        uploaded_data['full_contract_data'] = merged_data

        return JSONResponse({
            "success": True,
            "message": f"전체 진행 계약리스트 '{file.filename}' 업로드 완료",
            "periods": data.get('periods', []),
            "products": len(data.get('products', [])),
            "rows": len(data.get('loan_loss_rows') or []),
            "total_balance": data.get('section1', {}).get('융잔합계', 0),
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


@app.get("/api/data/full-contract")
async def get_full_contract_data(period: str = ""):
    """전체 진행 계약리스트 집계 데이터 반환."""
    d = uploaded_data.get("full_contract_data")
    if not d:
        return JSONResponse({"error": "데이터 없음"}, status_code=404)
    period_key = normalize_period_key(period) or str(period or '').strip()
    if period_key:
        period_data = build_full_contract_period(d, period_key)
        if not period_data.get('loan_loss_rows'):
            return JSONResponse({"error": "데이터 없음"}, status_code=404)
        return JSONResponse(period_data)
    return JSONResponse(d)

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



    K열(회계)='+' 필터 후



    N열(원금입금) → 원금회수액, Q열(이자입금) → 이자회수액 월별 집계



    """



    if not file.filename.lower().endswith(('.xlsx', '.xls')):



        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")



    try:



        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:



            content = await file.read()



            tmp.write(content)



            tmp_path = tmp.name







        product_groups = uploaded_data.get("product_groups", []) or []



        data = parse_payment(tmp_path, product_groups)



        if period:



            data['upload_period'] = period



        existing_payment_data = uploaded_data.get('payment_data') or {}



        incoming_periods = set(data.get('periods') or [])



        existing_raw_rows = existing_payment_data.get('raw_rows') or []



        preserved_raw_rows = [



            row for row in existing_raw_rows



            if not isinstance(row, dict) or row.get('period') not in incoming_periods



        ]



        merged_payment_data = _merge_uploaded_periodic_data('payment_data', data)



        if 'raw_rows' in data:



            merged_payment_data['raw_rows'] = preserved_raw_rows + (data.get('raw_rows') or [])



        uploaded_data['payment_data'] = merged_payment_data



        os.unlink(tmp_path)







        return JSONResponse({



            "success": True,



            "message": f"입금명세 '{file.filename}' 업로드 완료",



            "periods": data.get('periods', []),



            "rows": {



                "원금회수액": list(data.get('section1', {}).get('원금회수액', {}).values()),



                "이자회수액": list(data.get('section1', {}).get('이자회수액', {}).values()),



            },



            "raw_rows": len(data.get("raw_rows") or []),



            "group_column": data.get("group_column"),



            "product_column": data.get("product_column"),



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



        uploaded_data['writeoff_data'] = _merge_uploaded_periodic_data('writeoff_data', data)



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



        uploaded_data['sale_data'] = _merge_uploaded_periodic_data('sale_data', data)



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



    saq = uploaded_data.get("settlement_aq") or {}



    if not d and not saq:



        return JSONResponse({"error": "데이터 없음"}, status_code=404)



    if not d:



        periods = sorted({



            pkey[:7]



            for pkey in saq.keys()



            if isinstance(pkey, str) and len(pkey) >= 7 and pkey[4] == "-"



        })



        d = {"periods": periods, "sections": {}, "cb_products": {}}







    import copy



    result = copy.deepcopy(d)







    # ── settlement_aq section4 → sections['상품별'] 병합 구성 ───────────



    # 기본값: asset_data.sections['상품별'] (대출자산속성.xlsx 전체 기간)



    # 덮어쓰기: settlement_aq에 해당 기간이 있으면 그 기간만 결산자료 값으로 교체



    if saq:



        # ── 상품명 정규화 매핑 (오타/동의어 통합) ──────────────────────



        # 결산자료 상품명 → asset_data 표준 상품명



        PROD_ALIAS: dict[str, str] = {



            # 토마토토탈론이 표준명 (asset_data 및 결산자료 모두 동일)



        }







        sections = result.get("sections", {})







        def merge_saq_section(section_name: str, payload_key: str, multiplier: float = 1_000_000) -> None:



            section = {k: dict(v) for k, v in sections.get(section_name, {}).items()}



            for pkey, pdata in saq.items():



                ym = pkey[:7] if isinstance(pkey, str) and len(pkey) >= 7 and pkey[4] == "-" else None



                if not ym:



                    continue



                for row_key, val in (pdata.get(payload_key) or {}).items():



                    if val is None:



                        continue



                    if row_key not in section:



                        section[row_key] = {}



                    section[row_key][ym] = float(val) * multiplier



            if section:



                sections[section_name] = section







        merge_saq_section("대출만기", "maturity")



        merge_saq_section("상환방식", "repayment")



        merge_saq_section("금액별", "amount_band")







        merge_saq_section("접수경로", "channel_amount")



        rate_sec = {k: dict(v) for k, v in sections.get("평균이자율", {}).items()}



        rate_sec2 = {k: dict(v) for k, v in sections.get("평균이자율2", {}).items()}



        for pkey, pdata in saq.items():



            ym = pkey[:7] if isinstance(pkey, str) and len(pkey) >= 7 and pkey[4] == "-" else None



            if not ym:



                continue



            for row_key, val in (pdata.get("avg_rate") or {}).items():



                if val is None:



                    continue



                target = rate_sec if row_key == "대출자산평균" else rate_sec2



                if row_key not in target:



                    target[row_key] = {}



                target[row_key][ym] = float(val)



        if rate_sec:



            sections["평균이자율"] = rate_sec



        if rate_sec2:



            sections["평균이자율2"] = rate_sec2







        result["sections"] = sections







        # asset_data 기존 상품별 섹션을 기본으로 복사



        prod_sec: dict[str, dict[str, float]] = {



            k: dict(v) for k, v in sections.get("상품별", {}).items()



        }







        # settlement_aq에 있는 기간만 덮어씀



        for pkey, pdata in saq.items():



            ym = pkey[:7] if len(pkey) >= 7 and pkey[4] == '-' else None



            if not ym:



                continue



            s4 = pdata.get("section4") or {}



            if not s4:



                continue







            # 먼저 해당 ym의 기존 값을 모두 제거



            for prod in list(prod_sec.keys()):



                if ym in prod_sec[prod]:



                    prod_sec[prod].pop(ym, None)







            # settlement_aq section4 값으로 채움 (백만원 → 원 단위 복원)



            # 상품명 정규화 적용: 오타는 표준명으로 합산



            for prod_raw, buckets in s4.items():



                val = buckets.get("융잔합계")



                if val is None:



                    continue



                prod = PROD_ALIAS.get(prod_raw, prod_raw)  # 정규화



                if prod not in prod_sec:



                    prod_sec[prod] = {}



                # 같은 표준명으로 합산될 수 있으므로 += 처리



                prod_sec[prod][ym] = prod_sec[prod].get(ym, 0.0) + float(val) * 1_000_000







            # 해당 기간 전체융잔 합계 재계산



            total_ym = sum(



                prod_sec[p].get(ym, 0.0) or 0.0



                for p in prod_sec if p != "전체융잔 합계"



            )



            if "전체융잔 합계" not in prod_sec:



                prod_sec["전체융잔 합계"] = {}



            prod_sec["전체융잔 합계"][ym] = total_ym







        sections["상품별"] = prod_sec



        result["sections"] = sections







        # ── settlement_aq BW/BX → sections['상품구분화해채권'] 동적 구성 ──



        # sections['상품구분화해채권'][item_key][ym] = 원단위  (item_key = '[회생]xxx' or '[신복]xxx')



        # sections['상품구분화해채권_상품별'][item_key][상품명][ym] = 원단위  (상품별 분해용)



        hwahae_sec: dict[str, dict[str, float]] = {}



        hwahae_prod: dict[str, dict[str, dict[str, float]]] = {}  # {item_key: {상품명: {ym: 원}}}

        hwahae_rows_by_ym: dict[str, list[dict]] = {}



        for pkey, pdata in saq.items():



            ym = pkey[:7] if len(pkey) >= 7 and pkey[4] == '-' else None



            if not ym:



                continue



            source_defs = [

                ("[\uB7AD\uD06C]", pdata.get("section5_rank") or {}),

                ("[\uD68C\uC0DD]", pdata.get("section5_bw") or {}),

                ("[\uC2E0\uBCF5]", pdata.get("section5_bx") or {}),

            ]

            rows = pdata.get("section5_rows") or []

            if rows:

                hwahae_rows_by_ym[ym] = [

                    {

                        **row,

                        "amount": float(row.get("amount") or 0) * 1_000_000,

                    }

                    for row in rows

                ]



            for prefix, status_data in source_defs:



                for status_val, prod_dict in status_data.items():



                    item_key = f"{prefix}{status_val}"



                    if item_key not in hwahae_sec:



                        hwahae_sec[item_key] = {}



                    hwahae_sec[item_key][ym] = sum(float(v) * 1_000_000 for v in prod_dict.values())



                    if item_key not in hwahae_prod:



                        hwahae_prod[item_key] = {}



                    for prod_name, val in prod_dict.items():



                        if prod_name not in hwahae_prod[item_key]:



                            hwahae_prod[item_key][prod_name] = {}



                        hwahae_prod[item_key][prod_name][ym] = float(val) * 1_000_000







        if hwahae_sec:



            # 전체화해 합계 계산



            all_yms: set[str] = set()



            for v in hwahae_sec.values():



                all_yms.update(v.keys())



            total_hwahae: dict[str, float] = {}



            for ym in all_yms:



                total_hwahae[ym] = sum(



                    hwahae_sec[k].get(ym, 0.0) for k in hwahae_sec



                )



            hwahae_sec["전체화해 합계"] = total_hwahae



            sections["상품구분화해채권"] = hwahae_sec



            # 상품별 분해 데이터도 함께 응답



            sections["상품구분화해채권_상품별"] = hwahae_prod
            if hwahae_rows_by_ym:
                sections["상품구분화해채권_행별"] = hwahae_rows_by_ym



            result["sections"] = sections







        # ── settlement_aq gender/age_band/job_raw/region/cb_raw → 동적 섹션 구성 ──



        # 구조: sections['성별']['남성'][ym] = 원단위



        gender_sec_out: dict[str, dict[str, float]] = {}



        age_sec_out:    dict[str, dict[str, float]] = {}



        job_sec_out:    dict[str, dict[str, float]] = {}



        region_sec_out: dict[str, dict[str, float]] = {}



        cb_sec_out:     dict[str, dict[str, float]] = {}   # CB등급별







        for pkey, pdata in saq.items():



            ym = pkey[:7] if len(pkey) >= 7 and pkey[4] == '-' else None



            if not ym:



                continue







            # 성별



            gender_data = pdata.get("gender") or {}



            for row_key, val in gender_data.items():



                if val is None:



                    continue



                if row_key not in gender_sec_out:



                    gender_sec_out[row_key] = {}



                gender_sec_out[row_key][ym] = float(val) * 1_000_000







            # 연령별



            age_data = pdata.get("age_band") or {}



            for row_key, val in age_data.items():



                if val is None:



                    continue



                if row_key not in age_sec_out:



                    age_sec_out[row_key] = {}



                age_sec_out[row_key][ym] = float(val) * 1_000_000







            # 직업별



            job_data = pdata.get("job_raw") or {}



            for row_key, val in job_data.items():



                if val is None:



                    continue



                if row_key not in job_sec_out:



                    job_sec_out[row_key] = {}



                job_sec_out[row_key][ym] = float(val) * 1_000_000







            # 지역별



            region_data = pdata.get("region") or {}



            for row_key, val in region_data.items():



                if val is None:



                    continue



                if row_key not in region_sec_out:



                    region_sec_out[row_key] = {}



                region_sec_out[row_key][ym] = float(val) * 1_000_000







            # CB등급: cb_raw {score_str: 백만원} → cb_grade_config 구간 기반 등급별 집계



            cb_raw_pdata = pdata.get("cb_raw") or {}



            if cb_raw_pdata:



                # CB등급 설정 로드



                cb_cfg = uploaded_data.get("cb_grade_config")



                if not cb_cfg:



                    _cfg_path = _data_path("cb_grade_config")



                    if os.path.exists(_cfg_path):



                        with open(_cfg_path, encoding="utf-8") as _f:



                            cb_cfg = json.load(_f)



                        uploaded_data["cb_grade_config"] = cb_cfg



                if cb_cfg:



                    # 분위별 집계 버킷 — cfg에 있는 모든 분위를 0으로 미리 초기화 (데이터 없어도 표시)



                    grade_bucket: dict[str, float] = {item["grade"]: 0.0 for item in cb_cfg if item.get("grade")}



                    no_grade_total = 0.0  # 무분위: 어떤 구간에도 해당 없는 점수



                    total_cb = 0.0



                    for score_str, bal_m in cb_raw_pdata.items():



                        try:



                            score = int(score_str)



                        except (ValueError, TypeError):



                            continue



                        bal_won = float(bal_m) * 1_000_000



                        total_cb += bal_won



                        # 해당 분위 찾기



                        matched = False



                        for cfg_item in cb_cfg:



                            lo = cfg_item.get("lo", 0)



                            hi = cfg_item.get("hi", 9999)



                            grade = cfg_item.get("grade", "")



                            if lo <= score <= hi:



                                grade_bucket[grade] = grade_bucket.get(grade, 0.0) + bal_won



                                matched = True



                                break



                        if not matched:



                            no_grade_total += bal_won







                    # 섹션 삽입 — 10분위→1분위 내림차순, 무분위, 전체융잔 합계 순



                    GRADE_ORDER_DESC = list(reversed([item["grade"] for item in cb_cfg if item.get("grade")]))



                    for grade_key in GRADE_ORDER_DESC:



                        bal_won = grade_bucket.get(grade_key, 0.0)



                        if grade_key not in cb_sec_out:



                            cb_sec_out[grade_key] = {}



                        cb_sec_out[grade_key][ym] = cb_sec_out[grade_key].get(ym, 0.0) + bal_won



                    # 무분위



                    if "무분위" not in cb_sec_out:



                        cb_sec_out["무분위"] = {}



                    cb_sec_out["무분위"][ym] = cb_sec_out["무분위"].get(ym, 0.0) + no_grade_total



                    # 전체융잔 합계



                    if total_cb > 0:



                        if "전체융잔 합계" not in cb_sec_out:



                            cb_sec_out["전체융잔 합계"] = {}



                        cb_sec_out["전체융잔 합계"][ym] = cb_sec_out["전체융잔 합계"].get(ym, 0.0) + total_cb







        sections = result.get("sections", {})



        if gender_sec_out:



            sections["성별"] = gender_sec_out



        if age_sec_out:



            sections["연령별"] = age_sec_out



        if job_sec_out:



            sections["직업별"] = job_sec_out



        if region_sec_out:



            sections["지역별"] = region_sec_out



        if cb_sec_out:



            sections["CB분위"] = cb_sec_out



        result["sections"] = sections







    # contract avg_rate_deal -> loan asset average deal rate dynamic merge



    contract_data = uploaded_data.get("contract_data") or {}



    avg_rate_deal = contract_data.get("avg_rate_deal") or {}



    if avg_rate_deal:



        sections = result.setdefault("sections", {})



        rate_sec = {k: dict(v) for k, v in (sections.get("평균이자율") or {}).items()}



        deal_rate_row = rate_sec.setdefault("취급대출평균", {})



        for ym, val in avg_rate_deal.items():



            if val is None:



                continue



            deal_rate_row[str(ym)] = float(val)



            if isinstance(ym, str) and re.fullmatch(r"\d{4}-\d{2}", ym):



                periods = result.setdefault("periods", [])



                if ym not in periods:



                    periods.append(ym)



        if "periods" in result:



            result["periods"] = sorted(result.get("periods") or [])



        sections["평균이자율"] = rate_sec



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



        _saq_file = _data_path("settlement_aq")



        if os.path.exists(_saq_file):



            with open(_saq_file, encoding="utf-8") as _f:



                saq_store = json.load(_f)



            uploaded_data["settlement_aq"] = saq_store  # 메모리에도 캐시



    if not saq_store:



        return JSONResponse({"error": "결산자료 데이터가 없습니다. 결산자료를 먼저 업로드하세요."}, status_code=404)







    # asset_data: 메모리에 없으면 파일에서 로드



    asset_data = uploaded_data.get("asset_data")



    if asset_data is None:



        _asset_path = _data_path("asset_data")



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



        if not pdata.get("amount_band") or not pdata.get("channel_amount"):



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



                            _pg_path = _data_path("product_groups")



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
                pdata["channel_amount"] = _reparsed.get("channel_amount", pdata.get("channel_amount", {}))



                maturity = pdata["maturity"]



                saq_store[pkey] = pdata



                # saq_store 파일에도 저장



                _saq_save = _data_path("settlement_aq")



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



        # channel amount -> asset_data.sections

        channel_amount = pdata.get("channel_amount", {})

        if channel_amount:

            sections = asset_data.get("sections", {})

            sec_channel = sections.get("접수경로", {})

            for row_key, val in channel_amount.items():

                if val is not None:

                    if row_key not in sec_channel:

                        sec_channel[row_key] = {}

                    sec_channel[row_key][pkey] = float(val) * 1_000_000

            sections["접수경로"] = sec_channel

            asset_data["sections"] = sections

        results.append({"period": pkey, "status": "merged", "maturity": maturity, "avg_rate": avg_rate, "repayment": repayment, "amount_band": amount_band, "channel_amount": channel_amount})







    # ── rebuild 결과를 파일에 저장 ──────────────────────────────



    if any(r.get("status") == "merged" for r in results) and asset_data is not None:



        _asset_save_path = _data_path("asset_data")



        try:



            with open(_asset_save_path, "w", encoding="utf-8") as _f:



                json.dump(asset_data, _f, ensure_ascii=False, indent=2)



        except Exception as _e:



            pass  # 저장 실패는 무시 (메모리 반영은 이미 완료)







    return JSONResponse({"success": True, "results": results})











if __name__ == "__main__":



    import uvicorn



    uvicorn.run(app, host="0.0.0.0", port=3000)
