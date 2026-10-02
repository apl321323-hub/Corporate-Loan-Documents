from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request, Body



from fastapi.middleware.cors import CORSMiddleware



from fastapi.staticfiles import StaticFiles



from fastapi.responses import FileResponse, JSONResponse, StreamingResponse



import tempfile



import os



import json



import io



import re



import copy
import calendar
import threading
import time

import base64
import ipaddress



from datetime import date, datetime, timedelta



from excel_parser import (
    parse_excel_file,
    parse_settlement_asset_quality,
    parse_company_info,
    parse_bsis,
    parse_financial_statement_source,
)



from asset_quality_parser import parse_asset_quality_excel, SECTION1_ROWS, SECTION2_ROWS, SECTION3_GROUPS, SECTION4_PRODUCTS, PRODUCT_ROW_OFFSETS



from settlement_aq_parser import parse_settlement_asset_quality



from business_parser import parse_business_excel



from contract_parser import parse_contract_list, normalize_period_key



from full_contract_parser import parse_full_contract_list, build_full_contract_period
from fms_borrowing_parser import parse_fms_borrowing_list
from fms_payment_parser import parse_fms_payment_list
from borrowing_contract_parser import parse_borrowing_contract_list

from payment_parser import parse_payment, validate_payment_upload
from storage_guard import COMMON_UPLOAD_KEYS, StorageScopeError, validated_write_plan

from supabase_store import SupabaseStore



from writeoff_parser import parse_writeoff



from sale_parser import parse_sale



from asset_parser import parse_asset



from loan_count_parser import parse_loan_count



from pdf_parser import parse_pdf_fs, apply_pdf_to_bsis



import openpyxl


def _load_upload_workbook(tmp_path: str, filename: str | None = None, *, data_only: bool = True):
    """Load modern .xlsx files with openpyxl and legacy .xls files with xlrd."""
    ext = os.path.splitext(filename or tmp_path)[1].lower()
    if ext == ".xls":
        try:
            import xlrd
        except ImportError as exc:
            raise ValueError("구형 Excel(.xls) 파일을 읽기 위한 xlrd 라이브러리가 서버에 없습니다.") from exc

        book = xlrd.open_workbook(tmp_path)
        converted = openpyxl.Workbook()
        converted.remove(converted.active)

        for sheet_index in range(book.nsheets):
            src = book.sheet_by_index(sheet_index)
            ws = converted.create_sheet(title=src.name[:31] or f"Sheet{sheet_index + 1}")
            for row_idx in range(src.nrows):
                for col_idx in range(src.ncols):
                    cell = src.cell(row_idx, col_idx)
                    if cell.ctype == xlrd.XL_CELL_EMPTY:
                        continue
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            value = xlrd.xldate_as_datetime(value, book.datemode)
                        except Exception:
                            pass
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        value = bool(value)
                    elif cell.ctype == xlrd.XL_CELL_NUMBER and isinstance(value, float) and value.is_integer():
                        value = int(value)
                    ws.cell(row=row_idx + 1, column=col_idx + 1, value=value)
        return converted

    return openpyxl.load_workbook(tmp_path, data_only=data_only)







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


ASSET_REGION_ROWS = ["서울", "부산", "대구", "경기/인천", "강원", "충청/대전", "전라/광주", "경상/울산", "제주", "기타"]


def _asset_region_bucket(value) -> str:
    text = str(value or "").replace("\xa0", " ").strip()
    compact = re.sub(r"\s+", "", text)
    if not compact:
        return "기타"
    if "서울" in compact:
        return "서울"
    if "부산" in compact:
        return "부산"
    if "대구" in compact:
        return "대구"
    if "경기" in compact or "인천" in compact:
        return "경기/인천"
    if "강원" in compact:
        return "강원"
    if any(token in compact for token in ("충청", "충북", "충남", "대전", "세종")):
        return "충청/대전"
    if any(token in compact for token in ("전라", "전북", "전남", "광주")):
        return "전라/광주"
    if any(token in compact for token in ("경상", "경북", "경남", "울산")):
        return "경상/울산"
    if "제주" in compact:
        return "제주"
    return "기타"







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


def _bundled_template_path(filename: str) -> str | None:
    path = os.path.join(os.path.dirname(__file__), "templates", filename)
    if os.path.exists(path):
        return path
    return None


def _load_workbook_without_external_links(source):
    wb = openpyxl.load_workbook(source, keep_links=False)
    if hasattr(wb, "_external_links"):
        wb._external_links = []
    return wb







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
DATA_CACHE_TTL = max(0.0, float(os.environ.get("DATA_CACHE_TTL") or "5"))
_DATA_CACHE: dict[str, tuple[float, object]] = {}
_DATA_STORE_LOCK = threading.RLock()


def _save_local_data(key: str, value) -> None:
    with open(_data_path(key), "w", encoding="utf-8") as f:
        json.dump(value, f, ensure_ascii=False, indent=2)


def _atomic_save_local_many(values: dict) -> None:
    staged: dict[str, str] = {}
    previous: dict[str, bytes | None] = {}
    replaced: list[str] = []
    try:
        for key, value in values.items():
            target = _data_path(key)
            previous[target] = open(target, "rb").read() if os.path.exists(target) else None
            fd, temp_path = tempfile.mkstemp(prefix=f".{key}.", suffix=".tmp", dir=os.path.dirname(target))
            with os.fdopen(fd, "w", encoding="utf-8") as stream:
                json.dump(value, stream, ensure_ascii=False, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
            staged[target] = temp_path
        for target, temp_path in staged.items():
            os.replace(temp_path, target)
            replaced.append(target)
        staged.clear()
    except Exception:
        for target in reversed(replaced):
            old = previous[target]
            if old is None:
                if os.path.exists(target):
                    os.remove(target)
            else:
                fd, restore_path = tempfile.mkstemp(prefix=".restore.", suffix=".tmp", dir=os.path.dirname(target))
                with os.fdopen(fd, "wb") as stream:
                    stream.write(old)
                    stream.flush()
                    os.fsync(stream.fileno())
                os.replace(restore_path, target)
        raise
    finally:
        for temp_path in staged.values():
            if os.path.exists(temp_path):
                os.remove(temp_path)


def save_many_data(values: dict) -> None:
    if not isinstance(values, dict) or not values:
        return
    with _DATA_STORE_LOCK:
        if SUPABASE_STORE.enabled:
            SUPABASE_STORE.save_many(values)
        _atomic_save_local_many(values)
        now = time.monotonic()
        for key, value in values.items():
            _DATA_CACHE[str(key)] = (now, copy.deepcopy(value))


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



    save_many_data({str(key): value})







def load_data(key: str):



    """Supabase/로컬 JSON에서 데이터 로드 (없으면 None)"""



    cache_key = str(key)
    with _DATA_STORE_LOCK:
        cached = _DATA_CACHE.get(cache_key)
        if cached and time.monotonic() - cached[0] <= DATA_CACHE_TTL:
            return copy.deepcopy(cached[1])

    if SUPABASE_STORE.enabled and SUPABASE_READ_FIRST:
        try:
            remote = SUPABASE_STORE.load(key)
            if remote is not None:
                _save_local_data(key, remote)
                with _DATA_STORE_LOCK:
                    _DATA_CACHE[cache_key] = (time.monotonic(), copy.deepcopy(remote))
                return remote
        except Exception as ex:
            print(f"[supabase] 로드 실패({key}): {ex}")

    local = _load_local_data(key)
    if local is not None:
        with _DATA_STORE_LOCK:
            _DATA_CACHE[cache_key] = (time.monotonic(), copy.deepcopy(local))
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

    def get_many(self, keys) -> dict:
        clean = [str(key) for key in dict.fromkeys(keys or [])]
        result = {}
        missing = []
        now = time.monotonic()
        with _DATA_STORE_LOCK:
            for key in clean:
                cached = _DATA_CACHE.get(key)
                if cached and now - cached[0] <= DATA_CACHE_TTL:
                    result[key] = copy.deepcopy(cached[1])
                else:
                    missing.append(key)
        if missing and SUPABASE_STORE.enabled and SUPABASE_READ_FIRST:
            try:
                remote = SUPABASE_STORE.load_many(missing)
                if remote:
                    _atomic_save_local_many(remote)
                    with _DATA_STORE_LOCK:
                        stamp = time.monotonic()
                        for key, value in remote.items():
                            _DATA_CACHE[key] = (stamp, copy.deepcopy(value))
                    result.update(remote)
                    missing = [key for key in missing if key not in remote]
            except Exception as ex:
                print(f"[supabase] 일괄 로드 실패: {ex}")
        for key in missing:
            local = _load_local_data(key)
            if local is not None:
                result[key] = local
                with _DATA_STORE_LOCK:
                    _DATA_CACHE[key] = (time.monotonic(), copy.deepcopy(local))
        return result







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
        save_many_data(d)







    def keys(self):



        names = [f[:-5] for f in os.listdir(DATA_DIR) if f.endswith(".json")]



        if SUPABASE_STORE.enabled:
            try:
                names = sorted(set(names) | set(SUPABASE_STORE.keys()))
            except Exception as ex:
                print(f"[supabase] 키 목록 로드 실패: {ex}")
        return names







uploaded_data = PersistentStore()

UPLOAD_HISTORY_KEY = "upload_history"
UPLOAD_HISTORY_LIMIT = 5


def _clean_upload_history_record(value: dict) -> dict | None:
    if not isinstance(value, dict):
        return None
    sector = str(value.get("sector") or "").strip()
    filename = str(value.get("filename") or value.get("fileName") or "").strip()
    if not sector or not filename:
        return None
    uploader = str(value.get("uploader") or value.get("user") or "").strip() or "미지정"
    uploaded_at = str(value.get("uploadedAt") or value.get("uploaded_at") or "").strip()
    if not uploaded_at:
        uploaded_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    return {
        "sector": sector,
        "filename": filename,
        "uploader": uploader,
        "uploadedAt": uploaded_at,
        "period": str(value.get("period") or "").strip(),
    }


def _sanitize_upload_history(value) -> dict:
    if not isinstance(value, dict):
        return {}
    cleaned = {}
    for sector, rows in value.items():
        if not isinstance(rows, list):
            continue
        sector_key = str(sector or "").strip()
        if not sector_key:
            continue
        records = []
        for row in rows:
            record = _clean_upload_history_record({**row, "sector": row.get("sector") or sector_key} if isinstance(row, dict) else {})
            if record:
                records.append(record)
        cleaned[sector_key] = records[:UPLOAD_HISTORY_LIMIT]
    return cleaned


def _load_upload_history() -> dict:
    history = uploaded_data.get(UPLOAD_HISTORY_KEY) or {}
    history = _sanitize_upload_history(history)
    uploaded_data[UPLOAD_HISTORY_KEY] = history
    return history


@app.get("/api/upload-history")
async def get_upload_history():
    return JSONResponse({"history": _load_upload_history()})


@app.post("/api/upload-history")
async def save_upload_history_record(request: Request):
    body = await request.json()
    record = _clean_upload_history_record(body)
    if not record:
        raise HTTPException(status_code=400, detail="업로드 이력에는 sector와 filename이 필요합니다.")
    history = _load_upload_history()
    sector_rows = history.get(record["sector"], [])
    history[record["sector"]] = [record] + sector_rows
    history[record["sector"]] = history[record["sector"]][:UPLOAD_HISTORY_LIMIT]
    uploaded_data[UPLOAD_HISTORY_KEY] = history
    return JSONResponse({"success": True, "history": history, "record": record})


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



    filename = file.filename or ""
    if not filename.lower().endswith(('.xlsx', '.xls')):



        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")







    try:



        suffix = os.path.splitext(filename)[1].lower() or '.xlsx'
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:



            content = await file.read()



            tmp.write(content)



            tmp_path = tmp.name







        data = parse_excel_file(tmp_path)
        try:
            write_plan = validated_write_plan(data, COMMON_UPLOAD_KEYS)
        except StorageScopeError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc



        uploaded_data.update(write_plan)



        if period:



            uploaded_data["upload_period_main"] = period







        os.unlink(tmp_path)







        return JSONResponse({



            "success": True,



            "message": f"파일 '{file.filename}' 업로드 및 분석 완료",



            "sheets": list(write_plan.keys()),



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



        "business_status", "loan_asset", "borrowing", "fms_borrowing_data", "fms_payment_data", "borrowing_contract_data",



        "company_info", "company_profiles", "settlement_asset_quality",



        "asset_quality_detail",



    ]



    result = uploaded_data.get_many(_DATA_KEYS)



    return JSONResponse(result, headers={"Cache-Control": "no-store, no-cache, must-revalidate"})











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







def _ci_copy_row_style(ws, source_row: int, target_row: int, min_col: int = 1, max_col: int = 6) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(min_col, max_col + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if source.has_style:
            target._style = copy.copy(source._style)
        target.number_format = source.number_format
        target.alignment = copy.copy(source.alignment)
        target.protection = copy.copy(source.protection)


def _ci_range_coord(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    from openpyxl.utils import get_column_letter

    return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"


def _ci_insert_rows_preserving_merges(ws, insert_at: int, amount: int) -> None:
    if amount <= 0:
        return

    adjusted = []
    for rng in list(ws.merged_cells.ranges):
        min_row, min_col, max_row, max_col = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        if max_row >= insert_at:
            ws.unmerge_cells(str(rng))
        if min_row >= insert_at:
            min_row += amount
            max_row += amount
        elif min_row < insert_at <= max_row:
            max_row += amount
        adjusted.append((min_row, min_col, max_row, max_col))

    ws.insert_rows(insert_at, amount)

    for min_row, min_col, max_row, max_col in adjusted:
        try:
            ws.merge_cells(_ci_range_coord(min_row, min_col, max_row, max_col))
        except ValueError:
            pass


def _ci_remerge_vertical_label(ws, start_row: int, end_row: int, new_end_row: int) -> None:
    try:
        ws.unmerge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
    except ValueError:
        pass
    ws.merge_cells(start_row=start_row, start_column=1, end_row=new_end_row, end_column=1)


def _ci_update_print_area(ws, extra_rows: int) -> None:
    try:
        ws.print_area = f"A1:F{max(ws.max_row, 68 + extra_rows)}"
    except Exception:
        pass


def _ci_expand_ceo_history_rows(ws, ceo_history: list[dict]) -> int:
    base_rows = 8
    extra_rows = max(0, len(ceo_history) - base_rows)
    if extra_rows <= 0:
        return 0

    insert_at = 25
    _ci_insert_rows_preserving_merges(ws, insert_at, extra_rows)

    _ci_remerge_vertical_label(ws, 16, 24, 24 + extra_rows)

    for offset, item in enumerate(ceo_history[base_rows:]):
        row = insert_at + offset
        _ci_copy_row_style(ws, insert_at - 1, row)
        try:
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        except ValueError:
            pass
        _ci_clear(ws, [f"B{row}", f"C{row}"])
        item = item or {}
        _ci_set(ws, f"B{row}", item.get("date"))
        _ci_set(ws, f"C{row}", item.get("content"))

    _ci_update_print_area(ws, extra_rows)
    return extra_rows


def _ci_expand_company_history_rows(ws, company_history: list[dict], ceo_extra_rows: int = 0) -> int:
    base_rows = 15
    extra_rows = max(0, len(company_history) - base_rows)
    if extra_rows <= 0:
        return 0

    label_start = 25 + ceo_extra_rows
    label_end = 40 + ceo_extra_rows
    insert_at = label_end + 1
    _ci_insert_rows_preserving_merges(ws, insert_at, extra_rows)
    _ci_remerge_vertical_label(ws, label_start, label_end, label_end + extra_rows)

    for offset, item in enumerate(company_history[base_rows:]):
        row = insert_at + offset
        _ci_copy_row_style(ws, insert_at - 1, row)
        try:
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        except ValueError:
            pass
        _ci_clear(ws, [f"B{row}", f"C{row}"])
        item = item or {}
        _ci_set(ws, f"B{row}", item.get("date"))
        _ci_set(ws, f"C{row}", item.get("content"))

    _ci_update_print_area(ws, ceo_extra_rows + extra_rows)
    return extra_rows


def _ci_expand_guarantor_history_rows(ws, guarantor_history: list[dict], prior_extra_rows: int = 0) -> int:
    base_rows = 8
    extra_rows = max(0, len(guarantor_history) - base_rows)
    if extra_rows <= 0:
        return 0

    label_start = 55 + prior_extra_rows
    label_end = 63 + prior_extra_rows
    insert_at = label_end + 1
    _ci_insert_rows_preserving_merges(ws, insert_at, extra_rows)
    _ci_remerge_vertical_label(ws, label_start, label_end, label_end + extra_rows)

    for offset, item in enumerate(guarantor_history[base_rows:]):
        row = insert_at + offset
        _ci_copy_row_style(ws, insert_at - 1, row)
        try:
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        except ValueError:
            pass
        _ci_clear(ws, [f"B{row}", f"C{row}"])
        item = item or {}
        _ci_set(ws, f"B{row}", item.get("date"))
        _ci_set(ws, f"C{row}", item.get("content"))

    _ci_update_print_area(ws, prior_extra_rows + extra_rows)
    return extra_rows


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
    ceo_extra_rows = _ci_expand_ceo_history_rows(ws, ceo_history)
    company_extra_rows = _ci_expand_company_history_rows(ws, company_history, ceo_extra_rows)
    _ci_expand_guarantor_history_rows(ws, guarantor_history, ceo_extra_rows + company_extra_rows)







def _build_company_info_workbook(data: dict):
    """Build a styled company-info workbook when the uploaded template is unavailable."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "\uae30\uc5c5\uc815\ubcf4"
    ws.sheet_view.showGridLines = False

    labels = {
        "A1": "1. \uae30\uc5c5\uc815\ubcf4",
        "A3": "\u25a0 \ucc28\uc8fc\uc0ac",
        "A4": "\ucc28\uc8fc\uba85", "C4": "\ub300\ud45c\uc774\uc0ac", "E4": "\uc2e4\uacbd\uc601\uc790",
        "A5": "\uc0ac\uc5c5\uc790\ubc88\ud638", "C5": "\uc5c5\uc885", "E5": "\uae30\uc5c5\ud615\ud0dc",
        "A6": "\uc804\ud654\ubc88\ud638", "C6": "\ub300\ud45c\uc774\uc0ac \uc804\ud654\ubc88\ud638", "E6": "\uc0c1\uc7a5",
        "A7": "\uc784\uc9c1\uc6d0\uc218", "C7": "\uc0ac\uc5c5\uc7a5\uc8fc\uc18c",
        "A8": "\uc8fc\uc8fc\ud604\ud669", "B8": "\uc8fc\uc8fc\uba85", "C8": "\uc8fc\uc2dd\uc218",
        "D8": "\uc9c0\ubd84\uc728", "E8": "\uc8fc\uc2dd\uae08\uc561", "F8": "\uc774\ud574\uad00\uacc4",
        "A15": "\ub4f1\uae30\uc784\uc6d0",
        "A16": "\ub300\ud45c\uc774\uc0ac(\uc2e4\uacbd\uc601\uc790)\n\uc774\ub825", "B16": "\uc77c\uc790(\ub144/\uc6d4)", "C16": "\ub0b4\uc6a9",
        "A25": "\uc5f0\ud601", "B25": "\uc77c\uc790(\ub144/\uc6d4)", "C25": "\ub0b4\uc6a9",
        "A42": "\u25a0 \ubcf4\uc99d(\ubc95)\uc778",
        "A43": "\ubcf4\uc99d(\ubc95)\uc778\uba85", "E43": "\uc2e4\uacbd\uc601\uc790",
        "A44": "\uc0ac\uc5c5\uc790\ubc88\ud638", "C44": "\uc5c5\uc885", "E44": "\uae30\uc5c5\ud615\ud0dc",
        "A45": "\uc804\ud654\ubc88\ud638", "C45": "\ub300\ud45c\uc774\uc0ac \uc804\ud654\ubc88\ud638", "E45": "\uc0c1\uc7a5",
        "A46": "\uc784\uc9c1\uc6d0\uc218", "C46": "\uc0ac\uc5c5\uc7a5\uc8fc\uc18c",
        "A47": "\uc8fc\uc8fc\ud604\ud669", "B47": "\uc8fc\uc8fc\uba85", "C47": "\uc8fc\uc2dd\uc218",
        "D47": "\uc9c0\ubd84\uc728", "E47": "\uc8fc\uc2dd\uae08\uc561", "F47": "\uc774\ud574\uad00\uacc4",
        "A54": "\ub4f1\uae30\uc784\uc6d0",
        "A55": "\uc5f0\ud601(\uacbd\ub825)", "B55": "\uc77c\uc790(\ub144/\uc6d4)", "C55": "\ub0b4\uc6a9",
        "A65": "\u25a0 \uae30\ud0c0", "A66": "\uad6c\ubd84", "B66": "\ub0b4\uc6a9",
        "A67": "\ucc28\uc8fc\uc0ac \uc601\uc5c5\ubc29\ud5a5 \ubc0f \uc7a5\uc810", "A68": "\uad00\uacc4\uc0ac \ubc0f \uacc4\uc5f4\uc0ac",
    }
    for coord, value in labels.items():
        ws[coord] = value

    fixed_merges = (
        "D7:F7", "A8:A14", "A16:A24", "C16:F16", "A25:A40", "C25:F25",
        "D46:F46", "A47:A53", "A55:A63", "C55:F55",
        "B66:F66", "B67:F67", "B68:F68",
    )
    for cell_range in fixed_merges:
        ws.merge_cells(cell_range)
    for row in list(range(17, 25)) + list(range(26, 41)) + list(range(56, 64)):
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)

    _write_company_info_to_sheet(ws, data)

    navy, light_blue, pale_blue = "1F4E78", "D9EAF7", "EAF2F8"
    thin = Side(style="thin", color="AFC3D7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    default_font = Font(name="Malgun Gothic", size=10, color="17365D")
    max_style_row = max(68, ws.max_row)
    for row in ws.iter_rows(min_row=1, max_row=max_style_row, min_col=1, max_col=6):
        for cell in row:
            cell.font = default_font
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for row_num in (1, 3, 42, 65):
        for cell in ws[row_num][:6]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(name="Malgun Gothic", size=11, bold=True, color="FFFFFF")
    for row_num in (8, 16, 25, 47, 55, 66):
        for cell in ws[row_num][:6]:
            cell.fill = PatternFill("solid", fgColor=light_blue)
            cell.font = Font(name="Malgun Gothic", size=10, bold=True, color="17365D")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_num in (4, 5, 6, 7, 43, 44, 45, 46, 67, 68):
        for col_num in (1, 3, 5):
            cell = ws.cell(row_num, col_num)
            cell.fill = PatternFill("solid", fgColor=pale_blue)
            cell.font = Font(name="Malgun Gothic", size=10, bold=True, color="17365D")
    for coord in ("A8", "A16", "A25", "A47", "A55"):
        ws[coord].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col, width in {"A": 21, "B": 22, "C": 18, "D": 22, "E": 18, "F": 22}.items():
        ws.column_dimensions[col].width = width
    for row_num, height in {1: 25, 3: 22, 42: 22, 65: 22, 67: 90, 68: 60}.items():
        ws.row_dimensions[row_num].height = height
    ws.freeze_panes = "A3"
    ws.print_area = f"A1:F{max_style_row}"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return wb



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







        wb = _load_upload_workbook(tmp_path, file.filename, data_only=True)



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

            "content_base64": base64.b64encode(content).decode("ascii"),



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


COMPANY_PROFILE_KEY = "company_profiles"


def _normalize_company_profile_period(value) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    match = re.search(r"(\d{4})\D{0,4}(\d{1,2})", text)
    if not match:
        match = re.search(r"(\d{2})\D{0,4}(\d{1,2})", text)
    if not match and re.fullmatch(r"\d{6}", text):
        match = re.match(r"(\d{4})(\d{2})", text)
    if not match and re.fullmatch(r"\d{4}", text):
        match = re.match(r"(\d{2})(\d{2})", text)
    if not match:
        return ""
    year = int(match.group(1))
    month = int(match.group(2))
    if year < 100:
        year += 2000
    if month < 1 or month > 12:
        return ""
    return f"{year:04d}.{month:02d}"


def _company_profile_store() -> dict:
    profiles = uploaded_data.get(COMPANY_PROFILE_KEY, {})
    return profiles if isinstance(profiles, dict) else {}


def _normalized_company_profile_store() -> dict:
    normalized = {}
    for key, value in _company_profile_store().items():
        period = _normalize_company_profile_period(key)
        if period:
            normalized[period] = value
    return normalized


@app.get("/api/company-profile/periods")
async def get_company_profile_periods():
    profiles = _normalized_company_profile_store()
    periods = sorted(profiles.keys(), reverse=True)
    return {"periods": periods}


@app.get("/api/company-profile")
async def get_company_profile(period: str = ""):
    profiles = _normalized_company_profile_store()
    periods = sorted(profiles.keys(), reverse=True)
    normalized = _normalize_company_profile_period(period) or (periods[0] if periods else "")
    return {"period": normalized, "data": profiles.get(normalized, {}) if normalized else {}}


@app.post("/api/company-profile")
async def save_company_profile(payload: dict):
    data = payload.get("data") if isinstance(payload.get("data"), dict) else payload
    normalized = _normalize_company_profile_period(payload.get("period") or data.get("basedate"))
    if not normalized:
        raise HTTPException(status_code=400, detail="period is required")
    profiles = dict(_company_profile_store())
    profile = dict(data or {})
    profile["basedate"] = normalized
    profile["period"] = normalized
    profile["updated_at"] = datetime.now().isoformat(timespec="seconds")
    profiles[normalized] = profile
    uploaded_data[COMPANY_PROFILE_KEY] = profiles
    save_data(COMPANY_PROFILE_KEY, profiles)
    return {
        "success": True,
        "period": normalized,
        "data": profile,
        "periods": sorted(profiles.keys(), reverse=True),
    }















@app.get("/api/export/company_info")



async def export_company_info():



    """업로드한 기업현황 엑셀 서식을 보존하여 현재 기업정보를 다운로드"""



    data = uploaded_data.get("company_info")



    if not data:



        raise HTTPException(status_code=404, detail="기업정보 데이터가 없습니다. 기업현황 엑셀을 먼저 업로드하세요.")







    template_path = _stored_file_path("company_info_template.xlsx")



    try:



        meta = uploaded_data.get("company_info_template", {}) or {}



        encoded_template = meta.get("content_base64")

        bundled_template_path = _bundled_template_path("company_info_template.xlsx")

        if template_path:

            wb = _load_workbook_without_external_links(template_path)

        elif encoded_template:

            wb = _load_workbook_without_external_links(io.BytesIO(base64.b64decode(encoded_template)))

        elif bundled_template_path:

            wb = _load_workbook_without_external_links(bundled_template_path)

        else:

            wb = _build_company_info_workbook(data)



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



            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",
                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
                "Pragma": "no-cache",
            },



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


REPORT_DATA_MIN_PERIOD = "2020-12"


def _report_norm_period_key(value):
    m = re.search(r"(\d{4})[-.](\d{1,2})", str(value or "").strip())
    if not m:
        return ""
    return f"{m.group(1)}-{int(m.group(2)):02d}"


def _report_period_is_visible(value):
    period = _report_norm_period_key(value)
    return bool(period and period >= REPORT_DATA_MIN_PERIOD)


def _prune_period_store(store):
    if not isinstance(store, dict):
        return store
    return {
        key: value
        for key, value in store.items()
        if _report_period_is_visible(key)
    }


def _prune_indexed_period_data(data):
    if not isinstance(data, dict):
        return data
    raw_periods = list(data.get("periods") or [])
    if not raw_periods:
        return data
    keep = [
        (idx, period)
        for idx, period in enumerate(raw_periods)
        if _report_period_is_visible(period)
    ]
    data["periods"] = [period for _, period in keep]

    def prune(obj, key_name=""):
        if isinstance(obj, list):
            if key_name != "periods" and len(obj) == len(raw_periods):
                return [obj[idx] if idx < len(obj) else None for idx, _ in keep]
            return [prune(item) for item in obj]
        if isinstance(obj, dict):
            for key in list(obj.keys()):
                norm = _report_norm_period_key(key)
                if norm and norm < REPORT_DATA_MIN_PERIOD:
                    obj.pop(key, None)
                    continue
                obj[key] = prune(obj.get(key), str(key))
        return obj

    for key in list(data.keys()):
        if key != "periods":
            data[key] = prune(data.get(key), key)
    return data


def _periods_from_periodic_payload(data):
    if not isinstance(data, dict):
        return []
    periods = set()
    for period in data.get("periods") or []:
        norm = _report_norm_period_key(period)
        if norm and _report_period_is_visible(norm):
            periods.add(norm)
    for key in data.keys():
        norm = _report_norm_period_key(key)
        if norm and _report_period_is_visible(norm):
            periods.add(norm)
    return sorted(periods)


def _source_report_periods(*keys):
    periods = set()
    for key in keys:
        periods.update(_periods_from_periodic_payload(uploaded_data.get(key) or {}))
    return sorted(periods)


def _asset_data_is_cleared_marker(data):
    return isinstance(data, dict) and bool(data.get("_cleared_for_reupload"))


def _empty_asset_data_shell(periods=None):
    clean_periods = sorted({
        _report_norm_period_key(period)
        for period in (periods or [])
        if _report_norm_period_key(period) and _report_period_is_visible(period)
    })
    return {"periods": clean_periods, "sections": {}, "cb_products": {}}


_BLANK_CHANNEL_FALLBACK = "\uae30\ud0c0"
_LOAN_ASSET_TOTAL_KEY = "\uc804\uccb4\uc735\uc794 \ud569\uacc4"


def _settlement_number(value):
    if value is None or value == "":
        return 0.0
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


def _normalize_settlement_channel_amount(pdata):
    if not isinstance(pdata, dict):
        return pdata
    rows = pdata.get("loan_loss_rows") or []
    if not isinstance(rows, list) or not rows:
        return pdata

    channel_totals = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        balance = _settlement_number(row.get("balance"))
        if not balance and row.get("balance_million") is not None:
            balance = _settlement_number(row.get("balance_million")) * 1_000_000
        if not balance:
            continue
        channel = str(row.get("channel") or row.get("channel_raw") or "").strip()
        if not channel or channel == "None":
            channel = _BLANK_CHANNEL_FALLBACK
        channel_totals[channel] = channel_totals.get(channel, 0.0) + balance

    if not channel_totals:
        return pdata

    normalized = dict(pdata)
    normalized["channel_amount"] = {
        channel: amount / 1_000_000 for channel, amount in sorted(channel_totals.items())
    }
    normalized["channel_amount"][_LOAN_ASSET_TOTAL_KEY] = sum(channel_totals.values()) / 1_000_000
    normalized["channels"] = sorted(channel_totals.keys())
    return normalized


def _normalize_settlement_aq_store(store):
    if not isinstance(store, dict):
        return store
    return {
        key: _normalize_settlement_channel_amount(value)
        for key, value in store.items()
    }


def _ensure_asset_data_shell(periods=None):
    """Ensure derived loan-asset uploads have a writable asset_data target."""
    asset_data = uploaded_data.get("asset_data")
    if _asset_data_is_cleared_marker(asset_data) or not isinstance(asset_data, dict):
        asset_data = _empty_asset_data_shell(periods)
    else:
        asset_data.setdefault("periods", [])
        asset_data.setdefault("sections", {})
        asset_data.setdefault("cb_products", {})
        for period in periods or []:
            norm = _report_norm_period_key(period)
            if norm and _report_period_is_visible(norm) and norm not in asset_data["periods"]:
                asset_data["periods"].append(norm)
        asset_data["periods"] = sorted(
            {
                _report_norm_period_key(period)
                for period in (asset_data.get("periods") or [])
                if _report_norm_period_key(period) and _report_period_is_visible(period)
            }
        )
    uploaded_data["asset_data"] = asset_data
    return asset_data











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



        if not period:



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











AQ_31PLUS_OVERDUE_KEYS = ["31~60", "61~90", "91~120", "121~150", "151~180", "181~"]


def _aq_series_value(section, label, index):
    values = (section or {}).get(label)
    if not isinstance(values, list) or index >= len(values):
        return None
    return values[index]


def _aq_31plus_overdue_values(data):
    section1 = data.get("section1") or {}
    periods = data.get("periods") or []
    existing = section1.get("31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4") or (data.get("section2") or {}).get("31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4")
    result = []
    for idx, _ in enumerate(periods):
        if isinstance(existing, list) and idx < len(existing) and isinstance(existing[idx], (int, float)) and not isinstance(existing[idx], bool):
            result.append(existing[idx])
            continue
        amount_sum = 0
        has_value = False
        for label in AQ_31PLUS_OVERDUE_KEYS:
            value = _aq_series_value(section1, label, idx)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                amount_sum += value
                has_value = True
        result.append(amount_sum if has_value else None)
    return result


def _aq_apply_31plus_overdue_sum(data):
    if not isinstance(data, dict):
        return False
    values = _aq_31plus_overdue_values(data)
    if not values:
        return False
    section1 = data.setdefault("section1", {})
    section2 = data.setdefault("section2", {})
    section1["31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4"] = values
    section2["31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4"] = values
    return any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values)


def _aq_apply_product_overdue_totals(data):
    if not isinstance(data, dict):
        return False
    section4 = data.get("section4") or {}
    if not isinstance(section4, dict):
        return False
    period_count = len(data.get("periods") or [])
    if period_count <= 0:
        return False
    overdue_keys = ["1~30", "31~60", "61~90", "91~120", "121~150", "151~180", "181~"]
    applied = False
    for product_data in section4.values():
        if not isinstance(product_data, dict):
            continue
        direct = product_data.get("\uc5f0\uccb4\ud569\uacc4")
        needs_fill = not isinstance(direct, list) or len(direct) < period_count or any(v is None for v in direct[:period_count])
        if not needs_fill:
            continue
        totals = []
        has_any = False
        for idx in range(period_count):
            amount_sum = 0
            has_value = False
            for key in overdue_keys:
                value = _aq_series_value(product_data, key, idx)
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    amount_sum += value
                    has_value = True
            totals.append(amount_sum if has_value else None)
            has_any = has_any or has_value
        if has_any:
            product_data["\uc5f0\uccb4\ud569\uacc4"] = totals
            applied = True
    return applied


def _aq_norm_group_product_name(value):
    return str(value or "").strip().replace(" ", "")


def _aq_product_group_pairs(product_groups):
    pairs = []
    seen = set()
    for group in product_groups or []:
        if not isinstance(group, dict):
            continue
        name = str(group.get("name") or "").strip()
        if not name or name in seen:
            continue
        products = []
        for product in group.get("products") or []:
            product_name = str(product or "").strip()
            if product_name:
                products.append(product_name)
        pairs.append((name, products))
        seen.add(name)
    return pairs


def _aq_product_value(product_data, label, index):
    value = _aq_series_value(product_data, label, index)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value), True
    return 0.0, False


def _aq_product_overdue_91_180(product_data, index):
    direct, has_direct = _aq_product_value(product_data, "91~180", index)
    if has_direct:
        return direct, True
    total = 0.0
    has_any = False
    for label in ("91~120", "121~150", "151~180"):
        value, has_value = _aq_product_value(product_data, label, index)
        if has_value:
            total += value
            has_any = True
    return total, has_any


def _aq_product_overdue_total(product_data, index):
    direct, has_direct = _aq_product_value(product_data, "\uc5f0\uccb4\ud569\uacc4", index)
    if has_direct:
        return direct, True
    total = 0.0
    has_any = False
    for label in ("1~30", "31~60", "61~90", "181~"):
        value, has_value = _aq_product_value(product_data, label, index)
        if has_value:
            total += value
            has_any = True
    value, has_value = _aq_product_overdue_91_180(product_data, index)
    if has_value:
        total += value
        has_any = True
    return total, has_any


def _aq_product_balance_total(product_data, index):
    direct, has_direct = _aq_product_value(product_data, "\uc735\uc794\ud569\uacc4", index)
    if has_direct:
        return direct, True
    normal, has_normal = _aq_product_value(product_data, "\ubb34\uc5f0\uccb4", index)
    overdue, has_overdue = _aq_product_overdue_total(product_data, index)
    if has_normal or has_overdue:
        return normal + overdue, True
    return 0.0, False


def _aq_rebuild_group_section_from_products(data, product_groups):
    if not isinstance(data, dict):
        return False
    group_pairs = _aq_product_group_pairs(product_groups)
    section4 = data.get("section4") or {}
    period_count = len(data.get("periods") or [])
    if not group_pairs or not isinstance(section4, dict) or period_count <= 0:
        return False

    product_key_by_norm = {
        _aq_norm_group_product_name(product): product
        for product in section4.keys()
    }

    rebuilt_any = False
    section3 = data.setdefault("section3", {})
    group_labels = ("\ubb34\uc5f0\uccb4", "1~30", "31~60", "61~90", "91~180", "181~", "\uc5f0\uccb4\ud569\uacc4", "\uc735\uc794\ud569\uacc4", "\uc5f0\uccb4\uc728(%)")

    for group_name, product_names in group_pairs:
        product_keys = []
        for product_name in product_names:
            key = product_key_by_norm.get(_aq_norm_group_product_name(product_name))
            if key and key not in product_keys:
                product_keys.append(key)

        group_data = {label: [] for label in group_labels}
        for index in range(period_count):
            values_by_label = {}
            for label in group_labels:
                if label == "91~180":
                    getter = _aq_product_overdue_91_180
                elif label == "\uc5f0\uccb4\ud569\uacc4":
                    getter = _aq_product_overdue_total
                elif label == "\uc735\uc794\ud569\uacc4":
                    getter = _aq_product_balance_total
                elif label == "\uc5f0\uccb4\uc728(%)":
                    continue
                else:
                    getter = lambda product_data, idx, row_label=label: _aq_product_value(product_data, row_label, idx)

                total = 0.0
                has_any = False
                for product_key in product_keys:
                    product_data = section4.get(product_key) or {}
                    value, has_value = getter(product_data, index)
                    if has_value:
                        total += value
                        has_any = True
                values_by_label[label] = int(round(total)) if has_any else None

            overdue_total = values_by_label.get("\uc5f0\uccb4\ud569\uacc4")
            balance_total = values_by_label.get("\uc735\uc794\ud569\uacc4")
            if isinstance(overdue_total, (int, float)) and isinstance(balance_total, (int, float)) and balance_total:
                rate_value = (float(overdue_total) / float(balance_total)) * 100
            else:
                rate_value = None

            for label in group_labels:
                group_data[label].append(rate_value if label == "\uc5f0\uccb4\uc728(%)" else values_by_label.get(label))

        section3[group_name] = group_data
        rebuilt_any = True

    return rebuilt_any


def _aq_business_section_value_million(business_data, period, label):
    if not isinstance(business_data, dict):
        return None
    norm_period = _business_norm_period(period)
    if not norm_period:
        return None
    periods = [_business_norm_period(p) for p in (business_data.get("periods") or [])]
    idx = periods.index(norm_period) if norm_period in periods else None
    if idx is None:
        return None

    section1 = business_data.get("section1") or {}
    normalized_target = label.replace(" ", "")
    for key in (label, normalized_target):
        series = section1.get(key)
        if isinstance(series, list) and idx is not None and idx < len(series):
            value = series[idx]
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                return value
    for key, series in section1.items():
        if str(key).replace(" ", "") != normalized_target:
            continue
        if isinstance(series, list) and idx is not None and idx < len(series):
            value = series[idx]
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                return value
    return None


def _aq_apply_business_ending_allowance(merged, business_data):
    if not isinstance(merged, dict) or not isinstance(business_data, dict):
        return False
    section2 = merged.setdefault("section2", {})
    applied = False
    business_periods = None
    business_cache = {}
    raw_periods = [_business_norm_period(p) for p in (business_data.get("periods") or []) if _business_norm_period(p)]
    latest_raw_period = max(raw_periods, key=_business_period_sort_key) if raw_periods else None
    for period in list(merged.get("periods") or []):
        norm_period = _business_norm_period(period)
        idx = _business_period_index(business_data, norm_period)
        value_million = _aq_business_section_value_million(business_data, period, "\uae30\ub9d0 \ucda9\ub2f9\uae08")
        if not isinstance(value_million, (int, float)) or isinstance(value_million, bool):
            should_calculate = latest_raw_period and norm_period and _business_period_sort_key(norm_period) > _business_period_sort_key(latest_raw_period)
            if should_calculate:
                try:
                    if business_periods is None:
                        business_periods = _business_collect_periods(business_data)
                    value_million = _business_s1_value(
                        business_data,
                        "\uae30\ub9d0 \ucda9\ub2f9\uae08",
                        idx,
                        norm_period,
                        business_periods,
                        business_cache,
                    )
                except Exception:
                    value_million = None
        if not isinstance(value_million, (int, float)) or isinstance(value_million, bool):
            continue
        index = _aq_ensure_period(merged, period)
        period_count = len(merged.get("periods") or [])
        _aq_set_series_value(section2, "\ub300\uc190\ucda9\ub2f9\uae08", index, int(round(float(value_million) * 1_000_000)), period_count)
        applied = True
    return applied


def _merge_settlement_into_asset_quality_detail(aq_data, settlement_store, business_data=None, product_groups=None):



    if not aq_data:



        return aq_data



    merged = copy.deepcopy(aq_data)



    merged["periods"] = [_aq_norm_period(p) for p in (merged.get("periods") or [])]



    if not isinstance(settlement_store, dict):



        settlement_store = {}







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







    if _aq_apply_31plus_overdue_sum(merged):



        applied = True

    if _aq_apply_product_overdue_totals(merged):



        applied = True

    if _aq_rebuild_group_section_from_products(merged, product_groups):



        applied = True



    if _aq_apply_loan_loss_allowance(merged, settlement_store):



        applied = True



    elif _aq_apply_business_ending_allowance(merged, business_data):



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



    {"row": 36, "label": "31일 이상 충당금", "key": "__allowance_to_31plus_ratio__", "kind": "pct"},



]

AQ_SECTION2_EXPORT_ROWS = [
    {"row": 19, "label": "\ubb34\uc5f0\uccb4", "key": "\ubb34\uc5f0\uccb4", "kind": "pct"},
    {"row": 20, "label": "1~30", "key": "1~30", "kind": "pct"},
    {"row": 21, "label": "1~10", "key": "1~10", "kind": "pct"},
    {"row": 22, "label": "11~30", "key": "11~30", "kind": "pct"},
    {"row": 23, "label": "31~60", "key": "31~60", "kind": "pct"},
    {"row": 24, "label": "61~90", "key": "61~90", "kind": "pct"},
    {"row": 25, "label": "91~120", "key": "91~120", "kind": "pct"},
    {"row": 26, "label": "121~150", "key": "121~150", "kind": "pct"},
    {"row": 27, "label": "151~180", "key": "151~180", "kind": "pct"},
    {"row": 28, "label": "181~", "key": "181~", "kind": "pct"},
    {"row": 29, "label": "\ud569\uacc4(%)", "key": "__total_pct__", "kind": "pct"},
    {"row": 30, "label": "1\uc77c\uc774\uc0c1 \uc5f0\uccb4\uc728", "key": "\uc5f0\uccb4\uc728(%)", "kind": "pct"},
    {"row": 31, "label": "\uc5f0\uccb4\ud569\uacc4", "key": "\uc5f0\uccb4\ud569\uacc4", "kind": "amount"},
    {"row": 32, "label": "", "key": None, "kind": "blank"},
    {"row": 33, "label": "31\uc77c\uc774\uc0c1 \uc5f0\uccb4\uc728", "key": "30\uc77c\uc774\uc0c1\uc5f0\uccb4\uc728", "kind": "pct"},
    {"row": 34, "label": "31\uc77c\uc774\uc0c1 \uc5f0\uccb4\ud569\uacc4", "key": "31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4", "kind": "amount"},
    {"row": 35, "label": "\ub300\uc190\ucda9\ub2f9\uae08", "key": "\ub300\uc190\ucda9\ub2f9\uae08", "kind": "amount"},
    {"row": 36, "label": "31\uc77c \uc774\uc0c1 \ucda9\ub2f9\uae08", "key": "__allowance_to_31plus_ratio__", "kind": "pct"},
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



    updated = False



    for row in range(1, min(ws.max_row, 180) + 1):



        for col in range(1, min(ws.max_column, 30) + 1):



            cell = ws.cell(row=row, column=col)



            if isinstance(cell.value, str) and "\ub2e8\uc704" in cell.value:



                cell.value = unit_pattern.sub(lambda match: f"{match.group(1)}{label}{match.group(3)}", cell.value)



                updated = True



    return updated











def _aq_scale_amount(value, unit):



    if value is None:



        return None



    if not isinstance(value, (int, float)):



        return value



    factor = AQ_UNIT_OPTIONS[_aq_unit_key(unit)]["factor"]



    return value * factor







AQ_SECTION2_DETAIL_RATE_KEYS = ["\ubb34\uc5f0\uccb4", "1~30", "1~10", "11~30", "31~60", "61~90", "91~120", "121~150", "151~180", "181~"]
AQ_SECTION2_TOTAL_RATE_KEYS = ["\ubb34\uc5f0\uccb4", "1~30", "31~60", "61~90", "91~120", "121~150", "151~180", "181~"]


def _aq_number_at(values, index):
    if isinstance(values, list) and index < len(values):
        value = values[index]
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return value
    return None


def _aq_section1_ratio_value(data, amount_key, index):
    section1 = data.get("section1") or {}
    total = _aq_number_at(section1.get("\uc735\uc794\ud569\uacc4"), index)
    amount = _aq_number_at(section1.get(amount_key), index)
    if not total or amount is None:
        return None
    return amount / total


def _aq_section2_direct_rate(section2, key, index):
    direct = _aq_number_at((section2 or {}).get(key), index)
    if direct is not None:
        return direct
    if key == "\uc5f0\uccb4\uc728(%)":
        for alt_key in ("1\uc77c\uc774\uc0c1 \uc5f0\uccb4\uc728", "1\uc77c\uc774\uc0c1\uc5f0\uccb4\uc728"):
            alt = _aq_number_at((section2 or {}).get(alt_key), index)
            if alt is not None:
                return alt
    if key == "30\uc77c\uc774\uc0c1\uc5f0\uccb4\uc728":
        for alt_key in ("31\uc77c\uc774\uc0c1 \uc5f0\uccb4\uc728", "31\uc77c\uc774\uc0c1\uc5f0\uccb4\uc728", "30\uc77c\uc774\uc0c1 \uc5f0\uccb4\uc728"):
            alt = _aq_number_at((section2 or {}).get(alt_key), index)
            if alt is not None:
                return alt
    return None


def _aq_section2_rate_value(data, key, index):
    section2 = data.get("section2") or {}
    direct = _aq_section2_direct_rate(section2, key, index)
    if direct is not None:
        return direct
    if key in AQ_SECTION2_DETAIL_RATE_KEYS:
        return _aq_section1_ratio_value(data, key, index)
    if key == "\uc5f0\uccb4\uc728(%)":
        return _aq_section1_ratio_value(data, "\uc5f0\uccb4\ud569\uacc4", index)
    if key == "30\uc77c\uc774\uc0c1\uc5f0\uccb4\uc728":
        total = _aq_number_at((data.get("section1") or {}).get("\uc735\uc794\ud569\uacc4"), index)
        overdue31 = _aq_number_at(_aq_31plus_overdue_values(data), index)
        if not total or overdue31 is None:
            return None
        return overdue31 / total
    return None


def _aq_section2_rate_values(data, key):
    return [_aq_section2_rate_value(data, key, idx) for idx, _ in enumerate(data.get("periods") or [])]


def _aq_section2_amount_values(data, key):
    section1 = data.get("section1") or {}
    section2 = data.get("section2") or {}
    periods = data.get("periods") or []
    if key == "\uc5f0\uccb4\ud569\uacc4":
        result = []
        for idx, _ in enumerate(periods):
            direct = _aq_number_at(section2.get(key), idx)
            result.append(direct if direct is not None else _aq_number_at(section1.get(key), idx))
        return result
    if key == "31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4":
        return _aq_31plus_overdue_values(data)
    values = section2.get(key)
    return values if isinstance(values, list) else None


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



















def _aq_section2_total_pct_values(data):
    result = []
    for idx, _ in enumerate(data.get("periods") or []):
        rate_sum = 0
        has_value = False
        for key in AQ_SECTION2_TOTAL_RATE_KEYS:
            value = _aq_section2_rate_value(data, key, idx)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                rate_sum += value
                has_value = True
        result.append(rate_sum if has_value else None)
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











AQ_MODERN_SECTION1_ROWS = {
    5: "\ubb34\uc5f0\uccb4",
    6: "1~30",
    7: "1~10",
    8: "11~30",
    9: "31~60",
    10: "61~90",
    11: "91~120",
    12: "121~150",
    13: "151~180",
    14: "181~",
    15: "\uc5f0\uccb4\ud569\uacc4",
    16: "31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4",
    17: "\uc735\uc794\ud569\uacc4",
}


AQ_MODERN_SECTION2_EXPORT_ROWS = [
    {"row": 21, "label": "\ubb34\uc5f0\uccb4", "key": "\ubb34\uc5f0\uccb4", "kind": "pct"},
    {"row": 22, "label": "1~30", "key": "1~30", "kind": "pct"},
    {"row": 23, "label": "1~10", "key": "1~10", "kind": "pct"},
    {"row": 24, "label": "11~30", "key": "11~30", "kind": "pct"},
    {"row": 25, "label": "31~60", "key": "31~60", "kind": "pct"},
    {"row": 26, "label": "61~90", "key": "61~90", "kind": "pct"},
    {"row": 27, "label": "91~120", "key": "91~120", "kind": "pct"},
    {"row": 28, "label": "121~150", "key": "121~150", "kind": "pct"},
    {"row": 29, "label": "151~180", "key": "151~180", "kind": "pct"},
    {"row": 30, "label": "181~", "key": "181~", "kind": "pct"},
    {"row": 31, "label": "\ud569\uacc4(%)", "key": "__total_pct__", "kind": "pct"},
    {"row": 32, "label": "1\uc77c\uc774\uc0c1 \uc5f0\uccb4\uc728", "key": "\uc5f0\uccb4\uc728(%)", "kind": "pct"},
    {"row": 33, "label": "", "key": None, "kind": "blank"},
    {"row": 34, "label": "31\uc77c\uc774\uc0c1 \uc5f0\uccb4\uc728", "key": "30\uc77c\uc774\uc0c1\uc5f0\uccb4\uc728", "kind": "pct"},
    {"row": 35, "label": "31\uc77c\uc774\uc0c1 \uc5f0\uccb4\ud569\uacc4", "key": "31\uc77c\uc774\uc0c1\uc5f0\uccb4\ud569\uacc4", "kind": "amount"},
    {"row": 36, "label": "\ub300\uc190\ucda9\ub2f9\uae08", "key": "\ub300\uc190\ucda9\ub2f9\uae08", "kind": "amount"},
    {"row": 37, "label": "\uc5f0\uccb4\ub300\ube44 \ucda9\ub2f9\ube44\uc728", "key": "__allowance_to_31plus_ratio__", "kind": "pct"},
]


AQ_MODERN_GROUP_ROWS = {
    "\uc2e0\uc6a9": {
        43: "\ubb34\uc5f0\uccb4",
        44: "1~30",
        45: "31~60",
        46: "61~90",
        47: "91~180",
        48: "181~",
        49: "\uc5f0\uccb4\ud569\uacc4",
        50: "\uc735\uc794\ud569\uacc4",
        51: "\uc5f0\uccb4\uc728(%)",
    },
    "\ub2f4\ubcf4": {
        55: "\ubb34\uc5f0\uccb4",
        56: "1~30",
        57: "31~60",
        58: "61~90",
        59: "91~180",
        60: "181~",
        61: "\uc5f0\uccb4\ud569\uacc4",
        62: "\uc735\uc794\ud569\uacc4",
        63: "\uc5f0\uccb4\uc728(%)",
    },
    "\ubcf4\uc99d": {
        67: "\ubb34\uc5f0\uccb4",
        68: "1~30",
        69: "31~60",
        70: "61~90",
        71: "91~180",
        72: "181~",
        73: "\uc5f0\uccb4\ud569\uacc4",
        74: "\uc735\uc794\ud569\uacc4",
        75: "\uc5f0\uccb4\uc728(%)",
    },
}


AQ_MODERN_PRODUCT_TEMPLATE_ROW = 80
AQ_MODERN_PRODUCT_BLOCK_HEIGHT = 12
AQ_MODERN_PRODUCT_BLOCK_GAP = 1
AQ_EXPORT_TEMPLATE_VERSION = "asset-quality-template-20260729-gap"
AQ_MODERN_PRODUCT_ROW_OFFSETS = {
    1: "\ubb34\uc5f0\uccb4",
    2: "1~30",
    3: "31~60",
    4: "61~90",
    5: "91~120",
    6: "121~150",
    7: "151~180",
    8: "181~",
    9: "\uc5f0\uccb4\ud569\uacc4",
    10: "\uc735\uc794\ud569\uacc4",
    11: "\uc5f0\uccb4\uc728(%)",
}


def _aq_default_asset_quality_template_path():
    path = os.path.join(os.path.dirname(__file__), "templates", "asset_quality_template.xlsx")
    return path if os.path.exists(path) else None


def _aq_is_modern_asset_quality_template(ws):
    row3 = str(ws.cell(row=3, column=1).value or "")
    row78 = str(ws.cell(row=78, column=1).value or "")
    return "\uc804\uccb4 \ub300\ucd9c\ucc44\uad8c" in row3 and "\uc0c1\ud488\ubcc4 \ub300\ucd9c\ucc44\uad8c" in row78


def _aq_period_like(value):
    return bool(re.match(r"^\d{4}\.\d{2}$", _aq_norm_period(value)))


def _aq_copy_column_style(ws, source_col, target_col):
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


def _aq_prepare_period_columns(ws, periods):
    needed_max_col = max(2, len(periods) + 1)
    while ws.max_column < needed_max_col:
        source_col = max(2, ws.max_column)
        target_col = ws.max_column + 1
        _aq_copy_column_style(ws, source_col, target_col)
    return {period: idx + 2 for idx, period in enumerate(periods)}


def _aq_write_period_headers(ws, periods, header_rows):
    max_header_col = max(ws.max_column, len(periods) + 1)
    for row in header_rows:
        if not row:
            continue
        for idx, period in enumerate(periods, start=2):
            ws.cell(row=row, column=idx).value = period
        for col in range(len(periods) + 2, max_header_col + 1):
            if _aq_period_like(ws.cell(row=row, column=col).value):
                ws.cell(row=row, column=col).value = None


def _aq_visible_product_names(data):
    section4 = data.get("section4") or {}
    names = []
    for product in data.get("products") or []:
        product_name = str(product or "").strip()
        if product_name and product_name in section4 and product_name not in names:
            names.append(product_name)
    for product in section4.keys():
        product_name = str(product or "").strip()
        if product_name and product_name not in names:
            names.append(product_name)
    return names


def _aq_prepare_modern_product_blocks(ws, data):
    product_names = _aq_visible_product_names(data)
    base = AQ_MODERN_PRODUCT_TEMPLATE_ROW
    block_height = AQ_MODERN_PRODUCT_BLOCK_HEIGHT
    block_stride = block_height + AQ_MODERN_PRODUCT_BLOCK_GAP
    if len(product_names) > 1:
        ws.insert_rows(base + block_height, amount=(len(product_names) - 1) * block_stride)
    product_rows = {}
    if not product_names:
        for offset in range(block_height):
            _aq_copy_row_style(ws, base + offset, base + offset)
        return product_rows, product_names
    for index, product in enumerate(product_names):
        target_base = base + index * block_stride
        for offset in range(block_height):
            _aq_copy_row_style(ws, base + offset, target_base + offset)
        ws.cell(row=target_base, column=1).value = product
        for offset, label in AQ_MODERN_PRODUCT_ROW_OFFSETS.items():
            display_label = "\uc5f0\uccb4\uc728" if label == "\uc5f0\uccb4\uc728(%)" else label
            ws.cell(row=target_base + offset, column=1).value = display_label
        product_rows[product] = target_base
    return product_rows, product_names


def _aq_product_rate_values(product_data, period_count):
    values = []
    for index in range(period_count):
        overdue, has_overdue = _aq_product_overdue_total(product_data, index)
        balance, has_balance = _aq_product_balance_total(product_data, index)
        if has_overdue and has_balance and balance:
            values.append((overdue / balance) * 100)
        else:
            values.append(None)
    return values


def _aq_rate_output_value(value, cell):
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return value
    if "%" in str(cell.number_format or "") and abs(value) > 1:
        return value / 100
    return value


def _aq_write_rate_row(ws, row, periods, period_columns, values):
    if not isinstance(values, list):
        return
    for idx, value in enumerate(values):
        if idx >= len(periods):
            break
        col = period_columns.get(_aq_norm_period(periods[idx]))
        if col and value is not None:
            cell = ws.cell(row=row, column=col)
            cell.value = _aq_rate_output_value(value, cell)
            if "%" in str(cell.number_format or ""):
                cell.number_format = AQ_PERCENT_NUMBER_FORMAT


def _write_modern_asset_quality_sheet(ws, data, unit="million"):
    unit = _aq_unit_key(unit)
    periods = [_aq_norm_period(p) for p in (data.get("periods") or [])]
    period_columns = _aq_prepare_period_columns(ws, periods)
    product_rows, product_names = _aq_prepare_modern_product_blocks(ws, data)

    _aq_update_unit_labels(ws, unit)
    ws.cell(row=2, column=14).value = f"(\ub2e8\uc704: {_aq_unit_label(unit)})"
    _aq_write_period_headers(ws, periods, [4, 20, 42, 54, 66] + list(product_rows.values()))

    section1 = data.get("section1") or {}
    for row, label in AQ_MODERN_SECTION1_ROWS.items():
        _aq_write_series_row(ws, row, periods, period_columns, section1.get(label), unit=unit, number_format=AQ_AMOUNT_NUMBER_FORMAT)

    section2 = data.get("section2") or {}
    section2_total_pct = _aq_section2_total_pct_values(data)
    section2_allowance_to_31plus_ratio = _aq_section2_allowance_to_31plus_ratio_values(data)
    period_cols = list(period_columns.values())
    for item in AQ_MODERN_SECTION2_EXPORT_ROWS:
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
        elif kind == "pct":
            values = _aq_section2_rate_values(data, key)
        elif kind == "amount":
            values = _aq_section2_amount_values(data, key)
        else:
            values = section2.get(key)
        if kind == "pct":
            _aq_write_rate_row(ws, row, periods, period_columns, values)
        else:
            _aq_write_series_row(ws, row, periods, period_columns, values, unit=unit, scale_amount=True, number_format=AQ_AMOUNT_NUMBER_FORMAT)

    section3 = data.get("section3") or {}
    for group, rows in AQ_MODERN_GROUP_ROWS.items():
        group_data = section3.get(group) or {}
        for row, label in rows.items():
            if label == "\uc5f0\uccb4\uc728(%)":
                _aq_write_rate_row(ws, row, periods, period_columns, group_data.get(label))
            else:
                _aq_write_series_row(ws, row, periods, period_columns, group_data.get(label), unit=unit, scale_amount=True, number_format=AQ_AMOUNT_NUMBER_FORMAT)

    section4 = data.get("section4") or {}
    for product in product_names:
        base_row = product_rows.get(product)
        product_data = section4.get(product) or {}
        if not base_row:
            continue
        for offset, label in AQ_MODERN_PRODUCT_ROW_OFFSETS.items():
            row = base_row + offset
            if label == "\uc5f0\uccb4\uc728(%)":
                _aq_write_rate_row(ws, row, periods, period_columns, _aq_product_rate_values(product_data, len(periods)))
            else:
                _aq_write_series_row(ws, row, periods, period_columns, product_data.get(label), unit=unit, scale_amount=True, number_format=AQ_AMOUNT_NUMBER_FORMAT)


def _write_asset_quality_to_sheet(ws, data, unit="million"):



    unit = _aq_unit_key(unit)



    if _aq_is_modern_asset_quality_template(ws):



        _write_modern_asset_quality_sheet(ws, data, unit)



        return



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



        elif kind == "pct":



            values = _aq_section2_rate_values(data, key)



        elif kind == "amount":



            values = _aq_section2_amount_values(data, key)



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



        ws, 17, periods, period_columns, _aq_31plus_overdue_values(data),



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







def _create_asset_quality_fallback_workbook(data, unit="million"):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "\uc790\uc0b0\uac74\uc804\uc131"

    periods = [_aq_norm_period(p) for p in (data.get("periods") or [])]
    max_col = max(2, len(periods) + 1)
    max_row = max(
        120,
        max(SECTION1_ROWS.keys(), default=1),
        max((item["row"] for item in AQ_SECTION2_EXPORT_ROWS), default=1),
    )

    ws.cell(row=1, column=1).value = "\u25a0 \uc790\uc0b0\uac74\uc804\uc131"
    ws.cell(row=2, column=1).value = f"(\ub2e8\uc704: {_aq_unit_label(unit)})"
    ws.cell(row=3, column=1).value = "\uad6c\ubd84"
    for idx, period in enumerate(periods, start=2):
        ws.cell(row=3, column=idx).value = period

    for row, label in SECTION1_ROWS.items():
        ws.cell(row=row, column=1).value = label
    for item in AQ_SECTION2_EXPORT_ROWS:
        if item.get("kind") != "blank":
            ws.cell(row=item["row"], column=1).value = item.get("label") or item.get("key")

    for group_info in SECTION3_GROUPS.values():
        rows = group_info.get("rows", group_info) if isinstance(group_info, dict) else {}
        for row, label in rows.items():
            ws.cell(row=row, column=1).value = label

    header_fill = PatternFill("solid", fgColor="1F4E78")
    label_fill = PatternFill("solid", fgColor="DDEBF7")
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(horizontal="center" if col > 1 else "left", vertical="center")
            if row == 3:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
            elif row in {1, 2}:
                cell.font = Font(bold=True)
            elif col == 1:
                cell.fill = label_fill

    ws.column_dimensions["A"].width = 24
    for col in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.freeze_panes = "B4"
    return wb, ws


@app.get("/api/data/bs_is")



async def get_bs_is():



    """BS/IS 데이터"""



    return uploaded_data.get("bs_is", {})











PAYMENT_CASHFLOW_PRINCIPAL_KEY = "\uc6d0\uae08\ud68c\uc218\uc561"
PAYMENT_CASHFLOW_INTEREST_KEY = "\uc774\uc790\ud68c\uc218\uc561"
CF_LOAN_PRINCIPAL_COLLECTION_LABEL = "\ub300\ucd9c\uae08 \uc6d0\uae08\ud68c\uc218"
CF_LOAN_INTEREST_INCOME_LABEL = "\ub300\ucd9c\uae08 \uc774\uc790\uc218\uc785"
CF_LOAN_REPAYMENT_PARENT_LABEL = "1. \ub300\ucd9c\uae08 \uc6d0\ub9ac\uae08\uc0c1\ud658"


def _cf_number(value, default=0.0):
    if value is None or value == "":
        return default
    try:
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if not value:
                return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _cf_period_sort_key(period: str):
    text = str(period or "")
    match = re.match(r"^(\d{4})-(\d{2})$", text)
    if match:
        return (int(match.group(1)), int(match.group(2)), text)
    return (9999, 99, text)


def _cf_get_items_map(data):
    items = data.get("items") if isinstance(data, dict) else None
    if isinstance(items, dict):
        return copy.deepcopy(items)
    if isinstance(items, list):
        result = {}
        for row in items:
            if isinstance(row, dict) and row.get("label"):
                result[str(row.get("label"))] = copy.deepcopy(row.get("values") or {})
        return result
    return {}


def _cf_update_cashflow_row(rows, label, values):
    if not isinstance(rows, list):
        return rows
    found = False
    labels = {
        str(row.get("label") or "").strip()
        for row in rows
        if isinstance(row, dict)
    }
    result = []
    for row in rows:
        if not isinstance(row, dict):
            result.append(row)
            continue
        current = copy.deepcopy(row)
        current_label = str(current.get("label") or "").strip()
        if current_label == label:
            current_values = current.get("values")
            if not isinstance(current_values, dict):
                current_values = {}
            current_values.update(values)
            current["values"] = current_values
            found = True
        result.append(current)
        if current_label == CF_LOAN_REPAYMENT_PARENT_LABEL:
            for missing_label in (CF_LOAN_PRINCIPAL_COLLECTION_LABEL, CF_LOAN_INTEREST_INCOME_LABEL):
                if missing_label != label and missing_label in labels:
                    continue
                if missing_label == label and found:
                    continue
                insert_values = copy.deepcopy(values) if missing_label == label else {}
                result.append({"label": missing_label, "type": "detail", "values": insert_values})
                labels.add(missing_label)
                if missing_label == label:
                    found = True
    if not found:
        result.append({"label": label, "type": "detail", "values": copy.deepcopy(values)})
    return result


def _cashflow_with_payment_data(cashflow_data, payment_data):
    result = copy.deepcopy(cashflow_data) if isinstance(cashflow_data, dict) else {}
    if not isinstance(payment_data, dict):
        return result
    section1 = payment_data.get("section1") or {}
    if not isinstance(section1, dict):
        return result

    source_pairs = [
        (CF_LOAN_PRINCIPAL_COLLECTION_LABEL, section1.get(PAYMENT_CASHFLOW_PRINCIPAL_KEY) or {}),
        (CF_LOAN_INTEREST_INCOME_LABEL, section1.get(PAYMENT_CASHFLOW_INTEREST_KEY) or {}),
    ]
    period_values = {}
    for label, values in source_pairs:
        if not isinstance(values, dict):
            continue
        converted = {
            str(period): int(round(_cf_number(amount) * 1_000_000))
            for period, amount in values.items()
            if period
        }
        if converted:
            period_values[label] = converted

    if not period_values:
        return result

    dates = copy.deepcopy(result.get("dates") or {})
    if not isinstance(dates, dict):
        dates = {}
    existing_dates = {str(value) for value in dates.values()}
    all_periods = sorted(
        {period for values in period_values.values() for period in values.keys()},
        key=_cf_period_sort_key,
    )
    for period in all_periods:
        if period not in existing_dates:
            dates[f"payment_{str(period).replace('-', '_')}"] = period
            existing_dates.add(period)
    result["dates"] = dict(sorted(dates.items(), key=lambda item: _cf_period_sort_key(str(item[1]))))

    items = _cf_get_items_map(result)
    for label, values in period_values.items():
        current = items.get(label)
        if not isinstance(current, dict):
            current = {}
        current.update(values)
        items[label] = current
    result["items"] = items

    if isinstance(result.get("rows"), list):
        rows = result.get("rows")
        for label, values in period_values.items():
            rows = _cf_update_cashflow_row(rows, label, values)
        result["rows"] = rows

    return result


@app.get("/api/data/cashflow")



async def get_cashflow():



    """현금흐름 데이터"""



    return _cashflow_with_payment_data(
        uploaded_data.get("cashflow", {}),
        uploaded_data.get("payment_data", {}),
    )











@app.get("/api/data/asset_quality")



async def get_asset_quality():



    """자산건전성 데이터"""



    return uploaded_data.get("asset_quality", {})











@app.post("/api/upload/asset-quality")



async def upload_asset_quality(file: UploadFile = File(...), period: str = Form(default="")):



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
        data = _prune_indexed_period_data(data)
        _aq_apply_product_overdue_totals(data)



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
    if _asset_data_is_cleared_marker(d) or not isinstance(d, dict):
        d = None
    saq = _normalize_settlement_aq_store(uploaded_data.get("settlement_aq") or {})



    if not d and not saq:



        return JSONResponse({"error": "\ub370\uc774\ud130 \uc5c6\uc74c"}, status_code=404)


    if not d:
        d = {
            "periods": [],
            "section1": {},
            "section2": {},
            "section3": {},
            "section4": {},
            "products": [],
        }


    merged = _merge_settlement_into_asset_quality_detail(
        d,
        saq,
        uploaded_data.get("business_detail") or {},
        uploaded_data.get("product_groups") or [],
    )



    return JSONResponse(_prune_indexed_period_data(merged))











@app.get("/api/export/asset-quality-detail")



async def export_asset_quality_detail(unit: str = "million"):



    """Export current asset-quality data using the uploaded workbook template."""



    data = uploaded_data.get("asset_quality_detail")



    if not data:

        if not uploaded_data.get("settlement_aq"):



            raise HTTPException(status_code=404, detail="\uc790\uc0b0\uac74\uc804\uc131 \ub370\uc774\ud130\uac00 \uc5c6\uc2b5\ub2c8\ub2e4. \uc790\uc0b0\uac74\uc804\uc131 \uc5d1\uc140\uc744 \uba3c\uc800 \uc5c5\ub85c\ub4dc\ud558\uc138\uc694.")



        data = {
            "periods": [],
            "section1": {},
            "section2": {},
            "section3": {},
            "section4": {},
            "products": [],
        }







    template_path = _aq_default_asset_quality_template_path() or _stored_file_path("asset_quality_template.xlsx")



    if not template_path:



        template_path = None







    try:



        meta = uploaded_data.get("asset_quality_template", {}) or {}



        merged = _merge_settlement_into_asset_quality_detail(
            data,
            uploaded_data.get("settlement_aq") or {},
            uploaded_data.get("business_detail") or {},
            uploaded_data.get("product_groups") or [],
        )
        merged = _prune_indexed_period_data(merged)



        unit = _aq_unit_key(unit)



        wb = None



        ws = None



        if template_path:



            try:



                wb = openpyxl.load_workbook(template_path, data_only=True, keep_links=False)



                ws = _asset_quality_sheet(wb, meta.get("sheet_name"))



            except Exception:



                wb = None



                ws = None



        if wb is None or ws is None:



            wb, ws = _create_asset_quality_fallback_workbook(merged, unit=unit)



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



async def upload_settlement_aq(file: UploadFile = File(...), period: str = Form(default=""), confirm_period_mismatch: str = Form(default="")):



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

        selected_period = normalize_period_key(period)
        parsed_period = normalize_period_key(data.get("period"))
        confirmed_mismatch = str(confirm_period_mismatch).strip().lower() in {"1", "true", "yes", "y"}
        period_mismatch = bool(selected_period and parsed_period and selected_period != parsed_period)
        if selected_period:
            data["selected_period"] = selected_period
            data["upload_period"] = selected_period
        if period_mismatch and not confirmed_mismatch:
            os.unlink(tmp_path)
            return JSONResponse({
                "success": False,
                "code": "PERIOD_MISMATCH",
                "selected_period": selected_period,
                "parsed_periods": [parsed_period],
                "message": "Selected upload period differs from settlement period in the file.",
            }, status_code=409)
        if period_mismatch:
            data["period_mismatch_confirmed"] = True
        pkey = parsed_period or selected_period
        if not pkey:
            os.unlink(tmp_path)
            return JSONResponse({
                "success": False,
                "code": "PERIOD_REQUIRED",
                "message": "Upload period is required when the file does not contain a period.",
            }, status_code=400)
        data["period"] = pkey



        os.unlink(tmp_path)







        # 기존 asset_quality_detail 에 결산자료 기반 데이터를 병합



        # period 기준으로 단일 기간 데이터로 저장



        existing = uploaded_data.get("asset_quality_detail") or {}



        # settlement_aq 키에 별도 저장 (기간별 여러 건 누적 가능)



        saq_store = uploaded_data.get("settlement_aq") or {}



        pkey = data["period"]



        saq_store[pkey] = data
        saq_store = _prune_period_store(saq_store)



        uploaded_data["settlement_aq"] = saq_store







        # ── 대출만기 → asset_data.sections['대출만기'] 병합 ─────────



        # 결산자료의 maturity 집계값을 대출자산속성 '대출만기' 섹션에 반영



        # pkey: 'YYYY-MM' 형태 (대출자산속성 기간 키와 동일)



        # ※ 대출자산속성 기존 데이터는 원 단위, maturity는 백만원 단위



        #   → asset_data에 저장 시 백만원 × 1,000,000 = 원 단위로 변환



        maturity = data.get("maturity", {})



        if False and maturity and pkey and pkey != "unknown":  # derived on asset GET



            asset_data = _ensure_asset_data_shell([pkey])



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



        if False and avg_rate and pkey and pkey != "unknown":  # derived on asset GET



            asset_data = _ensure_asset_data_shell([pkey])



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



        if False and repayment and pkey and pkey != "unknown":  # derived on asset GET



            asset_data = _ensure_asset_data_shell([pkey])



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



        if False and amount_band and pkey and pkey != "unknown":  # derived on asset GET



            asset_data = _ensure_asset_data_shell([pkey])



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

        if False and channel_amount and pkey and pkey != "unknown":  # derived on asset GET

            asset_data = _ensure_asset_data_shell([pkey])

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


        cb_by_product_data = data.get("cb_by_product", {})

        if cb_by_product_data and pkey in saq_store:

            saq_store[pkey]["cb_by_product"] = {
                str(product): {str(score): value for score, value in (scores or {}).items()}
                for product, scores in cb_by_product_data.items()
            }

            uploaded_data["settlement_aq"] = saq_store







        return JSONResponse({



            "success": True,



            "message": f"결산자료 '{file.filename}' 업로드 완료",



            "period": data["period"],

            "selected_period": selected_period,

            "period_mismatch_confirmed": bool(data.get("period_mismatch_confirmed")),



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
            "cb_by_product_updated": bool(cb_by_product_data and pkey and pkey != "unknown"),



        })



    except Exception as e:



        import traceback



        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")











@app.get("/api/data/settlement-aq")



async def get_settlement_aq(period: str = ""):



    """결산자료 기반 자산건전성 데이터 반환



    period 지정 시 해당 기간, 미지정 시 최신 기간



    """



    store = _normalize_settlement_aq_store(
        _prune_period_store(uploaded_data.get("settlement_aq") or {})
    )



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



    store = _prune_period_store(uploaded_data.get("settlement_aq") or {})



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



    saq = _normalize_settlement_aq_store(uploaded_data.get("settlement_aq") or {})



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


def _sanitize_daily_close_collateral_ltv(value):
    if not isinstance(value, dict):
        value = {}

    def _clean_list(items):
        if not isinstance(items, list):
            items = []
        result = []
        seen = set()
        for item in items:
            text = str(item or "").strip()
            if text and text not in seen:
                result.append(text)
                seen.add(text)
        return result

    def _clean_rule(item, idx):
        if not isinstance(item, dict):
            item = {}
        raw_id = str(item.get("id") or f"ltv-{idx + 1}").strip()
        safe_id = re.sub(r"[^a-zA-Z0-9_-]", "", raw_id) or f"ltv-{idx + 1}"
        return {
            "id": safe_id,
            "name": str(item.get("name") or f"기준 {idx + 1}").strip(),
            "ltvMax": str(item.get("ltvMax") or "").strip(),
            "collateralTypes": _clean_list(item.get("collateralTypes")),
            "collateralDivisions": _clean_list(item.get("collateralDivisions")),
        }

    raw_groups = value.get("regionGroups") or []
    if not isinstance(raw_groups, list):
        raw_groups = []
    if not raw_groups:
        criteria = value.get("criteria") or []
        if not isinstance(criteria, list):
            criteria = []
        raw_groups = [
            {
                "id": item.get("id") if isinstance(item, dict) else f"ltv-migrated-{idx + 1}",
                "name": item.get("name") if isinstance(item, dict) else f"기준 {idx + 1}",
                "regions": item.get("regions") if isinstance(item, dict) else [],
                "criteria": [item] if isinstance(item, dict) else [],
            }
            for idx, item in enumerate(criteria)
        ]

    legacy_rules = []
    raw_criteria = value.get("criteria") or []
    if isinstance(raw_criteria, list):
        legacy_rules.extend([item for item in raw_criteria if isinstance(item, dict)])
    for group in raw_groups:
        if isinstance(group, dict) and isinstance(group.get("criteria"), list):
            legacy_rules.extend([item for item in group.get("criteria") if isinstance(item, dict)])

    common_source = value.get("common") if isinstance(value.get("common"), dict) else {}
    legacy_common = {}
    for item in legacy_rules:
        has_legacy_value = (
            str(item.get("ltvMax") or "").strip()
            or str(item.get("loanMinMan") or "").strip()
            or (isinstance(item.get("appraisers"), list) and item.get("appraisers"))
        )
        if has_legacy_value:
            legacy_common = item
            break
    common = {
        "ltvMax": str(common_source.get("ltvMax") or legacy_common.get("ltvMax") or "").strip(),
        "loanMinMan": str(common_source.get("loanMinMan") or legacy_common.get("loanMinMan") or "").strip(),
        "appraisers": _clean_list(common_source.get("appraisers") if isinstance(common_source.get("appraisers"), list) else legacy_common.get("appraisers")),
    }
    category_groups = _clean_list(value.get("categoryGroups"))

    cleaned_groups = []
    used_regions = set()
    for idx, group in enumerate(raw_groups):
        if not isinstance(group, dict):
            continue
        raw_id = str(group.get("id") or f"ltv-group-{idx + 1}").strip()
        safe_id = re.sub(r"[^a-zA-Z0-9_-]", "", raw_id) or f"ltv-group-{idx + 1}"
        rules = group.get("criteria") or []
        if not isinstance(rules, list):
            rules = []
        cleaned_regions = []
        for region in _clean_list(group.get("regions")):
            if region in used_regions:
                continue
            used_regions.add(region)
            cleaned_regions.append(region)
        cleaned_rules = []
        used_types = set()
        for rule_idx, rule in enumerate(rules):
            cleaned_rule = _clean_rule(rule, rule_idx)
            unique_types = []
            for collateral_type in cleaned_rule.get("collateralTypes") or []:
                if collateral_type in used_types:
                    continue
                used_types.add(collateral_type)
                unique_types.append(collateral_type)
            cleaned_rule["collateralTypes"] = unique_types
            cleaned_rules.append(cleaned_rule)
        cleaned_groups.append({
            "id": safe_id,
            "ltvMax": str(group.get("ltvMax") or "").strip(),
            "name": str(group.get("name") or f"지역그룹 {idx + 1}").strip(),
            "regions": cleaned_regions,
            "criteria": cleaned_rules,
        })
    return {"categoryGroups": category_groups, "common": common, "regionGroups": cleaned_groups}


def _sanitize_daily_close_joy_eligible(value):
    if not isinstance(value, dict):
        value = {}
    base = _sanitize_daily_close_collateral_ltv(value)

    def _clean_list(items):
        if not isinstance(items, list):
            items = []
        result = []
        seen = set()
        for item in items:
            text = str(item or "").strip()
            if text and text not in seen:
                result.append(text)
                seen.add(text)
        return result

    def _safe_id(value, fallback):
        raw = str(value or fallback).strip()
        return re.sub(r"[^a-zA-Z0-9_-]", "", raw) or fallback

    def _clean_rule(item, idx):
        if not isinstance(item, dict):
            item = {}
        return {
            "id": _safe_id(item.get("id"), f"joy-rule-{idx + 1}"),
            "name": str(item.get("name") or f"기준 {idx + 1}").strip(),
            "ltvMax": str(item.get("ltvMax") or "").strip(),
            "collateralTypes": _clean_list(item.get("collateralTypes")),
            "collateralDivisions": [],
        }

    def _clean_region_group(item, idx, used_regions):
        if not isinstance(item, dict):
            item = {}
        regions = []
        for region in _clean_list(item.get("regions")):
            if region in used_regions:
                continue
            used_regions.add(region)
            regions.append(region)
        rules = item.get("criteria") or []
        if not isinstance(rules, list):
            rules = []
        cleaned_rules = []
        used_types = set()
        for rule_idx, rule in enumerate(rules):
            cleaned_rule = _clean_rule(rule, rule_idx)
            unique_types = []
            for collateral_type in cleaned_rule.get("collateralTypes") or []:
                if collateral_type in used_types:
                    continue
                used_types.add(collateral_type)
                unique_types.append(collateral_type)
            cleaned_rule["collateralTypes"] = unique_types
            cleaned_rules.append(cleaned_rule)
        return {
            "id": _safe_id(item.get("id"), f"joy-region-{idx + 1}"),
            "name": str(item.get("name") or f"담보지역 {idx + 1}").strip(),
            "ltvMax": str(item.get("ltvMax") or "").strip(),
            "regions": regions,
            "criteria": cleaned_rules,
        }

    raw_division_groups = value.get("divisionGroups") or []
    if not isinstance(raw_division_groups, list):
        raw_division_groups = []
    if not raw_division_groups and base.get("regionGroups"):
        raw_division_groups = [{
            "id": "joy-division-migrated",
            "name": "담보구분 미지정",
            "collateralDivisions": [],
            "regionGroups": base.get("regionGroups") or [],
        }]

    cleaned_divisions = []
    used_divisions = set()
    for idx, group in enumerate(raw_division_groups):
        if not isinstance(group, dict):
            continue
        divisions = []
        for division in _clean_list(group.get("collateralDivisions") or group.get("divisions")):
            if division in used_divisions:
                continue
            used_divisions.add(division)
            divisions.append(division)
        raw_region_groups = group.get("regionGroups") or []
        if not isinstance(raw_region_groups, list):
            raw_region_groups = []
        used_regions = set()
        relationships = _clean_list(
            group.get("relationships")
            or group.get("relations")
            or group.get("collateralRelations")
        )
        cleaned_divisions.append({
            "id": _safe_id(group.get("id"), f"joy-division-{idx + 1}"),
            "name": str(group.get("name") or f"담보구분 {idx + 1}").strip(),
            "collateralDivisions": divisions,
            "relationships": relationships,
            "regionGroups": [
                _clean_region_group(region_group, region_idx, used_regions)
                for region_idx, region_group in enumerate(raw_region_groups)
                if isinstance(region_group, dict)
            ],
        })

    return {
        "categoryGroups": base.get("categoryGroups") or [],
        "common": base.get("common") or {"ltvMax": "", "loanMinMan": "", "appraisers": []},
        "regionGroups": base.get("regionGroups") or [],
        "divisionGroups": cleaned_divisions,
    }


def _sanitize_daily_close_config(value):
    if not isinstance(value, dict):
        value = {}
    return {
        "product": _sanitize_daily_close_groups(value.get("product")),
        "category": _sanitize_daily_close_groups(value.get("category")),
        "agent": _sanitize_daily_close_groups(value.get("agent")),
        "excellent": _sanitize_daily_close_excellent(value.get("excellent")),
        "collateralLtv": _sanitize_daily_close_collateral_ltv(value.get("collateralLtv")),
        "joyEligible": _sanitize_daily_close_joy_eligible(value.get("joyEligible")),
        "receivableCollateral": value.get("receivableCollateral") if isinstance(value.get("receivableCollateral"), dict) else {},
        "receivableJoyPledge": _sanitize_daily_close_joy_eligible(value.get("receivableJoyPledge")),
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


ACCOUNT_PERMISSIONS_KEY = "account_permissions"


def _sanitize_account_permissions(value):
    if not isinstance(value, dict):
        value = {}
    accounts = value.get("accounts")
    if not isinstance(accounts, list):
        accounts = []

    cleaned_accounts = []
    seen_ids = set()
    for idx, account in enumerate(accounts):
        if not isinstance(account, dict):
            continue
        account_id = str(account.get("id") or f"account-{idx + 1}").strip()
        if not account_id or account_id in seen_ids:
            continue
        seen_ids.add(account_id)
        permissions = account.get("permissions")
        if not isinstance(permissions, dict):
            permissions = {}
        clean_permissions = {
            str(page): str(mode)
            for page, mode in permissions.items()
            if str(mode) in {"none", "view", "edit"}
        }
        if clean_permissions.get("page-permission-config") == "edit":
            clean_permissions["page-ip-access"] = "edit"
        cleaned_accounts.append({
            "id": account_id,
            "name": str(account.get("name") or "").strip(),
            "email": str(account.get("email") or "").strip(),
            "loginId": str(account.get("loginId") or "").strip(),
            "passwordHash": str(account.get("passwordHash") or "").strip(),
            "passwordSalt": str(account.get("passwordSalt") or "").strip(),
            "passwordUpdatedAt": str(account.get("passwordUpdatedAt") or "").strip(),
            "authVersion": str(account.get("authVersion") or account.get("passwordUpdatedAt") or "").strip(),
            "permissions": clean_permissions,
        })

    if not any(account.get("id") == "admin" for account in cleaned_accounts):
        cleaned_accounts.insert(0, {
            "id": "admin",
            "name": "관리자",
            "email": "admin",
            "loginId": "admin",
            "passwordHash": "",
            "passwordSalt": "",
            "passwordUpdatedAt": "",
            "authVersion": "",
            "permissions": {},
        })

    return {
        "version": 1,
        "sessionVersion": str(value.get("sessionVersion") or "").strip(),
        "sessionUpdatedAt": str(value.get("sessionUpdatedAt") or "").strip(),
        "accounts": cleaned_accounts,
    }


def _load_account_permissions():
    config = uploaded_data.get(ACCOUNT_PERMISSIONS_KEY)
    if config is None:
        config = load_data(ACCOUNT_PERMISSIONS_KEY) or {}
        uploaded_data[ACCOUNT_PERMISSIONS_KEY] = config
    config = _sanitize_account_permissions(config)
    uploaded_data[ACCOUNT_PERMISSIONS_KEY] = config
    return config


@app.get("/api/account-permissions")
async def get_account_permissions():
    """계정별 사이드바 권한 설정 반환."""
    return JSONResponse(_load_account_permissions())


@app.post("/api/account-permissions")
async def save_account_permissions(request: Request):
    """계정별 사이드바 권한 설정 저장."""
    body = await request.json()
    config = _sanitize_account_permissions(body)
    uploaded_data[ACCOUNT_PERMISSIONS_KEY] = config
    save_data(ACCOUNT_PERMISSIONS_KEY, config)
    return JSONResponse({"success": True, "config": config})


IP_ACCESS_CONFIG_KEY = "ip_access_config"
IP_ACCESS_NO_CACHE_HEADERS = {
    "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
    "Pragma": "no-cache",
}


def _normalize_ip_rule(value):
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        if "/" in text:
            return str(ipaddress.ip_network(text, strict=False))
        return str(ipaddress.ip_address(text))
    except ValueError:
        return ""


def _sanitize_ip_access_config(value):
    if not isinstance(value, dict):
        value = {}
    rows = value.get("entries") if isinstance(value.get("entries"), list) else []
    entries = []
    seen = set()
    for idx, row in enumerate(rows):
        if not isinstance(row, dict):
            continue
        rule = _normalize_ip_rule(row.get("ip") or row.get("address"))
        if not rule or rule in seen:
            continue
        seen.add(rule)
        entries.append({
            "id": str(row.get("id") or f"ip-{idx + 1}").strip()[:80],
            "name": str(row.get("name") or row.get("label") or "").strip()[:80],
            "ip": rule,
            "enabled": bool(row.get("enabled", True)),
        })
    return {
        "version": 1,
        "enabled": bool(value.get("enabled", False)),
        "entries": entries,
        "updatedAt": str(value.get("updatedAt") or "").strip()[:40],
        "updatedBy": str(value.get("updatedBy") or "").strip()[:120],
    }


def _load_ip_access_config():
    config = uploaded_data.get(IP_ACCESS_CONFIG_KEY)
    if config is None:
        config = load_data(IP_ACCESS_CONFIG_KEY) or {}
    config = _sanitize_ip_access_config(config)
    uploaded_data[IP_ACCESS_CONFIG_KEY] = config
    return config


def _request_client_ip(request: Request):
    forwarded = str(request.headers.get("x-forwarded-for") or "")
    candidates = [part.strip() for part in forwarded.split(",") if part.strip()]
    if request.client and request.client.host:
        candidates.append(str(request.client.host).strip())
    for candidate in candidates:
        candidate = candidate.strip("[]")
        try:
            return str(ipaddress.ip_address(candidate))
        except ValueError:
            continue
    return ""


def _ip_access_result(request: Request, config=None):
    config = config or _load_ip_access_config()
    current_ip = _request_client_ip(request)
    try:
        current_address = ipaddress.ip_address(current_ip)
    except ValueError:
        current_address = None
    active_entries = [row for row in config.get("entries", []) if row.get("enabled")]
    matched_rule = ""
    if current_address:
        for row in active_entries:
            try:
                if current_address in ipaddress.ip_network(row.get("ip"), strict=False):
                    matched_rule = str(row.get("ip") or "")
                    break
            except ValueError:
                continue
    local_bypass = bool(current_address and current_address.is_loopback)
    allowed = not config.get("enabled") or local_bypass or bool(matched_rule)
    return {
        "allowed": allowed,
        "enabled": bool(config.get("enabled")),
        "current_ip": current_ip,
        "matched_rule": matched_rule,
        "configured_count": len(active_entries),
        "local_bypass": local_bypass,
    }


@app.get("/api/ip-access-config")
async def get_ip_access_config(request: Request):
    config = _load_ip_access_config()
    return JSONResponse(
        {"config": config, "access": _ip_access_result(request, config)},
        headers=IP_ACCESS_NO_CACHE_HEADERS,
    )


@app.post("/api/ip-access-config")
async def save_ip_access_config(request: Request):
    body = await request.json()
    config = _sanitize_ip_access_config(body)
    active_entries = [row for row in config.get("entries", []) if row.get("enabled")]
    if config.get("enabled") and not active_entries:
        raise HTTPException(status_code=400, detail="IP 제한을 사용하려면 허용 IP를 한 개 이상 등록하세요.")
    access = _ip_access_result(request, config)
    if config.get("enabled") and not access.get("allowed"):
        raise HTTPException(
            status_code=400,
            detail=f"현재 접속 IP({access.get('current_ip') or '확인 불가'})를 허용 목록에 먼저 등록하세요.",
        )
    uploaded_data[IP_ACCESS_CONFIG_KEY] = config
    save_data(IP_ACCESS_CONFIG_KEY, config)
    return JSONResponse(
        {"success": True, "config": config, "access": access},
        headers=IP_ACCESS_NO_CACHE_HEADERS,
    )


@app.get("/api/ip-access/check")
async def check_ip_access(request: Request):
    return JSONResponse(_ip_access_result(request), headers=IP_ACCESS_NO_CACHE_HEADERS)


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



    saq = _normalize_settlement_aq_store(uploaded_data.get("settlement_aq") or {})



    if not saq:



        _saq_file = _data_path("settlement_aq")



        if os.path.exists(_saq_file):



            with open(_saq_file, encoding="utf-8") as f:



                saq = _normalize_settlement_aq_store(json.load(f))







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



    for section3 in (contract_section3, (uploaded_data.get("business_detail") or {}).get("section3") or {}):



        if isinstance(section3, dict):



            for name, value in section3.items():



                if isinstance(value, dict) and any(row["key"] in value for row in BUSINESS_COLLATERAL_SUB_ROWS):



                    label = str(name).strip()



                    if label:



                        names.append(label)



    if not names:



        names = ["\uc2e0\uc6a9", "\ub2f4\ubcf4", "\ubcf4\uc99d"]



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











def _business_detail_group_data(business, group_name):



    section3 = (business or {}).get("section3") or {}



    if not isinstance(section3, dict):



        return None



    if group_name in section3 and isinstance(section3[group_name], dict):



        return section3[group_name]



    target = _business_normalize_group_name(group_name)



    for key, value in section3.items():



        if isinstance(value, dict) and _business_normalize_group_name(key) == target:



            return value



    return None


def _business_collateral_key_aliases(key):
    aliases = []
    text = str(key or "").strip()
    if text:
        aliases.append(text)
    if text == "\ucd94\uac00":
        aliases.append("\ucd94\uac00\ub300\ucd9c")
    elif text == "\ucd94\uac00\ub300\ucd9c":
        aliases.append("\ucd94\uac00")
    return list(dict.fromkeys(aliases))


def _business_nested_value(group, key, periods, idx=None, ym=None):
    if not isinstance(group, dict):
        return None
    for candidate in _business_collateral_key_aliases(key):
        value = _business_value_at(group.get(candidate), periods, idx, ym)
        if value is not None:
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



    contract_value = _business_nested_value(contract_group, sub_key, (uploaded_data.get("contract_data") or {}).get("periods") or [], None, ym)



    if contract_value is not None:



        return contract_value







    business_group = _business_detail_group_data(business, group_name)



    business_value = _business_nested_value(business_group, sub_key, (business or {}).get("periods") or [], idx, ym)



    if business_value is not None:



        return business_value



    prefix = _business_fallback_section3_prefix(group_name)



    for section_name in ("section3", "section2"):



        for candidate in _business_collateral_key_aliases(fallback_sub):



            key = f"{prefix}_{candidate}"



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











def _business_has_sections(data):
    if not isinstance(data, dict):
        return False
    return any(
        isinstance(data.get(section), dict) and bool(data.get(section))
        for section in ("section1", "section2", "section3")
    )


def _business_section3_from_contract_rows(periods):
    contract = uploaded_data.get("contract_data") or {}
    rows = contract.get("raw_rows") or []
    if not isinstance(rows, list) or not rows:
        return {}

    groups = _business_configured_collateral_group_names() or [
        "\uc2e0\uc6a9",
        "\ub2f4\ubcf4",
        "\ubcf4\uc99d",
    ]
    group_lookup = {_business_normalize_group_name(name): name for name in groups}
    product_to_group = {}
    for group in uploaded_data.get("product_groups") or []:
        if not isinstance(group, dict):
            continue
        group_name = str(group.get("name") or "").strip()
        display_group = group_lookup.get(_business_normalize_group_name(group_name), group_name)
        if not display_group:
            continue
        for product in group.get("products") or []:
            product_name = str(product or "").strip()
            if product_name:
                product_to_group[product_name] = display_group

    labels = ["\uc2e0\uaddc", "\ucd94\uac00\ub300\ucd9c", "\uc7ac\ub300\ucd9c"]
    total_key = "\ucde8\uae09\uc561(\ub300\ucd9c)"
    period_set = set(periods or [])
    buckets = {
        group: {label: {period: 0.0 for period in periods} for label in labels + [total_key]}
        for group in groups
    }
    has_value = False

    for row in rows:
        if not isinstance(row, dict):
            continue
        period = _business_norm_period(row.get("period") or row.get("date"))
        if not period or period not in period_set:
            continue
        deal_type = str(row.get("deal_type") or "").strip()
        if deal_type == "\ucd94\uac00":
            deal_type = "\ucd94\uac00\ub300\ucd9c"
        if deal_type not in labels:
            continue
        group_name = str(row.get("group") or "").strip()
        if not group_name:
            product_name = str(row.get("product") or "").strip()
            group_name = product_to_group.get(product_name, "")
        group_name = group_lookup.get(_business_normalize_group_name(group_name), group_name)
        if group_name not in buckets:
            continue
        amount = row.get("amount")
        if amount in (None, ""):
            try:
                amount = float(row.get("amount_won") or 0) / 1_000_000
            except Exception:
                amount = 0
        try:
            amount = float(amount or 0)
        except Exception:
            amount = 0
        if not amount:
            continue
        buckets[group_name][deal_type][period] += amount
        buckets[group_name][total_key][period] += amount
        has_value = True

    if not has_value:
        return {}
    return {
        group: {
            label: [values.get(period) or None for period in periods]
            for label, values in group_values.items()
        }
        for group, group_values in buckets.items()
    }


def _business_detail_from_source_uploads():
    periods = _source_report_periods(
        "contract_data",
        "payment_data",
        "writeoff_data",
        "sale_data",
        "settlement_aq",
        "loan_loss_data",
    )
    if not periods:
        return None
    result = {
        "periods": periods,
        "section1": {},
        "section2": {},
        "section3": {},
        "_source": "source_uploads",
    }
    contract = uploaded_data.get("contract_data") or {}
    if not isinstance(contract, dict):
        return result

    source_periods = [
        _business_norm_period(period)
        for period in (contract.get("periods") or [])
        if _business_norm_period(period)
    ]

    def series_to_map(value):
        if isinstance(value, list):
            return {
                source_periods[idx]: item
                for idx, item in enumerate(value)
                if idx < len(source_periods)
            }
        if isinstance(value, dict):
            mapped = {}
            for key, item in value.items():
                norm = _business_norm_period(key)
                if norm:
                    mapped[norm] = item
            if mapped:
                return mapped
        return None

    def align_value(value):
        mapped = series_to_map(value)
        if mapped is not None:
            return [mapped.get(period) for period in periods]
        if isinstance(value, dict):
            return {key: align_value(child) for key, child in value.items()}
        return value

    for section_name in ("section1", "section2", "section3"):
        section = contract.get(section_name)
        if isinstance(section, dict) and section:
            result[section_name] = {
                key: align_value(value)
                for key, value in section.items()
            }

    raw_section3 = _business_section3_from_contract_rows(periods)
    if raw_section3:
        section3 = result.get("section3") if isinstance(result.get("section3"), dict) else {}
        for group_name, group_values in raw_section3.items():
            if not isinstance(group_values, dict):
                continue
            current = section3.get(group_name)
            if not isinstance(current, dict):
                section3[group_name] = group_values
                continue
            for key, series in group_values.items():
                current_series = current.get(key)
                if not isinstance(current_series, list) or not any(v is not None for v in current_series):
                    current[key] = series
        result["section3"] = section3

    for key_name in ("s1_keys", "s2_keys", "s3_keys"):
        if isinstance(contract.get(key_name), list):
            result[key_name] = copy.deepcopy(contract.get(key_name))

    return result


@app.get("/api/data/business")



async def get_business_detail():



    """\uc601\uc5c5\ud604\ud669 3\uc139\uc158 \ub370\uc774\ud130 (\uc5d1\uc140 \ud30c\uc2f1 \uae30\ubc18)"""



    d = uploaded_data.get("business_detail")



    fallback = _business_detail_from_source_uploads()
    if fallback:
        d = _merge_business_detail(fallback, d or {}, overwrite_null=False) if isinstance(d, dict) else fallback

    if not d:

        return JSONResponse({"error": "\ub370\uc774\ud130 \uc5c6\uc74c"}, status_code=404)



    return JSONResponse(_prune_indexed_period_data(copy.deepcopy(d)))











def _merge_business_detail(existing, incoming, sections=None, overwrite_null=True):
    if not isinstance(existing, dict):
        existing = {}
    if not isinstance(incoming, dict):
        return existing
    sections = sections or ["section1", "section2", "section3"]
    old_periods = [_business_norm_period(p) for p in (existing.get("periods") or []) if _business_norm_period(p)]
    new_periods = [_business_norm_period(p) for p in (incoming.get("periods") or []) if _business_norm_period(p)]
    merged_periods = sorted(
        (period for period in (set(old_periods) | set(new_periods)) if _report_period_is_visible(period)),
        key=_business_period_sort_key,
    )

    def series_to_map(series, periods):
        if isinstance(series, dict):
            return {_business_norm_period(k): v for k, v in series.items() if _business_norm_period(k)}
        if isinstance(series, list):
            return {period: series[idx] for idx, period in enumerate(periods) if idx < len(series)}
        return {}

    def is_nested_series(value):
        if not isinstance(value, dict):
            return False
        has_period_key = any(_business_norm_period(key) for key in value.keys())
        return not has_period_key and any(isinstance(child, (dict, list)) for child in value.values())

    def merge_item(existing_item, incoming_item):
        if is_nested_series(existing_item) or is_nested_series(incoming_item):
            existing_node = existing_item if isinstance(existing_item, dict) else {}
            incoming_node = incoming_item if isinstance(incoming_item, dict) else {}
            merged_node = {}
            for child_key in set(existing_node.keys()) | set(incoming_node.keys()):
                merged_node[child_key] = merge_item(existing_node.get(child_key), incoming_node.get(child_key))
            return merged_node
        values = series_to_map(existing_item, old_periods)
        for period, value in series_to_map(incoming_item, new_periods).items():
            if overwrite_null or value is not None:
                values[period] = value
        return [values.get(period) for period in merged_periods]

    result = dict(existing)
    result["periods"] = merged_periods
    for section in sections:
        section_result = dict(existing.get(section) or {})
        incoming_section = incoming.get(section) or {}
        if not isinstance(incoming_section, dict):
            continue
        for key in set(section_result.keys()) | set(incoming_section.keys()):
            section_result[key] = merge_item(section_result.get(key), incoming_section.get(key))
        result[section] = section_result
    for key_name in ("s1_keys", "s2_keys", "s3_keys"):
        keys = []
        for source in (existing.get(key_name), incoming.get(key_name)):
            if isinstance(source, list):
                for key in source:
                    if key not in keys:
                        keys.append(key)
        if keys:
            result[key_name] = keys
    result["is_collateral_only"] = False
    return result


@app.post("/api/upload/business")



async def upload_business(file: UploadFile = File(...), period: str = Form(default="")):



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







        data = _prune_indexed_period_data(parse_business_excel(tmp_path))
        if data.get("is_collateral_only"):
            data = _merge_business_detail(uploaded_data.get("business_detail") or {}, data, sections=["section3"])



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







    data = _prune_indexed_period_data(copy.deepcopy(data))



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







BSIS_DEFAULT_LABELS = {
    "bs": [
        "Ⅰ. 자산총계", "1. 현금 및 예치금", "2. 대출자산", "대출채권", "대손충당금",
        "3. 기타자산", "단기대여금", "기업대출금", "기타",
        "Ⅱ. 부채총계", "1. 차입부채", "차입금", "사채", "2. 기타부채",
        "Ⅲ. 자본총계", "1. 자본금", "2. 자본잉여금(자본조정)", "3. 이익잉여금",
    ],
    "is_": [
        "Ⅰ. 영업수익", "1. 대출채권 이자수익", "2. 대출채권평가 및 처분이익",
        "3. 외환거래이익", "4. 기타 이자수익", "5. 부동산 매각수익",
        "Ⅱ. 매출원가", "1. 부동산매출원가",
        "Ⅲ. 영업비용", "1. 차입금 이자비용", "2. 모집비용", "중개수수료", "광고선전비",
        "3. 대출채권평가 및 처분손실", "대손상각비", "대출채권처분손실",
        "4. 기타 판매관리비", "인건비", "기타",
        "Ⅳ. 영업이익", "1. 영업외수익", "2. 영업외비용",
        "Ⅴ. 법인세비용차감전순이익", "1. 법인세 등", "Ⅵ. 당기순이익",
    ],
    "summary": [
        "자산총계", "미수수익", "미수금", "선급금", "가지급금", "선급비용", "선납세금",
        "유형자산", "무형자산", "기타비유동자산", "기타자산", "현금및현금성자산",
        "대출채권", "대손충당금", "대출평잔", "부채총계", "미지급금", "예수금", "가수금",
        "미지급세금", "미지급비용", "당기법인세부채", "선수금", "기타부채",
        "단기차입금", "장기차입금", "사채", "자본총계", "자본금", "자본잉여금", "이익잉여금",
        "영업수익", "지급이자", "중개수수료", "판관비", "대손상각비", "영업이익", "당기손익",
        "영업활동 현금흐름", "투자활동 현금흐름", "재무활동 현금흐름",
        "부채비율", "차입금의존도", "ROA", "ROE", "영업현황 대출채권",
    ],
}


def _ensure_bsis_scaffold(bsis_data: dict | None, period: str, parsed: dict | None = None) -> tuple[dict, bool]:
    """Create the minimal BS/IS structure needed when a PDF is uploaded first."""
    if not isinstance(bsis_data, dict):
        bsis_data = {}
    added_period = False
    parsed = parsed or {}

    for sheet_key, default_labels in BSIS_DEFAULT_LABELS.items():
        sheet = bsis_data.get(sheet_key)
        if not isinstance(sheet, dict):
            sheet = {"headers": [], "rows": []}
            bsis_data[sheet_key] = sheet

        headers = sheet.get("headers")
        if not isinstance(headers, list):
            headers = []
        sheet["headers"] = headers

        rows = sheet.get("rows")
        if not isinstance(rows, list):
            rows = []
        sheet["rows"] = rows

        existing = {str(row.get("label", "")) for row in rows if isinstance(row, dict)}
        labels_to_add = list(default_labels)
        labels_to_add.extend([str(label) for label in (parsed.get(sheet_key) or {}).keys()])
        for label in labels_to_add:
            if not label or label in existing:
                continue
            rows.append({"label": label, "values": [None] * len(headers)})
            existing.add(label)

        if period and period not in headers:
            headers.append(period)
            headers.sort()
            insert_idx = headers.index(period)
            for row in rows:
                values = row.setdefault("values", [])
                while len(values) < len(headers) - 1:
                    values.append(None)
                values.insert(insert_idx, None)
            added_period = True

        for row in rows:
            values = row.setdefault("values", [])
            while len(values) < len(headers):
                values.append(None)
            if len(values) > len(headers):
                del values[len(headers):]

    return bsis_data, added_period


def _normalize_financial_mapping_period(period: str | None) -> str:
    value = str(period or "").strip().replace("-", ".")
    match = re.fullmatch(r"(\d{4})\.(0?[1-9]|1[0-2])", value)
    if not match:
        return ""
    return f"{match.group(1)}.{int(match.group(2)):02d}"


def _available_pdf_mapping_periods() -> list[str]:
    mappings = uploaded_data.get("pdf_mapping_by_period", {})
    if not isinstance(mappings, dict):
        return []
    return sorted(
        period
        for period, config in mappings.items()
        if _normalize_financial_mapping_period(period) and isinstance(config, dict)
    )


def _default_financial_mapping_period(period: str | None = None) -> str:
    return (
        _normalize_financial_mapping_period(period)
        or _normalize_financial_mapping_period(uploaded_data.get("financial_source_period"))
        or _normalize_financial_mapping_period(uploaded_data.get("bsis", {}).get("upload_period"))
    )


def _get_pdf_mapping_for_period(period: str | None = None, allow_previous: bool = True):
    target_period = _default_financial_mapping_period(period)
    mappings = uploaded_data.get("pdf_mapping_by_period", {})
    if isinstance(mappings, dict) and target_period:
        exact = mappings.get(target_period)
        if isinstance(exact, dict):
            return copy.deepcopy(exact), target_period, target_period, True
        if allow_previous:
            previous = [
                p for p, config in mappings.items()
                if isinstance(config, dict)
                and _normalize_financial_mapping_period(p)
                and _normalize_financial_mapping_period(p) < target_period
            ]
            if previous:
                source_period = sorted(previous)[-1]
                return copy.deepcopy(mappings[source_period]), target_period, source_period, False

    legacy = uploaded_data.get("pdf_mapping", None)
    if isinstance(legacy, dict):
        source_period = (
            _normalize_financial_mapping_period(legacy.get("_period"))
            or _normalize_financial_mapping_period(uploaded_data.get("financial_source_period"))
            or "legacy"
        )
        return copy.deepcopy(legacy), target_period, source_period, False
    return None, target_period, "", False


def _mapping_aux_config(mapping: dict | None):
    mapping = mapping if isinstance(mapping, dict) else {}
    return (
        mapping.get("_custom_labels") or uploaded_data.get("custom_excel_labels", {"bs": [], "is_": [], "summary": []}),
        mapping.get("_ordered_labels") or uploaded_data.get("ordered_excel_labels", {"bs": None, "is_": None, "summary": None}),
        mapping.get("_label_renames") or {},
    )


@app.post("/api/upload/bsis")



async def upload_bsis(
    file: UploadFile = File(...),
    period: str = Form(default=""),
):



    """BS/IS 엑셀 파일 업로드 (재무상태표 + 손익계산서 + BSPL_요약)"""



    period = str(period or "").strip().replace("-", ".")
    if not re.fullmatch(r"\d{4}\.(0[1-9]|1[0-2])", period):
        raise HTTPException(
            status_code=400,
            detail="재무제표를 반영할 기준 년월을 선택하세요 (예: 2026.06).",
        )

    if not file.filename.endswith(('.xlsx', '.xls')):



        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")



    try:



        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:



            content = await file.read()



            tmp.write(content)



            tmp_path = tmp.name







        wb = openpyxl.load_workbook(tmp_path, data_only=True)



        data = parse_bsis(wb)

        if not data:
            source = parse_financial_statement_source(wb)
            raw = source.get("raw") or {}
            if not raw:
                sheet_names = ", ".join(wb.sheetnames)
                raise ValueError(
                    "재무상태표/손익계산서 과목을 찾지 못했습니다. "
                    f"확인된 시트: {sheet_names or '없음'}"
                )

            # 인쇄형 재무제표 엑셀은 매핑 원본으로 사용한다. 기존 BS/IS
            # 대상 레이아웃은 유지하고 원본 과목 및 금액만 교체한다.
            uploaded_data["pdf_raw"] = raw
            uploaded_data["financial_source_options"] = source.get("options", {})
            uploaded_data["financial_source_type"] = "excel"
            uploaded_data["financial_source_filename"] = file.filename
            uploaded_data["financial_source_period"] = period

            target_period = period
            bsis_data = uploaded_data.get("bsis", {})
            mapping, _, _, _ = _get_pdf_mapping_for_period(target_period)
            mapped = {"bs": {}, "is_": {}, "summary": {}}
            if mapping:
                from pdf_parser import apply_user_mapping
                mapping_for_apply = dict(mapping)
                mapping_for_apply["_ordered_labels"] = mapping.get("_ordered_labels") or uploaded_data.get("ordered_excel_labels", {})
                mapped = apply_user_mapping(raw, mapping_for_apply)

            bsis_data, _ = _ensure_bsis_scaffold(bsis_data, target_period, mapped)
            if mapping:
                bsis_data, sheets_updated = apply_pdf_to_bsis(bsis_data, mapped, target_period)
            else:
                sheets_updated = []
            bsis_data["upload_period"] = target_period
            bsis_data["_source_type"] = "excel"
            bsis_data["_source_filename"] = file.filename
            uploaded_data["bsis"] = bsis_data

            os.unlink(tmp_path)
            counts = source.get("counts", {})
            return JSONResponse({
                "success": True,
                "source_type": "excel",
                "message": f"재무제표 원본 '{file.filename}' 항목 불러오기 완료",
                "sheets": [
                    f"재무상태표 {counts.get('bs', 0)}개 항목",
                    f"손익계산서 {counts.get('is_', 0)}개 항목",
                ],
                "sheets_updated": sheets_updated,
                "source_counts": counts,
                "upload_period": target_period,
            })



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



    return JSONResponse(
        uploaded_data.get("bsis", {}),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate"},
    )











@app.post("/api/upload/pdf-fs")



async def upload_pdf_fs(



    file: UploadFile = File(...),



    period: str = "2026.05"



):



    """



    재무제표 PDF 업로드 → BS/IS/Summary 특정 기간 컬럼 업데이트



    period: "YYYY.MM" 형식 (예: "2026.05")



    """



    period = _normalize_financial_mapping_period(period) or str(period or "").strip().replace("-", ".")

    if not file.filename.lower().endswith('.pdf'):



        raise HTTPException(status_code=400, detail="PDF 파일만 업로드 가능합니다.")







    bsis_data = uploaded_data.get("bsis", {})







    try:



        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:



            content = await file.read()



            tmp.write(content)



            tmp_path = tmp.name







        # PDF 파싱 (저장된 사용자 매핑 우선 적용)



        user_mapping, _, _, _ = _get_pdf_mapping_for_period(period)



        if user_mapping:
            user_mapping = dict(user_mapping)
            user_mapping["_ordered_labels"] = user_mapping.get("_ordered_labels") or uploaded_data.get("ordered_excel_labels", {})

        parsed = parse_pdf_fs(tmp_path, user_mapping=user_mapping)



        os.unlink(tmp_path)







        # raw 값 저장 (매핑 설정 UI의 '당기값' 표시용)



        if "raw" in parsed:



            uploaded_data["pdf_raw"] = parsed["raw"]
            uploaded_data["financial_source_type"] = "pdf"
            uploaded_data["financial_source_options"] = {}
            uploaded_data["financial_source_filename"] = file.filename
            uploaded_data["financial_source_period"] = period







        # BS/IS 엑셀이 없어도 PDF 결과를 담을 기본 구조를 만든다.
        bsis_data, added_period = _ensure_bsis_scaffold(bsis_data, period, parsed)







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



            },

            "bsis": updated_bsis



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



async def get_pdf_mapping(period: str = ""):



    """



    PDF 과목 → 엑셀 과목 매핑 설정 반환



    저장된 설정이 없으면 pdf_mapping.py 기본값으로 초기화



    """



    saved, target_period, mapping_source_period, period_has_own_mapping = _get_pdf_mapping_for_period(period)



    pdf_raw = uploaded_data.get("pdf_raw", {})







    # raw 값에서 당기값(cur) 추출 헬퍼



    def get_pdf_val(pdf_key):



        entry = pdf_raw.get(pdf_key)



        if entry and isinstance(entry, (list, tuple)) and len(entry) > 0 and entry[0] is not None:



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



    def build_pdf_options_from_raw():



        options = []



        seen = set()



        for raw_key, entry in pdf_raw.items():



            key = str(raw_key or "").strip()



            if not key or key in seen:



                continue



            seen.add(key)



            value = None



            if entry and isinstance(entry, (list, tuple)) and len(entry) > 0 and entry[0] is not None:



                value = entry[0]



            options.append({"pdf_key": key, "pdf_val": value})



        return options



    source_type = uploaded_data.get("financial_source_type", "pdf")
    source_options = uploaded_data.get("financial_source_options", {})
    pdf_options = build_pdf_options_from_raw()
    if source_type == "excel" and isinstance(source_options, dict):
        source_all = source_options.get("all")
        if isinstance(source_all, list) and source_all:
            pdf_options = source_all

    def source_options_for(sheet_key):
        if source_type == "excel" and isinstance(source_options, dict):
            values = source_options.get(sheet_key)
            if isinstance(values, list) and values:
                return values
        return pdf_options







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



        response = dict(saved)



        response["_pdf_options"] = {



            "all": pdf_options,



            "bs": source_options_for("bs"),



            "is_": source_options_for("is_"),



            "summary": source_options_for("summary"),



        }



        response["_pdf_raw_count"] = len(pdf_options)
        response["_source_type"] = source_type
        response["_source_filename"] = uploaded_data.get("financial_source_filename", "")
        response["_source_period"] = target_period or uploaded_data.get("financial_source_period", "")
        response["_financial_source_period"] = uploaded_data.get("financial_source_period", "")
        response["_period"] = target_period
        response["_mapping_source_period"] = mapping_source_period
        response["_period_has_own_mapping"] = period_has_own_mapping
        response["_available_mapping_periods"] = _available_pdf_mapping_periods()



        return JSONResponse(response)







    config = {



        "bs":      build_rows_from_raw(PDF_BS_MAPPING),



        "is_":     build_rows_from_raw(PDF_IS_MAPPING),



        "summary": build_rows_from_raw(PDF_SUMMARY_MAPPING),



        "excel_options": excel_opts,



        "_pdf_options": {



            "all": pdf_options,



            "bs": source_options_for("bs"),



            "is_": source_options_for("is_"),



            "summary": source_options_for("summary"),



        },



        "_pdf_raw_count": len(pdf_options),
        "_source_type": source_type,
        "_source_filename": uploaded_data.get("financial_source_filename", ""),
        "_source_period": target_period or uploaded_data.get("financial_source_period", ""),
        "_financial_source_period": uploaded_data.get("financial_source_period", ""),
        "_period": target_period,
        "_mapping_source_period": mapping_source_period,
        "_period_has_own_mapping": period_has_own_mapping,
        "_available_mapping_periods": _available_pdf_mapping_periods(),



    }



    return JSONResponse(config)











def _sync_bsis_labels_and_order(bsis_data: dict, ordered_labels: dict | None, label_renames: dict | None, custom_labels: dict | None = None) -> bool:
    """Apply mapping-screen target label edits to the persisted BS/IS rows."""
    if not isinstance(bsis_data, dict):
        return False
    changed = False
    ordered_labels = ordered_labels or {}
    label_renames = label_renames or {}
    custom_labels = custom_labels or {}

    for sheet_key in ["bs", "is_", "summary"]:
        sheet = bsis_data.get(sheet_key)
        if not isinstance(sheet, dict):
            continue
        rows = sheet.get("rows")
        if not isinstance(rows, list):
            continue

        renames = label_renames.get(sheet_key) or {}
        if isinstance(renames, dict) and renames:
            existing = {str(row.get("label", "")) for row in rows if isinstance(row, dict)}
            for old_label, new_label in renames.items():
                old_label = str(old_label or "").strip()
                new_label = str(new_label or "").strip()
                if not old_label or not new_label or old_label == new_label:
                    continue
                if old_label not in existing:
                    continue
                if new_label in existing:
                    raise HTTPException(status_code=400, detail=f"'{new_label}' 과목이 이미 존재합니다.")
                for row in rows:
                    if isinstance(row, dict) and row.get("label") == old_label:
                        row["label"] = new_label
                        changed = True
                        existing.remove(old_label)
                        existing.add(new_label)
                        break

        headers = sheet.get("headers") or []
        n_cols = len(headers)
        existing = {str(row.get("label", "")) for row in rows if isinstance(row, dict)}
        for label in custom_labels.get(sheet_key, []) or []:
            label = str(label or "").strip()
            if label and label not in existing:
                rows.append({"label": label, "values": [None] * n_cols, "indent": 0, "bold": False})
                existing.add(label)
                changed = True

        ordered = ordered_labels.get(sheet_key)
        if isinstance(ordered, list) and ordered:
            order_idx = {str(label): idx for idx, label in enumerate(ordered)}
            before = [row.get("label") for row in rows if isinstance(row, dict)]
            rows.sort(key=lambda row: order_idx.get(str(row.get("label", "")), len(order_idx)))
            after = [row.get("label") for row in rows if isinstance(row, dict)]
            if before != after:
                changed = True

    return changed


@app.post("/api/pdf-mapping")



async def save_pdf_mapping(request: Request):



    """



    PDF 매핑 설정 저장



    Body: { bs: [{pdf_key, excel_label, enabled}], is_: [...], summary: [...] }



    { _reset: true } 전달 시 저장된 설정 삭제 → 기본값으로 복원



    """



    body = await request.json()
    period = _normalize_financial_mapping_period(
        body.get("period") or request.query_params.get("period")
    )







    # _reset: true 처리 — 저장된 매핑 삭제 (기본값으로 복원)



    if body.get("_reset"):
        if period:
            mappings = uploaded_data.get("pdf_mapping_by_period", {})
            if isinstance(mappings, dict) and period in mappings:
                mappings = copy.deepcopy(mappings)
                mappings.pop(period, None)
                uploaded_data["pdf_mapping_by_period"] = mappings
            return JSONResponse({"success": True, "message": "mapping reset", "period": period})



        if "pdf_mapping" in uploaded_data:



            del uploaded_data["pdf_mapping"]



        return JSONResponse({"success": True, "message": "매핑 설정이 기본값으로 복원되었습니다."})







    body.pop("excel_options", None)



    body.pop("_reset", None)
    body.pop("period", None)



    # _custom_labels, _ordered_labels, _label_renames는 별도 키로 분리 저장



    custom_labels  = body.get("_custom_labels", None)



    ordered_labels = body.get("_ordered_labels", None)



    label_renames = body.get("_label_renames", None)







    # 모든 시트가 빈 배열이면 저장하지 않음 (기본 매핑 덮어쓰기 방지)



    total_rows = sum(len(body.get(k, [])) for k in ["bs", "is_", "summary"])
    child_sum_labels = body.get("_child_sum_labels") or {}
    total_child_sum_rows = sum(len(child_sum_labels.get(k, []) or []) for k in ["bs", "is_", "summary"])



    if total_rows > 0 or total_child_sum_rows > 0:



        if period:
            body["_period"] = period
            mappings = uploaded_data.get("pdf_mapping_by_period", {})
            if not isinstance(mappings, dict):
                mappings = {}
            mappings = copy.deepcopy(mappings)
            mappings[period] = body
            uploaded_data["pdf_mapping_by_period"] = mappings
        uploaded_data["pdf_mapping"] = body



    # custom_labels / ordered_labels는 내용과 관계없이 항상 저장



    if custom_labels is not None and not period:



        uploaded_data["custom_excel_labels"] = custom_labels



    if ordered_labels is not None and not period:



        uploaded_data["ordered_excel_labels"] = ordered_labels



    if any(value for value in (label_renames or {}).values()) or ordered_labels is not None or custom_labels is not None:
        bsis_data = uploaded_data.get("bsis", {})
        if _sync_bsis_labels_and_order(bsis_data, ordered_labels, label_renames, custom_labels):
            uploaded_data["bsis"] = bsis_data



    return JSONResponse({"success": True, "message": "매핑 설정이 저장되었습니다."})











@app.get("/api/pdf-mapping/excel-options")



async def get_excel_options(period: str = ""):



    """엑셀 BS/IS/Summary 과목 목록 반환 (매핑 드롭다운용)



    - bsis rows 전체 label 반환



    - 수기 추가 과목(custom_excel_labels) 별도 포함



    """



    bsis_data     = uploaded_data.get("bsis", {})



    mapping, target_period, mapping_source_period, period_has_own_mapping = _get_pdf_mapping_for_period(period)



    custom_labels, ordered_labels, _ = _mapping_aux_config(mapping)



    return JSONResponse({



        "bs":      [r["label"] for r in bsis_data.get("bs",      {}).get("rows", [])],



        "is_":     [r["label"] for r in bsis_data.get("is_",     {}).get("rows", [])],



        "summary": [r["label"] for r in bsis_data.get("summary", {}).get("rows", [])],



        "_custom_labels":  custom_labels,



        "_ordered_labels": ordered_labels,



        "_period": target_period,



        "_mapping_source_period": mapping_source_period,



        "_period_has_own_mapping": period_has_own_mapping,



        "_available_mapping_periods": _available_pdf_mapping_periods(),



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
    period = _normalize_financial_mapping_period(body.get("period"))
    period_mappings = None
    period_mapping = None
    if period:
        existing_mappings = uploaded_data.get("pdf_mapping_by_period", {})
        if isinstance(existing_mappings, dict) and isinstance(existing_mappings.get(period), dict):
            period_mappings = copy.deepcopy(existing_mappings)
            period_mapping = copy.deepcopy(period_mappings.get(period) or {})



    if not sheet or not label:



        raise HTTPException(status_code=400, detail="sheet, label 필수")







    # 1. custom_excel_labels 업데이트



    if period_mapping is not None:
        custom = copy.deepcopy(period_mapping.get("_custom_labels") or {"bs": [], "is_": [], "summary": []})
    else:
        custom = uploaded_data.get("custom_excel_labels", {"bs": [], "is_": [], "summary": []})



    if label in custom.get(sheet, []):



        custom[sheet] = [l for l in custom[sheet] if l != label]



        if period_mapping is not None:
            period_mapping["_custom_labels"] = custom
        else:
            uploaded_data["custom_excel_labels"] = custom







    # 2. ordered_excel_labels 업데이트



    if period_mapping is not None:
        ordered = copy.deepcopy(period_mapping.get("_ordered_labels") or {"bs": None, "is_": None, "summary": None})
    else:
        ordered = uploaded_data.get("ordered_excel_labels", {"bs": None, "is_": None, "summary": None})



    if ordered.get(sheet) and label in ordered[sheet]:



        ordered[sheet] = [l for l in ordered[sheet] if l != label]



        if period_mapping is not None:
            period_mapping["_ordered_labels"] = ordered
        else:
            uploaded_data["ordered_excel_labels"] = ordered



    if period_mapping is not None and period_mappings is not None:



        period_mappings[period] = period_mapping



        uploaded_data["pdf_mapping_by_period"] = period_mappings







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



    period = _normalize_financial_mapping_period(body.get("period", None))







    # 1. 저장된 raw 확인



    pdf_raw = uploaded_data.get("pdf_raw", {})



    if not pdf_raw:



        raise HTTPException(status_code=400, detail="저장된 PDF raw 데이터가 없습니다. PDF를 먼저 업로드하세요.")







    # 2. 매핑 확인



    user_mapping, mapping_target_period, _, _ = _get_pdf_mapping_for_period(period)



    if not user_mapping:



        raise HTTPException(status_code=400, detail="저장된 매핑 설정이 없습니다.")







    # 3. bsis 확인



    bsis_data = uploaded_data.get("bsis", {})







    # 4. period 자동 결정 (지정 안 하면 _pdf_updated에서 가장 최근 기간)



    if not period:



        period = mapping_target_period or uploaded_data.get("financial_source_period") or bsis_data.get("upload_period")



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



    period = _normalize_financial_mapping_period(period) or period



    bsis_data, _ = _ensure_bsis_scaffold(bsis_data, period)







    # 5. 수기 추가 과목을 bsis rows에 없으면 삽입



    custom_labels, ordered_labels, _ = _mapping_aux_config(user_mapping)







    sheet_map = {"bs": "bs", "is_": "is_", "summary": "summary"}



    for sh_key, sh_name in sheet_map.items():



        sheet = bsis_data.get(sh_key)



        if not sheet:



            continue



        existing_labels = {r["label"] for r in sheet.get("rows", [])}



        custom_list = custom_labels.get(sh_key, []) or []







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



    mapping_for_apply = dict(user_mapping)
    mapping_for_apply["_ordered_labels"] = ordered_labels
    mapped = apply_user_mapping(pdf_raw, mapping_for_apply)

    bsis_data, _ = _ensure_bsis_scaffold(bsis_data, period, mapped)



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


PAYMENT_RAW_ROWS_PREFIX = "payment_raw_rows"


def _payment_raw_rows_key(period: str) -> str:
    norm = _business_norm_period(period) or str(period or "").strip()
    safe = re.sub(r"[^0-9A-Za-z_-]", "_", norm)
    return f"{PAYMENT_RAW_ROWS_PREFIX}_{safe}"


def _group_payment_raw_rows(rows: list | None) -> dict[str, list]:
    grouped: dict[str, list] = {}
    if not isinstance(rows, list):
        return grouped
    for row in rows:
        if not isinstance(row, dict):
            continue
        period = _business_norm_period(row.get("period")) or str(row.get("period") or "").strip()
        if not period:
            continue
        clean_row = dict(row)
        clean_row["period"] = period
        grouped.setdefault(period, []).append(clean_row)
    return grouped


def _save_payment_raw_rows(rows: list | None, periods: list | set | tuple | None = None) -> dict[str, int]:
    grouped = _group_payment_raw_rows(rows)
    target_periods = set(grouped.keys())
    if isinstance(periods, (list, set, tuple)):
        for period in periods:
            norm = _business_norm_period(period) or str(period or "").strip()
            if norm:
                target_periods.add(norm)
    saved_counts: dict[str, int] = {}
    for period in sorted(target_periods, key=_business_period_sort_key):
        period_rows = grouped.get(period, [])
        save_data(_payment_raw_rows_key(period), period_rows)
        saved_counts[period] = len(period_rows)
    return saved_counts


def _load_payment_raw_rows(periods: list | tuple | set | None = None) -> list:
    rows: list = []
    if not isinstance(periods, (list, tuple, set)):
        return rows
    for period in sorted({_business_norm_period(p) or str(p or "").strip() for p in periods if p}, key=_business_period_sort_key):
        if not period:
            continue
        period_rows = load_data(_payment_raw_rows_key(period)) or []
        if isinstance(period_rows, list):
            rows.extend([row for row in period_rows if isinstance(row, dict)])
    return rows







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

        if not parsed_periods:

            os.unlink(tmp_path)

            return JSONResponse({

                "success": False,

                "code": "CONTRACT_PARSE_EMPTY",

                "selected_period": selected_period,

                "message": "계약일 기준으로 읽힌 계약리스트 행이 없습니다. 계약일/상품명/계약구분/대출액/대출잔액 열을 확인해 주세요.",

                "header_row": data.get("header_row"),

                "column_map": data.get("column_map"),

            }, status_code=422)



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
        merge_source = {k: v for k, v in data.items() if k != 'raw_rows'}
        merged_contract_data = _merge_uploaded_periodic_data('contract_data', merge_source)
        if 'raw_rows' in data:
            merged_contract_data['raw_rows'] = preserved_raw_rows + (data.get('raw_rows') or [])
        uploaded_data['contract_data'] = merged_contract_data



        os.unlink(tmp_path)







        # ── asset_data['sections']['대출채권잔액'] 자동 업데이트 ──



        section_balance = data.get("section_balance", {})



        if False and section_balance:  # derived at GET /api/data/asset; do not cross-write asset_data



            asset_data = _ensure_asset_data_shell(data.get("periods", []))



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



                        normalized_vals = {}
                        for ym, val in (ym_vals or {}).items():
                            normalized_vals[ym] = None if val is None else val * 1_000_000

                        existing_deal[row_label].update(normalized_vals)



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



        if False and avg_rate_deal:  # derived at GET /api/data/asset; do not cross-write asset_data



            asset_data = _ensure_asset_data_shell(data.get("periods", []))



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

            "selected_period": selected_period,

            "period_mismatch_confirmed": bool(data.get("period_mismatch_confirmed")),



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
        merge_source = {k: v for k, v in data.items() if k != 'loan_loss_rows'}
        merged_data = _merge_uploaded_periodic_data('full_contract_data', merge_source)
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


def _repair_fms_text(value):
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        repaired = text.encode("latin-1").decode("utf-8")
        return repaired if repaired else text
    except UnicodeError:
        return text


def _normalize_fms_borrowing_rows(rows):
    normalized = []
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        clean = dict(row)
        for key in ("investor_name", "borrower", "contract_name", "round_label"):
            clean[key] = _repair_fms_text(clean.get(key))
        normalized.append(clean)
    return normalized


@app.post("/api/upload/fms-borrowing")
async def upload_fms_borrowing(file: UploadFile = File(...), period: str = Form(default="")):
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Excel file only.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        selected_period = normalize_period_key(period) or str(period or "").strip()
        data = parse_fms_borrowing_list(tmp_path, selected_period)
        data["rows"] = _normalize_fms_borrowing_rows(data.get("rows") or [])
        existing_data = uploaded_data.get("fms_borrowing_data") or {}
        incoming_periods = set(data.get("periods") or [])
        existing_rows = existing_data.get("rows") or []
        preserved_rows = [
            row for row in existing_rows
            if not isinstance(row, dict) or (incoming_periods and row.get("period") not in incoming_periods)
        ] if incoming_periods else list(existing_rows)
        rows = preserved_rows + (data.get("rows") or [])
        periods = sorted({
            str(row.get("period") or "").strip()
            for row in rows
            if isinstance(row, dict) and str(row.get("period") or "").strip()
        })
        merged_data = {
            "period": selected_period,
            "periods": periods,
            "rows": rows,
            "total_balance": sum(float(row.get("borrow_balance") or 0) for row in rows if isinstance(row, dict)),
            "source": "fms_borrowing",
            "source_file": file.filename,
        }
        uploaded_data["fms_borrowing_data"] = merged_data

        return JSONResponse({
            "success": True,
            "message": f"FMS borrowing list '{file.filename}' uploaded",
            "periods": data.get("periods", []),
            "rows": len(data.get("rows") or []),
            "total_balance": data.get("total_balance", 0),
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"file processing error: {str(e)}\n{traceback.format_exc()}")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


@app.get("/api/data/fms-borrowing")
async def get_fms_borrowing_data(period: str = ""):
    data = uploaded_data.get("fms_borrowing_data")
    if not data:
        return JSONResponse({"error": "no data"}, status_code=404)
    data = {**data, "rows": _normalize_fms_borrowing_rows(data.get("rows") or [])}
    period_key = normalize_period_key(period) or str(period or "").strip()
    if period_key:
        rows = [
            row for row in (data.get("rows") or [])
            if isinstance(row, dict) and row.get("period") == period_key
        ]
        return JSONResponse({
            **data,
            "period": period_key,
            "periods": [period_key] if rows else [],
            "rows": rows,
            "total_balance": sum(float(row.get("borrow_balance") or 0) for row in rows),
        })
    return JSONResponse(data)


def _fms_payment_row_period(row):
    if not isinstance(row, dict):
        return ""
    raw = row.get("raw") if isinstance(row.get("raw"), dict) else {}
    for value in (row.get("payment_date"), raw.get("__col_C"), row.get("period")):
        text = str(value or "").strip()
        if not text:
            continue
        period_key = normalize_period_key(text)
        if period_key:
            return period_key
        digits = re.sub(r"\D", "", text)
        if len(digits) >= 6:
            return f"{digits[:4]}-{digits[4:6]}"
    return ""


def _normalize_fms_payment_rows(rows):
    normalized = []
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        clean = dict(row)
        raw = clean.get("raw") if isinstance(clean.get("raw"), dict) else {}
        for key in (
            "investor_name", "borrower", "contract_name", "round_label", "payment_date",
            "contract_date", "maturity_date", "agreement_day", "category",
        ):
            clean[key] = _repair_fms_text(clean.get(key))
        if raw and not clean.get("category"):
            clean["category"] = _repair_fms_text(
                raw.get("차입처구분")
                or raw.get("__col_T")
                or raw.get("유형")
                or raw.get("구분")
                or ""
            )
        for key in ("borrow_amount", "payment_amount", "principal_amount", "interest_amount", "balance_amount"):
            try:
                clean[key] = float(clean.get(key) or 0)
            except (TypeError, ValueError):
                clean[key] = 0.0
        period_key = _fms_payment_row_period(clean)
        if period_key:
            clean["period"] = period_key
        normalized.append(clean)
    return normalized


def _fms_payment_totals(rows):
    totals = {
        "borrow_amount": 0.0,
        "payment_amount": 0.0,
        "principal_amount": 0.0,
        "interest_amount": 0.0,
        "balance_amount": 0.0,
    }
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        for key in totals:
            try:
                totals[key] += float(row.get(key) or 0)
            except (TypeError, ValueError):
                pass
    return totals


FMS_PAYMENT_RAW_KEYS = {
    "__col_A",
    "__col_B",
    "__col_C",
    "__col_F",
    "__col_G",
    "__col_O",
    "__col_P",
    "__col_Q",
    "__col_S",
    "__col_T",
    "업체명",
    "계약명",
    "지급일자",
    "지급이자",
    "지급원금",
    "차입금액",
    "계약일",
    "만료일",
    "약정일",
    "차입처구분",
    "유형",
    "구분",
}


def _compact_fms_payment_row(row):
    if not isinstance(row, dict):
        return row
    compact = dict(row)
    raw = compact.get("raw")
    if isinstance(raw, dict):
        compact["raw"] = {key: value for key, value in raw.items() if key in FMS_PAYMENT_RAW_KEYS}
    return compact


def _compact_fms_payment_data(data):
    if not isinstance(data, dict):
        return data
    compact = dict(data)
    compact["rows"] = [_compact_fms_payment_row(row) for row in (compact.get("rows") or [])]
    return compact


def _fms_payment_period_totals(rows):
    result = {}
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        period_key = _fms_payment_row_period(row)
        if not period_key:
            continue
        if period_key not in result:
            result[period_key] = {
                "borrow_amount": 0.0,
                "payment_amount": 0.0,
                "principal_amount": 0.0,
                "interest_amount": 0.0,
                "balance_amount": 0.0,
                "rows": 0,
            }
        result[period_key]["rows"] += 1
        for key in ("borrow_amount", "payment_amount", "principal_amount", "interest_amount", "balance_amount"):
            try:
                result[period_key][key] += float(row.get(key) or 0)
            except (TypeError, ValueError):
                pass
    return result


@app.post("/api/upload/fms-payment")
async def upload_fms_payment(file: UploadFile = File(...), period: str = Form(default="")):
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Excel file only.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        selected_period = normalize_period_key(period) or str(period or "").strip()
        data = parse_fms_payment_list(tmp_path, selected_period)
        data["rows"] = _normalize_fms_payment_rows(data.get("rows") or [])
        existing_data = uploaded_data.get("fms_payment_data") or {}
        new_rows = data.get("rows") or []
        incoming_periods = {
            _fms_payment_row_period(row)
            for row in new_rows
            if isinstance(row, dict) and _fms_payment_row_period(row)
        }
        if not incoming_periods:
            incoming_periods = {
                str(item or "").strip()
                for item in (data.get("periods") or [])
                if str(item or "").strip()
            }
        existing_rows = existing_data.get("rows") or []
        preserved_rows = [
            row for row in existing_rows
            if not isinstance(row, dict) or _fms_payment_row_period(row) not in incoming_periods
        ] if incoming_periods else list(existing_rows)
        rows = preserved_rows + new_rows
        periods = sorted({
            _fms_payment_row_period(row)
            for row in rows
            if isinstance(row, dict) and _fms_payment_row_period(row)
        })
        merged_data = {
            "period": selected_period,
            "periods": periods,
            "rows": rows,
            "totals": _fms_payment_totals(rows),
            "period_totals": _fms_payment_period_totals(rows),
            "total_borrow_amount": sum(float(row.get("borrow_amount") or 0) for row in rows if isinstance(row, dict)),
            "total_payment_amount": sum(float(row.get("payment_amount") or 0) for row in rows if isinstance(row, dict)),
            "source": "fms_payment",
            "source_file": file.filename,
        }
        uploaded_data["fms_payment_data"] = merged_data

        return JSONResponse({
            "success": True,
            "message": f"FMS payment list '{file.filename}' uploaded",
            "periods": sorted(incoming_periods) or data.get("periods", []),
            "rows": len(data.get("rows") or []),
            "totals": data.get("totals", {}),
            "total_borrow_amount": data.get("total_borrow_amount", 0),
            "total_payment_amount": data.get("total_payment_amount", 0),
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"file processing error: {str(e)}\n{traceback.format_exc()}")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


@app.get("/api/data/fms-payment")
async def get_fms_payment_data(period: str = "", compact: str = "1"):
    data = uploaded_data.get("fms_payment_data")
    if not data:
        return JSONResponse({"error": "no data"}, status_code=404)
    data = {**data, "rows": _normalize_fms_payment_rows(data.get("rows") or [])}
    use_compact = str(compact or "1").lower() not in {"0", "false", "no"}
    period_key = normalize_period_key(period) or str(period or "").strip()
    if period_key:
        rows = [
            row for row in (data.get("rows") or [])
            if isinstance(row, dict) and _fms_payment_row_period(row) == period_key
        ]
        totals = _fms_payment_totals(rows)
        response = {
            "period": period_key,
            "periods": [period_key] if rows else [],
            "rows": rows,
            "totals": totals,
            "period_totals": {period_key: {**totals, "rows": len(rows)}} if rows else {},
            "total_borrow_amount": totals["borrow_amount"],
            "total_payment_amount": totals["payment_amount"],
            "uploaded_file": data.get("uploaded_file"),
            "updated_at": data.get("updated_at"),
        }
        if use_compact:
            response = _compact_fms_payment_data(response)
        return JSONResponse(response)
    if use_compact:
        data = _compact_fms_payment_data(data)
    return JSONResponse(data)


def _borrowing_schedule_number(value):
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


def _normalize_borrowing_schedule(schedule):
    normalized = []
    for item in schedule or []:
        if not isinstance(item, dict):
            continue
        due_date = str(item.get("due_date") or item.get("payment_date") or "").strip()
        round_no = item.get("round")
        try:
            round_no = int(float(str(round_no).replace(",", "")))
        except (TypeError, ValueError):
            round_no = len(normalized) + 1
        clean = {
            "round": round_no,
            "due_date": due_date,
            "payment_amount": _borrowing_schedule_number(item.get("payment_amount")),
            "principal_amount": _borrowing_schedule_number(item.get("principal_amount")),
            "interest_amount": _borrowing_schedule_number(item.get("interest_amount")),
            "balance_amount": _borrowing_schedule_number(item.get("balance_amount")),
        }
        for key in ("period_start", "period_end", "use_days", "income_tax", "local_income_tax", "auto_generated"):
            if key in item:
                clean[key] = item.get(key)
        normalized.append(clean)
    return normalized


def _borrowing_schedule_totals(schedule):
    rows = _normalize_borrowing_schedule(schedule)
    return {
        "rows": len(rows),
        "payment_amount": sum(float(row.get("payment_amount") or 0) for row in rows),
        "principal_amount": sum(float(row.get("principal_amount") or 0) for row in rows),
        "interest_amount": sum(float(row.get("interest_amount") or 0) for row in rows),
        "last_balance_amount": float(rows[-1].get("balance_amount") or 0) if rows else 0.0,
    }


def _borrowing_parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    compact = re.sub(r"\D", "", text)
    candidates = []
    if len(compact) == 8:
        candidates.append(f"{compact[:4]}-{compact[4:6]}-{compact[6:8]}")
    elif len(compact) == 6:
        candidates.append(f"20{compact[:2]}-{compact[2:4]}-{compact[4:6]}")
    candidates.append(text[:10])
    for candidate in candidates:
        try:
            return datetime.strptime(candidate, "%Y-%m-%d").date()
        except (TypeError, ValueError):
            continue
    return None


def _borrowing_is_leap_year(year):
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


def _borrowing_month_end(day):
    return date(day.year, day.month, calendar.monthrange(day.year, day.month)[1])


def _borrowing_next_month_payment_date(period_end):
    year = period_end.year + (1 if period_end.month == 12 else 0)
    month = 1 if period_end.month == 12 else period_end.month + 1
    payment_date = date(year, month, 10)
    if payment_date.weekday() == 5:
        return payment_date - timedelta(days=1)
    if payment_date.weekday() == 6:
        return payment_date - timedelta(days=2)
    return payment_date


def _borrowing_round_won(value):
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return 0


def _borrowing_floor_10(value):
    try:
        return int(float(value) // 10) * 10
    except (TypeError, ValueError):
        return 0


def _borrowing_contract_identity(row, include_source=True):
    if not isinstance(row, dict):
        return ""
    parts = [
        normalize_period_key(row.get("period")) or str(row.get("period") or "").strip(),
        str(row.get("investor_name") or "").strip(),
        str(row.get("contract_name") or "").strip(),
        str(row.get("key") or "").strip(),
    ]
    if include_source:
        parts.append(str(row.get("source_row") or "").strip())
    return "|".join(parts)


def _copy_existing_borrowing_schedules(incoming_rows, existing_rows):
    exact = {}
    loose = {}
    for row in existing_rows or []:
        schedule = _normalize_borrowing_schedule(row.get("repayment_schedule") or [])
        if not schedule:
            continue
        exact[_borrowing_contract_identity(row, include_source=True)] = schedule
        loose[_borrowing_contract_identity(row, include_source=False)] = schedule
    for row in incoming_rows or []:
        schedule = exact.get(_borrowing_contract_identity(row, include_source=True)) or loose.get(_borrowing_contract_identity(row, include_source=False))
        if schedule:
            row["repayment_schedule"] = schedule
            row["repayment_schedule_totals"] = _borrowing_schedule_totals(schedule)
    return incoming_rows


def _normalize_borrowing_contract_category(row):
    def _coerce(value):
        compact_value = str(value or "").strip().replace(" ", "")
        if "금융기관" in compact_value:
            return "금융기관"
        if "사모사채" in compact_value or compact_value == "사채":
            return "사모사채"
        if "일반차입" in compact_value:
            return "일반차입금"
        return ""

    # T열(차입처구분)을 서버 저장값보다 우선한다.
    t_category = _coerce((row or {}).get("investor_type")) or _coerce((row or {}).get("borrower_type"))
    if t_category:
        return t_category

    saved_category = _coerce((row or {}).get("category"))
    if saved_category:
        return saved_category

    contract_category = _coerce((row or {}).get("contract_kind"))
    if contract_category:
        return contract_category

    values = [
        (row or {}).get("investor_name"),
        (row or {}).get("contract_name"),
    ]
    text = " ".join(str(value or "") for value in values).replace(" ", "")
    if "금융기관" in text or "은행" in text or "저축" in text or "캐피탈" in text or "카드" in text:
        return "금융기관"
    if "사모사채" in text or text == "사채":
        return "사모사채"
    if "일반차입" in text:
        return "일반차입금"
    category = str((row or {}).get("category") or "").strip()
    return category or "일반차입금"


def _normalize_borrowing_repayment_method(value, category):
    text = str(value or "").strip()
    compact = text.replace(" ", "")
    if compact == "만기일시(익월)":
        return "만기일시"
    if text:
        return text
    return "상환스케줄" if category == "금융기관" else "만기일시"


def _borrowing_auto_schedule_required(row):
    category = str((row or {}).get("category") or "").strip()
    method = str((row or {}).get("repayment_method") or "").replace(" ", "")
    return category in ("사모사채", "일반차입금") and (not method or "만기일시" in method)


_BORROWING_FIXED_PUBLIC_HOLIDAYS = {
    "01-01",
    "03-01",
    "05-05",
    "06-06",
    "08-15",
    "10-03",
    "10-09",
    "12-25",
}


def _borrowing_is_non_business_day(day):
    if not day:
        return False
    return day.weekday() >= 5 or day.strftime("%m-%d") in _BORROWING_FIXED_PUBLIC_HOLIDAYS


def _borrowing_previous_business_day(day):
    adjusted = day
    while _borrowing_is_non_business_day(adjusted):
        adjusted -= timedelta(days=1)
    return adjusted


def _build_borrowing_bullet_schedule(row):
    if not _borrowing_auto_schedule_required(row):
        return None
    contract_date = _borrowing_parse_date((row or {}).get("contract_date"))
    maturity_date = _borrowing_parse_date((row or {}).get("maturity_date"))
    maturity_date = _borrowing_previous_business_day(maturity_date) if maturity_date else None
    if not contract_date or not maturity_date or maturity_date <= contract_date:
        return []

    try:
        principal = float((row or {}).get("borrow_amount") or (row or {}).get("borrow_balance") or 0)
    except (TypeError, ValueError):
        principal = 0.0
    try:
        raw_rate = float((row or {}).get("interest_rate") or 0)
    except (TypeError, ValueError):
        raw_rate = 0.0
    annual_rate = raw_rate / 100 if raw_rate > 1 else raw_rate
    if principal <= 0 or annual_rate <= 0:
        return []

    category = str((row or {}).get("category") or "").strip()
    tax_rate = 0.14 if category == "사모사채" else 0.25
    balance_amount = float((row or {}).get("borrow_balance") or principal)
    schedule = []
    period_start = contract_date + timedelta(days=1)
    round_no = 1
    while period_start <= maturity_date:
        month_end = _borrowing_month_end(period_start)
        regular_due_date = _borrowing_next_month_payment_date(month_end)
        if regular_due_date >= maturity_date:
            period_end = maturity_date
            due_date = maturity_date
        else:
            period_end = month_end
            due_date = regular_due_date
        use_days = (period_end - period_start).days + 1
        days_in_year = 366 if _borrowing_is_leap_year(period_start.year) else 365
        interest_amount = _borrowing_round_won(principal * annual_rate / days_in_year * use_days)
        income_tax = _borrowing_floor_10(interest_amount * tax_rate)
        local_income_tax = _borrowing_floor_10(income_tax * 0.10)
        payment_amount = interest_amount - income_tax - local_income_tax
        schedule.append({
            "round": round_no,
            "due_date": due_date.strftime("%Y-%m-%d"),
            "payment_amount": float(payment_amount),
            "principal_amount": 0.0,
            "interest_amount": float(interest_amount),
            "balance_amount": balance_amount,
            "period_start": period_start.strftime("%Y-%m-%d"),
            "period_end": period_end.strftime("%Y-%m-%d"),
            "use_days": use_days,
            "income_tax": float(income_tax),
            "local_income_tax": float(local_income_tax),
            "auto_generated": True,
            "effective_maturity_date": maturity_date.strftime("%Y-%m-%d"),
        })
        period_start = period_end + timedelta(days=1)
        round_no += 1
    return schedule


def _normalize_borrowing_contract_rows(rows):
    normalized = []
    for row in rows or []:
        if not isinstance(row, dict):
            continue
        clean = dict(row)
        clean["period"] = normalize_period_key(clean.get("period")) or str(clean.get("period") or "").strip()
        clean["category"] = _normalize_borrowing_contract_category(clean)
        clean["investor_type"] = str(clean.get("investor_type") or clean.get("category") or "").strip()
        clean["repayment_method"] = _normalize_borrowing_repayment_method(clean.get("repayment_method"), clean["category"])
        for key in ("borrow_amount", "borrow_balance", "interest_rate", "prepayment_fee_rate", "monthly_interest_estimate", "next_payment_amount"):
            try:
                clean[key] = float(clean.get(key) or 0)
            except (TypeError, ValueError):
                clean[key] = 0.0
        auto_schedule = _build_borrowing_bullet_schedule(clean)
        clean["repayment_schedule"] = _normalize_borrowing_schedule(
            auto_schedule if auto_schedule is not None else (clean.get("repayment_schedule") or [])
        )
        clean["repayment_schedule_totals"] = _borrowing_schedule_totals(clean["repayment_schedule"])
        normalized.append(clean)
    return normalized


def _build_borrowing_contract_snapshot(rows, period: str = "", source_file: str = ""):
    rows = _normalize_borrowing_contract_rows(rows)
    periods = sorted({
        str(row.get("period") or "").strip()
        for row in rows
        if str(row.get("period") or "").strip()
    })
    categories = ("금융기관", "사모사채", "일반차입금")
    by_category = {}
    for category in categories:
        category_rows = [row for row in rows if row.get("category") == category]
        by_category[category] = {
            "rows": len(category_rows),
            "borrow_amount": sum(float(row.get("borrow_amount") or 0) for row in category_rows),
            "borrow_balance": sum(float(row.get("borrow_balance") or 0) for row in category_rows),
            "monthly_interest_estimate": sum(float(row.get("monthly_interest_estimate") or 0) for row in category_rows),
        }
    return {
        "period": period or (periods[-1] if periods else ""),
        "periods": periods,
        "rows": rows,
        "totals": {
            "rows": len(rows),
            "borrow_amount": sum(float(row.get("borrow_amount") or 0) for row in rows),
            "borrow_balance": sum(float(row.get("borrow_balance") or 0) for row in rows),
            "monthly_interest_estimate": sum(float(row.get("monthly_interest_estimate") or 0) for row in rows),
        },
        "by_category": by_category,
        "source": "borrowing_contracts",
        "source_file": source_file,
    }


@app.post("/api/upload/borrowing-contracts")
async def upload_borrowing_contracts(file: UploadFile = File(...), period: str = Form(default="")):
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Excel .xlsx file only.")

    tmp_path = None
    try:
        suffix = os.path.splitext(filename)[1] or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        selected_period = normalize_period_key(period) or normalize_period_key(filename) or str(period or "").strip()
        data = parse_borrowing_contract_list(tmp_path, selected_period)
        incoming_rows = _normalize_borrowing_contract_rows(data.get("rows") or [])
        incoming_periods = {
            str(row.get("period") or "").strip()
            for row in incoming_rows
            if str(row.get("period") or "").strip()
        }
        existing_data = uploaded_data.get("borrowing_contract_data") or {}
        existing_rows = _normalize_borrowing_contract_rows(existing_data.get("rows") or [])
        incoming_rows = _copy_existing_borrowing_schedules(incoming_rows, existing_rows)
        preserved_rows = [
            row for row in existing_rows
            if not incoming_periods or str(row.get("period") or "").strip() not in incoming_periods
        ]
        merged_data = _build_borrowing_contract_snapshot(
            preserved_rows + incoming_rows,
            selected_period or data.get("period") or "",
            filename,
        )
        uploaded_data["borrowing_contract_data"] = merged_data
        return JSONResponse({
            "success": True,
            "message": f"Borrowing contract list '{filename}' uploaded",
            "period": merged_data.get("period") or "",
            "periods": sorted(incoming_periods),
            "rows": len(incoming_rows),
            "totals": data.get("totals") or {},
            "data": merged_data,
        })
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"file processing error: {str(e)}\n{traceback.format_exc()}")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


@app.post("/api/data/borrowing-contracts/schedule")
async def save_borrowing_contract_schedule(payload: dict = Body(...)):
    data = uploaded_data.get("borrowing_contract_data")
    if not data:
        raise HTTPException(status_code=404, detail="차입 계약리스트 데이터가 없습니다.")

    rows = _normalize_borrowing_contract_rows(data.get("rows") or [])
    period = normalize_period_key(payload.get("period")) or str(payload.get("period") or "").strip()
    source_row = str(payload.get("source_row") or "").strip()
    key = str(payload.get("key") or "").strip()
    investor_name = str(payload.get("investor_name") or "").strip()
    contract_name = str(payload.get("contract_name") or "").strip()
    schedule = _normalize_borrowing_schedule(payload.get("schedule") or [])

    def matches(row):
        if period and str(row.get("period") or "").strip() != period:
            return False
        if source_row and str(row.get("source_row") or "").strip() == source_row:
            return True
        if key and str(row.get("key") or "").strip() == key and investor_name and str(row.get("investor_name") or "").strip() == investor_name:
            return True
        return (
            investor_name
            and contract_name
            and str(row.get("investor_name") or "").strip() == investor_name
            and str(row.get("contract_name") or "").strip() == contract_name
        )

    updated = False
    for row in rows:
        if matches(row):
            row["repayment_schedule"] = schedule
            row["repayment_schedule_totals"] = _borrowing_schedule_totals(schedule)
            updated = True
            break
    if not updated:
        raise HTTPException(status_code=404, detail="상환스케줄을 저장할 계약을 찾지 못했습니다.")

    snapshot = _build_borrowing_contract_snapshot(rows, data.get("period") or period, data.get("source_file") or "")
    uploaded_data["borrowing_contract_data"] = snapshot
    return JSONResponse({
        "success": True,
        "message": "상환스케줄이 저장되었습니다.",
        "schedule_rows": len(schedule),
        "data": snapshot,
    })


@app.get("/api/data/borrowing-contracts")
async def get_borrowing_contracts(period: str = ""):
    data = uploaded_data.get("borrowing_contract_data")
    if not data:
        return JSONResponse({"error": "no data"}, status_code=404)
    data = _build_borrowing_contract_snapshot(data.get("rows") or [], data.get("period") or "", data.get("source_file") or "")
    period_key = normalize_period_key(period) or str(period or "").strip()
    if period_key:
        rows = [
            row for row in (data.get("rows") or [])
            if isinstance(row, dict) and row.get("period") == period_key
        ]
        return JSONResponse(_build_borrowing_contract_snapshot(rows, period_key, data.get("source_file") or ""))
    return JSONResponse(data)


def _borrowing_category_label(row):
    for key in ("investor_type", "borrower_type", "category"):
        category = str((row or {}).get(key) or "").strip()
        category_compact = category.replace(" ", "")
        if "금융기관" in category_compact:
            return "금융기관"
        if "사모사채" in category_compact or category_compact == "사채":
            return "사모사채"
        if "일반차입" in category_compact:
            return "일반차입금"

    text = " ".join([
        str((row or {}).get("contract_kind") or ""),
        str((row or {}).get("investor_name") or ""),
        str((row or {}).get("contract_name") or ""),
    ])
    compact = text.replace(" ", "")
    if "금융기관" in compact or "은행" in compact or "저축" in compact or "캐피탈" in compact or "카드" in compact:
        return "금융기관"
    if "사모사채" in compact or "사채" in compact:
        return "사모사채"
    if "일반차입" in compact:
        return "일반차입금"
    category = str((row or {}).get("category") or "").strip()
    return category or "일반차입금"


def _borrowing_excel_number(value):
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return 0
    if abs(number - round(number)) < 0.000001:
        return int(round(number))
    return number


def _borrowing_excel_rate(value):
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return 0
    if number > 1:
        return number / 100
    return number


def _borrowing_excel_text(value, fallback=""):
    text = str(value or "").strip()
    return text if text else fallback


def _borrowing_excel_date(value):
    text = str(value or "").strip()
    if not text:
        return ""
    compact = re.sub(r"\D", "", text)
    if len(compact) == 8:
        return f"{compact[:4]}.{compact[4:6]}.{compact[6:8]}"
    match = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text)
    if match:
        return f"{match.group(1)}.{int(match.group(2)):02d}.{int(match.group(3)):02d}"
    return text


def _borrowing_excel_sort_date(value):
    text = _borrowing_excel_date(value)
    compact = re.sub(r"\D", "", text)
    return compact if len(compact) == 8 else "99999999"


def _borrowing_effective_maturity_value(row):
    raw_value = (row or {}).get("maturity_date")
    maturity_date = _borrowing_parse_date(raw_value)
    if maturity_date and _borrowing_auto_schedule_required(row):
        return _borrowing_previous_business_day(maturity_date).strftime("%Y-%m-%d")
    return raw_value


def _borrowing_period_end_label(period):
    import calendar
    period_key = normalize_period_key(period) or str(period or "").strip()
    match = re.search(r"(\d{4})-(\d{2})", period_key)
    if not match:
        return period_key.replace("-", ".") if period_key else ""
    year = int(match.group(1))
    month = int(match.group(2))
    last_day = calendar.monthrange(year, month)[1]
    return f"{year}.{month:02d}.{last_day:02d}"


def _borrowing_financial_kind(row):
    kind = _borrowing_excel_text((row or {}).get("contract_kind"))
    if kind:
        return kind
    name = str((row or {}).get("investor_name") or "")
    if "저축" in name:
        return "저축은행"
    if "은행" in name:
        return "은행"
    if "캐피탈" in name:
        return "캐피탈"
    if "카드" in name:
        return "카드"
    if "대부" in name:
        return "대부"
    return "금융기관"


def _borrowing_contract_display_name(row):
    investor = _borrowing_excel_text((row or {}).get("investor_name"))
    contract = _borrowing_excel_text((row or {}).get("contract_name"))
    return " ".join(part for part in (investor, contract) if part) or "-"


def _borrowing_weighted_rate(rows):
    weighted = 0.0
    total = 0.0
    rates = []
    for row in rows or []:
        rate = _borrowing_excel_rate(row.get("interest_rate"))
        balance = float(row.get("borrow_balance") or 0)
        if rate:
            rates.append(rate)
        if rate and balance:
            weighted += rate * balance
            total += balance
    if total:
        return weighted / total
    return sum(rates) / len(rates) if rates else 0


def _borrowing_pledge_ratio(row):
    for key in ("pledge_ratio", "collateral_ratio", "collateral_rate"):
        if key in (row or {}) and row.get(key) not in (None, ""):
            value = row.get(key)
            try:
                number = float(value)
                return number / 100 if number > 10 else number
            except (TypeError, ValueError):
                return str(value).strip()
    collateral = str((row or {}).get("collateral_type") or "").strip()
    if "신용" in collateral:
        return "신용"
    return ""


def _borrowing_rows_for_statement(rows):
    normalized = _normalize_borrowing_contract_rows(rows or [])
    return sorted(
        [row for row in normalized if float(row.get("borrow_balance") or 0) > 0],
        key=lambda row: (
            _borrowing_excel_sort_date(row.get("contract_date")),
            str(row.get("investor_name") or ""),
            str(row.get("contract_name") or ""),
        ),
    )


def _build_borrowing_statement_workbook(rows, period):
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    label = _borrowing_period_end_label(period)
    sheet_suffix = (label[:7] if label else "차입현황").replace("-", ".")
    ws.title = f"차입현황 {sheet_suffix}"[:31]

    widths = [28.875, 14.25, 10.625, 13, 13, 12.75, 10.625, 15.625, 13, 13, 10.625]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    font_name = "맑은 고딕"
    title_font = Font(name=font_name, bold=True, size=13)
    header_font = Font(name=font_name, bold=True, size=10)
    normal_font = Font(name=font_name, size=10)
    total_font = Font(name=font_name, bold=True, size=10)
    header_fill = PatternFill("solid", fgColor="BFBFBF")
    section_fill = PatternFill("solid", fgColor="D9EAF7")
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    def style_row(row_idx, fill=None, font=None):
        for col in range(1, 12):
            cell = ws.cell(row=row_idx, column=col)
            cell.border = border
            cell.font = font or normal_font
            cell.alignment = center
            if fill:
                cell.fill = fill

    def write_headers(row_idx, headers):
        for col, value in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
    ws["A1"] = f"■ 차입금명세 (기준일 {label})" if label else "■ 차입금명세"
    ws["A1"].font = title_font
    ws["A1"].alignment = left

    headers = ["차입처", "금융기관", "구분", "차입일시", "만기일시", "상환방식", "금리", "차입원금", "차입잔액", "담보종류", "약정담보비율"]
    write_headers(3, headers)

    financial_rows = [row for row in rows if _borrowing_category_label(row) == "금융기관"]
    other_rows = [row for row in rows if _borrowing_category_label(row) in ("사모사채", "일반차입금")]

    row_idx = 4
    for item in financial_rows:
        values = [
            _borrowing_contract_display_name(item),
            _borrowing_financial_kind(item),
            "원화",
            _borrowing_excel_date(item.get("contract_date")),
            _borrowing_excel_date(_borrowing_effective_maturity_value(item)),
            _borrowing_excel_text(item.get("repayment_method"), "상환스케줄"),
            _borrowing_excel_rate(item.get("interest_rate")),
            _borrowing_excel_number(item.get("borrow_amount")),
            _borrowing_excel_number(item.get("borrow_balance")),
            _borrowing_excel_text(item.get("collateral_type"), "-"),
            _borrowing_pledge_ratio(item),
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.fill = white_fill
            cell.font = normal_font
            cell.alignment = right if col in (7, 8, 9, 11) and isinstance(value, (int, float)) else center
            cell.border = border
        ws.cell(row=row_idx, column=7).number_format = "0.00%"
        for col in (8, 9):
            ws.cell(row=row_idx, column=col).number_format = "#,##0"
        if isinstance(ws.cell(row=row_idx, column=11).value, (int, float)):
            ws.cell(row=row_idx, column=11).number_format = "0%"
        row_idx += 1

    if not financial_rows:
        ws.cell(row=row_idx, column=1, value="금융기관 잔액 데이터 없음")
        style_row(row_idx, white_fill, normal_font)
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=11)
        row_idx += 1

    financial_total_row = row_idx
    style_row(financial_total_row, section_fill, total_font)
    ws.cell(row=financial_total_row, column=6, value="가중평균금리")
    ws.cell(row=financial_total_row, column=7, value=_borrowing_weighted_rate(financial_rows))
    ws.cell(row=financial_total_row, column=8, value="소계")
    ws.cell(row=financial_total_row, column=9, value=sum(float(row.get("borrow_balance") or 0) for row in financial_rows))
    ws.cell(row=financial_total_row, column=7).number_format = "0.00%"
    ws.cell(row=financial_total_row, column=9).number_format = "#,##0"
    row_idx += 2

    other_header_row = row_idx
    other_headers = ["차입처", "", "구분", "발행일", "만기일", "상환방식", "금리", "차입원금", "차입잔액", "담보종류", "약정담보비율"]
    write_headers(other_header_row, other_headers)
    row_idx += 1

    grouped = {}
    for item in other_rows:
        category = _borrowing_category_label(item)
        investor = _borrowing_excel_text(item.get("investor_name"), "미지정")
        key = (category, investor)
        if key not in grouped:
            grouped[key] = {
                "category": category,
                "investor": investor,
                "rows": [],
                "borrow_amount": 0.0,
                "borrow_balance": 0.0,
            }
        grouped[key]["rows"].append(item)
        grouped[key]["borrow_amount"] += float(item.get("borrow_amount") or 0)
        grouped[key]["borrow_balance"] += float(item.get("borrow_balance") or 0)

    def group_sort_key(group):
        first_date = min((_borrowing_excel_sort_date(row.get("contract_date")) for row in group["rows"]), default="99999999")
        category_order = 0 if group["category"] == "사모사채" else 1
        major_order = 0 if group["investor"] == "김진규" else 1
        return (category_order, major_order, first_date, group["investor"])

    for group in sorted(grouped.values(), key=group_sort_key):
        group_rows = group["rows"]
        contract_dates = [_borrowing_excel_date(row.get("contract_date")) for row in group_rows if _borrowing_excel_date(row.get("contract_date"))]
        maturity_dates = [
            _borrowing_excel_date(_borrowing_effective_maturity_value(row))
            for row in group_rows
            if _borrowing_excel_date(_borrowing_effective_maturity_value(row))
        ]
        collateral_types = sorted({_borrowing_excel_text(row.get("collateral_type")) for row in group_rows if _borrowing_excel_text(row.get("collateral_type"))})
        sub_type = f"{group['category']}-대주주" if group["investor"] == "김진규" else group["category"]
        values = [
            group["investor"],
            sub_type,
            "원화",
            min(contract_dates) if contract_dates else "",
            max(maturity_dates) if maturity_dates else "",
            "만기일시상환",
            _borrowing_weighted_rate(group_rows),
            _borrowing_excel_number(group["borrow_amount"]),
            _borrowing_excel_number(group["borrow_balance"]),
            ", ".join(collateral_types) or "무담보",
            "",
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.fill = total_fill if group["investor"] == "김진규" else white_fill
            cell.font = total_font if group["investor"] == "김진규" else normal_font
            cell.alignment = right if col in (7, 8, 9) else center
            cell.border = border
        ws.cell(row=row_idx, column=7).number_format = "0.00%"
        for col in (8, 9):
            ws.cell(row=row_idx, column=col).number_format = "#,##0"
        row_idx += 1

    if not grouped:
        ws.cell(row=row_idx, column=1, value="사모사채/일반차입금 잔액 데이터 없음")
        style_row(row_idx, white_fill, normal_font)
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=11)
        row_idx += 1

    other_total_row = row_idx
    style_row(other_total_row, section_fill, total_font)
    ws.cell(row=other_total_row, column=6, value="가중평균금리")
    ws.cell(row=other_total_row, column=7, value=_borrowing_weighted_rate(other_rows))
    ws.cell(row=other_total_row, column=8, value="소계")
    ws.cell(row=other_total_row, column=9, value=sum(float(row.get("borrow_balance") or 0) for row in other_rows))
    ws.cell(row=other_total_row, column=7).number_format = "0.00%"
    ws.cell(row=other_total_row, column=9).number_format = "#,##0"
    row_idx += 1

    total_row = row_idx
    style_row(total_row, total_fill, total_font)
    ws.cell(row=total_row, column=8, value="총계")
    ws.cell(row=total_row, column=9, value=sum(float(row.get("borrow_balance") or 0) for row in rows))
    ws.cell(row=total_row, column=9).number_format = "#,##0"

    for row in ws.iter_rows(min_row=1, max_row=total_row, min_col=1, max_col=11):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value is None:
                cell.value = ""
            if cell.row not in (1,):
                cell.border = border

    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    return wb


@app.get("/api/download/borrowing-statement")
async def download_borrowing_statement(period: str = ""):
    from urllib.parse import quote

    try:
        data = uploaded_data.get("borrowing_contract_data")
        if not data:
            raise HTTPException(status_code=404, detail="차입 계약리스트 업로드 데이터가 없습니다.")

        snapshot = _build_borrowing_contract_snapshot(data.get("rows") or [], data.get("period") or "", data.get("source_file") or "")
        period_key = normalize_period_key(period) or str(period or "").strip() or snapshot.get("period") or ""
        rows = [
            row for row in (snapshot.get("rows") or [])
            if not period_key or str(row.get("period") or "").strip() == period_key
        ]
        statement_rows = _borrowing_rows_for_statement(rows)
        if not statement_rows:
            raise HTTPException(status_code=404, detail="차입금명세로 다운로드할 잔액 데이터가 없습니다.")

        wb = _build_borrowing_statement_workbook(statement_rows, period_key)
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        filename_period = (period_key or datetime.now().strftime("%Y-%m")).replace("-", "")
        filename = f"기업여신자료_차입금명세_{filename_period}.xlsx"
        encoded_filename = quote(filename)
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"},
        )
    except HTTPException:
        raise
    except Exception as exc:
        import traceback
        print("borrowing statement download failed", traceback.format_exc(), flush=True)
        raise HTTPException(status_code=500, detail=f"차입금명세 다운로드 처리 오류: {exc}") from exc

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
        validation = validate_payment_upload(tmp_path, data)
        if not validation.get("ok"):
            first_validation = validation
            data = parse_payment(tmp_path, product_groups, read_only=False)
            validation = validate_payment_upload(tmp_path, data)
            if validation.get("ok"):
                validation["recovered_by"] = "normal_workbook_reader"
                validation["first_attempt"] = first_validation
            else:
                validation["first_attempt"] = first_validation
        if not validation.get("ok"):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            return JSONResponse({
                "success": False,
                "code": "PAYMENT_VALIDATION_MISMATCH",
                "message": "입금명세 검수 실패: 원본 엑셀 합계와 저장 합계가 다릅니다.",
                "validation": validation,
            }, status_code=409)



        if period:



            data['upload_period'] = period



        existing_payment_data = uploaded_data.get('payment_data') or {}
        incoming_periods = set(data.get('periods') or [])
        existing_raw_rows = existing_payment_data.get('raw_rows') or []
        migrated_counts = {}
        if isinstance(existing_raw_rows, list) and existing_raw_rows:
            migrated_counts = _save_payment_raw_rows(existing_raw_rows)
        raw_row_counts = _save_payment_raw_rows(data.get('raw_rows') or [], incoming_periods)



        merge_source = {k: v for k, v in data.items() if k != 'raw_rows'}
        existing_for_merge = {k: v for k, v in existing_payment_data.items() if k != 'raw_rows'}
        merged_payment_data = _merge_period_values(existing_for_merge, merge_source) if existing_for_merge else merge_source
        merged_payment_data.pop('raw_rows', None)



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
            "raw_rows_saved": raw_row_counts,
            "raw_rows_migrated": migrated_counts,
            "validation": validation,



            "group_column": data.get("group_column"),



            "product_column": data.get("product_column"),



        })



    except Exception as e:



        import traceback



        raise HTTPException(status_code=500, detail=f"파일 처리 오류: {str(e)}\n{traceback.format_exc()}")











@app.get("/api/data/payment")



async def get_payment_data(period: str = ""):



    """입금명세 집계 데이터 반환 (원금회수액 / 이자회수액)"""



    d = uploaded_data.get("payment_data")



    if not d:



        return JSONResponse({"error": "데이터 없음"}, status_code=404)



    response = copy.deepcopy(d)
    period_key = normalize_period_key(period) or str(period or "").strip()
    periods = [period_key] if period_key else (response.get("periods") or [])
    if period_key:
        response["periods"] = [period_key] if period_key in (d.get("periods") or [period_key]) else []
    legacy_rows = response.get("raw_rows")
    if isinstance(legacy_rows, list) and legacy_rows:
        if period_key:
            response["raw_rows"] = [
                row for row in legacy_rows
                if isinstance(row, dict) and (_business_norm_period(row.get("period")) or str(row.get("period") or "").strip()) == period_key
            ]
    else:
        response["raw_rows"] = _load_payment_raw_rows(periods)

    return JSONResponse(response)











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











def _merge_contract_deal_into_asset_data(asset_data: dict) -> dict:
    """계약리스트 기반 대출취급액을 대출자산속성 데이터에 병합한다."""
    if not isinstance(asset_data, dict):
        return asset_data
    contract = uploaded_data.get("contract_data") or {}
    section_deal = contract.get("section_deal") or {}
    if not isinstance(section_deal, dict) or not section_deal:
        return asset_data

    sections = asset_data.setdefault("sections", {})
    deal_section = sections.get("대출취급액")
    if not isinstance(deal_section, dict):
        deal_section = {}

    for row_label, ym_vals in section_deal.items():
        if not isinstance(ym_vals, dict):
            continue
        row_map = deal_section.setdefault(row_label, {})
        for ym, val in ym_vals.items():
            if val is None:
                continue
            raw_val = float(val)
            row_map[ym] = raw_val * 1_000_000 if abs(raw_val) < 1_000_000 else raw_val

    sections["대출취급액"] = deal_section
    periods = set(asset_data.get("periods") or [])
    periods.update(contract.get("periods") or [])
    asset_data["periods"] = sorted(periods)
    return asset_data


def _merge_source_views_into_asset_data(asset_data: dict) -> dict:
    """Merge contract/loan-count source documents into an in-memory response only."""
    if not isinstance(asset_data, dict):
        return asset_data
    sections = asset_data.setdefault("sections", {})
    periods = set(asset_data.get("periods") or [])

    contract = uploaded_data.get("contract_data") or {}
    balance = contract.get("section_balance") or {}
    if isinstance(balance, dict):
        target = {key: dict(value) for key, value in
                  (sections.get("대출채권잔액") or {}).items() if isinstance(value, dict)}
        for row_label, values in balance.items():
            if isinstance(values, dict):
                target.setdefault(row_label, {}).update(values)
        if target:
            sections["대출채권잔액"] = target

    rates = contract.get("avg_rate_deal") or {}
    if isinstance(rates, dict) and rates:
        rate_section = {key: dict(value) for key, value in
                        (sections.get("평균이자율") or {}).items() if isinstance(value, dict)}
        rate_row = rate_section.setdefault("취급대출평균", {})
        for ym, value in rates.items():
            if value is not None:
                rate_row[str(ym)] = float(value)
        sections["평균이자율"] = rate_section
    periods.update(contract.get("periods") or [])

    loan_count = uploaded_data.get("loan_count_data") or {}
    count_rows = loan_count.get("rows") or {}
    if isinstance(count_rows, dict) and count_rows:
        target = {key: dict(value) for key, value in
                  (sections.get("대출취급건수") or {}).items() if isinstance(value, dict)}
        for row_label, values in count_rows.items():
            if not isinstance(values, dict):
                continue
            row = target.setdefault(row_label, {})
            for ym, value in values.items():
                if value is not None:
                    row[str(ym)] = float(value)
        sections["대출취급건수"] = target
    periods.update(loan_count.get("periods") or [])

    asset_data["sections"] = sections
    asset_data["periods"] = sorted(periods)
    return asset_data


def _load_product_category_groups_for_asset() -> dict:
    """대출자산속성 상품구분/상세 계산에 사용할 상품구분 설정을 로드한다."""
    groups = uploaded_data.get("product_category_groups")
    if not groups:
        _path = _data_path("product_category_groups")
        if os.path.exists(_path):
            try:
                with open(_path, encoding="utf-8") as f:
                    groups = json.load(f)
                uploaded_data["product_category_groups"] = groups
            except Exception:
                groups = {}
    return groups if isinstance(groups, dict) else {}


def _product_group_items_for_asset(group: dict) -> list[str]:
    raw = group.get("items")
    if raw is None:
        raw = group.get("products")
    return [str(v).strip() for v in (raw or []) if str(v).strip()]


ASSET_DATA_MIN_PERIOD = REPORT_DATA_MIN_PERIOD


def _is_asset_period_key(value) -> bool:
    text = str(value or "")
    return len(text) == 7 and text[4] == "-" and text[:4].isdigit() and text[5:].isdigit()


def _prune_asset_data_periods(asset_data: dict) -> dict:
    """Keep loan-asset monthly data from the configured report start month onward."""
    if not isinstance(asset_data, dict):
        return asset_data

    asset_data["periods"] = sorted(
        p for p in (asset_data.get("periods") or [])
        if _is_asset_period_key(p) and str(p) >= ASSET_DATA_MIN_PERIOD
    )

    def prune(obj):
        if isinstance(obj, dict):
            for key in list(obj.keys()):
                if _is_asset_period_key(key) and str(key) < ASSET_DATA_MIN_PERIOD:
                    obj.pop(key, None)
                    continue
                prune(obj.get(key))
        elif isinstance(obj, list):
            for item in obj:
                prune(item)

    prune(asset_data.get("sections"))
    prune(asset_data.get("cb_products"))
    return asset_data


def _merge_settlement_loan_balance_into_asset_data(asset_data: dict, saq: dict) -> dict:
    """결산자료 행 기준으로 대출채권잔액의 그룹별 건수/잔액을 재계산한다."""
    if not isinstance(asset_data, dict) or not isinstance(saq, dict) or not saq:
        return asset_data

    groups = _load_product_category_groups_for_asset()
    group_rows = groups.get("a") or groups.get("상품구분") or []
    if not isinstance(group_rows, list):
        group_rows = []

    product_to_group: dict[str, str] = {}
    for group in group_rows:
        if not isinstance(group, dict):
            continue
        group_name = str(group.get("name") or "").strip()
        if not group_name:
            continue
        for product_name in _product_group_items_for_asset(group):
            product_to_group[product_name] = group_name

    target_groups = ("신용", "담보", "보증")
    sections = asset_data.setdefault("sections", {})
    loan_section = {
        key: dict(value)
        for key, value in (sections.get("대출채권잔액") or {}).items()
        if isinstance(value, dict)
    }
    periods = set(asset_data.get("periods") or [])

    for pkey, pdata in saq.items():
        ym = pkey[:7] if isinstance(pkey, str) and len(pkey) >= 7 and pkey[4] == "-" else None
        if not ym or not isinstance(pdata, dict):
            continue

        counts = {group_name: 0 for group_name in target_groups}
        balances = {group_name: 0.0 for group_name in target_groups}
        total_count = 0
        total_balance = 0.0
        saw_rows = False

        rows = pdata.get("loan_loss_rows") or []
        if isinstance(rows, list) and rows:
            for row in rows:
                if not isinstance(row, dict):
                    continue
                product_name = str(row.get("product") or "").strip()
                try:
                    balance = float(row.get("balance") or 0)
                except (TypeError, ValueError):
                    continue
                if balance <= 0:
                    continue

                saw_rows = True
                total_count += 1
                total_balance += balance

                group_name = product_to_group.get(product_name) or str(row.get("group") or "").strip()
                if group_name in balances:
                    counts[group_name] += 1
                    balances[group_name] += balance

        if not saw_rows:
            section4 = pdata.get("section4") or {}
            if isinstance(section4, dict):
                for product_name, bucket in section4.items():
                    if not isinstance(bucket, dict):
                        continue
                    try:
                        balance = float(bucket.get("융잔합계") or 0) * 1_000_000
                    except (TypeError, ValueError):
                        continue
                    if balance <= 0:
                        continue
                    total_balance += balance
                    group_name = product_to_group.get(str(product_name).strip())
                    if group_name in balances:
                        balances[group_name] += balance

            section1_total = (pdata.get("section1") or {}).get("융잔합계")
            if section1_total is not None:
                try:
                    total_balance = float(section1_total) * 1_000_000
                except (TypeError, ValueError):
                    pass

        if total_balance <= 0 and total_count <= 0:
            continue

        periods.add(ym)
        for group_name in target_groups:
            if saw_rows:
                loan_section.setdefault(f"{group_name}채권 건수", {})[ym] = counts[group_name]
            loan_section.setdefault(f"{group_name}채권 잔액", {})[ym] = balances[group_name]

        if saw_rows:
            loan_section.setdefault("전체채권 건수", {})[ym] = total_count
        loan_section.setdefault("전체채권 잔액", {})[ym] = total_balance

    sections["대출채권잔액"] = loan_section
    asset_data["sections"] = sections
    asset_data["periods"] = sorted(periods)
    return asset_data


def _apply_asset_manual_overrides(asset_data: dict) -> dict:
    """사용자가 확인한 단건 보정값을 조회 응답에 반영한다."""
    if not isinstance(asset_data, dict):
        return asset_data
    sections = asset_data.setdefault("sections", {})
    periods = set(asset_data.get("periods") or [])

    deal_section = sections.get("대출취급액")
    if not isinstance(deal_section, dict):
        deal_section = {}
    deal_section.setdefault("추가재대출", {})["2025-10"] = 1_224_000_000
    sections["대출취급액"] = deal_section
    periods.add("2025-10")

    asset_data["sections"] = sections
    asset_data["periods"] = sorted(periods)
    return asset_data


def _merge_product_category_sections_from_products(asset_data: dict) -> dict:
    """상품별 융잔에서 상품구분/상품구분상세 섹션을 재계산한다.

    결산자료 업로드 후 상품별 섹션만 갱신되는 기간도 화면 상단 분류표가
    비지 않도록 응답 직전에 파생 계산한다.
    """
    if not isinstance(asset_data, dict):
        return asset_data

    sections = asset_data.setdefault("sections", {})
    product_section = sections.get("상품별") or {}
    if not isinstance(product_section, dict) or not product_section:
        return asset_data

    groups = _load_product_category_groups_for_asset()
    if not groups:
        return asset_data

    imported_cutoff = "2026-04"
    periods = set(asset_data.get("periods") or [])
    for row in product_section.values():
        if isinstance(row, dict):
            periods.update(str(ym) for ym in row.keys())

    def _sum_groups(source_key: str, target_key: str) -> None:
        group_rows = groups.get(source_key) or groups.get(target_key) or []
        if not isinstance(group_rows, list):
            return

        existing = sections.get(target_key) or {}
        target_section: dict[str, dict] = {
            k: dict(v) for k, v in existing.items() if isinstance(v, dict)
        }

        for group in group_rows:
            if not isinstance(group, dict):
                continue
            name = str(group.get("name") or "").strip()
            if not name:
                continue
            row = target_section.setdefault(name, {})
            products = _product_group_items_for_asset(group)
            for ym in periods:
                found = False
                total = 0.0
                for product in products:
                    value = (product_section.get(product) or {}).get(ym)
                    if value is None:
                        continue
                    found = True
                    try:
                        total += float(value)
                    except (TypeError, ValueError):
                        pass
                if found:
                    # 2026.04까지는 대출자산속성-상품구분 원본 엑셀 이력값을 우선한다.
                    if ym <= imported_cutoff and row.get(ym) is not None:
                        continue
                    row[ym] = total

        total_row = target_section.setdefault("전체융잔 합계", {})
        source_total = product_section.get("전체융잔 합계") or product_section.get("융잔합계") or {}
        for ym in periods:
            if ym <= imported_cutoff and total_row.get(ym) is not None:
                continue
            if isinstance(source_total, dict) and source_total.get(ym) is not None:
                total_row[ym] = source_total.get(ym)
                continue
            found = False
            total = 0.0
            for product, row in product_section.items():
                if product in ("전체융잔 합계", "융잔합계") or not isinstance(row, dict):
                    continue
                value = row.get(ym)
                if value is None:
                    continue
                found = True
                try:
                    total += float(value)
                except (TypeError, ValueError):
                    pass
            if found:
                total_row[ym] = total

        sections[target_key] = target_section

    _sum_groups("a", "상품구분")
    _sum_groups("b", "상품구분상세")

    asset_data["sections"] = sections
    asset_data["periods"] = sorted(periods)
    return asset_data


def _load_cb_grade_config_for_asset() -> list[dict]:
    cfg = uploaded_data.get("cb_grade_config")
    if not cfg:
        _path = _data_path("cb_grade_config")
        if os.path.exists(_path):
            try:
                with open(_path, encoding="utf-8") as f:
                    cfg = json.load(f)
                uploaded_data["cb_grade_config"] = cfg
            except Exception:
                cfg = None
    if isinstance(cfg, list) and cfg:
        return cfg
    return [
        {"grade": "1분위", "lo": 0, "hi": 724},
        {"grade": "2분위", "lo": 725, "hi": 749},
        {"grade": "3분위", "lo": 750, "hi": 789},
        {"grade": "4분위", "lo": 790, "hi": 839},
        {"grade": "5분위", "lo": 835, "hi": 889},
        {"grade": "6분위", "lo": 890, "hi": 914},
        {"grade": "7분위", "lo": 910, "hi": 944},
        {"grade": "8분위", "lo": 945, "hi": 959},
        {"grade": "9분위", "lo": 955, "hi": 984},
        {"grade": "10분위", "lo": 985, "hi": 1000},
    ]


def _cb_product_grade_label(item: dict) -> str:
    grade = str(item.get("grade") or "")
    lo = item.get("lo")
    hi = item.get("hi")
    if grade and lo is not None and hi is not None:
        return f"{grade}({lo}~{hi})"
    return grade


def _grade_for_score_key(score_key, cb_cfg: list[dict]) -> str:
    text = str(score_key or "").strip()
    if text in ("", "None", "무등급", "무분위"):
        return "1분위"
    try:
        score = int(float(text))
    except (TypeError, ValueError):
        return "1분위"
    for item in cb_cfg:
        try:
            lo = int(item.get("lo", 0))
            hi = int(item.get("hi", 9999))
        except (TypeError, ValueError):
            continue
        if lo <= score <= hi:
            return str(item.get("grade") or "1분위")
    return "1분위"


def _merge_saq_cb_products_into_asset_data(asset_data: dict, saq: dict) -> dict:
    """결산자료의 상품+CB점수 집계가 있으면 상품별 CB분위에 병합한다."""
    if not isinstance(asset_data, dict) or not isinstance(saq, dict) or not saq:
        return asset_data

    cb_cfg = _load_cb_grade_config_for_asset()
    label_by_grade = {
        str(item.get("grade") or ""): _cb_product_grade_label(item)
        for item in cb_cfg
        if item.get("grade")
    }
    first_label = label_by_grade.get("1분위", "1분위")
    cb_products = {
        product: {grade: dict(values) for grade, values in (grades or {}).items() if isinstance(values, dict)}
        for product, grades in (asset_data.get("cb_products") or {}).items()
        if isinstance(grades, dict)
    }
    product_section = (asset_data.get("sections") or {}).get("상품별") or {}
    touched_periods: set[str] = set()

    for pkey, pdata in saq.items():
        ym = pkey[:7] if isinstance(pkey, str) and len(pkey) >= 7 and pkey[4] == "-" else None
        if not ym or not isinstance(pdata, dict):
            continue
        by_product = pdata.get("cb_by_product") or {}
        if not isinstance(by_product, dict) or not by_product:
            continue
        touched_periods.add(ym)

        for product, score_values in by_product.items():
            if not isinstance(score_values, dict):
                continue
            grades = cb_products.setdefault(str(product), {})
            for score_key, bal_m in score_values.items():
                if bal_m is None:
                    continue
                grade = _grade_for_score_key(score_key, cb_cfg)
                label = label_by_grade.get(grade, first_label if grade == "1분위" else grade)
                row = grades.setdefault(label, {})
                row[ym] = row.get(ym, 0.0) + float(bal_m) * 1_000_000

    if not touched_periods:
        return asset_data

    # 같은 기간 재업로드 시 오래된 상품별 CB 값을 남기지 않는다.
    for grades in cb_products.values():
        for row_key, row in list(grades.items()):
            if isinstance(row, dict):
                for ym in touched_periods:
                    if row_key not in ("융잔합계", "전체융잔 합계"):
                        row.pop(ym, None)

    for pkey, pdata in saq.items():
        ym = pkey[:7] if isinstance(pkey, str) and len(pkey) >= 7 and pkey[4] == "-" else None
        by_product = (pdata or {}).get("cb_by_product") if isinstance(pdata, dict) else None
        if not ym or not isinstance(by_product, dict) or not by_product:
            continue
        for product, score_values in by_product.items():
            if not isinstance(score_values, dict):
                continue
            grades = cb_products.setdefault(str(product), {})
            for score_key, bal_m in score_values.items():
                if bal_m is None:
                    continue
                grade = _grade_for_score_key(score_key, cb_cfg)
                label = label_by_grade.get(grade, first_label if grade == "1분위" else grade)
                row = grades.setdefault(label, {})
                row[ym] = row.get(ym, 0.0) + float(bal_m) * 1_000_000

            product_total = (product_section.get(str(product)) or {}).get(ym)
            if product_total is None:
                product_total = sum(
                    (row.get(ym) or 0.0)
                    for key, row in grades.items()
                    if key not in ("융잔합계", "전체융잔 합계") and isinstance(row, dict)
                )
            grades.setdefault("융잔합계", {})[ym] = float(product_total or 0.0)

    asset_data["cb_products"] = cb_products
    periods = set(asset_data.get("periods") or [])
    periods.update(touched_periods)
    asset_data["periods"] = sorted(periods)
    return asset_data


def _merge_cb_total_from_products(asset_data: dict) -> dict:
    """상품별 CB분위가 있는 월은 전체 CB분위의 빈 월을 상품별 합산으로 보완한다."""
    if not isinstance(asset_data, dict):
        return asset_data
    cb_products = asset_data.get("cb_products") or {}
    if not isinstance(cb_products, dict) or not cb_products:
        return asset_data

    def _simple_grade(label) -> str | None:
        raw = str(label or "").strip()
        if not raw or "합계" in raw:
            return None
        if "무등급" in raw or "무분위" in raw:
            return "1분위"
        m = re.search(r"(\d+)\s*분위", raw)
        if not m:
            return None
        return f"{int(m.group(1))}분위"

    sections = asset_data.setdefault("sections", {})
    cb_sec = {
        str(row): dict(values)
        for row, values in (sections.get("CB분위") or {}).items()
        if isinstance(values, dict)
    }
    existing_periods: set[str] = set()
    for row_name, row_values in cb_sec.items():
        if row_name == "전체융잔 합계":
            continue
        for ym, value in row_values.items():
            if value not in (None, ""):
                existing_periods.add(str(ym))

    summed: dict[str, dict[str, float]] = {}
    periods = set(asset_data.get("periods") or [])
    for grades in cb_products.values():
        if not isinstance(grades, dict):
            continue
        for row_label, row_values in grades.items():
            grade = _simple_grade(row_label)
            if not grade or not isinstance(row_values, dict):
                continue
            for ym, value in row_values.items():
                if value in (None, "") or str(ym) in existing_periods:
                    continue
                try:
                    amount = float(value)
                except (TypeError, ValueError):
                    continue
                periods.add(str(ym))
                summed.setdefault(grade, {})[str(ym)] = summed.setdefault(grade, {}).get(str(ym), 0.0) + amount

    if not summed:
        return asset_data

    for grade, row_values in summed.items():
        row = cb_sec.setdefault(grade, {})
        for ym, value in row_values.items():
            row[ym] = value

    total_row = cb_sec.setdefault("전체융잔 합계", {})
    fill_periods = sorted({ym for row in summed.values() for ym in row})
    for ym in fill_periods:
        total_row[ym] = sum((cb_sec.get(f"{idx}분위") or {}).get(ym) or 0.0 for idx in range(1, 11))

    sections["CB분위"] = cb_sec
    asset_data["sections"] = sections
    asset_data["periods"] = sorted(periods)
    return asset_data


@app.post("/api/upload/asset")



async def upload_asset(file: UploadFile = File(...)):



    import tempfile, os



    suffix = os.path.splitext(file.filename)[1] if file.filename else ".xlsx"



    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:



        tmp.write(await file.read())



        tmp_path = tmp.name



    try:



        data = parse_asset(tmp_path)
        data = _merge_contract_deal_into_asset_data(data)
        data = _merge_product_category_sections_from_products(data)
        data = _prune_asset_data_periods(data)



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

    if _asset_data_is_cleared_marker(d):
        d = None



    saq = _normalize_settlement_aq_store(uploaded_data.get("settlement_aq") or {})
    contract = uploaded_data.get("contract_data") or {}



    if not d and not saq and not contract:



        return JSONResponse({"error": "데이터 없음"}, status_code=404)



    if not d:
        d = _empty_asset_data_shell(
            _source_report_periods("settlement_aq", "contract_data", "loan_count_data")
        )







    import copy



    result = _merge_contract_deal_into_asset_data(copy.deepcopy(d))
    result = _merge_source_views_into_asset_data(result)







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
        result = _merge_product_category_sections_from_products(result)
        sections = result.get("sections", {})







        # ── settlement_aq BW/BX → sections['상품구분화해채권'] 동적 구성 ──



        # sections['상품구분화해채권'][item_key][ym] = 원단위  (item_key = '[회생]xxx' or '[신복]xxx')



        # sections['상품구분화해채권_상품별'][item_key][상품명][ym] = 원단위  (상품별 분해용)



        hwahae_sec: dict[str, dict[str, float]] = {
            key: dict(values)
            for key, values in (sections.get("상품구분화해채권") or {}).items()
        }



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



                    hwahae_sec[k].get(ym) or 0.0
                    for k in hwahae_sec
                    if k != "전체화해 합계"



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



        base_asset_sections = result.get("sections", {})



        gender_sec_out: dict[str, dict[str, float]] = {
            key: dict(values)
            for key, values in (base_asset_sections.get("성별") or {}).items()
        }



        age_sec_out: dict[str, dict[str, float]] = {
            key: dict(values)
            for key, values in (base_asset_sections.get("연령별") or {}).items()
        }



        job_sec_out: dict[str, dict[str, float]] = {
            key: dict(values)
            for key, values in (base_asset_sections.get("직업별") or {}).items()
        }



        region_sec_out: dict[str, dict[str, float]] = {
            key: dict(values)
            for key, values in (base_asset_sections.get("지역별") or {}).items()
        }
        for row_key in ASSET_REGION_ROWS:
            region_sec_out.setdefault(row_key, {})



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



            age_period_totals: dict[str, float] = {}

            for row_key, val in age_data.items():

                if val is None:

                    continue

                normalized_key = "20대" if str(row_key).strip().startswith("10") else row_key
                age_period_totals[normalized_key] = age_period_totals.get(normalized_key, 0.0) + float(val) * 1_000_000

            for row_key, amount in age_period_totals.items():

                if row_key not in age_sec_out:

                    age_sec_out[row_key] = {}

                age_sec_out[row_key][ym] = amount







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



            for raw_key, val in region_data.items():



                if val is None:



                    continue



                if raw_key == "전체융잔 합계":
                    region_sec_out.setdefault("전체융잔 합계", {})
                    region_sec_out["전체융잔 합계"][ym] = float(val) * 1_000_000
                    continue

                row_key = _asset_region_bucket(raw_key)

                if row_key not in region_sec_out:



                    region_sec_out[row_key] = {}



                prev = region_sec_out[row_key].get(ym)
                region_sec_out[row_key][ym] = (float(prev) if prev is not None else 0.0) + float(val) * 1_000_000







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



                    first_grade_key = next((item.get("grade") for item in cb_cfg if item.get("grade") == "1분위"), "1분위")



                    total_cb = 0.0



                    for score_str, bal_m in cb_raw_pdata.items():



                        try:



                            score = int(score_str)



                        except (ValueError, TypeError):



                            bal_won = float(bal_m) * 1_000_000
                            total_cb += bal_won
                            grade_bucket[first_grade_key] = grade_bucket.get(first_grade_key, 0.0) + bal_won
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



                            grade_bucket[first_grade_key] = grade_bucket.get(first_grade_key, 0.0) + bal_won







                    # 섹션 삽입 — 10분위→1분위 내림차순. 무분위는 1분위에 합산.



                    GRADE_ORDER_DESC = list(reversed([item["grade"] for item in cb_cfg if item.get("grade")]))



                    for grade_key in GRADE_ORDER_DESC:



                        bal_won = grade_bucket.get(grade_key, 0.0)



                        if grade_key not in cb_sec_out:



                            cb_sec_out[grade_key] = {}



                        cb_sec_out[grade_key][ym] = cb_sec_out[grade_key].get(ym, 0.0) + bal_won



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
        result = _merge_saq_cb_products_into_asset_data(result, saq)







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







    result = _merge_product_category_sections_from_products(result)
    result = _merge_settlement_loan_balance_into_asset_data(result, saq)
    result = _merge_saq_cb_products_into_asset_data(result, saq)
    result = _merge_cb_total_from_products(result)
    result = _apply_asset_manual_overrides(result)
    result = _prune_asset_data_periods(result)

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







        # asset_data의 대출취급건수는 GET /api/data/asset에서 파생한다.



        if False and ym:  # source ownership: loan_count_data only



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
