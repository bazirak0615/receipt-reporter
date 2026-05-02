"""
출장 영수증 보고서 생성기 — 메인 서버
FastAPI 로컬 서버 + 브라우저 UI
"""
import os
import sys

# 한글 Windows에서 EasyOCR 프로그레스 바 크래시 방지
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"
if sys.stdout.encoding != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import json
import uuid
import shutil
from pathlib import Path
from datetime import datetime

import uvicorn
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from modules.ocr_engine import OCREngine
from modules.parser import ReceiptParser
from modules.categorizer import calculate_summary, check_qualified_receipt, is_vat_deductible
from modules.report_generator import ExcelReportGenerator, WordReportGenerator, PDFReportGenerator

# 경로 설정
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "output"
SESSIONS_FILE = DATA_DIR / "sessions.json"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 설정 로드
with open(BASE_DIR / "config.json", "r", encoding="utf-8") as f:
    CONFIG = json.load(f)

MAX_FILE_SIZE = CONFIG["upload"]["max_file_size_mb"] * 1024 * 1024  # bytes

# 앱 초기화
app = FastAPI(title=CONFIG["app_name"])
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

# 엔진 초기화 (지연 로딩)
ocr_engine = None
parser = ReceiptParser()


# ===== 세션 영속화 (Issue #1 Major 수정) =====
def load_sessions() -> dict:
    """JSON 파일에서 세션 데이터 로드"""
    if SESSIONS_FILE.exists():
        try:
            with open(SESSIONS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}


def save_sessions():
    """세션 데이터를 JSON 파일에 저장"""
    with open(SESSIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(sessions, f, ensure_ascii=False, indent=2, default=str)


sessions = load_sessions()


def get_ocr_engine(languages=None):
    global ocr_engine
    langs = languages or CONFIG["ocr"]["languages"]
    if ocr_engine is None:
        ocr_engine = OCREngine(languages=langs, gpu=CONFIG["ocr"]["gpu"])
    else:
        ocr_engine.set_languages(langs)
    return ocr_engine


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {
        "request": request,
        "config": CONFIG,
        "sessions": sessions,
    })


# ===== 대시보드용 세션 목록 API (미구현 #2 수정) =====
@app.get("/api/sessions")
async def list_sessions():
    """이전 보고서 목록 반환 (대시보드용)"""
    session_list = []
    for sid, data in sorted(sessions.items(), key=lambda x: x[1].get("created_at", ""), reverse=True):
        ti = data.get("trip_info", {})
        session_list.append({
            "id": sid,
            "report_id": ti.get("report_id", ""),
            "employee_name": ti.get("employee_name", ""),
            "destination": ti.get("destination", ""),
            "start_date": ti.get("start_date", ""),
            "end_date": ti.get("end_date", ""),
            "trip_type": ti.get("trip_type", "domestic"),
            "receipt_count": len(data.get("receipts", [])),
            "created_at": data.get("created_at", ""),
            "status": "완료" if data.get("receipts") else "작성중",
        })
    return session_list


@app.get("/api/session/{session_id}")
async def get_session(session_id: str):
    """기존 세션 데이터 반환 (이어쓰기용)"""
    if session_id not in sessions:
        return JSONResponse({"error": "세션을 찾을 수 없습니다"}, status_code=404)
    return sessions[session_id]


@app.post("/api/session")
async def create_session(
    employee_name: str = Form(""),
    department: str = Form(""),
    position: str = Form(""),
    trip_type: str = Form("domestic"),
    destination: str = Form(""),
    start_date: str = Form(""),
    end_date: str = Form(""),
    purpose: str = Form(""),
    visit_company: str = Form(""),
    attendees: str = Form(""),
    result: str = Form(""),
    follow_up: str = Form(""),
    exchange_rate: float = Form(1.0),
    exchange_rates_json: str = Form("{}"),
    ocr_languages: str = Form("ko,en"),
):
    session_id = str(uuid.uuid4())[:8]
    report_id = f"BT-{datetime.now().strftime('%Y%m%d')}-{session_id[:4].upper()}"

    # 복수 환율 파싱
    try:
        exchange_rates = json.loads(exchange_rates_json)
    except (json.JSONDecodeError, TypeError):
        exchange_rates = {}

    sessions[session_id] = {
        "id": session_id,
        "report_id": report_id,
        "trip_info": {
            "report_id": report_id,
            "employee_name": employee_name,
            "department": department,
            "position": position,
            "trip_type": trip_type,
            "destination": destination,
            "start_date": start_date,
            "end_date": end_date,
            "purpose": purpose,
            "visit_company": visit_company,
            "attendees": attendees,
            "result": result,
            "follow_up": follow_up,
            "exchange_rate": exchange_rate,
            "exchange_rates": exchange_rates,
        },
        "ocr_languages": ocr_languages.split(","),
        "receipts": [],
        "created_at": datetime.now().isoformat(),
    }

    save_sessions()
    return {"session_id": session_id, "report_id": report_id}


# ===== 파일 크기 검증 추가 (Issue #2 Minor 수정) =====
@app.post("/api/upload/{session_id}")
async def upload_receipts(session_id: str, files: list[UploadFile] = File(...)):
    if session_id not in sessions:
        return JSONResponse({"error": "세션을 찾을 수 없습니다"}, status_code=404)

    session = sessions[session_id]
    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(exist_ok=True)

    uploaded = []
    skipped = []
    for file in files:
        ext = Path(file.filename).suffix.lower()
        if ext.lstrip(".") not in CONFIG["upload"]["allowed_extensions"]:
            skipped.append({"name": file.filename, "reason": "지원하지 않는 파일 형식"})
            continue

        content = await file.read()

        if len(content) > MAX_FILE_SIZE:
            skipped.append({"name": file.filename, "reason": f"파일 크기 초과 ({len(content) // 1024 // 1024}MB > {CONFIG['upload']['max_file_size_mb']}MB)"})
            continue

        file_id = str(uuid.uuid4())[:8]
        filename = f"{file_id}{ext}"
        filepath = session_dir / filename

        with open(filepath, "wb") as f:
            f.write(content)

        uploaded.append({
            "file_id": file_id,
            "original_name": file.filename,
            "saved_path": str(filepath),
        })

    return {"uploaded": len(uploaded), "skipped": skipped, "files": uploaded}


@app.post("/api/ocr/{session_id}")
async def run_ocr(session_id: str):
    if session_id not in sessions:
        return JSONResponse({"error": "세션을 찾을 수 없습니다"}, status_code=404)

    session = sessions[session_id]
    session_dir = UPLOAD_DIR / session_id

    if not session_dir.exists():
        return JSONResponse({"error": "업로드된 파일이 없습니다"}, status_code=400)

    engine = get_ocr_engine(session.get("ocr_languages"))
    trip_type = session["trip_info"].get("trip_type", "domestic")
    default_currency = "KRW" if trip_type == "domestic" else "USD"
    receipt_parser = ReceiptParser(default_currency=default_currency)

    receipts = []
    image_files = sorted(session_dir.glob("*"))
    valid_files = [f for f in image_files if f.suffix.lower().lstrip(".") in CONFIG["upload"]["allowed_extensions"]]

    for idx, img_path in enumerate(valid_files):
        try:
            ocr_result = engine.extract_from_image(str(img_path))
            parsed = receipt_parser.parse(ocr_result)
            parsed["file_id"] = img_path.stem
            parsed["image_filename"] = img_path.name

            qual = check_qualified_receipt(parsed)
            parsed["is_qualified"] = qual["is_qualified"]
            parsed["tax_risk"] = qual["risk_level"]
            parsed["tax_message"] = qual["message"]
            parsed["vat_deductible"] = is_vat_deductible(parsed, trip_type)

            conf = parsed.get("confidence", 0)
            parsed["confidence_level"] = "high" if conf >= 0.8 else "medium" if conf >= 0.6 else "low"

            receipts.append(parsed)
        except Exception as e:
            receipts.append({
                "file_id": img_path.stem,
                "image_filename": img_path.name,
                "error": str(e),
                "confidence_level": "low",
            })

    session["receipts"] = receipts
    save_sessions()
    return {"total": len(receipts), "receipts": receipts}


@app.put("/api/receipt/{session_id}/{index}")
async def update_receipt(session_id: str, index: int, request: Request):
    if session_id not in sessions:
        return JSONResponse({"error": "세션을 찾을 수 없습니다"}, status_code=404)

    data = await request.json()
    receipts = sessions[session_id]["receipts"]

    if 0 <= index < len(receipts):
        receipts[index].update(data)
        save_sessions()
        return {"updated": True}

    return JSONResponse({"error": "잘못된 인덱스"}, status_code=400)


@app.get("/api/summary/{session_id}")
async def get_summary(session_id: str):
    if session_id not in sessions:
        return JSONResponse({"error": "세션을 찾을 수 없습니다"}, status_code=404)

    session = sessions[session_id]
    trip_info = session["trip_info"]
    all_receipts = session["receipts"]
    active_receipts = [r for r in all_receipts if not r.get("excluded")]

    summary = calculate_summary(
        active_receipts,
        trip_info.get("trip_type", "domestic"),
        trip_info.get("exchange_rate", 1.0),
        trip_info.get("exchange_rates", {}),
    )

    return {
        "trip_info": trip_info,
        "summary": summary,
        "receipts": all_receipts,
        "active_count": len(active_receipts),
        "excluded_count": len(all_receipts) - len(active_receipts),
    }


@app.post("/api/report/{session_id}")
async def generate_report(session_id: str, request: Request):
    if session_id not in sessions:
        return JSONResponse({"error": "세션을 찾을 수 없습니다"}, status_code=404)

    body = await request.json()
    fmt = body.get("format", "excel")

    session = sessions[session_id]
    trip_info = session["trip_info"]
    receipts = [r for r in session["receipts"] if not r.get("excluded")]
    report_id = trip_info.get("report_id", session_id)

    output_dir = OUTPUT_DIR / session_id
    output_dir.mkdir(exist_ok=True)

    if fmt == "excel":
        path = str(output_dir / f"{report_id}.xlsx")
        ExcelReportGenerator().generate(trip_info, receipts, path)
    elif fmt == "word":
        path = str(output_dir / f"{report_id}.docx")
        WordReportGenerator().generate(trip_info, receipts, path)
    elif fmt == "pdf":
        path = str(output_dir / f"{report_id}.pdf")
        PDFReportGenerator().generate(trip_info, receipts, path)
    else:
        return JSONResponse({"error": "지원하지 않는 형식"}, status_code=400)

    return {"file": path, "filename": Path(path).name}


@app.get("/api/download/{session_id}/{filename}")
async def download_file(session_id: str, filename: str):
    filepath = OUTPUT_DIR / session_id / filename
    if not filepath.exists():
        return JSONResponse({"error": "파일을 찾을 수 없습니다"}, status_code=404)
    return FileResponse(str(filepath), filename=filename)


@app.get("/api/image/{session_id}/{filename}")
async def get_image(session_id: str, filename: str):
    filepath = UPLOAD_DIR / session_id / filename
    if not filepath.exists():
        return JSONResponse({"error": "이미지를 찾을 수 없습니다"}, status_code=404)
    return FileResponse(str(filepath))


if __name__ == "__main__":
    print(f"\n  {CONFIG['app_name']} v{CONFIG['version']}")
    print(f"  http://{CONFIG['server']['host']}:{CONFIG['server']['port']}\n")
    uvicorn.run(app, host=CONFIG["server"]["host"], port=CONFIG["server"]["port"])
