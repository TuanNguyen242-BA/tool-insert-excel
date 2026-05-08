import importlib.util
import json
import os
import re
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from cloud_run_executor import execute_worker_job
from generation_service import (
    GenerationConfigError,
    generate_docx_from_payload,
    payload_from_form,
    validate_payload,
)
from job_store import get_job_store
from storage_backend import (
    GCSStorage,
    cloud_mode,
    job_manifest_object,
    job_output_object,
    session_manifest_object,
)


BASE_DIR = Path(__file__).resolve().parent
SESSIONS_DIR = Path(os.getenv("DOCX_BUILDER_SESSIONS", tempfile.gettempdir())) / "docx_builder_web"
SESSIONS_DIR.mkdir(parents=True, exist_ok=True)


def _load_builder_module():
    source_path = BASE_DIR / "docx_template_builder_ver0.9.2.py"
    spec = importlib.util.spec_from_file_location("docx_template_builder_core", source_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


builder = _load_builder_module()
DocxModel = builder.DocxModel
parse_excel_column_spec = builder.parse_excel_column_spec
w = builder.w

app = FastAPI(title="DOCX Template Builder Web")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def render_template(request: Request, name: str, context: dict | None = None):
    """Render template theo signature mới của Starlette/Jinja2Templates."""
    data = dict(context or {})
    data["request"] = request
    return templates.TemplateResponse(request=request, name=name, context=data)


def safe_filename(name: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9_.() \-\[\]\u00C0-\u1EF9]+", "_", name or "file")
    return clean[:180] or "file"


def session_dir(session_id: str) -> Path:
    if not re.fullmatch(r"[a-f0-9-]{36}", session_id):
        raise HTTPException(status_code=400, detail="Session không hợp lệ")
    path = SESSIONS_DIR / session_id
    if not path.exists():
        raise HTTPException(status_code=404, detail="Session đã hết hạn hoặc không tồn tại")
    return path


def save_upload(upload: UploadFile, dst_dir: Path) -> Path:
    dst = dst_dir / safe_filename(upload.filename)
    with dst.open("wb") as f:
        shutil.copyfileobj(upload.file, f)
    return dst


def upload_cloud_session(session_id: str, docx_path: Path, excel_path: Path):
    storage = GCSStorage()
    docx_object = f"sessions/{session_id}/{docx_path.name}"
    excel_object = f"sessions/{session_id}/{excel_path.name}"
    storage.upload_file(
        docx_path,
        docx_object,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    storage.upload_file(
        excel_path,
        excel_object,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    manifest = {
        "session_id": session_id,
        "docx_name": docx_path.name,
        "excel_name": excel_path.name,
        "docx_object": docx_object,
        "excel_object": excel_object,
    }
    storage.upload_json(session_manifest_object(session_id), manifest)
    return manifest


def load_cloud_session(session_id: str):
    if not re.fullmatch(r"[a-f0-9-]{36}", session_id):
        raise HTTPException(status_code=400, detail="Session không hợp lệ")
    storage = GCSStorage()
    try:
        return storage.download_json(session_manifest_object(session_id))
    except Exception as exc:
        raise HTTPException(status_code=404, detail="Session cloud không tồn tại") from exc


def detect_header_row(ws):
    max_scan = min(ws.max_row or 1, 50)
    max_col = 1
    for (_r, _c), cell in getattr(ws, "_cells", {}).items():
        if _r > max_scan:
            continue
        v = cell.value
        if v is None or str(v).strip() == "":
            continue
        max_col = max(max_col, _c)
    max_col = max(max_col, ws.max_column or 1)

    best_row = 1
    best_score = -1
    for r in range(1, max_scan + 1):
        nonempty = 0
        text_cells = 0
        numeric_cells = 0
        total_text_len = 0
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or str(v).strip() == "":
                continue
            nonempty += 1
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                numeric_cells += 1
            else:
                text_cells += 1
                total_text_len += min(len(str(v).strip()), 40)
        if nonempty < 2:
            continue
        score = nonempty * 10 + text_cells * 6 + total_text_len - numeric_cells * 8
        if score > best_score:
            best_score = score
            best_row = r
    return best_row, max_col


def excel_metadata(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        preview_rows = list(ws.iter_rows(min_row=1, max_row=50, values_only=True))
        header_row, max_col = detect_header_row_from_values(preview_rows, ws.max_column or 1)
        header_values = preview_rows[header_row - 1] if header_row - 1 < len(preview_rows) else ()
        columns = []
        for c in range(1, max_col + 1):
            header = header_values[c - 1] if c - 1 < len(header_values) else None
            header_text = str(header).strip() if header is not None else ""
            if not header_text:
                continue
            columns.append({
                "index": c,
                "letter": get_column_letter(c),
                "header": header_text,
            })
        sheets.append({
            "name": ws.title,
            "header_row": header_row,
            "max_row": ws.max_row or 1,
            "columns": columns,
        })
    wb.close()
    return sheets


def detect_header_row_from_values(rows, max_col_hint=1):
    max_col = max([len(row or ()) for row in rows] + [max_col_hint or 1])
    best_row = 1
    best_score = -1
    for idx, row in enumerate(rows, start=1):
        nonempty = 0
        text_cells = 0
        numeric_cells = 0
        total_text_len = 0
        for c in range(max_col):
            v = row[c] if row and c < len(row) else None
            if v is None or str(v).strip() == "":
                continue
            nonempty += 1
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                numeric_cells += 1
            else:
                text_cells += 1
                total_text_len += min(len(str(v).strip()), 40)
        if nonempty < 2:
            continue
        score = nonempty * 10 + text_cells * 6 + total_text_len - numeric_cells * 8
        if score > best_score:
            best_score = score
            best_row = idx
    return best_row, max_col


def load_docx_metadata(path: Path):
    model = DocxModel()
    model.load(str(path))
    intro = model.get_effective_intro_end_idx()
    all_headings = model.get_all_headings()
    headings = [h for h in model.get_all_headings() if h["original_idx"] >= intro]
    return {
        "model": model,
        "all_headings": all_headings,
        "headings": headings,
        "pages": model.get_total_pages(),
        "intro": intro,
        "doc_defaults": model.config.get("doc_defaults", {}),
        "sections": model.config.get("sections", []),
        "heading_styles": model.config.get("heading_styles", {}),
        "intro_text_blocks": model.get_intro_text_blocks(),
        "header_footer_text": model.get_header_footer_files(),
        "na_placeholder_heading_indexes": detect_na_placeholder_heading_indexes(model, headings),
    }


def detect_na_placeholder_heading_indexes(model, headings):
    """Heading có nội dung gốc chỉ là N/A thì mặc định nên để trống khi build template."""
    result = []
    body_elements = getattr(model, "body_elements", []) or []
    original_indexes = [h.get("original_idx") for h in headings]
    for i, heading in enumerate(headings):
        cur_idx = heading.get("original_idx")
        if cur_idx is None:
            continue
        next_idx = original_indexes[i + 1] if i + 1 < len(original_indexes) else len(body_elements)
        texts = []
        for k in range(cur_idx + 1, min(next_idx, len(body_elements))):
            text = "".join(t.text or "" for t in body_elements[k].iter(w("t"))).strip()
            if text:
                texts.append(text)
        if texts and all(re.sub(r"[\s./_-]+", "", text).upper() == "NA" for text in texts):
            result.append(cur_idx)
    return result


def parse_json_form(raw: str, default, field_name: str):
    if raw is None or str(raw).strip() == "":
        return default
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"{field_name} không phải JSON hợp lệ") from exc


def as_bool(value, default=False):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on", "y"}


def as_int(value, default=None):
    try:
        if value is None or str(value).strip() == "":
            return default
        return int(value)
    except (TypeError, ValueError):
        return default


def as_float(value, default=None):
    try:
        if value is None or str(value).strip() == "":
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def find_session_sources(sdir: Path):
    docx_files = [
        p for p in sdir.glob("*.docx")
        if not p.name.endswith("_web_template.docx") and not p.name.startswith("~$")
    ]
    excel_files = list(sdir.glob("*.xlsx")) + list(sdir.glob("*.xlsm"))
    if not docx_files or not excel_files:
        raise HTTPException(status_code=404, detail="Không tìm thấy file upload trong session")
    return sorted(docx_files)[0], sorted(excel_files)[0]


def apply_doc_defaults_config(model, raw_defaults):
    if not isinstance(raw_defaults, dict):
        return
    cfg = model.config.get("doc_defaults", {})
    text_keys = {"font", "lang", "line_spacing", "jc"}
    number_keys = {"size_pt", "space_before", "space_after", "indent_left_cm"}
    for key in text_keys:
        if key in raw_defaults:
            cfg[key] = str(raw_defaults.get(key) or "").strip()
    for key in number_keys:
        if key in raw_defaults:
            fallback = cfg.get(key, 0)
            cfg[key] = as_float(raw_defaults.get(key), fallback)
    model.config["doc_defaults"] = cfg


def apply_sections_config(model, raw_sections):
    if not isinstance(raw_sections, list):
        return
    sections = model.config.get("sections", [])
    numeric_keys = {
        "pg_w_cm", "pg_h_cm", "top_cm", "right_cm", "bottom_cm",
        "left_cm", "header_cm", "footer_cm", "gutter_cm",
    }
    for i, incoming in enumerate(raw_sections):
        if i >= len(sections) or not isinstance(incoming, dict):
            continue
        cfg = sections[i]
        for key in numeric_keys:
            if key in incoming:
                cfg[key] = as_float(incoming.get(key), cfg.get(key, 0))
        if "orient" in incoming:
            orient = str(incoming.get("orient") or "portrait").strip().lower()
            cfg["orient"] = "landscape" if orient == "landscape" else "portrait"
        if "titlePg" in incoming:
            cfg["titlePg"] = as_bool(incoming.get("titlePg"), cfg.get("titlePg", False))
    model.config["sections"] = sections


def apply_heading_styles_config(model, raw_styles):
    if not isinstance(raw_styles, dict):
        return
    styles = model.config.get("heading_styles", {})
    for style_id, incoming in raw_styles.items():
        if style_id not in styles or not isinstance(incoming, dict):
            continue
        cfg = styles[style_id]
        for key in ("font", "color"):
            if key in incoming:
                cfg[key] = str(incoming.get(key) or "").strip()
        for key in ("size_pt", "space_before_pt", "space_after_pt"):
            if key in incoming:
                cfg[key] = as_float(incoming.get(key), cfg.get(key, 0))
        for key in ("bold", "italic"):
            if key in incoming:
                cfg[key] = as_bool(incoming.get(key), cfg.get(key, False))
    model.config["heading_styles"] = styles


def build_headings_config(raw_headings, excel_path: Path):
    if not isinstance(raw_headings, list):
        raise HTTPException(status_code=400, detail="Danh sách heading không hợp lệ")

    result = []
    for item in raw_headings:
        if not isinstance(item, dict):
            continue
        level = as_int(item.get("level"), 1) or 1
        level = min(9, max(1, level))
        original_idx = as_int(item.get("original_idx"), None)
        action = str(item.get("action") or "keep").strip().lower()
        entry = {
            "level": level,
            "text": str(item.get("text") or "").strip(),
            "original_idx": original_idx,
            "keep_content": True,
            "auto_number": as_bool(item.get("auto_number"), False),
            "insert_source": None,
            "insert_selection": None,
        }

        if action == "empty":
            entry["keep_content"] = False
        elif action == "insert":
            selection = item.get("insert_selection") or {}
            if not isinstance(selection, dict):
                selection = {}
            mode = str(selection.get("mode") or "custom").strip().lower()
            if mode not in {"custom", "range", "sheet", "all"}:
                mode = "custom"

            insert_selection = {"mode": mode}
            if mode in {"custom", "range", "sheet"}:
                insert_selection["sheet"] = str(selection.get("sheet") or "").strip()
            if mode == "custom":
                cols = str(selection.get("columns") or "").strip()
                if not cols:
                    raise HTTPException(
                        status_code=400,
                        detail=f'Heading "{entry["text"]}" đang chèn Excel nhưng chưa chọn cột',
                    )
                if not parse_excel_column_spec(cols):
                    raise HTTPException(
                        status_code=400,
                        detail=f'Danh sách cột không hợp lệ ở heading "{entry["text"]}"',
                    )
                insert_selection.update({
                    "columns": cols,
                    "row_start": max(1, as_int(selection.get("row_start"), 1) or 1),
                    "row_end": as_int(selection.get("row_end"), None),
                })
            elif mode == "range":
                insert_selection["range"] = str(selection.get("range") or "").strip()

            entry["keep_content"] = False
            entry["insert_source"] = str(excel_path)
            entry["insert_selection"] = insert_selection

        result.append(entry)
    return result


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return render_template(request, "index.html")


@app.post("/configure", response_class=HTMLResponse)
async def configure(
    request: Request,
    docx_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
):
    if not docx_file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Vui lòng upload file .docx")
    if not excel_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Vui lòng upload file .xlsx/.xlsm")

    sid = str(uuid.uuid4())
    sdir = SESSIONS_DIR / sid
    sdir.mkdir(parents=True, exist_ok=True)

    docx_path = save_upload(docx_file, sdir)
    excel_path = save_upload(excel_file, sdir)
    meta = load_docx_metadata(docx_path)
    sheets = excel_metadata(excel_path)
    if cloud_mode():
        upload_cloud_session(sid, docx_path, excel_path)

    return render_template(
        request,
        "configure.html",
        {
            "session_id": sid,
            "docx_name": docx_path.name,
            "excel_name": excel_path.name,
            "all_headings": meta["all_headings"],
            "headings": meta["headings"],
            "sheets": sheets,
            "pages": meta["pages"],
            "intro": meta["intro"],
            "doc_defaults": meta["doc_defaults"],
            "sections": meta["sections"],
            "heading_styles": meta["heading_styles"],
            "intro_text_blocks": meta["intro_text_blocks"],
            "header_footer_text": meta["header_footer_text"],
            "na_placeholder_heading_indexes": meta["na_placeholder_heading_indexes"],
            "cloud_mode": cloud_mode(),
        },
    )


def submit_cloud_generation(request: Request, session_id: str, payload: dict):
    session_manifest = load_cloud_session(session_id)
    job_id = str(uuid.uuid4())
    output_name = f"{Path(session_manifest['docx_name']).stem}_web_template.docx"
    output_object = job_output_object(job_id, output_name)
    manifest = {
        "job_id": job_id,
        "session_id": session_id,
        "docx_name": session_manifest["docx_name"],
        "excel_name": session_manifest["excel_name"],
        "docx_object": session_manifest["docx_object"],
        "excel_object": session_manifest["excel_object"],
        "output_name": output_name,
        "output_object": output_object,
        "payload": payload,
    }

    storage = GCSStorage()
    storage.upload_json(job_manifest_object(job_id), manifest)

    store = get_job_store()
    store.create(
        job_id,
        {
            "status": "queued",
            "progress": 0,
            "message": "Đã tạo job",
            "session_id": session_id,
            "docx_name": session_manifest["docx_name"],
            "excel_name": session_manifest["excel_name"],
            "output_name": output_name,
            "output_object": output_object,
        },
    )
    try:
        operation = execute_worker_job(job_id)
        store.update(
            job_id,
            status="submitted",
            progress=1,
            message="Đã gửi job, Cloud Run đang cấp worker mới (thường 15-30 giây)",
            operation=operation.get("name", ""),
        )
    except Exception as exc:
        store.update(job_id, status="failed", error=str(exc), message="Không gọi được Cloud Run Job")

    return render_template(request, "job_status.html", {"job_id": job_id})


@app.post("/generate")
async def generate(
    request: Request,
    session_id: str = Form(...),
    intro_end_idx: str = Form(""),
    doc_defaults_json: str = Form("{}"),
    sections_json: str = Form("[]"),
    intro_replacements_json: str = Form("{}"),
    heading_styles_json: str = Form("{}"),
    header_footer_text_json: str = Form("{}"),
    headings_config_json: str = Form("[]"),
    preserve_inline_formatting: str = Form("true"),
    excel_font_name: str = Form(""),
    excel_font_size: float = Form(11),
):
    try:
        payload = payload_from_form(
            intro_end_idx=intro_end_idx,
            doc_defaults_json=doc_defaults_json,
            sections_json=sections_json,
            intro_replacements_json=intro_replacements_json,
            heading_styles_json=heading_styles_json,
            header_footer_text_json=header_footer_text_json,
            headings_config_json=headings_config_json,
            preserve_inline_formatting=preserve_inline_formatting,
            excel_font_name=excel_font_name,
            excel_font_size=excel_font_size,
        )
        validate_payload(payload)
    except GenerationConfigError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if cloud_mode():
        return submit_cloud_generation(request, session_id, payload)

    sdir = session_dir(session_id)
    docx_path, excel_path = find_session_sources(sdir)

    out_path = sdir / f"{docx_path.stem}_web_template.docx"
    generate_docx_from_payload(docx_path, excel_path, out_path, payload)
    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=out_path.name,
    )


@app.get("/jobs/{job_id}", response_class=HTMLResponse)
def job_page(request: Request, job_id: str):
    return render_template(request, "job_status.html", {"job_id": job_id})


@app.get("/jobs/{job_id}/status")
def job_status(job_id: str):
    record = get_job_store().get(job_id)
    if not record:
        raise HTTPException(status_code=404, detail="Job không tồn tại")
    return record


@app.get("/jobs/{job_id}/download")
def job_download(job_id: str):
    record = get_job_store().get(job_id)
    if not record:
        raise HTTPException(status_code=404, detail="Job không tồn tại")
    if record.get("status") != "done":
        raise HTTPException(status_code=409, detail="Job chưa hoàn tất")
    output_name = record.get("output_name") or f"{job_id}.docx"
    if cloud_mode():
        output_object = record.get("output_object")
        if not output_object:
            raise HTTPException(status_code=404, detail="Không tìm thấy output")
        dst = SESSIONS_DIR / "job_downloads" / job_id / safe_filename(output_name)
        GCSStorage().download_file(output_object, dst)
        return FileResponse(
            dst,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=output_name,
        )
    raise HTTPException(status_code=400, detail="Download job chỉ dùng ở cloud mode")


@app.get("/health")
def health():
    return {"ok": True}
