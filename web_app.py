import importlib.util
import json
import os
import re
import shutil
import tempfile
import unicodedata
import uuid
import zipfile
from pathlib import Path

from env_loader import load_dotenv_file
from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
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
from preview_service import PreviewConversionError, convert_docx_to_pdf
from storage_backend import (
    GCSStorage,
    cloud_mode,
    job_manifest_object,
    job_output_object,
    job_preview_object,
    session_manifest_object,
)
from ulnl_tool.core.classifier import Classifier, ClassificationRule, DEFAULT_RULES
from ulnl_tool.core.data_model import Project
from ulnl_tool.core.excel_io import (
    build_ulnl_file,
    enrich_with_classifier,
    import_functions_from_excel,
)
from ulnl_tool.core.libreoffice import xlsx_to_pdf
from data_zone_builder import (
    DEFAULT_OPTIONS as DATA_ZONE_DEFAULT_OPTIONS,
    build_raw_workbook,
    build_work_workbook,
    coerce_options as coerce_data_zone_options,
    list_sheet_names as data_zone_sheet_names,
    parse_source_workbook,
    preview_tables as data_zone_preview_tables,
)


load_dotenv_file()

BASE_DIR = Path(__file__).resolve().parent
SESSIONS_DIR = Path(os.getenv("DOCX_BUILDER_SESSIONS", tempfile.gettempdir())) / "docx_builder_web"
SESSIONS_DIR.mkdir(parents=True, exist_ok=True)
ULNL_BASE_DIR = BASE_DIR / "ulnl_tool"
ULNL_TEMPLATE_PATH = ULNL_BASE_DIR / "resources" / "ULNL_template_base.xlsx"
ULNL_SCRIPTS_DIR = ULNL_BASE_DIR / "scripts"


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


def _load_csdl_module():
    source_path = BASE_DIR / "tk_csdl_tool.py"
    spec = importlib.util.spec_from_file_location("tk_csdl_tool_core", source_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


csdl_tool = _load_csdl_module()

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
    if dst.exists():
        stem = dst.stem
        suffix = dst.suffix
        for i in range(2, 1000):
            candidate = dst_dir / f"{stem}_{i}{suffix}"
            if not candidate.exists():
                dst = candidate
                break
    with dst.open("wb") as f:
        shutil.copyfileobj(upload.file, f)
    return dst


def excel_file_id(path: Path) -> str:
    return path.name


def build_excel_sources(excel_paths):
    return [
        {
            "id": excel_file_id(path),
            "name": path.name,
            "path": str(path),
        }
        for path in excel_paths
    ]


def write_local_session_manifest(session_id: str, docx_path: Path, excel_paths):
    manifest = {
        "session_id": session_id,
        "docx_name": docx_path.name,
        "excel_name": ", ".join(path.name for path in excel_paths),
        "excel_files": build_excel_sources(excel_paths),
    }
    (docx_path.parent / "session.json").write_text(
        json.dumps(manifest, ensure_ascii=False),
        encoding="utf-8",
    )
    return manifest


def session_excel_payload(excel_paths):
    excel_sources = build_excel_sources(excel_paths)
    sheets = []
    for source in excel_sources:
        sheets.extend(excel_metadata(Path(source["path"]), source["id"], source["name"]))
    return excel_sources, sheets


def render_configure_session(request: Request, session_id: str, payload: dict | None = None):
    sdir = session_dir(session_id)
    docx_path, excel_paths = find_session_sources(sdir)
    meta = load_docx_metadata(docx_path)
    excel_sources, sheets = session_excel_payload(excel_paths)
    return render_template(
        request,
        "configure.html",
        {
            "session_id": session_id,
            "docx_name": docx_path.name,
            "excel_name": ", ".join(path.name for path in excel_paths),
            "excel_files": excel_sources,
            "all_headings": meta["all_headings"],
            "headings": meta["headings"],
            "sheets": sheets,
            "pages": meta["pages"],
            "intro": meta["intro"],
            "doc_defaults": meta["doc_defaults"],
            "sections": meta["sections"],
            "heading_styles": meta["heading_styles"],
            "toc_settings": meta["toc_settings"],
            "intro_text_blocks": meta["intro_text_blocks"],
            "header_footer_text": meta["header_footer_text"],
            "na_placeholder_heading_indexes": meta["na_placeholder_heading_indexes"],
            "cloud_mode": cloud_mode(),
            "saved_payload": payload or {},
        },
    )


def upload_cloud_session(session_id: str, docx_path: Path, excel_paths):
    storage = GCSStorage()
    docx_object = f"sessions/{session_id}/{docx_path.name}"
    excel_entries = []
    storage.upload_file(
        docx_path,
        docx_object,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    for excel_path in excel_paths:
        excel_object = f"sessions/{session_id}/excel/{excel_path.name}"
        storage.upload_file(
            excel_path,
            excel_object,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        excel_entries.append({
            "id": excel_file_id(excel_path),
            "name": excel_path.name,
            "object": excel_object,
        })
    manifest = {
        "session_id": session_id,
        "docx_name": docx_path.name,
        "excel_name": excel_entries[0]["name"] if excel_entries else "",
        "excel_files": excel_entries,
        "docx_object": docx_object,
        "excel_object": excel_entries[0]["object"] if excel_entries else "",
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
        values = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        score = score_header_row(values, r)
        if score is not None and score > best_score:
            best_score = score
            best_row = r
    return best_row, max_col


def normalize_header_text(value):
    text = str(value or "").strip().lower()
    text = "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )
    return re.sub(r"\s+", " ", text)


def score_header_row(values, row_index):
    nonempty = 0
    text_cells = 0
    numeric_cells = 0
    header_keyword_hits = 0
    short_text_cells = 0
    keywords = {
        "stt", "ma", "ten", "mo ta", "tan suat", "loai", "nguon", "dich",
        "bang", "cot", "field", "table", "column", "name", "description",
        "key", "type",
    }
    for v in values:
        if v is None or str(v).strip() == "":
            continue
        nonempty += 1
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            numeric_cells += 1
            continue
        text_cells += 1
        text = normalize_header_text(v)
        if len(text) <= 28:
            short_text_cells += 1
        if any(k in text for k in keywords):
            header_keyword_hits += 1
    if nonempty < 2:
        return None
    return (
        nonempty * 16
        + text_cells * 8
        + short_text_cells * 5
        + header_keyword_hits * 35
        - numeric_cells * 18
        - row_index * 2
    )


def excel_metadata(path: Path, file_id: str | None = None, file_name: str | None = None):
    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = []
    file_id = file_id or excel_file_id(path)
    file_name = file_name or path.name
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
            "file_id": file_id,
            "file_name": file_name,
            "name": ws.title,
            "display_name": f"{file_name} / {ws.title}",
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
        values = [
            row[c] if row and c < len(row) else None
            for c in range(max_col)
        ]
        score = score_header_row(values, idx)
        if score is not None and score > best_score:
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
        "toc_settings": model.config.get("toc_settings", {}),
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
        has_content = False
        for k in range(cur_idx + 1, min(next_idx, len(body_elements))):
            el = body_elements[k]
            text = "".join(t.text or "" for t in el.iter(w("t"))).strip()
            if text:
                has_content = True
                break
            for child in el.iter():
                local_name = child.tag.rsplit("}", 1)[-1] if isinstance(child.tag, str) else ""
                if local_name in {"tbl", "drawing", "pict", "object"}:
                    has_content = True
                    break
            if has_content:
                break
        if not has_content:
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
    manifest_path = sdir / "session.json"
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            docx_path = sdir / manifest["docx_name"]
            excel_paths = [
                sdir / item["name"]
                for item in manifest.get("excel_files", [])
                if item.get("name")
            ]
            if docx_path.exists() and all(path.exists() for path in excel_paths) and excel_paths:
                return docx_path, excel_paths
        except (OSError, KeyError, TypeError, json.JSONDecodeError):
            pass

    docx_files = [
        p for p in sdir.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".docx"
        and not p.name.lower().endswith("_web_template.docx")
        and not p.name.startswith("~$")
    ]
    excel_files = [
        p for p in sdir.iterdir()
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"} and not p.name.startswith("~$")
    ]
    if not docx_files or not excel_files:
        raise HTTPException(status_code=404, detail="Không tìm thấy file upload trong session")
    return sorted(docx_files)[0], sorted(excel_files)


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


def build_headings_config(raw_headings, excel_paths):
    if not isinstance(raw_headings, list):
        raise HTTPException(status_code=400, detail="Danh sách heading không hợp lệ")

    excel_sources = {
        str(source["id"]): source
        for source in build_excel_sources(excel_paths)
    }
    default_excel = next(iter(excel_sources.values()), None)

    def normalize_source_key(value):
        text = unicodedata.normalize("NFD", str(value or ""))
        text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
        return re.sub(r"\s+", " ", text).strip().lower()

    def resolve_excel_source(src, default_file_id=""):
        src = src or {}
        candidates = [
            src.get("file_id"),
            src.get("file_name"),
            src.get("excel_file_name"),
            Path(str(src.get("source_path") or "")).name if src.get("source_path") else "",
            default_file_id,
        ]
        for candidate in candidates:
            candidate = str(candidate or "").strip()
            if candidate and candidate in excel_sources:
                return candidate, excel_sources[candidate]

        source_by_basename = {
            Path(source.get("path") or "").name: (file_id, source)
            for file_id, source in excel_sources.items()
        }
        for candidate in candidates:
            candidate = str(candidate or "").strip()
            if candidate and candidate in source_by_basename:
                return source_by_basename[candidate]

        normalized_sources = {
            normalize_source_key(file_id): (file_id, source)
            for file_id, source in excel_sources.items()
        }
        normalized_sources.update({
            normalize_source_key(source.get("name") or ""): (file_id, source)
            for file_id, source in excel_sources.items()
        })
        normalized_sources.update({
            normalize_source_key(Path(source.get("path") or "").name): (file_id, source)
            for file_id, source in excel_sources.items()
        })
        for candidate in candidates:
            key = normalize_source_key(candidate)
            if key and key in normalized_sources:
                return normalized_sources[key]

        raw_path = Path(str(src.get("source_path") or ""))
        if raw_path.exists():
            return raw_path.name, {"id": raw_path.name, "name": raw_path.name, "path": str(raw_path)}
        return "", {}

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
            file_id = str(selection.get("file_id") or item.get("excel_file_id") or "").strip()
            if not file_id and default_excel:
                file_id = default_excel["id"]
            if file_id not in excel_sources:
                raise HTTPException(
                    status_code=400,
                    detail=f'Heading "{entry["text"]}" chon file Excel khong hop le',
                )
            mode = str(selection.get("mode") or "custom").strip().lower()
            if mode not in {"custom", "range", "sheet", "all"}:
                mode = "custom"

            insert_selection = {"mode": mode, "file_id": file_id}
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
                header_overrides = selection.get("header_overrides") or {}
                if isinstance(header_overrides, dict):
                    insert_selection["header_overrides"] = {
                        str(k): str(v)
                        for k, v in header_overrides.items()
                        if str(v).strip()
                    }
                advanced = selection.get("advanced") or {}
                if isinstance(advanced, dict):
                    split_column = as_int(advanced.get("split_column"), None)
                    merge_columns = []
                    for col in advanced.get("merge_columns") or []:
                        parsed = as_int(col, None)
                        if parsed:
                            merge_columns.append(parsed)
                    split_sources = []
                    skipped_sources = []
                    for src in advanced.get("split_sources") or []:
                        if not isinstance(src, dict):
                            continue
                        src_file_id, src_source = resolve_excel_source(src, file_id)
                        src_entry = {
                            "file_id": src_file_id,
                            "source_path": src_source.get("path", ""),
                            "sheet": str(src.get("sheet") or "").strip(),
                            "columns": str(src.get("columns") or cols).strip(),
                            "row_start": max(1, as_int(src.get("row_start"), insert_selection["row_start"]) or 1),
                            "row_end": as_int(src.get("row_end"), None),
                            "split_column": as_int(src.get("split_column"), split_column),
                        }
                        if src_entry["source_path"] and src_entry["sheet"] and src_entry["columns"]:
                            split_sources.append(src_entry)
                        else:
                            skipped_sources.append(src.get("file_name") or src.get("file_id") or src.get("sheet") or "(không rõ)")
                    if advanced.get("split_sources") and skipped_sources:
                        raise HTTPException(
                            status_code=400,
                            detail=(
                                f'Không map được nguồn phụ ở heading "{entry["text"]}": '
                                + ", ".join(str(x) for x in skipped_sources)
                            ),
                        )
                    split_text_rules = []
                    for rule in advanced.get("split_text_rules") or []:
                        if not isinstance(rule, dict):
                            continue
                        rule_file_id, rule_source = resolve_excel_source(rule, file_id)
                        rule_entry = {
                            "enabled": as_bool(rule.get("enabled"), True),
                            "file_id": rule_file_id,
                            "source_path": rule_source.get("path", ""),
                            "sheet": str(rule.get("sheet") or "").strip(),
                            "key_column": as_int(rule.get("key_column"), None),
                            "value_column": as_int(rule.get("value_column"), None),
                            "prefix": str(rule.get("prefix") or ""),
                            "prefix_bold": as_bool(rule.get("prefix_bold"), False),
                            "prefix_italic": as_bool(rule.get("prefix_italic"), False),
                            "prefix_underline": as_bool(rule.get("prefix_underline"), False),
                            "force_left": as_bool(rule.get("force_left"), True),
                            "row_start": max(1, as_int(rule.get("row_start"), 1) or 1),
                            "row_end": as_int(rule.get("row_end"), None),
                        }
                        if rule_entry["source_path"] and rule_entry["sheet"] and rule_entry["key_column"] and rule_entry["value_column"]:
                            split_text_rules.append(rule_entry)
                    insert_selection["advanced"] = {
                        "split_enabled": as_bool(advanced.get("split_enabled"), False),
                        "split_column": split_column,
                        "add_split_heading": as_bool(advanced.get("add_split_heading"), False),
                        "heading_prefix": str(advanced.get("heading_prefix", "Bảng ") or ""),
                        "heading_suffix": str(advanced.get("heading_suffix", "") or ""),
                        "heading_font": str(advanced.get("heading_font") or "").strip(),
                        "heading_size_pt": as_float(advanced.get("heading_size_pt"), None),
                        "heading_bold": as_bool(advanced.get("heading_bold"), False),
                        "heading_italic": as_bool(advanced.get("heading_italic"), False),
                        "heading_underline": as_bool(advanced.get("heading_underline"), False),
                        "heading_color": str(advanced.get("heading_color") or "").strip().lstrip("#"),
                        "auto_stt_enabled": as_bool(advanced.get("auto_stt_enabled"), False),
                        "internal_link_enabled": as_bool(advanced.get("internal_link_enabled"), False),
                        "internal_link_column": as_int(advanced.get("internal_link_column"), None),
                        "internal_link_target_heading_original_idx": as_int(
                            advanced.get("internal_link_target_heading_original_idx"), None
                        ),
                        "multi_source_mode": str(advanced.get("multi_source_mode") or "sequential").strip(),
                        "interleave_group_heading_enabled": as_bool(
                            advanced.get("interleave_group_heading_enabled"), True
                        ),
                        "interleave_group_heading_prefix": str(
                            advanced.get("interleave_group_heading_prefix", "Cụm ") or ""
                        ),
                        "interleave_group_heading_suffix": str(
                            advanced.get("interleave_group_heading_suffix", "") or ""
                        ),
                        "split_sources": split_sources,
                        "split_text_rules": split_text_rules,
                        "merge_columns": merge_columns,
                    }
            elif mode == "range":
                insert_selection["range"] = str(selection.get("range") or "").strip()

            entry["keep_content"] = False
            entry["insert_source"] = excel_sources[file_id]["path"]
            entry["insert_selection"] = insert_selection

        result.append(entry)
    return result


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return render_template(request, "index.html")


@app.get("/tktkl", response_class=HTMLResponse)
def tktkl_index(request: Request):
    return render_template(request, "tktkl.html")


@app.get("/csdl", response_class=HTMLResponse)
def csdl_index(request: Request):
    return render_template(
        request,
        "csdl_tool.html",
        {
            "cloud_mode": cloud_mode(),
            "default_config": csdl_tool.DEFAULT_CONFIG,
            "headers": csdl_tool.HEADERS,
        },
    )


@app.get("/data-zone", response_class=HTMLResponse)
def data_zone_index(request: Request):
    return render_template(
        request,
        "data_zone_tool.html",
        {
            "default_options": DATA_ZONE_DEFAULT_OPTIONS,
        },
    )


def _data_zone_load_options(options_json: str | None):
    if not options_json or not options_json.strip():
        return coerce_data_zone_options({})
    try:
        raw = json.loads(options_json)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Options JSON khong hop le: {exc.msg}") from exc
    if not isinstance(raw, dict):
        raise HTTPException(status_code=400, detail="Options JSON phai la object")
    return coerce_data_zone_options(raw)


@app.post("/data-zone/inspect")
async def data_zone_inspect(
    excel_file: UploadFile = File(...),
    sheet_name: str = Form(""),
    options_json: str = Form("{}"),
):
    if not excel_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Vui long upload file .xlsx/.xlsm")
    options = _data_zone_load_options(options_json)
    sid = str(uuid.uuid4())
    sdir = SESSIONS_DIR / "data_zone_inspect" / sid
    sdir.mkdir(parents=True, exist_ok=True)
    excel_path = save_upload(excel_file, sdir)
    try:
        sheets = data_zone_sheet_names(excel_path)
        if not sheets:
            raise ValueError("File Excel khong co sheet nao")
        if not sheet_name.strip():
            selected_sheet = sheets[0]
            tables = []
        elif sheet_name in sheets:
            selected_sheet = sheet_name
            tables = parse_source_workbook(excel_path, selected_sheet)
        else:
            raise ValueError(f"Sheet khong ton tai: {sheet_name}")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Khong doc duoc file nguon: {exc}") from exc
    return {
        "sheets": sheets,
        "selected_sheet": selected_sheet,
        "summary": data_zone_preview_tables(tables, options),
    }


@app.post("/data-zone/generate")
async def data_zone_generate(
    excel_file: UploadFile = File(...),
    sheet_name: str = Form(...),
    options_json: str = Form("{}"),
):
    if not excel_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Vui long upload file .xlsx/.xlsm")
    options = _data_zone_load_options(options_json)
    sid = str(uuid.uuid4())
    sdir = SESSIONS_DIR / "data_zone" / sid
    sdir.mkdir(parents=True, exist_ok=True)
    excel_path = save_upload(excel_file, sdir)
    try:
        tables = parse_source_workbook(excel_path, sheet_name)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Khong doc duoc sheet nguon: {exc}") from exc
    if not tables:
        raise HTTPException(status_code=400, detail="Sheet nguon khong co bang/truong hop le")

    stem = safe_filename(Path(excel_file.filename).stem)
    raw_path = sdir / f"{stem}_phu_luc_1_raw.xlsx"
    work_path = sdir / f"{stem}_phu_luc_2_work.xlsx"
    zip_path = sdir / f"{stem}_phu_luc_raw_work.zip"
    build_raw_workbook(tables, raw_path, options)
    build_work_workbook(tables, work_path, options)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(raw_path, arcname=raw_path.name)
        zf.write(work_path, arcname=work_path.name)
    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=zip_path.name,
    )


def _ulnl_rules_payload():
    return [rule.to_dict() for rule in DEFAULT_RULES]


def _load_ulnl_project(project_json: str | None) -> Project:
    if not project_json or not project_json.strip():
        project = Project()
        project.classifier_rules = _ulnl_rules_payload()
        return project
    try:
        data = json.loads(project_json)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Project JSON khong hop le: {exc.msg}") from exc
    if not isinstance(data, dict):
        raise HTTPException(status_code=400, detail="Project JSON phai la object")
    project = Project.from_dict(data)
    if not project.classifier_rules:
        project.classifier_rules = _ulnl_rules_payload()
    return project


def _ulnl_classifier(project: Project) -> Classifier:
    rules = []
    for raw in project.classifier_rules or _ulnl_rules_payload():
        if not isinstance(raw, dict):
            continue
        try:
            rules.append(ClassificationRule.from_dict(raw))
        except TypeError:
            continue
    return Classifier(rules or None)


def _ulnl_project_stats(project: Project) -> dict:
    leafs = [f for f in project.functions if f.cap == 3]
    return {
        "functions": len(project.functions),
        "leaf_functions": len(leafs),
        "web_enabled": sum(1 for f in leafs if f.include_web),
        "mobile_enabled": sum(1 for f in leafs if f.include_mobile),
        "rules": len(project.classifier_rules or []),
    }


def _ulnl_formula_count(path: Path) -> int:
    try:
        wb = load_workbook(path, data_only=False, read_only=True)
        count = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        count += 1
        wb.close()
        return count
    except Exception:
        return 0


@app.get("/ulnl", response_class=HTMLResponse)
def ulnl_index(request: Request):
    project = Project()
    project.classifier_rules = _ulnl_rules_payload()
    return render_template(
        request,
        "ulnl_tool.html",
        {
            "default_project": project.to_dict(),
            "default_rules": _ulnl_rules_payload(),
        },
    )


@app.post("/ulnl/import")
async def ulnl_import(
    excel_file: UploadFile | None = File(None),
    settings_file: UploadFile | None = File(None),
):
    project = Project()
    project.classifier_rules = _ulnl_rules_payload()

    if settings_file and settings_file.filename:
        try:
            raw = (await settings_file.read()).decode("utf-8-sig")
            loaded = json.loads(raw)
            project = Project.from_dict(loaded)
            if not project.classifier_rules:
                project.classifier_rules = _ulnl_rules_payload()
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Khong doc duoc settings JSON: {exc}") from exc

    if excel_file and excel_file.filename:
        sdir = SESSIONS_DIR / str(uuid.uuid4())
        sdir.mkdir(parents=True, exist_ok=True)
        excel_path = save_upload(excel_file, sdir)
        try:
            project.functions = import_functions_from_excel(excel_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Khong import duoc Excel danh sach CN: {exc}") from exc

    return JSONResponse({
        "project": project.to_dict(),
        "stats": _ulnl_project_stats(project),
    })


@app.post("/ulnl/classify")
async def ulnl_classify(
    project_json: str = Form(...),
    force: bool = Form(False),
):
    project = _load_ulnl_project(project_json)
    if force:
        for func in project.functions:
            func.ma_cn_web = ""
            func.ma_cn_mobile = ""
    count = enrich_with_classifier(project, _ulnl_classifier(project))
    return JSONResponse({
        "project": project.to_dict(),
        "classified": count,
        "stats": _ulnl_project_stats(project),
    })


@app.post("/ulnl/test-rule")
async def ulnl_test_rule(
    project_json: str = Form(...),
    name: str = Form(""),
    desc: str = Form(""),
    platform: str = Form("web"),
):
    project = _load_ulnl_project(project_json)
    classifier = _ulnl_classifier(project)
    ma_cn, rule_id = classifier.classify(name, desc, platform if platform == "mobile" else "web")
    return JSONResponse({"ma_cn": ma_cn, "rule_id": rule_id})


def _ulnl_build_to_path(project: Project, output_path: Path) -> dict:
    if not ULNL_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="Thieu ULNL_template_base.xlsx trong resources")
    if not project.enable_web and not project.enable_mobile:
        leafs = [func for func in project.functions if func.cap == 3]
        if any(func.include_web for func in leafs):
            project.enable_web = True
        if any(func.include_mobile for func in leafs):
            project.enable_mobile = True
    enrich_with_classifier(project, _ulnl_classifier(project))
    result = build_ulnl_file(
        project,
        output_path,
        ULNL_TEMPLATE_PATH,
        ULNL_SCRIPTS_DIR,
        apply_format_overrides=True,
    )
    if not result.get("success"):
        raise HTTPException(status_code=500, detail=result.get("error") or "Build ULNL loi")
    result["formula_count"] = _ulnl_formula_count(output_path)
    return result


@app.post("/ulnl/generate")
async def ulnl_generate(project_json: str = Form(...)):
    project = _load_ulnl_project(project_json)
    sdir = SESSIONS_DIR / str(uuid.uuid4())
    sdir.mkdir(parents=True, exist_ok=True)
    output_path = sdir / "ULNL_output.xlsx"
    _ulnl_build_to_path(project, output_path)
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="ULNL_output.xlsx",
    )


@app.post("/ulnl/preview-pdf")
async def ulnl_preview_pdf(project_json: str = Form(...)):
    project = _load_ulnl_project(project_json)
    sdir = SESSIONS_DIR / str(uuid.uuid4())
    sdir.mkdir(parents=True, exist_ok=True)
    output_path = sdir / "ULNL_output.xlsx"
    _ulnl_build_to_path(project, output_path)
    pdf_path = xlsx_to_pdf(output_path, sdir)
    if not pdf_path:
        raise HTTPException(status_code=500, detail="Khong convert duoc PDF bang LibreOffice")
    return FileResponse(pdf_path, media_type="application/pdf", filename="ULNL_output.pdf")


def _load_csdl_config(config_json: str | None):
    if not config_json or not config_json.strip():
        return json.loads(json.dumps(csdl_tool.DEFAULT_CONFIG, ensure_ascii=False))
    try:
        config = json.loads(config_json)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Config JSON không hợp lệ: {exc.msg}") from exc
    if not isinstance(config, dict) or "rules" not in config:
        raise HTTPException(status_code=400, detail="Config JSON phải có key 'rules'")
    return config


def _parse_csdl_documents(input_paths, config):
    documents = []
    errors = []
    ref_patterns = config.get("reference_table_patterns", [])
    for path in input_paths:
        try:
            doc = csdl_tool.parse_document(str(path))
            csdl_tool.assign_rules(doc.tables, config)
            csdl_tool.apply_classification(doc.tables, None, ref_patterns)
            doc.source_file = path.name
            documents.append(doc)
        except Exception as exc:
            errors.append({"file": path.name, "error": str(exc)})
    return documents, errors


def _build_csdl_preview_rows(documents):
    rows = []
    stt = 0
    for doc in documents:
        for table in doc.tables:
            for field in table.fields:
                stt += 1
                rows.append({
                    "stt": stt,
                    "source_file": getattr(doc, "source_file", ""),
                    "system_name": getattr(doc, "system_name", ""),
                    "table": table.name,
                    "table_type": table.table_type,
                    "field": field.name,
                    "datatype": field.datatype,
                    "nullable": field.nullable,
                    "pk_fk": field.pk_fk,
                    "description": field.description,
                    "dqc": getattr(field, "dqc_codes", []),
                    "std": getattr(field, "std_codes", []),
                    "ref_table": field.ref_table,
                })
    return rows


def _csdl_summary(documents, rows):
    total_tables = sum(len(doc.tables) for doc in documents)
    fields_with_rules = sum(1 for row in rows if row["dqc"] or row["std"])
    rule_counts = {}
    for row in rows:
        for code in row["dqc"] + row["std"]:
            rule_counts[code] = rule_counts.get(code, 0) + 1
    return {
        "documents": len(documents),
        "tables": total_tables,
        "fields": len(rows),
        "fields_with_rules": fields_with_rules,
        "rule_counts": dict(sorted(rule_counts.items())),
    }


def _compute_csdl_diff(documents, existing_excel_path: Path | None):
    field_notes = {}
    table_notes = {}
    if not existing_excel_path:
        return field_notes, table_notes
    try:
        wb = load_workbook(existing_excel_path)
        ws = wb.active
    except Exception:
        return field_notes, table_notes

    headers_old = {
        ws.cell(1, c).value: c
        for c in range(1, ws.max_column + 1)
        if ws.cell(1, c).value
    }
    table_header = csdl_tool.HEADERS[1]
    field_header = csdl_tool.HEADERS[2]
    datatype_header = csdl_tool.HEADERS[3]
    pk_fk_header = csdl_tool.HEADERS[10]
    description_header = csdl_tool.HEADERS[4]
    if table_header not in headers_old or field_header not in headers_old:
        return field_notes, table_notes

    old_fields = {}
    old_tables = set()
    for row_idx in range(2, ws.max_row + 1):
        table_name = ws.cell(row_idx, headers_old[table_header]).value
        field_name = ws.cell(row_idx, headers_old[field_header]).value
        if not table_name or not field_name:
            continue
        table_name = str(table_name).strip()
        field_name = str(field_name).strip()
        old_tables.add(table_name)
        old_fields[(table_name, field_name)] = {
            "datatype": ws.cell(row_idx, headers_old.get(datatype_header, 4)).value,
            "pk_fk": ws.cell(row_idx, headers_old.get(pk_fk_header, 11)).value,
            "description": ws.cell(row_idx, headers_old.get(description_header, 5)).value,
        }

    new_tables = set()
    new_field_keys = set()
    for doc in documents:
        for table in doc.tables:
            new_tables.add(table.name)
            if table.name not in old_tables:
                table_notes[table.name] = "[THÊM MỚI] Bảng mới"
            for field in table.fields:
                key = (table.name, field.name)
                new_field_keys.add(key)
                if key not in old_fields:
                    field_notes[key] = "[THÊM MỚI] Trường mới"
                    continue
                notes = []
                old = old_fields[key]
                if str(old.get("datatype") or "").strip() != field.datatype.strip():
                    notes.append(f"Đổi định dạng: '{old.get('datatype')}' -> '{field.datatype}'")
                if str(old.get("pk_fk") or "").strip() != field.pk_fk.strip():
                    notes.append(f"Đổi PK/FK: '{old.get('pk_fk')}' -> '{field.pk_fk}'")
                if str(old.get("description") or "").strip() != field.description.strip():
                    notes.append("Đổi mô tả")
                if notes:
                    field_notes[key] = " | ".join(notes)

    for key in old_fields:
        if key not in new_field_keys and key[0] in new_tables:
            field_notes[key] = "[ĐÃ XÓA] Trường này không còn trong tài liệu mới"
    for table_name in old_tables:
        if table_name not in new_tables:
            table_notes[table_name] = "[ĐÃ XÓA] Bảng này không còn trong tài liệu mới"
    return field_notes, table_notes


def _postprocess_csdl_excel(path: Path, options: dict, field_notes: dict, table_notes: dict):
    wb = load_workbook(path)
    ws = wb.active
    header_index = {
        ws.cell(1, col_idx).value: col_idx
        for col_idx in range(1, ws.max_column + 1)
        if ws.cell(1, col_idx).value
    }

    if options.get("add_diff_columns") and (field_notes or table_notes):
        field_note_col = ws.max_column + 1
        table_note_col = ws.max_column + 2
        ws.cell(1, field_note_col, "Note trường")
        ws.cell(1, table_note_col, "Note bảng")
        table_col = header_index.get(csdl_tool.HEADERS[1], 2)
        field_col = header_index.get(csdl_tool.HEADERS[2], 3)
        for row_idx in range(2, ws.max_row + 1):
            table_name = ws.cell(row_idx, table_col).value
            field_name = ws.cell(row_idx, field_col).value
            if table_name and field_name and (table_name, field_name) in field_notes:
                ws.cell(row_idx, field_note_col, field_notes[(table_name, field_name)])
            if table_name and table_name in table_notes:
                ws.cell(row_idx, table_note_col, table_notes[table_name])

    selected_columns = set(options.get("columns") or csdl_tool.HEADERS)
    selected_columns.update(["Note trường", "Note bảng"])
    cols_to_delete = [
        col_idx
        for header, col_idx in header_index.items()
        if header not in selected_columns
    ]
    for col_idx in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(col_idx)

    if options.get("only_with_rules"):
        header_index = {
            ws.cell(1, col_idx).value: col_idx
            for col_idx in range(1, ws.max_column + 1)
            if ws.cell(1, col_idx).value
        }
        dqc_col = header_index.get(csdl_tool.HEADERS[14])
        std_col = header_index.get(csdl_tool.HEADERS[15])
        if dqc_col or std_col:
            rows_to_delete = []
            for row_idx in range(2, ws.max_row + 1):
                has_dqc = dqc_col and ws.cell(row_idx, dqc_col).value
                has_std = std_col and ws.cell(row_idx, std_col).value
                if not has_dqc and not has_std:
                    rows_to_delete.append(row_idx)
            for row_idx in sorted(rows_to_delete, reverse=True):
                ws.delete_rows(row_idx)
    wb.save(path)


def _csdl_excel_preview(path: Path, limit: int = 150):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    columns = [str(cell.value or "") for cell in ws[1]]
    rows = [
        [str(value) if value is not None else "" for value in row]
        for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row, limit + 1), values_only=True)
    ]
    row_count = max(ws.max_row - 1, 0)
    wb.close()
    return {"columns": columns, "rows": rows, "row_count": row_count, "truncated": row_count > limit}


@app.post("/csdl/preview")
async def csdl_preview(
    docx_files: list[UploadFile] = File(...),
    config_json: str = Form(""),
):
    uploads = [upload for upload in docx_files if upload.filename]
    if not uploads:
        raise HTTPException(status_code=400, detail="Vui lòng upload ít nhất một file .docx")
    for upload in uploads:
        if not upload.filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="Vui lòng upload đúng file .docx")

    config = _load_csdl_config(config_json)
    sid = str(uuid.uuid4())
    sdir = SESSIONS_DIR / "csdl_preview" / sid
    sdir.mkdir(parents=True, exist_ok=True)
    input_paths = [save_upload(upload, sdir) for upload in uploads]
    documents, errors = _parse_csdl_documents(input_paths, config)
    rows = _build_csdl_preview_rows(documents)
    return {
        "summary": _csdl_summary(documents, rows),
        "rows": rows,
        "errors": errors,
        "dqc_short": csdl_tool.DQC_SHORT,
        "std_short": csdl_tool.STD_SHORT,
    }


@app.post("/csdl/generate-direct")
async def csdl_generate_direct(
    docx_files: list[UploadFile] = File(...),
    existing_excel: UploadFile | None = File(None),
    config_json: str = Form(""),
    options_json: str = Form("{}"),
    preview_output: bool = Form(False),
):
    uploads = [upload for upload in docx_files if upload.filename]
    if not uploads:
        raise HTTPException(status_code=400, detail="Vui lòng upload ít nhất một file .docx")
    for upload in uploads:
        if not upload.filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="Vui lòng upload đúng file .docx")
    try:
        options = json.loads(options_json or "{}")
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Options JSON không hợp lệ: {exc.msg}") from exc

    mode = options.get("mode", "overwrite")
    if mode not in {"overwrite", "merge", "diff"}:
        raise HTTPException(status_code=400, detail="Mode không hợp lệ")

    config = _load_csdl_config(config_json)
    sid = str(uuid.uuid4())
    sdir = SESSIONS_DIR / "csdl_direct" / sid
    sdir.mkdir(parents=True, exist_ok=True)
    input_paths = [save_upload(upload, sdir) for upload in uploads]
    existing_path = None
    if existing_excel and existing_excel.filename:
        if not existing_excel.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail="File DQC/STD cũ phải là .xlsx")
        existing_path = save_upload(existing_excel, sdir)

    documents, errors = _parse_csdl_documents(input_paths, config)
    if not documents:
        detail = "; ".join(f"{err['file']}: {err['error']}" for err in errors) or "Không parse được tài liệu"
        raise HTTPException(status_code=400, detail=detail)

    output_name = f"{input_paths[0].stem}_dqc_std.xlsx" if len(input_paths) == 1 else "dqc_std_output.xlsx"
    out_path = sdir / output_name
    field_notes, table_notes = {}, {}
    write_mode = "overwrite"
    if mode == "merge" and existing_path:
        shutil.copyfile(existing_path, out_path)
        write_mode = "merge"
    elif mode == "diff" and existing_path:
        field_notes, table_notes = _compute_csdl_diff(documents, existing_path)

    csdl_tool.write_excel(documents, out_path, mode=write_mode)
    _postprocess_csdl_excel(out_path, options, field_notes, table_notes)

    if cloud_mode():
        output_object = job_output_object(sid, output_name)
        GCSStorage().upload_file(
            out_path,
            output_object,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if preview_output:
        return {
            "filename": output_name,
            "session_id": sid,
            "preview": _csdl_excel_preview(out_path),
        }
    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_name,
    )


@app.get("/csdl/generated/{session_id}/{filename}")
def csdl_generated_download(session_id: str, filename: str):
    if not re.fullmatch(r"[a-f0-9-]{36}", session_id):
        raise HTTPException(status_code=400, detail="Session không hợp lệ")
    safe_name = safe_filename(filename)
    if safe_name != filename:
        raise HTTPException(status_code=404, detail="File Excel không tồn tại hoặc đã hết hạn")
    if cloud_mode():
        path = SESSIONS_DIR / "csdl_direct_downloads" / session_id / safe_name
        GCSStorage().download_file(job_output_object(session_id, safe_name), path)
    else:
        path = SESSIONS_DIR / "csdl_direct" / session_id / safe_name
        if not path.exists():
            raise HTTPException(status_code=404, detail="File Excel không tồn tại hoặc đã hết hạn")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=safe_name,
    )


def run_csdl_generation_job(
    job_id: str,
    input_paths,
    out_path: Path,
    rule_config_path: Path | None,
    table_config_path: Path | None,
    system_name: str | None,
    mode: str,
    verbose: bool,
):
    import contextlib
    import io

    store = get_job_store()
    log_buffer = io.StringIO()
    try:
        store.update(job_id, status="running", progress=5, message="Đang đọc file DOCX", error="")
        with contextlib.redirect_stdout(log_buffer), contextlib.redirect_stderr(log_buffer):
            result = csdl_tool.process(
                [str(path) for path in input_paths],
                str(out_path),
                rule_config_path=str(rule_config_path) if rule_config_path else None,
                table_config_path=str(table_config_path) if table_config_path else None,
                system_name=system_name or None,
                mode=mode,
                verbose=verbose,
            )
        if not result:
            raise RuntimeError("Không có document nào parse thành công. Xem log để biết chi tiết.")

        output_object = ""
        if cloud_mode():
            output_object = job_output_object(job_id, out_path.name)
            GCSStorage().upload_file(
                out_path,
                output_object,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        store.update(
            job_id,
            status="done",
            progress=100,
            message="Hoàn tất",
            output_path=str(out_path),
            output_name=out_path.name,
            output_object=output_object,
            log=log_buffer.getvalue(),
        )
    except Exception as exc:
        store.update(
            job_id,
            status="failed",
            progress=0,
            message="Tạo file Excel lỗi",
            error=str(exc),
            log=log_buffer.getvalue(),
        )


@app.post("/csdl/generate", response_class=HTMLResponse)
async def csdl_generate(
    request: Request,
    background_tasks: BackgroundTasks,
    docx_files: list[UploadFile] = File(...),
    rule_config: UploadFile | None = File(None),
    table_config: UploadFile | None = File(None),
    system_name: str = Form(""),
    mode: str = Form("overwrite"),
    verbose: str = Form("false"),
):
    uploads = [upload for upload in docx_files if upload.filename]
    if not uploads:
        raise HTTPException(status_code=400, detail="Vui lòng upload ít nhất một file .docx")
    for upload in uploads:
        if not upload.filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="Vui lòng upload đúng file .docx")
    if mode not in {"overwrite", "merge"}:
        raise HTTPException(status_code=400, detail="Mode không hợp lệ")

    sid = str(uuid.uuid4())
    sdir = SESSIONS_DIR / "csdl" / sid
    sdir.mkdir(parents=True, exist_ok=True)

    input_paths = [save_upload(upload, sdir) for upload in uploads]
    rule_config_path = None
    table_config_path = None
    if rule_config and rule_config.filename:
        if not rule_config.filename.lower().endswith(".json"):
            raise HTTPException(status_code=400, detail="Rule config phải là file .json")
        rule_config_path = save_upload(rule_config, sdir)
    if table_config and table_config.filename:
        if not table_config.filename.lower().endswith(".json"):
            raise HTTPException(status_code=400, detail="Table config phải là file .json")
        table_config_path = save_upload(table_config, sdir)

    output_name = f"{input_paths[0].stem}_csdl.xlsx" if len(input_paths) == 1 else "csdl_output.xlsx"
    out_path = sdir / output_name
    job_id = str(uuid.uuid4())

    get_job_store().create(
        job_id,
        {
            "status": "queued",
            "progress": 0,
            "message": "Đã tạo job CSDL",
            "tool": "csdl",
            "session_id": sid,
            "docx_name": ", ".join(path.name for path in input_paths),
            "output_name": out_path.name,
            "output_path": str(out_path),
        },
    )
    background_tasks.add_task(
        run_csdl_generation_job,
        job_id,
        input_paths,
        out_path,
        rule_config_path,
        table_config_path,
        system_name.strip() or None,
        mode,
        str(verbose).lower() in {"1", "true", "yes", "on"},
    )
    response = render_template(request, "csdl_job_status.html", {"job_id": job_id})
    response.background = background_tasks
    return response


@app.post("/configure", response_class=HTMLResponse)
async def configure(
    request: Request,
    docx_file: UploadFile = File(...),
    excel_file: list[UploadFile] = File(...),
):
    if not docx_file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Vui lòng upload file .docx")
    excel_uploads = [upload for upload in excel_file if upload.filename]
    if not excel_uploads:
        raise HTTPException(status_code=400, detail="Vui lòng upload ít nhất một file .xlsx/.xlsm")
    for upload in excel_uploads:
        if not upload.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Vui lòng upload file .xlsx/.xlsm")

    sid = str(uuid.uuid4())
    sdir = SESSIONS_DIR / sid
    sdir.mkdir(parents=True, exist_ok=True)

    docx_path = save_upload(docx_file, sdir)
    excel_paths = [save_upload(upload, sdir) for upload in excel_uploads]
    write_local_session_manifest(sid, docx_path, excel_paths)
    meta = load_docx_metadata(docx_path)
    excel_sources, sheets = session_excel_payload(excel_paths)
    if cloud_mode():
        upload_cloud_session(sid, docx_path, excel_paths)

    return render_configure_session(request, sid)


@app.post("/sessions/{session_id}/excel-files")
async def add_session_excel_files(
    session_id: str,
    excel_file: list[UploadFile] = File(...),
):
    uploads = [upload for upload in excel_file if upload.filename]
    if not uploads:
        raise HTTPException(status_code=400, detail="Vui lòng chọn ít nhất một file Excel")
    for upload in uploads:
        if not upload.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Vui lòng upload file .xlsx/.xlsm")

    sdir = session_dir(session_id)
    docx_path, existing_excel_paths = find_session_sources(sdir)
    new_excel_paths = [save_upload(upload, sdir) for upload in uploads]
    excel_paths = existing_excel_paths + new_excel_paths
    write_local_session_manifest(session_id, docx_path, excel_paths)
    if cloud_mode():
        upload_cloud_session(session_id, docx_path, excel_paths)

    excel_sources, sheets = session_excel_payload(excel_paths)
    return JSONResponse({
        "excel_name": ", ".join(path.name for path in excel_paths),
        "excel_files": excel_sources,
        "sheets": sheets,
    })


def submit_cloud_generation(request: Request, session_id: str, payload: dict):
    session_manifest = load_cloud_session(session_id)
    job_id = str(uuid.uuid4())
    output_name = f"{Path(session_manifest['docx_name']).stem}_web_template.docx"
    preview_name = f"{Path(session_manifest['docx_name']).stem}_web_template_preview.pdf"
    output_object = job_output_object(job_id, output_name)
    preview_object = job_preview_object(job_id, preview_name)
    excel_files = session_manifest.get("excel_files") or [
        {
            "id": session_manifest.get("excel_name", ""),
            "name": session_manifest.get("excel_name", ""),
            "object": session_manifest.get("excel_object", ""),
        }
    ]
    excel_name = ", ".join(item.get("name", "") for item in excel_files if item.get("name"))
    manifest = {
        "job_id": job_id,
        "session_id": session_id,
        "docx_name": session_manifest["docx_name"],
        "excel_name": excel_name,
        "excel_files": excel_files,
        "docx_object": session_manifest["docx_object"],
        "excel_object": session_manifest.get("excel_object", ""),
        "output_name": output_name,
        "output_object": output_object,
        "preview_name": preview_name,
        "preview_object": preview_object,
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
            "excel_name": excel_name,
            "output_name": output_name,
            "output_object": output_object,
            "preview_name": preview_name,
            "preview_object": preview_object,
            "preview_available": False,
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


def run_local_generation_job(job_id: str, docx_path: Path, excel_paths, out_path: Path, payload: dict):
    store = get_job_store()
    try:
        store.update(job_id, status="running", progress=2, message="Đang nạp engine generate", error="")

        def progress(percent, message):
            store.update(job_id, status="running", progress=percent, message=message)

        generate_docx_from_payload(
            docx_path=docx_path,
            excel_path=excel_paths,
            output_path=out_path,
            payload=payload,
            progress_callback=progress,
        )
        preview_path = out_path.with_name(f"{out_path.stem}_preview.pdf")
        preview_error = ""
        try:
            store.update(job_id, status="running", progress=98, message="Đang tạo preview PDF")
            generated_preview = convert_docx_to_pdf(out_path, preview_path.parent)
            if generated_preview != preview_path:
                generated_preview.replace(preview_path)
        except PreviewConversionError as exc:
            preview_error = str(exc)
        store.update(
            job_id,
            status="done",
            progress=100,
            message="Hoàn tất",
            output_path=str(out_path),
            output_name=out_path.name,
            preview_path=str(preview_path) if preview_path.exists() else "",
            preview_name=preview_path.name,
            preview_available=preview_path.exists(),
            preview_error=preview_error,
        )
    except Exception as exc:
        store.update(
            job_id,
            status="failed",
            progress=0,
            message="Generate lỗi",
            error=str(exc),
        )


def submit_local_generation(
    request: Request,
    background_tasks: BackgroundTasks,
    session_id: str,
    payload: dict,
):
    sdir = session_dir(session_id)
    docx_path, excel_paths = find_session_sources(sdir)
    out_path = sdir / f"{docx_path.stem}_web_template.docx"
    preview_path = sdir / f"{docx_path.stem}_web_template_preview.pdf"
    job_id = str(uuid.uuid4())
    (sdir / f"{job_id}_payload.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    get_job_store().create(
        job_id,
        {
            "status": "queued",
            "progress": 0,
            "message": "Đã tạo job local",
            "session_id": session_id,
            "docx_name": docx_path.name,
            "excel_name": ", ".join(path.name for path in excel_paths),
            "output_name": out_path.name,
            "output_path": str(out_path),
            "preview_name": preview_path.name,
            "preview_path": str(preview_path),
            "preview_available": False,
        },
    )
    background_tasks.add_task(run_local_generation_job, job_id, docx_path, excel_paths, out_path, payload)
    response = render_template(request, "job_status.html", {"job_id": job_id})
    response.background = background_tasks
    return response


@app.post("/generate")
async def generate(
    request: Request,
    background_tasks: BackgroundTasks,
    session_id: str = Form(...),
    intro_end_idx: str = Form(""),
    doc_defaults_json: str = Form("{}"),
    sections_json: str = Form("[]"),
    intro_replacements_json: str = Form("{}"),
    heading_styles_json: str = Form("{}"),
    toc_settings_json: str = Form("{}"),
    header_footer_text_json: str = Form("{}"),
    headings_config_json: str = Form("[]"),
    step6_setup_json: str = Form("{}"),
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
            toc_settings_json=toc_settings_json,
            header_footer_text_json=header_footer_text_json,
            headings_config_json=headings_config_json,
            step6_setup_json=step6_setup_json,
            preserve_inline_formatting=preserve_inline_formatting,
            excel_font_name=excel_font_name,
            excel_font_size=excel_font_size,
        )
        validate_payload(payload)
    except GenerationConfigError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if cloud_mode():
        return submit_cloud_generation(request, session_id, payload)

    return submit_local_generation(request, background_tasks, session_id, payload)


@app.get("/jobs/{job_id}", response_class=HTMLResponse)
def job_page(request: Request, job_id: str):
    return render_template(request, "job_status.html", {"job_id": job_id})


@app.get("/jobs/{job_id}/edit", response_class=HTMLResponse)
def job_edit(request: Request, job_id: str):
    record = get_job_store().get(job_id)
    if not record:
        raise HTTPException(status_code=404, detail="Job không tồn tại")
    session_id = record.get("session_id")
    if not session_id:
        raise HTTPException(status_code=404, detail="Không tìm thấy session của job")
    payload = {}
    payload_path = session_dir(session_id) / f"{job_id}_payload.json"
    if payload_path.exists():
        payload = json.loads(payload_path.read_text(encoding="utf-8"))
    elif cloud_mode():
        # Cloud jobs keep the payload in the job manifest. If it is unavailable,
        # falling back to an empty payload still renders the original session.
        try:
            manifest = GCSStorage().download_json(job_manifest_object(job_id))
            payload = manifest.get("payload") or {}
        except Exception:
            payload = {}
    return render_configure_session(request, session_id, payload)


@app.get("/jobs/{job_id}/status")
def job_status(job_id: str):
    record = get_job_store().get(job_id)
    if not record:
        raise HTTPException(status_code=404, detail="Job không tồn tại")
    return record


@app.get("/jobs/{job_id}/preview")
def job_preview(job_id: str):
    record = get_job_store().get(job_id)
    if not record:
        raise HTTPException(status_code=404, detail="Job không tồn tại")
    if record.get("status") != "done":
        raise HTTPException(status_code=409, detail="Job chưa hoàn tất")
    if not record.get("preview_available"):
        raise HTTPException(status_code=404, detail=record.get("preview_error") or "Không có preview PDF")

    preview_name = record.get("preview_name") or f"{job_id}_preview.pdf"
    if cloud_mode():
        preview_object = record.get("preview_object")
        if not preview_object:
            raise HTTPException(status_code=404, detail="Không tìm thấy preview")
        dst = SESSIONS_DIR / "job_previews" / job_id / safe_filename(preview_name)
        GCSStorage().download_file(preview_object, dst)
        get_job_store().update(job_id, preview_seen=True)
        return FileResponse(
            dst,
            media_type="application/pdf",
            filename=preview_name,
            content_disposition_type="inline",
        )

    preview_path = record.get("preview_path")
    if not preview_path:
        raise HTTPException(status_code=404, detail="Không tìm thấy preview")
    path = Path(preview_path)
    if not path.exists():
        raise HTTPException(status_code=404, detail="File preview không tồn tại")
    get_job_store().update(job_id, preview_seen=True)
    return FileResponse(
        path,
        media_type="application/pdf",
        filename=preview_name,
        content_disposition_type="inline",
    )


@app.get("/jobs/{job_id}/download")
def job_download(job_id: str):
    record = get_job_store().get(job_id)
    if not record:
        raise HTTPException(status_code=404, detail="Job không tồn tại")
    if record.get("status") != "done":
        raise HTTPException(status_code=409, detail="Job chưa hoàn tất")
    if record.get("preview_available") and not record.get("preview_seen"):
        raise HTTPException(status_code=409, detail="Vui lòng xem preview trước khi tải DOCX")
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
    output_path = record.get("output_path")
    if not output_path:
        raise HTTPException(status_code=404, detail="Không tìm thấy output")
    path = Path(output_path)
    if not path.exists():
        raise HTTPException(status_code=404, detail="File output không tồn tại")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=output_name,
    )


@app.get("/csdl/jobs/{job_id}", response_class=HTMLResponse)
def csdl_job_page(request: Request, job_id: str):
    return render_template(request, "csdl_job_status.html", {"job_id": job_id})


@app.get("/csdl/jobs/{job_id}/download")
def csdl_job_download(job_id: str):
    record = get_job_store().get(job_id)
    if not record:
        raise HTTPException(status_code=404, detail="Job không tồn tại")
    if record.get("tool") != "csdl":
        raise HTTPException(status_code=404, detail="Không phải job CSDL")
    if record.get("status") != "done":
        raise HTTPException(status_code=409, detail="Job chưa hoàn tất")

    output_name = record.get("output_name") or f"{job_id}.xlsx"
    if cloud_mode():
        output_object = record.get("output_object")
        if not output_object:
            raise HTTPException(status_code=404, detail="Không tìm thấy output")
        dst = SESSIONS_DIR / "csdl_job_downloads" / job_id / safe_filename(output_name)
        GCSStorage().download_file(output_object, dst)
        return FileResponse(
            dst,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=output_name,
        )

    output_path = record.get("output_path")
    if not output_path:
        raise HTTPException(status_code=404, detail="Không tìm thấy output")
    path = Path(output_path)
    if not path.exists():
        raise HTTPException(status_code=404, detail="File output không tồn tại")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_name,
    )


@app.get("/health")
def health():
    return {"ok": True}
