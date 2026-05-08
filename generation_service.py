import importlib.util
import json
import re
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent


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


class GenerationConfigError(ValueError):
    pass


def parse_json_value(raw, default, field_name):
    if raw is None or str(raw).strip() == "":
        return default
    if isinstance(raw, (dict, list)):
        return raw
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise GenerationConfigError(f"{field_name} không phải JSON hợp lệ") from exc


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


def load_docx_metadata(path):
    model = DocxModel()
    model.load(str(path))
    intro = model.get_effective_intro_end_idx()
    all_headings = model.get_all_headings()
    headings = [h for h in all_headings if h["original_idx"] >= intro]
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


def build_headings_config(raw_headings, excel_path):
    if not isinstance(raw_headings, list):
        raise GenerationConfigError("Danh sách heading không hợp lệ")

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
                    raise GenerationConfigError(
                        f'Heading "{entry["text"]}" đang chèn Excel nhưng chưa chọn cột'
                    )
                if not parse_excel_column_spec(cols):
                    raise GenerationConfigError(
                        f'Danh sách cột không hợp lệ ở heading "{entry["text"]}"'
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


def payload_from_form(
    *,
    intro_end_idx="",
    doc_defaults_json="{}",
    sections_json="[]",
    intro_replacements_json="{}",
    heading_styles_json="{}",
    header_footer_text_json="{}",
    headings_config_json="[]",
    preserve_inline_formatting="true",
    excel_font_name="",
    excel_font_size=11,
):
    return {
        "intro_end_idx": intro_end_idx,
        "doc_defaults": parse_json_value(doc_defaults_json, {}, "Document defaults"),
        "sections": parse_json_value(sections_json, [], "Sections"),
        "intro_replacements": parse_json_value(intro_replacements_json, {}, "Intro replacements"),
        "heading_styles": parse_json_value(heading_styles_json, {}, "Heading styles"),
        "header_footer_text": parse_json_value(header_footer_text_json, {}, "Header/Footer"),
        "headings": parse_json_value(headings_config_json, [], "Danh sách heading"),
        "preserve_inline_formatting": as_bool(preserve_inline_formatting, True),
        "excel_font_name": str(excel_font_name or "").strip(),
        "excel_font_size": as_float(excel_font_size, 11) or 11,
    }


def validate_payload(payload):
    build_headings_config(payload.get("headings") or [], "excel.xlsx")


def configure_model_from_payload(model, excel_path, payload):
    intro_idx = as_int(payload.get("intro_end_idx"), model.get_intro_end_idx())
    model.config["intro_end_idx"] = max(0, intro_idx or 0)

    apply_doc_defaults_config(model, payload.get("doc_defaults") or {})
    apply_sections_config(model, payload.get("sections") or [])

    intro_replacements = payload.get("intro_replacements") or {}
    if isinstance(intro_replacements, dict):
        model.config["intro_replacements"] = {
            str(k): str(v)
            for k, v in intro_replacements.items()
            if str(k) != str(v)
        }

    apply_heading_styles_config(model, payload.get("heading_styles") or {})

    header_footer_text = payload.get("header_footer_text") or {}
    if isinstance(header_footer_text, dict):
        model.config["header_footer_text"] = {
            str(k): str(v)
            for k, v in header_footer_text.items()
        }

    headings_list = build_headings_config(payload.get("headings") or [], excel_path)
    if not headings_list:
        headings_list = [
            {
                **h,
                "keep_content": True,
                "insert_source": None,
                "insert_selection": None,
                "auto_number": False,
            }
            for h in model.get_all_headings()
            if h["original_idx"] >= model.config["intro_end_idx"]
        ]

    model.config["headings_list"] = headings_list
    model.config["preserve_inline_formatting"] = bool(payload.get("preserve_inline_formatting", True))
    model.config["excel_font_name"] = str(payload.get("excel_font_name") or "").strip()
    model.config["excel_font_size"] = payload.get("excel_font_size") or 11


def generate_docx_from_payload(docx_path, excel_path, output_path, payload, progress_callback=None):
    model = DocxModel()
    model.load(str(docx_path))
    configure_model_from_payload(model, str(excel_path), payload)
    model.generate(str(output_path), progress_callback=progress_callback)
    return output_path
