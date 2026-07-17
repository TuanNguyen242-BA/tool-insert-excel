import importlib.util
import json
import re
import unicodedata
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
        "toc_settings": model.config.get("toc_settings", {}),
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


def apply_toc_settings_config(model, raw_settings):
    if not isinstance(raw_settings, dict):
        return
    cfg = model._ensure_toc_settings_defaults()
    if "enabled" in raw_settings:
        cfg["enabled"] = as_bool(raw_settings.get("enabled"), cfg.get("enabled", True))
    if "update_on_open" in raw_settings:
        cfg["update_on_open"] = as_bool(raw_settings.get("update_on_open"), cfg.get("update_on_open", True))
    if "levels" in raw_settings:
        cfg["levels"] = min(9, max(1, as_int(raw_settings.get("levels"), cfg.get("levels", 4)) or 4))
    if "tab_leader" in raw_settings:
        leader = str(raw_settings.get("tab_leader") or "none").strip()
        cfg["tab_leader"] = leader if leader in {"none", "dot", "hyphen", "underscore"} else "none"
    if "right_tab_cm" in raw_settings:
        cfg["right_tab_cm"] = max(0.0, as_float(raw_settings.get("right_tab_cm"), cfg.get("right_tab_cm", 0)) or 0)

    incoming_styles = raw_settings.get("styles") or {}
    if isinstance(incoming_styles, dict):
        for style_id, incoming in incoming_styles.items():
            if style_id not in cfg.get("styles", {}) or not isinstance(incoming, dict):
                continue
            style = cfg["styles"][style_id]
            for key in ("font", "color", "line_spacing"):
                if key in incoming:
                    style[key] = str(incoming.get(key) or "").strip()
            for key in ("size_pt", "left_indent_cm", "text_tab_cm", "space_before_pt", "space_after_pt"):
                if key in incoming:
                    style[key] = as_float(incoming.get(key), style.get(key, 0))
            for key in ("bold", "italic"):
                if key in incoming:
                    style[key] = as_bool(incoming.get(key), style.get(key, False))
    model.config["toc_settings"] = cfg


def normalize_excel_sources(excel_paths):
    if isinstance(excel_paths, (str, Path)):
        paths = [Path(excel_paths)]
    elif isinstance(excel_paths, dict):
        return {
            str(file_id): str(path)
            for file_id, path in excel_paths.items()
        }
    else:
        paths = [Path(path) for path in (excel_paths or [])]

    return {path.name: str(path) for path in paths}


def normalize_source_key(value):
    text = unicodedata.normalize("NFD", str(value or ""))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", text).strip().lower()


def resolve_excel_source_path(src, excel_sources, default_file_id=""):
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
        Path(path).name: (file_id, path)
        for file_id, path in excel_sources.items()
    }
    for candidate in candidates:
        candidate = str(candidate or "").strip()
        if candidate and candidate in source_by_basename:
            return source_by_basename[candidate]

    normalized_sources = {
        normalize_source_key(file_id): (file_id, path)
        for file_id, path in excel_sources.items()
    }
    normalized_sources.update({
        normalize_source_key(Path(path).name): (file_id, path)
        for file_id, path in excel_sources.items()
    })
    for candidate in candidates:
        key = normalize_source_key(candidate)
        if key and key in normalized_sources:
            return normalized_sources[key]

    raw_path = Path(str(src.get("source_path") or ""))
    if raw_path.exists():
        return raw_path.name, str(raw_path)
    return "", ""


def build_headings_config(raw_headings, excel_paths):
    if not isinstance(raw_headings, list):
        raise GenerationConfigError("Danh sách heading không hợp lệ")

    excel_sources = normalize_excel_sources(excel_paths)
    default_file_id = next(iter(excel_sources.keys()), "")

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
            if not file_id:
                file_id = default_file_id
            if excel_sources and file_id not in excel_sources:
                raise GenerationConfigError(
                    f'Heading "{entry["text"]}" chon file Excel khong hop le'
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
                        src_file_id, src_path = resolve_excel_source_path(src, excel_sources, file_id)
                        src_entry = {
                            "file_id": src_file_id,
                            "source_path": src_path,
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
                    if excel_sources and advanced.get("split_sources") and skipped_sources:
                        raise GenerationConfigError(
                            f'Không map được nguồn phụ ở heading "{entry["text"]}": '
                            + ", ".join(str(x) for x in skipped_sources)
                        )
                    split_text_rules = []
                    for rule in advanced.get("split_text_rules") or []:
                        if not isinstance(rule, dict):
                            continue
                        rule_file_id, rule_path = resolve_excel_source_path(rule, excel_sources, file_id)
                        rule_entry = {
                            "enabled": as_bool(rule.get("enabled"), True),
                            "file_id": rule_file_id,
                            "source_path": rule_path,
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
            entry["insert_source"] = excel_sources.get(file_id) or "excel.xlsx"
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
    toc_settings_json="{}",
    header_footer_text_json="{}",
    headings_config_json="[]",
    step6_setup_json="{}",
    preserve_inline_formatting="true",
    excel_font_name="",
    excel_font_size=11,
):
    step6_setup_payload = parse_json_value(step6_setup_json, {}, "Step 6 setup")
    return {
        "intro_end_idx": intro_end_idx,
        "doc_defaults": parse_json_value(doc_defaults_json, {}, "Document defaults"),
        "sections": parse_json_value(sections_json, [], "Sections"),
        "intro_replacements": parse_json_value(intro_replacements_json, {}, "Intro replacements"),
        "heading_styles": parse_json_value(heading_styles_json, {}, "Heading styles"),
        "toc_settings": parse_json_value(toc_settings_json, {}, "TOC settings"),
        "header_footer_text": parse_json_value(header_footer_text_json, {}, "Header/Footer"),
        "headings": parse_json_value(headings_config_json, [], "Danh sách heading"),
        "_step6_setup_payload": step6_setup_payload if isinstance(step6_setup_payload, dict) else {},
        "preserve_inline_formatting": as_bool(preserve_inline_formatting, True),
        "excel_font_name": str(excel_font_name or "").strip(),
        "excel_font_size": as_float(excel_font_size, 11) or 11,
    }


def validate_payload(payload):
    build_headings_config(payload.get("headings") or [], {})


def configure_model_from_payload(model, excel_paths, payload):
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
    apply_toc_settings_config(model, payload.get("toc_settings") or {})

    header_footer_text = payload.get("header_footer_text") or {}
    if isinstance(header_footer_text, dict):
        model.config["header_footer_text"] = {
            str(k): str(v)
            for k, v in header_footer_text.items()
        }

    headings_list = build_headings_config(payload.get("headings") or [], excel_paths)
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
    configure_model_from_payload(model, excel_path, payload)
    model.generate(str(output_path), progress_callback=progress_callback)
    return output_path
