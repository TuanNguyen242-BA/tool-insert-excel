"""
tk_csdl_tool.py - Tool sinh Excel mô tả thiết kế CSDL từ tài liệu .docx
========================================================================

Một file duy nhất, gộp toàn bộ module: parser, rules, classifier,
templates, excel_writer, config, CLI.

CÀI ĐẶT:
    pip install openpyxl python-docx

(Không cần pandoc — đọc .docx trực tiếp qua python-docx)

CÁCH DÙNG:
    # Kiểm tra môi trường trước (không xử lý gì cả)
    python tk_csdl_tool.py --check

    # Cơ bản
    python tk_csdl_tool.py input.docx -o output.xlsx

    # Verbose (in chi tiết từng bước)
    python tk_csdl_tool.py input.docx -o output.xlsx --verbose

    # Nhiều file gộp vào 1 Excel
    python tk_csdl_tool.py f1.docx f2.docx -o output.xlsx

    # Xuất config mặc định ra file để chỉnh
    python tk_csdl_tool.py --dump-config my_rules.json

    # Dùng config tùy chỉnh
    python tk_csdl_tool.py input.docx -o output.xlsx --config my_rules.json

    # Merge mode: giữ chỉnh sửa thủ công, chỉ thêm code mới
    python tk_csdl_tool.py input.docx -o output.xlsx --mode merge

    # Override tên hệ thống
    python tk_csdl_tool.py input.docx -o output.xlsx --system-name "HT X"

    # Override loại bảng qua JSON riêng
    python tk_csdl_tool.py input.docx -o output.xlsx --table-config types.json

NẾU BỊ LỖI:
    python tk_csdl_tool.py --check
    # Xem output, gửi toàn bộ log cho người support
"""
# ============================================================================
# BOOTSTRAP - Phần này chạy trước tất cả để bắt lỗi sớm nhất có thể
# ============================================================================
import sys
import os
import traceback
import platform

# 1. Check Python version (cần 3.7+ vì dùng from __future__ import annotations và f-string)
if sys.version_info < (3, 7):
    sys.stderr.write(
        f"[FATAL] Cần Python 3.7+, hiện đang dùng Python "
        f"{sys.version_info.major}.{sys.version_info.minor}."
        f"{sys.version_info.micro}\n"
    )
    sys.stderr.write("Vui lòng nâng cấp Python: https://www.python.org/downloads/\n")
    sys.exit(1)

# 2. Fix Windows console encoding (in được tiếng Việt mà không crash)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    # Python cũ không có reconfigure, không quan trọng
    pass


def _print_env_info():
    """In thông tin môi trường để debug."""
    print("=" * 70)
    print("THÔNG TIN MÔI TRƯỜNG")
    print("=" * 70)
    print(f"Python:        {sys.version.split(chr(10))[0]}")
    print(f"Executable:    {sys.executable}")
    print(f"Platform:      {platform.platform()}")
    print(f"OS:            {sys.platform}")
    print(f"Encoding (stdout): {sys.stdout.encoding}")
    print(f"Working dir:   {os.getcwd()}")
    print(f"Script:        {os.path.abspath(__file__) if '__file__' in dir() else '<unknown>'}")
    print()


def _check_dependencies(verbose=True):
    """Kiểm tra các package cần thiết."""
    missing = []
    versions = {}

    # python-docx
    try:
        import docx
        versions["python-docx"] = getattr(docx, "__version__", "(no __version__)")
    except ImportError as e:
        missing.append(("python-docx", str(e)))

    # openpyxl
    try:
        import openpyxl
        versions["openpyxl"] = openpyxl.__version__
    except ImportError as e:
        missing.append(("openpyxl", str(e)))

    if verbose:
        print("DEPENDENCY CHECK")
        print("=" * 70)
        for name, ver in versions.items():
            print(f"  [OK]      {name}: {ver}")
        for name, err in missing:
            print(f"  [MISSING] {name}: {err}")
        print()

    if missing:
        print("[FATAL] Thiếu package cần thiết.", file=sys.stderr)
        print("Vui lòng cài đặt:", file=sys.stderr)
        print(f"  pip install {' '.join(n for n, _ in missing)}", file=sys.stderr)
        sys.exit(1)

    return versions


# 3. Check deps ngay (nếu thiếu thì in lỗi rõ ràng)
#    Chạy silent ở bootstrap; nếu user truyền --check/--verbose thì main() sẽ
#    in chi tiết sau
_check_dependencies(verbose=False)


# ============================================================================
# IMPORTS chính (sau khi đã verify deps)
# ============================================================================
import argparse
import json
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import List

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================================
# DEFAULT CONFIG (embed dưới dạng Python dict, có thể dump ra JSON)
# ============================================================================

DEFAULT_CONFIG = {
    "$schema_version": "2.0",
    "$description": (
        "Cấu hình rule DQC/STD cho tool sinh Excel mô tả CSDL. "
        "Tất cả rule đều có thể bật/tắt bằng 'enabled'. "
        "Các mảng datatype và điều kiện text tương đương lựa chọn tại màn hình setup."
    ),
    "reference_table_patterns": [
        "^DanhMuc", "^DM_", "^reference_", "^ref_",
        "^category_", "^dm_",
    ],
    "skip_fields": [
        "created_date", "last_modified_date",
        "created_by", "last_modified_by",
        "NgayTao", "NgayCapNhat", "NguoiTao", "NguoiCapNhat",
        "schedule_created_date", "assignment_created_date",
        "decision_created_date", "decision_date",
    ],
    "datatype_groups": {
        "date": ["DATE", "DATETIME", "TIMESTAMP", "DATETIME2",
                 "SMALLDATETIME", "TIME"],
        "numeric": ["INT", "BIGINT", "TINYINT", "SMALLINT", "DECIMAL",
                    "NUMERIC", "FLOAT", "DOUBLE", "REAL", "MONEY",
                    "SMALLMONEY", "INTEGER", "NUMBER", "BIT"],
        "text": ["VARCHAR", "NVARCHAR", "TEXT", "CHAR", "NCHAR",
                 "NTEXT", "CLOB", "STRING"],
        "uuid": ["UUID", "UNIQUEIDENTIFIER", "GUID"],
        "bool": ["BOOLEAN", "BOOL"],
    },
    "rules": {
        "DQC_003": {
            "enabled": True,
            "short_name": "Kiểm tra tính đồng bộ của dữ liệu",
            "apply_to_datatypes": ["*"],
            "apply_to_reference_tables": False,
            "exclude_pk": True,
            "trigger_pk_fk_fk": True,
            "trigger_name_id_uuid": True,
            "include_keywords": [
                "liên kết", "tham chiếu", "khóa ngoại",
                "foreign key", "reference", "linked to", "links to",
            ],
            "include_targets": ["description"],
        },
        "DQC_005": {
            "enabled": True,
            "short_name": "Kiểm tra giá trị ngày ngoài phạm vi nghiệp vụ",
            "apply_to_datatypes": ["date"],
            "apply_to_reference_tables": True,
        },
        "DQC_006": {
            "enabled": True,
            "short_name": "Kiểm tra định dạng số",
            "apply_to_datatypes": ["text"],
            "apply_to_reference_tables": True,
            "include_keywords": [
                "%", "tỷ lệ", "ty_le", "tyle", "ratio", "percent", "percentage",
                "hệ số", "he_so", "amount", "total", "quantity", "qty",
                "số tiền", "so_tien", "money", "price", "cost",
                "số lượng", "so_luong", "tổng số", "tong_so", "count",
            ],
            "include_targets": ["field_name", "description"],
        },
        "DQC_008": {
            "enabled": True,
            "short_name": "Kiểm tra định dạng ngày tháng",
            "apply_to_datatypes": ["date", "text"],
            "apply_to_reference_tables": True,
            "include_keywords": [
                "ngày", "ngay", "date", "time", "thời gian", "thoi gian",
                "tháng", "thang", "năm", "month", "year", "datetime",
            ],
            "include_targets": ["field_name", "description"],
        },
        "DQC_010": {
            "enabled": True,
            "short_name": "Kiểm tra logic nghiệp vụ giữa các trường",
            "apply_to_datatypes": ["*"],
            "apply_to_reference_tables": True,
            "assign_to_both": True,
            "pair_terms": [
                ["start_", "end_"],
                ["from_", "to_"],
                ["begin_", "end_"],
                ["tu_", "den_"],
                ["ngay_bat_dau", "ngay_ket_thuc"],
                ["NgayBatDau", "NgayKetThuc"],
                ["ThoiGianBatDau", "ThoiGianKetThuc"],
            ],
            "pair_targets": ["field_name"],
        },
        "STD_002": {
            "enabled": True,
            "short_name": "Chuẩn hóa dữ liệu số",
            "apply_to_datatypes": ["text"],
            "apply_to_reference_tables": True,
            "requires": "DQC_006",
            "not_if": "STD_003",
        },
        "STD_003": {
            "enabled": True,
            "short_name": "Chuẩn hóa trường tỷ lệ",
            "apply_to_datatypes": ["text"],
            "apply_to_reference_tables": True,
            "requires": "DQC_006",
            "include_keywords": [
                "tỷ lệ", "ty_le", "tyle", "ratio", "percent", "percentage",
                "%", "hệ số", "he_so",
            ],
            "include_targets": ["field_name", "description"],
        },
        "STD_004": {
            "enabled": True,
            "short_name": "Chuẩn hóa và hiển thị ngày tháng năm (đầy đủ)",
            "apply_to_datatypes": ["date"],
            "apply_to_reference_tables": True,
            "requires": "DQC_008",
        },
        "STD_009": {
            "enabled": True,
            "short_name": "Chuẩn hóa tên người",
            "apply_to_datatypes": ["text"],
            "apply_to_reference_tables": True,
            "name_patterns": [
                r"^Ho\w*$", r"^Ten\w*$", r"^Name$", r"^FullName$",
                r"^full_name$", r"^ho_ten$", r"^DemKhaiSinh$", r"^TenKhaiSinh$",
                r"^TenChu$", r"^TenHo$", r"^TenDem$", r"^TenLot$",
                r"^first_name$", r"^last_name$", r"^middle_name$",
                r".*_name$",
            ],
            "include_keywords": [
                "họ tên", "ho ten", "tên người", "ten nguoi",
                "họ và tên", "full name", "person name",
            ],
            "include_targets": ["description"],
            "blacklist_suffix": [
                "_org", "_dept", "_department", "_team", "_unit",
                "_chuc_vu", "_position", "_organization", "_company",
                "_bang", "_table", "_role",
            ],
            "exclude_keywords": [
                "tổ chức", "đơn vị", "phòng ban", "tổ công tác",
                "organization", "department", "unit", "team",
                "bảng", "table",
            ],
            "exclude_targets": ["description"],
        },
        "STD_011": {
            "enabled": True,
            "short_name": "Chuẩn hóa tên phân cấp cha-con (tổ chức/đơn vị)",
            "apply_to_datatypes": ["text"],
            "apply_to_reference_tables": True,
            "name_patterns": [
                r"^TenDonVi\w*$", r"^TenToChuc\w*$", r"^TenBoPhan\w*$",
                r"^TenNhom\w*$", r"^TenCoQuan\w*$", r"^TenPhongBan\w*$",
                r"^TenBan\w*$", r"^TenCucVu\w*$",
                r"^org_name$", r"^dept_name$", r"^department_name$",
                r"^unit_name$", r"^team_name$", r"^organization_name$",
                r".*_org_name$", r".*_dept_name$", r".*_unit_name$",
                r".*_organization$",
            ],
            "include_keywords": [
                "phân cấp", "cha-con", "cấu trúc phân cấp",
                "hierarchy", "parent-child", "phân cấp cha",
            ],
            "include_targets": ["description"],
        },
        "STD_012": {
            "enabled": True,
            "short_name": "Chuẩn hóa ngày tháng không đầy đủ (text/precision)",
            "apply_to_datatypes": ["text"],
            "apply_to_reference_tables": True,
            "requires": "DQC_008",
        },
        "STD_013": {
            "enabled": True,
            "short_name": "Chuẩn hóa khoảng trắng",
            "apply_to_datatypes": ["text"],
            "apply_to_reference_tables": True,
        },
    },
}


# ============================================================================
# DATACLASSES
# ============================================================================

@dataclass
class FieldInfo:
    name: str = ""
    datatype: str = ""
    nullable: str = ""
    unique: str = ""
    pk_fk: str = ""
    default: str = ""
    description: str = ""
    ref_table: str = ""
    ref_field: str = ""
    dqc_codes: List[str] = field(default_factory=list)
    std_codes: List[str] = field(default_factory=list)


@dataclass
class TableInfo:
    name: str = ""
    description: str = ""
    fields: List[FieldInfo] = field(default_factory=list)
    constraints_raw: List[str] = field(default_factory=list)
    table_type: str = ""


@dataclass
class DocumentInfo:
    system_name: str = ""
    format_type: str = ""
    tables: List[TableInfo] = field(default_factory=list)


# ============================================================================
# CONFIG HELPERS
# ============================================================================

def load_config(path=None):
    """Load config tu file JSON. Neu path=None, dung DEFAULT_CONFIG."""
    if not path:
        # Deep copy de tranh mutate
        return json.loads(json.dumps(DEFAULT_CONFIG))
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def dump_config(output_path):
    """Xuat DEFAULT_CONFIG ra file JSON."""
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
    return output_path


def _tokenize(text):
    """'NgayBatDau' -> {'ngay','bat','dau'}. Tach camelCase + snake/kebab."""
    s = re.sub(r"(?<!^)(?=[A-Z])", " ", text)
    tokens = re.split(r"[\s_\-/.,()]+", s)
    return {t.lower() for t in tokens if t}


def has_keyword(text, keywords):
    """Token-match (tranh false-positive substring).
    'nam' khong match 'name'; 'date' van match 'created_date'."""
    if not text:
        return False
    text_lower = text.lower()
    tokens = _tokenize(text)
    for kw in keywords:
        kw_lower = kw.lower()
        if not any(c.isalpha() for c in kw_lower):
            if kw_lower in text_lower:
                return True
            continue
        if " " in kw_lower:
            if kw_lower in text_lower:
                return True
            continue
        if kw_lower in tokens:
            return True
    return False


def datatype_matches(datatype, datatype_groups, allowed):
    """Check datatype thuoc nhom allowed (vd ['text','date'])."""
    if not datatype:
        return "*" in allowed
    if "*" in allowed:
        return True
    dt_upper = datatype.upper()
    for group_name in allowed:
        types = datatype_groups.get(group_name, [group_name])
        for t in types:
            if re.match(rf"^\s*{re.escape(t)}\b", dt_upper):
                return True
    return False


def rule_datatype_matches(field, rule, config):
    """Áp dụng đồng thời danh sách datatype được chọn và bị loại trừ."""
    datatype_groups = config.get("datatype_groups", {})
    allowed = rule.get("apply_to_datatypes", ["*"])
    excluded = rule.get("exclude_datatypes", [])
    if not datatype_matches(field.datatype, datatype_groups, allowed):
        return False
    return not excluded or not datatype_matches(field.datatype, datatype_groups, excluded)


def rule_reference_table_matches(code, rule, table, config):
    """Check rule co ap dung cho bang Reference hay khong."""
    ref_patterns = config.get("reference_table_patterns", [])
    if not ref_patterns or not is_reference_table(table.name, ref_patterns):
        return True
    if "apply_to_reference_tables" in rule:
        return bool(rule.get("apply_to_reference_tables"))
    if "exclude_reference_tables" in rule:
        return not bool(rule.get("exclude_reference_tables"))
    return True


def _rule_values(rule, modern_key, *legacy_keys):
    if modern_key in rule:
        return rule.get(modern_key, [])
    for key in legacy_keys:
        if key in rule:
            return rule.get(key, [])
    return []


def _target_texts(field, table, targets):
    values = {
        "table_name": table.name,
        "field_name": field.name,
        "description": field.description,
    }
    return [values[target] for target in targets if target in values and values[target]]


def _matches_text_condition(field, table, rule, value_key, target_key,
                            legacy_keys=(), legacy_targets=None):
    values = _rule_values(rule, value_key, *legacy_keys)
    if not values:
        return False
    targets = rule.get(target_key)
    if targets is None:
        targets = legacy_targets or ["field_name", "description"]
    return any(has_keyword(text, values) for text in _target_texts(field, table, targets))


def _is_text_excluded(field, table, rule):
    return _matches_text_condition(
        field, table, rule, "exclude_keywords", "exclude_targets",
        ("blacklist_description_keywords",), ["description"],
    )


def is_reference_table(table_name, patterns):
    """Match prefix bang danh muc."""
    for pat in patterns:
        if re.search(pat, table_name, re.IGNORECASE):
            return True
    return False


def match_any_pattern(value, patterns):
    """Match regex pattern."""
    for pat in patterns:
        if re.match(pat, value):
            return True
    return False


# ============================================================================
# PARSER (.docx -> Tables)
# ============================================================================

def _extract_markdown(docx_path):
    """Đọc file .docx bằng python-docx, convert sang markdown-like text
    để parser xử lý. Không cần pandoc/extract-text bên ngoài.

    Quy tắc convert:
    - Heading 1/2/3 -> #/##/###
    - Paragraph có numPr (bullet) -> '- {text}'
    - Paragraph thường -> giữ nguyên
    - Table -> markdown table | col1 | col2 |
    - Bold run -> **text**
    """
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    doc = Document(docx_path)
    lines = []

    def render_paragraph(p):
        # Lay text + xu ly bold
        parts = []
        for run in p.runs:
            txt = run.text or ""
            if not txt:
                continue
            if run.bold:
                parts.append(f"**{txt}**")
            else:
                parts.append(txt)
        text = "".join(parts).strip()
        if not text:
            return ""

        # Kiem tra style heading
        style_name = (p.style.name if p.style else "") or ""
        if style_name.startswith("Heading 1"):
            return f"# {text}"
        if style_name.startswith("Heading 2"):
            return f"## {text}"
        if style_name.startswith("Heading 3"):
            return f"### {text}"
        if style_name.startswith("Heading 4"):
            return f"#### {text}"
        if style_name == "Title":
            return f"# {text}"

        # Kiem tra bullet/numbered list (numPr)
        try:
            numPr = p._p.find(f".//{WORD_NS}numPr")
            if numPr is not None:
                return f"- {text}"
        except Exception:
            pass

        return text

    def render_table(tbl):
        out = []
        for i, row in enumerate(tbl.rows):
            cells = []
            for cell in row.cells:
                # Cell text: ghep cac paragraph trong cell
                cell_text = " ".join(p.text.strip() for p in cell.paragraphs)
                cell_text = re.sub(r"\s+", " ", cell_text).strip()
                cells.append(cell_text)
            out.append("| " + " | ".join(cells) + " |")
            # Sau dong header, them separator
            if i == 0:
                out.append("|" + "|".join(["---"] * len(cells)) + "|")
        return out

    # Duyet phan tu theo thu tu trong body (paragraph + table xen ke)
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag
        if tag.endswith("}p"):
            p = Paragraph(child, doc)
            line = render_paragraph(p)
            lines.append(line)
        elif tag.endswith("}tbl"):
            tbl = Table(child, doc)
            lines.extend(render_table(tbl))
            lines.append("")  # dong trong sau bang

    return "\n".join(lines)


def _clean_cell(s):
    if s is None:
        return ""
    s = s.strip()
    s = re.sub(r"\*+", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _norm_text(s):
    """Normalize Vietnamese/markdown text for tolerant label/header matching."""
    s = _clean_cell(s or "").lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.replace("đ", "d")
    s = re.sub(r"[^a-z0-9/]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _line_value_after_label(line, labels):
    """Return value after labels like 'Tên bảng:' with optional bullet/bold."""
    raw = _clean_cell(line or "")
    raw = re.sub(r"^\s*[-+*]\s*", "", raw)
    raw = re.sub(r"^\s*\d+[\.)]\s*", "", raw)
    raw = re.sub(r"^\s*[a-z]\)\s*", "", raw, flags=re.IGNORECASE)
    for label in labels:
        m = re.match(rf"^\s*{re.escape(label)}\s*[:：-]\s*(.+?)\s*$", raw,
                     re.IGNORECASE)
        if m:
            return m.group(1).strip()

        norm_raw = _norm_text(raw)
        norm_label = _norm_text(label)
        if norm_raw == norm_label:
            return ""
        if norm_raw.startswith(norm_label + " "):
            value = raw[len(label):].lstrip(" :：-\t")
            return value.strip()
    return None


def _header_index(cells, candidates):
    norm_candidates = [_norm_text(c) for c in candidates]
    for idx, cell in enumerate(cells):
        norm = _norm_text(cell)
        if any(cand and cand in norm for cand in norm_candidates):
            return idx
    return None


def _looks_like_field_header(cells):
    name_idx = _header_index(cells, [
        "Tên trường", "Trường", "Field", "Column name", "Tên cột",
    ])
    type_idx = _header_index(cells, [
        "Kiểu dữ liệu", "Định dạng", "Data type", "Datatype", "Type",
    ])
    desc_idx = _header_index(cells, [
        "Mô tả", "Ý nghĩa", "Description", "Meaning",
    ])
    return name_idx is not None and (type_idx is not None or desc_idx is not None)


def _looks_like_format_b_header(cells):
    if not _looks_like_field_header(cells):
        return False
    joined = " ".join(_norm_text(c) for c in cells)
    return (
        len(cells) >= 7
        and any(token in joined for token in (
            "nullable", "null", "pk", "fk", "default", "mac dinh", "unique"
        ))
    )


def _cell_at(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return _clean_cell(row[idx])


def _is_blank_row(row):
    return not any(_clean_cell(cell) for cell in row)


def _is_stt_text(text):
    norm = _norm_text(text)
    return norm in {"stt", "tt", "no", "no"} or norm.startswith("stt ")


def _is_index_header(cells):
    joined = " ".join(_norm_text(c) for c in cells)
    return "ten index" in joined or "cot tham chieu" in joined


def _field_header_map(cells):
    """Map flexible field-table headers to canonical column names."""
    if _is_index_header(cells):
        return None

    mapping = {}
    for idx, cell in enumerate(cells):
        norm = _norm_text(cell)
        if not norm:
            continue
        if norm in {"stt", "tt", "no", "no"}:
            mapping.setdefault("stt", idx)
        elif any(token in norm for token in (
            "ten truong", "field name", "filed name", "column name", "ten cot",
        )):
            mapping.setdefault("name", idx)
        elif any(token in norm for token in (
            "kieu du lieu", "data type", "datatype", "type", "dinh dang",
        )):
            mapping.setdefault("datatype", idx)
        elif norm in {"length", "len", "do dai"} or "length" in norm:
            mapping.setdefault("length", idx)
        elif "nullable" in norm or norm in {"null", "null able"}:
            mapping.setdefault("nullable", idx)
        elif "unique" in norm or "uni que" in norm:
            mapping.setdefault("unique", idx)
        elif "p/f" in cell.lower() or "pk/fk" in norm or "pf key" in norm:
            mapping.setdefault("pk_fk", idx)
        elif norm in {"p f key", "pf", "key"}:
            mapping.setdefault("pk_fk", idx)
        elif "mac dinh" in norm or "default" in norm or norm == "mac":
            mapping.setdefault("default", idx)
        elif any(token in norm for token in ("mo ta", "note", "y nghia", "description")):
            mapping.setdefault("description", idx)

    has_name = "name" in mapping
    has_type = "datatype" in mapping
    if has_name and has_type:
        return mapping
    return None


def _looks_like_field_name(value):
    value = _clean_cell(value)
    if not value:
        return False
    norm = _norm_text(value)
    if norm in {"na", "n/a", "constraint", "index", "trigger"}:
        return False
    if len(value) > 80:
        return False
    return bool(re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", value))


def _looks_like_datatype(value):
    value = _clean_cell(value)
    if not value:
        return False
    compact = re.sub(r"\s+", "", value).upper()
    return bool(re.match(
        r"^(N?VAR)?CHAR(?:\(\d+\))?$|^N?TEXT$|^STRING$|^UUID$|^GUID$|"
        r"^UNIQUEIDENTIFIER$|^INT$|^INTEGER$|^BIGINT$|^SMALLINT$|^TINYINT$|"
        r"^DECIMAL(?:\(\d+,\d+\))?$|^NUMERIC(?:\(\d+,\d+\))?$|^NUMBER(?:\([^)]*\))?$|"
        r"^FLOAT$|^DOUBLE$|^REAL$|^MONEY$|^DATE$|^DATETIME$|^DATETIME2$|"
        r"^TIMESTAMP$|^TIME$|^BOOLEAN$|^BOOL$|^BIT$|^ENUM$",
        compact,
    ))


def _infer_field_map_from_rows(rows):
    """Infer map for continuation tables that lost their header on page split."""
    if not rows:
        return None
    max_cols = max(len(r) for r in rows)
    best = None
    best_score = -1
    for name_idx in range(max_cols):
        for type_idx in range(max_cols):
            if name_idx == type_idx:
                continue
            score = 0
            for row in rows[:8]:
                if _is_blank_row(row):
                    continue
                if _looks_like_field_name(_cell_at(row, name_idx)):
                    score += 1
                if _looks_like_datatype(_cell_at(row, type_idx)):
                    score += 1
            if score > best_score:
                best_score = score
                best = {"name": name_idx, "datatype": type_idx}
    if best_score >= 4:
        if max_cols >= 8:
            best.update({"nullable": 3, "unique": 4, "pk_fk": 5,
                         "default": 6, "description": 7})
        elif max_cols >= 6:
            best.update({"length": 3, "description": max_cols - 1})
        elif max_cols >= 4:
            best.update({"description": 3})
        return best
    return None


def _normalize_pk_fk(value):
    raw = _clean_cell(value)
    norm = _norm_text(raw).upper()
    if norm in {"P", "PK", "PRIMARY KEY"}:
        return "PK"
    if norm in {"F", "FK", "FOREIGN KEY"}:
        return "FK"
    return raw


def _normalize_nullable(value):
    raw = _clean_cell(value)
    norm = _norm_text(raw)
    if norm in {"n", "no", "not null"}:
        return "NO"
    if norm in {"y", "yes", "x", "true"}:
        return "YES"
    return raw


def _datatype_with_length(datatype, length):
    datatype = _clean_cell(datatype)
    length = _clean_cell(length)
    if datatype and length and re.match(r"^\d+$", length) and "(" not in datatype:
        if datatype.lower() in {"varchar", "nvarchar", "char", "nchar"}:
            return f"{datatype}({length})"
    return datatype


def _field_from_row(row, mapping):
    name = _cell_at(row, mapping.get("name"))
    datatype = _datatype_with_length(
        _cell_at(row, mapping.get("datatype")),
        _cell_at(row, mapping.get("length")),
    )
    if not _looks_like_field_name(name):
        return None
    return FieldInfo(
        name=name,
        datatype=datatype,
        nullable=_normalize_nullable(_cell_at(row, mapping.get("nullable"))),
        unique=_cell_at(row, mapping.get("unique")),
        pk_fk=_normalize_pk_fk(_cell_at(row, mapping.get("pk_fk"))),
        default=_cell_at(row, mapping.get("default")),
        description=_cell_at(row, mapping.get("description")),
    )


def _extract_table_name_from_heading(text):
    raw = _clean_cell(text)
    raw = re.sub(r"^#+\s*", "", raw)
    raw = re.sub(r"^\s*[a-z]\)\s*", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", raw)

    labeled = _line_value_after_label(raw, ["Tên bảng", "Table name"])
    if labeled:
        return labeled.strip(), ""

    m = re.search(
        r"\bBảng\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:[:\-]\s*|\((.*?)\))?",
        raw,
        re.IGNORECASE,
    )
    if m:
        return m.group(1).strip(), _clean_cell(m.group(2) or "")

    m = re.match(r"([A-Za-z_][A-Za-z0-9_]*)\s*(?:[:\-]\s*|\((.*?)\))?", raw)
    if m and ("_" in m.group(1) or raw.lower().startswith("table ")):
        return m.group(1).strip(), _clean_cell(m.group(2) or "")

    norm = _norm_text(raw)
    if raw and len(raw) <= 120 and not any(token in norm for token in (
        "danh sach truong", "constraint", "index", "trigger", "nullable",
        "kieu du lieu", "mo ta", "muc dich", "na", "n/a",
    )):
        return raw.rstrip(":"), ""

    return "", raw.rstrip(":")


def _parse_docx_flexible(docx_path):
    """Parse many CSDL document variants directly from DOCX block order."""
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    docx = Document(docx_path)
    blocks = []

    def paragraph_text(p):
        return re.sub(r"\s+", " ", p.text or "").strip()

    def table_rows(tbl):
        rows = []
        for row in tbl.rows:
            values = []
            for cell in row.cells:
                text = " ".join(paragraph_text(p) for p in cell.paragraphs)
                values.append(_clean_cell(text))
            rows.append(values)
        return rows

    body = docx.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            text = paragraph_text(Paragraph(child, docx))
            if text:
                blocks.append(("p", text))
        elif child.tag.endswith("}tbl"):
            blocks.append(("table", table_rows(Table(child, docx))))

    tables = []
    current = None
    pending_name = ""
    pending_desc = ""
    recent_paragraphs = []

    for kind, value in blocks:
        if kind == "p":
            text = value
            label_name = _line_value_after_label(text, ["Tên bảng", "Table name"])
            if label_name:
                pending_name = label_name
                pending_desc = ""
                current = None
            else:
                label_desc = _line_value_after_label(
                    text, ["Mô tả", "Mục đích", "Description", "Purpose"]
                )
                if label_desc is not None:
                    pending_desc = label_desc
                    if current and not current.description:
                        current.description = label_desc
                else:
                    name, desc = _extract_table_name_from_heading(text)
                    norm = _norm_text(text)
                    if name and "index" not in norm and "constraint" not in norm and "trigger" not in norm:
                        pending_name = name
                        pending_desc = desc
                        current = None

            recent_paragraphs.append(text)
            recent_paragraphs = recent_paragraphs[-6:]
            continue

        rows = [row for row in value if not _is_blank_row(row)]
        if not rows:
            continue

        mapping = None
        data_rows = []
        for row_idx, row in enumerate(rows[:4]):
            mapping = _field_header_map(row)
            if mapping:
                data_rows = rows[row_idx + 1:]
                break

        continuation = False
        if not mapping and current:
            mapping = _infer_field_map_from_rows(rows)
            if mapping:
                data_rows = rows
                continuation = True

        if not mapping:
            continue

        if not continuation:
            table_name = pending_name
            table_desc = pending_desc
            if not table_name:
                for prev in reversed(recent_paragraphs):
                    table_name, table_desc = _extract_table_name_from_heading(prev)
                    if table_name:
                        break
            if not table_name:
                table_name = f"Table_{len(tables) + 1}"
            current = TableInfo(name=table_name, description=table_desc)
            tables.append(current)

        for row in data_rows:
            fld = _field_from_row(row, mapping)
            if fld:
                current.fields.append(fld)

    cleaned = []
    seen_tables = set()
    for table in tables:
        if not table.fields:
            continue
        key = table.name.lower()
        if key in seen_tables:
            existing = next(t for t in cleaned if t.name.lower() == key)
            existing.fields.extend(table.fields)
            continue
        seen_tables.add(key)
        _infer_meta_flexible(table)
        _parse_constraints(table)
        cleaned.append(table)
    return cleaned


def _infer_meta_flexible(t):
    """Fill missing PK/FK/nullable hints from common CSDL conventions."""
    fk_pat = re.compile(
        r"(?:liên kết|tham chiếu|fk)\s+(?:sang|đến|với|tới)?\s*(?:bảng\s+)?(\w+)",
        re.IGNORECASE,
    )
    for idx, f in enumerate(t.fields):
        if not f.nullable:
            f.nullable = "YES"
        if f.pk_fk.upper() in {"P", "PK"}:
            f.pk_fk = "PK"
        elif f.pk_fk.upper() in {"F", "FK"}:
            f.pk_fk = "FK"
        elif idx == 0 and f.name.lower() in {"id", t.name.lower() + "_id"}:
            f.pk_fk = "PK"
        elif idx == 0 and f.datatype.lower() in {"uniqueidentifier", "uuid", "guid"}:
            f.pk_fk = "PK"

        if f.unique.upper() in {"X", "TRUE", "YES"} and not f.unique == "x":
            f.unique = "x"

        if f.pk_fk == "FK" and not f.ref_table:
            m = fk_pat.search(f.description)
            if m:
                f.ref_table = m.group(1)
                f.ref_field = "id"
            elif f.name.lower().endswith("_id"):
                f.ref_table = f.name[:-3]
                f.ref_field = "id"


def _parse_md_table(lines, start):
    rows = []
    i = start
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.startswith("|"):
            break
        if re.match(r"^\s*\|\s*[-:\s|]+\s*\|\s*$", line):
            i += 1
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        cells = [_clean_cell(c) for c in cells]
        rows.append(cells)
        i += 1
    return rows, i


def _detect_format(text):
    """Format A = 4 cot, Format B = 8 cot (co Nullable, PK/FK...)."""
    for line in text.split("\n"):
        if _line_value_after_label(line, ["Tên bảng", "Table name"]) is not None:
            return "A"

    for line in text.split("\n"):
        if not line.strip().startswith("|"):
            continue
        cells = [_clean_cell(c) for c in line.strip().strip("|").split("|")]
        if _looks_like_format_b_header(cells):
            return "B"
        if _looks_like_field_header(cells):
            return "A"
    return "A"


def _extract_system_name(text):
    m = re.search(r"\*\*(?:HỆ THỐNG\s+)?(?:PHẦN MỀM\s+)?"
                  r"(QUẢN LÝ[^*\n]+?)\s*\*\*",
                  text, re.IGNORECASE)
    if m:
        name = m.group(1).strip()
        return "Hệ thống " + " ".join(w.capitalize() for w in name.split())
    m = re.search(r"\*\*CƠ SỞ DỮ LIỆU[^*]*?(?:NGHIỆP VỤ\s+)?"
                  r"([^*\n]+?)\*\*", text)
    if m:
        return "Hệ thống quản lý " + " ".join(
            w.capitalize() for w in m.group(1).strip().split())
    m = re.search(r"^#\s+(.+?)$", text, re.MULTILINE)
    if m:
        return m.group(1).strip()
    return ""


def _parse_format_a(text):
    """Format A: '- Tên bảng: X' + bảng 4 cột."""
    tables = []
    lines = text.split("\n")
    cur = None
    i = 0
    while i < len(lines):
        line = lines[i]
        table_name = _line_value_after_label(line, ["Tên bảng", "Table name"])
        if table_name:
            cur = TableInfo(name=table_name.strip())
            tables.append(cur)
            i += 1
            continue

        desc = _line_value_after_label(line, ["Mô tả", "Mục đích", "Description", "Purpose"])
        if desc is not None and cur:
            cur.description = desc.strip()
            i += 1
            continue

        if line.strip().startswith("|") and cur:
            rows, ni = _parse_md_table(lines, i)
            if rows:
                if _looks_like_field_header(rows[0]) and not _looks_like_format_b_header(rows[0]):
                    name_idx = _header_index(rows[0], [
                        "Tên trường", "Trường", "Field", "Column name", "Tên cột",
                    ])
                    type_idx = _header_index(rows[0], [
                        "Kiểu dữ liệu", "Định dạng", "Data type", "Datatype", "Type",
                    ])
                    desc_idx = _header_index(rows[0], [
                        "Mô tả", "Ý nghĩa", "Description", "Meaning",
                    ])
                    name_idx = 1 if name_idx is None and len(rows[0]) > 1 else name_idx
                    type_idx = 2 if type_idx is None and len(rows[0]) > 2 else type_idx
                    desc_idx = 3 if desc_idx is None and len(rows[0]) > 3 else desc_idx
                    for r in rows[1:]:
                        if not any(c.strip() for c in r):
                            continue
                        fn = r[name_idx].strip() if name_idx is not None and len(r) > name_idx else ""
                        if not fn:
                            continue
                        cur.fields.append(FieldInfo(
                            name=fn,
                            datatype=(r[type_idx].strip()
                                      if type_idx is not None and len(r) > type_idx else ""),
                            description=(r[desc_idx].strip()
                                         if desc_idx is not None and len(r) > desc_idx else ""),
                        ))
            i = ni
            continue
        i += 1

    for t in tables:
        _infer_meta_a(t)
    return tables


def _infer_meta_a(t):
    """Suy luan PK/FK/nullable cho format A tu mo ta."""
    fk_pat = re.compile(r"liên kết\s+(?:sang|đến|với|tới)\s+bảng\s+(\w+)",
                        re.IGNORECASE)
    for idx, f in enumerate(t.fields):
        dt = f.datatype.lower()
        if idx == 0 and "uniqueidentifier" in dt:
            f.pk_fk = "PK"
            f.nullable = "NO"
            f.unique = "x"
            continue
        m = fk_pat.search(f.description)
        if m:
            f.pk_fk = "FK"
            f.nullable = "YES"
            f.ref_table = m.group(1)
            f.ref_field = f.name
            continue
        f.nullable = "YES"


def _parse_format_b(text):
    """Format B: '## Bảng X (mô tả)' + bảng 8 cột."""
    tables = []
    lines = text.split("\n")
    cur = None
    in_constraint = False
    table_re = re.compile(
        r"^#{1,3}\s*Bảng\s+([A-Za-z_][A-Za-z0-9_]*)"
        r"\s*(?:\((.+?)\))?\s*$"
    )
    i = 0
    while i < len(lines):
        line = lines[i]
        m = table_re.match(line.strip())
        if m:
            in_constraint = False
            cur = TableInfo(name=m.group(1).strip(),
                            description=(m.group(2) or "").strip())
            tables.append(cur)
            i += 1
            continue
        if cur and re.match(r"^#{2,4}\s*Constraint",
                            line.strip(), re.IGNORECASE):
            in_constraint = True
            i += 1
            continue
        if in_constraint and re.match(r"^#{1,4}\s+", line.strip()):
            in_constraint = False
        if in_constraint and cur and line.strip().startswith("-"):
            cur.constraints_raw.append(line.strip()[1:].strip())
            i += 1
            continue
        if line.strip().startswith("|") and cur:
            rows, ni = _parse_md_table(lines, i)
            if rows:
                hdr = " ".join(rows[0]).lower()
                if "trường" in hdr or "ten truong" in hdr:
                    for r in rows[1:]:
                        if len(r) < 8 or not any(c.strip() for c in r):
                            continue
                        fn = r[1].strip()
                        if not fn:
                            continue
                        n_raw = r[3].strip().upper()
                        if n_raw == "N":
                            nullable = "NO"
                        elif n_raw == "Y":
                            nullable = "YES"
                        else:
                            nullable = "YES" if not n_raw else n_raw
                        p_raw = r[5].strip().upper()
                        if p_raw == "P":
                            pk_fk = "PK"
                        elif p_raw == "F":
                            pk_fk = "FK"
                        else:
                            pk_fk = p_raw
                        cur.fields.append(FieldInfo(
                            name=fn,
                            datatype=r[2].strip(),
                            nullable=nullable,
                            unique=r[4].strip(),
                            pk_fk=pk_fk,
                            default=r[6].strip(),
                            description=r[7].strip(),
                        ))
            i = ni
            continue
        i += 1
    for t in tables:
        _parse_constraints(t)
        # Suy luan ref_table cho FK chua co constraint
        for f in t.fields:
            if (f.pk_fk == "FK" and not f.ref_table
                    and f.name.lower().endswith("_id")):
                f.ref_table = f.name[:-3]
                f.ref_field = "id"
    return tables


def _parse_constraints(t):
    pat = re.compile(
        r"(?:FK[:\s]+|PK[:\s]+)?(\w+)\s+liên kết\s+"
        r"(?:với|đến bảng|đến|tới|sang)\s+(?:bảng\s+)?(\w+)\.(\w+)",
        re.IGNORECASE,
    )
    for c in t.constraints_raw:
        m = pat.search(c)
        if m:
            fn, rt, rf = m.group(1), m.group(2), m.group(3)
            for f in t.fields:
                if f.name == fn:
                    f.ref_table = rt
                    f.ref_field = rf
                    if not f.pk_fk:
                        f.pk_fk = "FK"
                    break


def parse_document(docx_path):
    text = _extract_markdown(docx_path)
    # Do not infer source system from document titles/headings. The source
    # system must be explicitly provided; otherwise the Excel column stays blank.
    sys_name = ""

    tables = _parse_docx_flexible(docx_path)
    if tables:
        return DocumentInfo(system_name=sys_name, format_type="FLEX", tables=tables)

    fmt = _detect_format(text)
    tables = _parse_format_a(text) if fmt == "A" else _parse_format_b(text)
    return DocumentInfo(system_name=sys_name, format_type=fmt, tables=tables)


# ============================================================================
# RULE ENGINE
# ============================================================================

def _check_dqc_003(field, table, config):
    rule = config["rules"]["DQC_003"]
    if not rule.get("enabled"):
        return False
    if rule.get("exclude_pk", True) and field.pk_fk == "PK":
        return False
    if not rule_datatype_matches(field, rule, config):
        return False
    if rule.get("trigger_pk_fk_fk") and field.pk_fk:
        if "F" in field.pk_fk.upper():
            return True
    if rule.get("trigger_name_id_uuid"):
        if (field.name.lower().endswith("_id")
                and datatype_matches(field.datatype,
                                     config["datatype_groups"], ["uuid"])
                and field.pk_fk != "PK"):
            return True
    if _matches_text_condition(
            field, table, rule, "include_keywords", "include_targets",
            ("description_keywords",), ["description"]):
        return True
    return False


def _check_dqc_005(field, table, config):
    rule = config["rules"]["DQC_005"]
    if not rule.get("enabled"):
        return False
    if not rule_datatype_matches(field, rule, config):
        return False
    if rule.get("include_keywords"):
        return _matches_text_condition(field, table, rule, "include_keywords", "include_targets")
    return True


def _check_dqc_006(field, table, config):
    rule = config["rules"]["DQC_006"]
    if not rule.get("enabled"):
        return False
    if not rule_datatype_matches(field, rule, config):
        return False
    keywords = _rule_values(rule, "include_keywords", "keywords")
    if not keywords:
        return True
    return _matches_text_condition(
        field, table, rule, "include_keywords", "include_targets",
        ("keywords",), ["field_name", "description"],
    )


def _check_dqc_008(field, table, config):
    rule = config["rules"]["DQC_008"]
    if not rule.get("enabled"):
        return False
    if not rule_datatype_matches(field, rule, config):
        return False
    is_date = datatype_matches(field.datatype,
                                config["datatype_groups"], ["date"])
    is_text = datatype_matches(field.datatype,
                                config["datatype_groups"], ["text"])
    if is_date:
        return True
    if is_text:
        keywords = _rule_values(rule, "include_keywords", "keywords")
        if not keywords:
            return True
        return _matches_text_condition(
            field, table, rule, "include_keywords", "include_targets",
            ("keywords",), ["field_name", "description"],
        )
    return False


def _find_pair_fields(table, rule):
    result = set()
    names_lower = {f.name.lower(): f.name for f in table.fields}
    pairs = rule.get("pair_terms", rule.get("pair_prefixes", []))
    targets = rule.get("pair_targets", ["field_name"])
    for pair in pairs:
        if len(pair) < 2:
            continue
        p1, p2 = pair[0].lower(), pair[1].lower()
        if "field_name" in targets:
            for f in table.fields:
                nl = f.name.lower()
                if nl.startswith(p2):
                    partner = p1 + nl[len(p2):]
                    if partner in names_lower:
                        result.add(f.name)
                        if rule.get("assign_to_both", True):
                            result.add(names_lower[partner])
                elif nl.startswith(p1):
                    partner = p2 + nl[len(p1):]
                    if partner in names_lower and rule.get("assign_to_both", True):
                        result.add(f.name)
                        result.add(names_lower[partner])
        context_values = []
        if "table_name" in targets:
            context_values.append(table.name.lower())
        if "description" in targets:
            context_values.extend(f.description.lower() for f in table.fields if f.description)
        if any(p1 in text and p2 in text for text in context_values):
            result.update(f.name for f in table.fields)
    return result


def _find_dqc_010_pairs(table, config):
    rule = config["rules"]["DQC_010"]
    if not rule.get("enabled"):
        return set()
    return _find_pair_fields(table, rule)


def _check_std_003(field, table, config):
    rule = config["rules"].get("STD_003", {})
    if not rule.get("enabled"):
        return False
    return _matches_text_condition(
        field, table, rule, "include_keywords", "include_targets",
        ("keywords",), ["field_name", "description"],
    )


def _check_std_009(field, table, config):
    rule = config["rules"].get("STD_009", {})
    if not rule.get("enabled"):
        return False
    if not rule_datatype_matches(field, rule, config):
        return False
    nl = field.name.lower()
    for suf in rule.get("blacklist_suffix", []):
        if nl.endswith(suf.lower()):
            return False
    if _is_text_excluded(field, table, rule):
        return False
    if match_any_pattern(field.name, rule.get("name_patterns", [])):
        return True
    if _matches_text_condition(
            field, table, rule, "include_keywords", "include_targets",
            ("description_keywords",), ["description"]):
        return True
    return False


def _check_std_011(field, table, config):
    rule = config["rules"].get("STD_011", {})
    if not rule.get("enabled"):
        return False
    if not rule_datatype_matches(field, rule, config):
        return False
    if match_any_pattern(field.name, rule.get("name_patterns", [])):
        return True
    if _matches_text_condition(
            field, table, rule, "include_keywords", "include_targets",
            ("description_keywords",), ["description"]):
        return True
    return False


def _check_std_013(field, table, config):
    rule = config["rules"].get("STD_013", {})
    if not rule.get("enabled"):
        return False
    if not rule_datatype_matches(field, rule, config):
        return False
    if rule.get("include_keywords"):
        return _matches_text_condition(field, table, rule, "include_keywords", "include_targets")
    return True


def _get_std_004_or_012(field, config):
    std_004 = config["rules"].get("STD_004", {})
    std_012 = config["rules"].get("STD_012", {})
    is_date = datatype_matches(field.datatype,
                                config["datatype_groups"], ["date"])
    is_text = datatype_matches(field.datatype,
                                config["datatype_groups"], ["text"])
    if (std_004.get("enabled") and is_date
            and rule_datatype_matches(field, std_004, config)):
        return "STD_004"
    if (std_012.get("enabled") and is_text
            and rule_datatype_matches(field, std_012, config)):
        return "STD_012"
    return None


def _rule_references(rule, modern_key, legacy_key):
    if modern_key in rule:
        refs = rule.get(modern_key, [])
        return refs if isinstance(refs, list) else [refs]
    value = rule.get(legacy_key)
    return [value] if value else []


def _filter_candidate_rules(dqc, std, field, table, config):
    """Lọc rule theo loại trừ text và điều kiện phụ thuộc do người dùng cấu hình."""
    candidates = dqc + std
    include_handled_in_trigger = {
        "DQC_003", "DQC_005", "DQC_006", "DQC_008",
        "STD_003", "STD_009", "STD_011", "STD_013",
    }
    retained = []
    for code in candidates:
        rule = config["rules"].get(code, {})
        if not rule_reference_table_matches(code, rule, table, config):
            continue
        if _is_text_excluded(field, table, rule):
            continue
        if not rule_datatype_matches(field, rule, config):
            continue
        if (code != "DQC_010" and rule.get("pair_terms")
                and field.name not in _find_pair_fields(table, rule)):
            continue
        if (code not in include_handled_in_trigger
                and rule.get("include_keywords")
                and not _matches_text_condition(
                    field, table, rule, "include_keywords", "include_targets")):
            continue
        retained.append(code)

    changed = True
    while changed:
        changed = False
        active = set(retained)
        next_retained = []
        for code in retained:
            rule = config["rules"].get(code, {})
            required = _rule_references(rule, "requires_rules", "requires")
            excluded_by = _rule_references(rule, "exclude_if_rules", "not_if")
            if any(required_code not in active for required_code in required):
                changed = True
                continue
            if any(excluded_code in active for excluded_code in excluded_by):
                changed = True
                continue
            next_retained.append(code)
        retained = next_retained

    dqc_set = set(dqc)
    return ([code for code in retained if code in dqc_set],
            [code for code in retained if code not in dqc_set])


def assign_rules_for_table(table, config):
    skip_set = {s.lower() for s in config.get("skip_fields", [])}
    pair_fields = _find_dqc_010_pairs(table, config)

    for f in table.fields:
        if f.name.lower() in skip_set:
            f.dqc_codes = []
            f.std_codes = []
            continue

        dqc, std = [], []

        if _check_dqc_003(f, table, config):
            dqc.append("DQC_003")
        if _check_dqc_005(f, table, config):
            dqc.append("DQC_005")
        if _check_dqc_006(f, table, config):
            dqc.append("DQC_006")
            if _check_std_003(f, table, config):
                std.append("STD_003")
            else:
                if config["rules"].get("STD_002", {}).get("enabled"):
                    std.append("STD_002")
        if _check_dqc_008(f, table, config):
            dqc.append("DQC_008")
            sdt = _get_std_004_or_012(f, config)
            if sdt:
                std.append(sdt)
        if f.name in pair_fields:
            dqc.append("DQC_010")
        if _check_std_009(f, table, config):
            std.append("STD_009")
        if _check_std_011(f, table, config):
            std.append("STD_011")
        if _check_std_013(f, table, config):
            std.append("STD_013")

        dqc, std = _filter_candidate_rules(dqc, std, f, table, config)
        seen = set()
        dqc = [x for x in dqc if not (x in seen or seen.add(x))]
        seen = set()
        std = [x for x in std if not (x in seen or seen.add(x))]

        f.dqc_codes = dqc
        f.std_codes = std


def assign_rules(tables, config):
    for t in tables:
        assign_rules_for_table(t, config)


# ============================================================================
# TABLE CLASSIFIER (Master / Event / Reference)
# ============================================================================

def classify_table(table, ref_patterns=None):
    """Uu tien: prefix DanhMuc/DM_ -> Reference. Sau do fallback heuristic."""
    name_lower = table.name.lower()
    if ref_patterns and is_reference_table(table.name, ref_patterns):
        return "Reference"
    event_kws = ["_history", "_log", "_case", "_report", "_event",
                 "_record", "_transaction", "_audit", "_trail",
                 "lich_su", "_progress"]
    if any(kw in name_lower for kw in event_kws):
        return "Event"

    field_names = [f.name.lower() for f in table.fields]
    has_code = any(n == "code" or n.endswith("_code") for n in field_names)
    has_name = any(n == "name" or n == "ten" or n.endswith("_name")
                   for n in field_names)
    has_parent = any(n in ("parent_id", "macha", "ma_cha", "parentid")
                     for n in field_names)
    technical = {"id", "created_by", "created_date", "last_modified_by",
                 "last_modified_date", "deleted", "ngaytao", "nguoitao",
                 "ngaycapnhat", "nguoicapnhat"}
    n_business = sum(1 for f in table.fields
                     if f.name.lower() not in technical)

    if has_parent and has_code and has_name and n_business < 10:
        return "Reference"
    if has_code and has_name and n_business < 8:
        return "Reference"
    return "Master"


def apply_classification(tables, table_config_path=None, ref_patterns=None):
    """Phan loai bang. Co the override qua file JSON {table_name: type}."""
    overrides = {}
    if table_config_path:
        try:
            with open(table_config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "table_types" in data:
                    overrides = data["table_types"]
                else:
                    overrides = {k: v for k, v in data.items()
                                 if isinstance(v, str)}
        except FileNotFoundError:
            pass
    for t in tables:
        if t.name in overrides:
            t.table_type = overrides[t.name]
        else:
            t.table_type = classify_table(t, ref_patterns)


# ============================================================================
# TEMPLATES (mo ta DQC / STD)
# ============================================================================

DQC_SHORT = {
    "DQC_003": "Kiểm tra tính đồng bộ của dữ liệu",
    "DQC_005": "Kiểm tra giá trị ngày ngoài phạm vi nghiệp vụ",
    "DQC_006": "Kiểm tra định dạng số",
    "DQC_008": "Kiểm tra định dạng ngày tháng",
    "DQC_010": "Kiểm tra logic nghiệp vụ giữa các trường",
}

STD_SHORT = {
    "STD_002": "Chuẩn hóa dữ liệu số",
    "STD_003": "Chuẩn hóa trường tỷ lệ",
    "STD_004": "Chuẩn hóa và hiển thị ngày tháng năm",
    "STD_009": "Chuẩn hóa tên người",
    "STD_011": "Chuẩn hóa tên phân cấp cha-con",
    "STD_012": "Chuẩn hóa ngày tháng không đầy đủ",
    "STD_013": "Chuẩn hóa khoảng trắng",
}

DQC_DETAIL_TPL = {
    "DQC_003": (
        "DQC_003: Kiểm tra tính đồng bộ của dữ liệu\n\n"
        "Kiểm tra giá trị trường {field_name} có tồn tại trong bảng tham chiếu {ref_info}\n"
        "- Nếu tồn tại -> Pass\n"
        "- Nếu không tồn tại -> Gắn flag & gửi về Bad Zone"
    ),
    "DQC_005": (
        "DQC_005: Kiểm tra giá trị ngày ngoài phạm vi nghiệp vụ\n\n"
        "Kiểm tra giá trị trường {field_name} có nằm trong phạm vi nghiệp vụ hợp lệ\n"
        "(ví dụ: ngày sinh không vượt quá ngày hiện tại, ngày bắt đầu <= ngày kết thúc)\n"
        "- Nếu hợp lệ -> Pass\n"
        "- Nếu vi phạm -> Gắn flag & gửi về Bad Zone"
    ),
    "DQC_006": (
        "DQC_006: Kiểm tra định dạng số\n\n"
        "Kiểm tra giá trị trường {field_name} có đúng định dạng số\n"
        "- Nếu không vi phạm -> Pass\n"
        "- Nếu vi phạm -> Gắn flag & gửi về Bad Zone"
    ),
    "DQC_008": (
        "DQC_008: Kiểm tra định dạng ngày tháng\n\n"
        "Kiểm tra giá trị trường {field_name} có đúng định dạng ngày tháng\n"
        "- Nếu không vi phạm -> Pass\n"
        "- Nếu vi phạm -> Gắn flag & gửi về Bad Zone"
    ),
    "DQC_010": (
        "DQC_010: Kiểm tra logic nghiệp vụ giữa các trường trong cùng bảng\n\n"
        "Kiểm tra giá trị trường {field_name} đảm bảo logic với các trường liên quan\n"
        "(ví dụ: ngày kết thúc >= ngày bắt đầu)\n"
        "- Nếu hợp lệ -> Pass\n"
        "- Nếu vi phạm -> Gắn flag & gửi về Bad Zone"
    ),
}

STD_DETAIL_TPL = {
    "STD_002": (
        "STD_002: Chuẩn hóa dữ liệu số\n\n"
        "1. Trim khoảng trắng đầu/cuối chuỗi\n"
        "2. Loại bỏ ký tự không phải số (giữ dấu - cho số âm, dấu . cho thập phân)\n"
        "3. Chuẩn hóa dấu phân tách: dấu , (nghìn) loại bỏ, dấu . cho thập phân\n"
        "4. Nếu giá trị bản ghi đang có giá trị là 'NULL' thì đưa về NULL\n"
        "5. Trả về giá trị đã chuẩn hóa dạng số"
    ),
    "STD_003": (
        "STD_003: Chuẩn hóa trường tỷ lệ\n\n"
        "1. Chuẩn hóa theo STD_002\n"
        "2. Nếu giá trị chứa '%' -> loại bỏ và chia cho 100 (hoặc giữ nguyên tùy nghiệp vụ)\n"
        "3. Giới hạn giá trị trong khoảng [0, 1] hoặc [0, 100] tùy quy ước\n"
        "4. Làm tròn theo số chữ số thập phân quy định"
    ),
    "STD_004": (
        "STD_004: Chuẩn hóa và hiển thị ngày tháng năm\n\n"
        "Đảm bảo dữ liệu ngày tháng theo đúng định dạng cho các trường ngày "
        "tháng trong bảng dữ liệu, ví dụ: Ngày sinh, Ngày phát hành, Ngày cập nhật...\n"
        "Định dạng nguồn có thể là: dd/MM/yyyy (hoặc MM/yyyy, hoặc dd/MM/yyyy hh:mm)\n\n"
        "1. Convert dữ liệu về dạng Datetime\n"
        "2. Chuẩn hóa theo định dạng: YYYY-MM-DD (tùy nghiệp vụ)"
    ),
    "STD_009": (
        "STD_009: Chuẩn hóa tên người\n\n"
        "1. Trim khoảng trắng đầu/cuối\n"
        "2. Có nhiều hơn 1 khoảng trắng -> chuyển thành 1 khoảng trắng\n"
        "   VD: Nguyễn Văn  An -> Nguyễn Văn An\n"
        "3. Viết hoa toàn bộ ký tự (hoặc viết hoa chữ cái đầu mỗi từ - tùy quy ước)\n"
        "4. Loại bỏ ký tự không hợp lệ\n"
        "   VD: NgUyễn! VăN A -> NGUYỄN VĂN A"
    ),
    "STD_011": (
        "STD_011: Chuẩn hóa tên phân cấp cha-con\n\n"
        "1. Nhận diện cấu trúc phân cấp trong chuỗi dữ liệu theo ký tự phân "
        "tách, xuống dòng, thụt đầu dòng hoặc ký hiệu đầu mục (A., I., 1., ...)\n"
        "2. Xác định level và tách dữ liệu thành từng cấp cha/con tương ứng\n"
        "3. Trim khoảng trắng đầu/cuối tại từng cấp dữ liệu\n"
        "4. Chuẩn hóa nội dung text theo quy tắc đặt tên\n"
        "5. Loại bỏ ký hiệu đánh số hoặc ký tự chỉ mang tính trình bày\n"
        "6. Ghép lại thành đường dẫn chuẩn bằng dấu phân tách >"
    ),
    "STD_012": (
        "STD_012: Chuẩn hóa ngày tháng không đầy đủ\n\n"
        "Chuẩn hóa dữ liệu ngày tháng đã hợp lệ về định dạng thống nhất "
        "để phục vụ lưu trữ, tra cứu và tính toán.\n"
        "Ngày được chuẩn hóa phù hợp với nghiệp vụ khách hàng.\n\n"
        "1. Nếu precision = FULL → chuyển về YYYY-MM-DD\n"
        "2. Nếu precision = MONTH_YEAR → chuyển về YYYY-MM\n"
        "3. Nếu precision = YEAR_ONLY → giữ dạng YYYY\n"
        "4. Chuẩn hóa ký tự phân tách về dấu -\n"
        "5. Lưu thêm Precision_Level nếu cần\n\n"
        "Ví dụ: 01/02/2023 → 2023-02-01, 02/2023 → 2023-02, 2023 → 2023"
    ),
    "STD_013": (
        "STD_013: Chuẩn hóa khoảng trắng\n\n"
        "1. Trim khoảng trắng đầu chuỗi\n"
        "2. Trim khoảng trắng cuối chuỗi\n"
        "3. Nếu có nhiều hơn 1 khoảng trắng liên tiếp -> chuẩn hóa về 1\n"
        "4. Chuẩn hóa tab, xuống dòng nội bộ thành khoảng trắng đơn\n"
        "5. Giữ nguyên ký tự đặc biệt, dấu phân tách\n"
        "6. Nếu giá trị bản ghi đang là 'NULL' thì đưa về NULL\n"
        "7. Trả về giá trị đã chuẩn hóa"
    ),
}


def get_dqc_detail(code, field_name, ref_table="", ref_field=""):
    tpl = DQC_DETAIL_TPL.get(code, "")
    if not tpl:
        return ""
    ref_info = f"{ref_table}.{ref_field}" if ref_table else "(cần xác định)"
    return tpl.format(field_name=field_name, ref_info=ref_info)


def get_std_detail(code, field_name):
    return STD_DETAIL_TPL.get(code, "")


# ============================================================================
# EXCEL WRITER
# ============================================================================

HEADERS = [
    "STT", "Bảng đích", "Trường đích", "Định dạng", "Ý nghĩa",
    "Hệ thống nguồn", "Bảng nguồn", "Trường nguồn",
    "Nullable", "Unique", "PK / FK", "Mặc định", "Định dạng dữ liệu",
    "ETL Rules", "Mã kiểm tra", "Mã chuẩn hóa",
    "Đề xuất phương pháp kiểm tra", "Đề xuất phương pháp kiểm tra chi tiết",
    "Đề xuất phương pháp chuẩn hoá", "Đề xuất phương pháp chuẩn hoá chi tiết",
    "Ghi chú", "Loại bảng", "Mô tả bảng",
]

HEADER_FILLS = {
    range(0, 5): "F2F2F2",
    range(5, 8): "E7E6E6",
    range(8, 13): "FFF2CC",
    range(13, 14): "FFE699",
    range(14, 16): "F4B084",
    range(16, 20): "C6E0B4",
    range(20, 21): "BDD7EE",
    range(21, 23): "FFC7CE",
}

HIGHLIGHT_AUTO_FILL = "9FC5E8"


def _get_header_fill(col_idx):
    for r, color in HEADER_FILLS.items():
        if col_idx in r:
            return color
    return "FFFFFF"


def _load_existing(output_path):
    """Doc Excel cu de merge. Tra ve {(bang, truong): {col: val}}"""
    p = Path(output_path)
    if not p.exists():
        return {}
    try:
        wb = load_workbook(p)
        ws = wb.active
    except Exception:
        return {}
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v:
            headers[v.strip()] = c
    if "Bảng đích" not in headers or "Trường đích" not in headers:
        return {}
    existing = {}
    col_table = headers["Bảng đích"]
    col_field = headers["Trường đích"]
    for r in range(2, ws.max_row + 1):
        tn = ws.cell(r, col_table).value
        fn = ws.cell(r, col_field).value
        if not tn or not fn:
            continue
        row_data = {h: ws.cell(r, c).value for h, c in headers.items()}
        existing[(str(tn).strip(), str(fn).strip())] = row_data
    return existing


def _merge_codes(old_value, new_codes):
    """Giu thu tu code cu, them code moi chua co o cuoi."""
    if old_value is None:
        old_value = ""
    old_str = str(old_value).strip()
    old_list = ([c.strip() for c in old_str.split("\n") if c.strip()]
                if old_str else [])
    new_added = [c for c in new_codes if c not in old_list]
    return "\n".join(old_list + new_added), len(new_added) > 0


def write_excel(documents, output_path, mode="overwrite"):
    """Sinh Excel. mode: 'overwrite' / 'merge' (giu user edit)."""
    existing = _load_existing(output_path) if mode == "merge" else {}
    wb = Workbook()
    ws = wb.active
    ws.title = "Mô tả thiết kế"

    bold = Font(bold=True, size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
    thin = Side(border_style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    highlight = PatternFill("solid", fgColor=HIGHLIGHT_AUTO_FILL)

    for col_idx, header in enumerate(HEADERS):
        cell = ws.cell(row=1, column=col_idx + 1, value=header)
        cell.font = bold
        cell.alignment = center
        cell.fill = PatternFill("solid", fgColor=_get_header_fill(col_idx))
        cell.border = border

    stt = 0
    row_idx = 2
    COL_ETL, COL_DQ, COL_STD = 14, 15, 16
    for doc in documents:
        for table in doc.tables:
            for fld in table.fields:
                stt += 1
                dqc_codes = getattr(fld, "dqc_codes", [])
                std_codes = getattr(fld, "std_codes", [])
                key = (table.name, fld.name)
                old_row = existing.get(key, {})
                changed_cells = set()

                if mode == "merge" and old_row:
                    new_dq_str, added_dq = _merge_codes(
                        old_row.get("Mã kiểm tra", ""), dqc_codes)
                    new_std_str, added_std = _merge_codes(
                        old_row.get("Mã chuẩn hóa", ""), std_codes)
                    dqc_final = [c.strip() for c in new_dq_str.split("\n")
                                 if c.strip()]
                    std_final = [c.strip() for c in new_std_str.split("\n")
                                 if c.strip()]
                    if added_dq:
                        changed_cells.update([COL_DQ, COL_ETL])
                    if added_std:
                        changed_cells.update([COL_STD, COL_ETL])
                else:
                    dqc_final = dqc_codes
                    std_final = std_codes
                    if dqc_final:
                        changed_cells.update([COL_DQ, COL_ETL])
                    if std_final:
                        changed_cells.update([COL_STD, COL_ETL])

                etl_rules = "\n".join(dqc_final + std_final)
                ma_kt = "\n".join(dqc_final)
                ma_ch = "\n".join(std_final)
                pp_kt = "\n".join(f"{c}: {DQC_SHORT.get(c, '')}"
                                  for c in dqc_final)
                pp_ch = "\n".join(f"{c}: {STD_SHORT.get(c, '')}"
                                  for c in std_final)
                pp_kt_dt = "\n\n".join(
                    get_dqc_detail(c, fld.name,
                                   fld.ref_table, fld.ref_field)
                    for c in dqc_final
                )
                pp_ch_dt = "\n\n".join(
                    get_std_detail(c, fld.name) for c in std_final
                )

                he_thong_nguon = (
                    old_row.get("Hệ thống nguồn")
                    if mode == "merge" and old_row.get("Hệ thống nguồn")
                    else doc.system_name
                )
                bang_nguon = fld.ref_table or table.name
                truong_nguon = fld.ref_field or fld.name
                ghi_chu = (old_row.get("Ghi chú", "")
                           if mode == "merge" else "")

                row = [
                    stt, table.name, fld.name, fld.datatype, fld.description,
                    he_thong_nguon, bang_nguon, truong_nguon,
                    fld.nullable, fld.unique, fld.pk_fk, fld.default,
                    fld.datatype,
                    etl_rules, ma_kt, ma_ch,
                    pp_kt, pp_kt_dt, pp_ch, pp_ch_dt,
                    ghi_chu, table.table_type, table.description,
                ]
                for col_idx, val in enumerate(row):
                    cell = ws.cell(row=row_idx, column=col_idx + 1, value=val)
                    cell.alignment = left_top
                    cell.border = border
                    if (col_idx + 1) in changed_cells and val:
                        cell.fill = highlight
                row_idx += 1

    widths = {
        "A": 5, "B": 22, "C": 22, "D": 16, "E": 35,
        "F": 25, "G": 22, "H": 22,
        "I": 10, "J": 8, "K": 8, "L": 12, "M": 16,
        "N": 14, "O": 14, "P": 14,
        "Q": 28, "R": 50, "S": 28, "T": 50,
        "U": 18, "V": 12, "W": 30,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{row_idx - 1}"

    wb.save(output_path)
    return output_path


# ============================================================================
# CLI ORCHESTRATOR
# ============================================================================

def process(input_files, output_path, rule_config_path=None,
            table_config_path=None, system_name=None, mode="overwrite",
            verbose=False):
    """Pipeline xu ly co try/except tung buoc de bao loi cu the."""

    # Bước 1: Load config
    if verbose:
        print(f"[1] Loading config từ: {rule_config_path or '(default)'}")
    try:
        config = load_config(rule_config_path)
    except FileNotFoundError as e:
        print(f"[FATAL] Không tìm thấy file config: {rule_config_path}",
              file=sys.stderr)
        raise
    except json.JSONDecodeError as e:
        print(f"[FATAL] File config JSON không hợp lệ: {rule_config_path}",
              file=sys.stderr)
        print(f"        Dòng {e.lineno}, cột {e.colno}: {e.msg}",
              file=sys.stderr)
        raise
    ref_patterns = config.get("reference_table_patterns", [])

    # Bước 2: Parse từng file
    documents = []
    for idx, fp in enumerate(input_files, 1):
        if verbose:
            print(f"[2.{idx}] Đang xử lý: {fp}")
        else:
            print(f"  Đang đọc: {fp}")

        # 2a. Kiểm tra file tồn tại
        if not os.path.exists(fp):
            print(f"  [LỖI] File không tồn tại: {fp}", file=sys.stderr)
            continue
        if not os.path.isfile(fp):
            print(f"  [LỖI] Không phải file: {fp}", file=sys.stderr)
            continue
        if not fp.lower().endswith(".docx"):
            print(f"  [CẢNH BÁO] File không có đuôi .docx: {fp}",
                  file=sys.stderr)

        # 2b. Parse
        try:
            if verbose:
                print(f"      - parse_document...")
            doc = parse_document(fp)
        except Exception as e:
            print(f"  [LỖI khi đọc file {fp}]: {type(e).__name__}: {e}",
                  file=sys.stderr)
            if verbose:
                traceback.print_exc()
            continue

        if system_name:
            doc.system_name = system_name

        n_fields = sum(len(t.fields) for t in doc.tables)
        print(f"    Format: {doc.format_type} | "
              f"Hệ thống: {doc.system_name} | "
              f"Bảng: {len(doc.tables)} | Trường: {n_fields}")

        # 2c. Assign rules
        try:
            if verbose:
                print(f"      - assign_rules...")
            assign_rules(doc.tables, config)
        except Exception as e:
            print(f"  [LỖI khi gán rules cho {fp}]: {type(e).__name__}: {e}",
                  file=sys.stderr)
            traceback.print_exc()
            continue

        # 2d. Classify tables
        try:
            if verbose:
                print(f"      - apply_classification...")
            apply_classification(doc.tables, table_config_path, ref_patterns)
        except Exception as e:
            print(f"  [LỖI khi phân loại bảng]: {type(e).__name__}: {e}",
                  file=sys.stderr)
            traceback.print_exc()
            continue

        documents.append(doc)

    if not documents:
        print("\n[LỖI] Không có document nào parse thành công.",
              file=sys.stderr)
        return None

    # Bước 3: Write Excel
    if verbose:
        print(f"[3] Ghi Excel ra: {output_path} (mode={mode})")
    try:
        # Đảm bảo thư mục output tồn tại
        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
            if verbose:
                print(f"      - Đã tạo thư mục: {out_dir}")

        output = write_excel(documents, output_path, mode=mode)
    except PermissionError as e:
        print(f"[FATAL] Không có quyền ghi file: {output_path}",
              file=sys.stderr)
        print(f"        (File có thể đang mở trong Excel? "
              f"Đóng Excel rồi thử lại.)", file=sys.stderr)
        raise
    except Exception as e:
        print(f"[FATAL] Lỗi khi ghi Excel: {type(e).__name__}: {e}",
              file=sys.stderr)
        raise

    n_rows = sum(len(t.fields) for d in documents for t in d.tables)
    print(f"\n[OK] Đã sinh: {output} ({n_rows} dòng, mode={mode})")
    return output


def main():
    """Entry point voi error handling toan dien."""
    ap = argparse.ArgumentParser(
        description="Sinh Excel mô tả thiết kế CSDL từ tài liệu .docx",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("inputs", nargs="*", help="Một hoặc nhiều file .docx")
    ap.add_argument("-o", "--output", help="File Excel đầu ra (.xlsx)")
    ap.add_argument("--config", dest="rule_config",
                    help="File JSON config rule DQC/STD tùy chỉnh")
    ap.add_argument("--table-config", dest="table_config",
                    help="File JSON map {table_name: loại bảng}")
    ap.add_argument("--system-name", dest="system_name",
                    help="Tên hệ thống (override auto-detect)")
    ap.add_argument("--mode", choices=["overwrite", "merge"],
                    default="overwrite",
                    help="overwrite (mặc định) / merge (giữ user edit)")
    ap.add_argument("--dump-config", dest="dump_config",
                    help="Xuất config mặc định ra file JSON rồi thoát")
    ap.add_argument("--check", action="store_true",
                    help="Chỉ kiểm tra môi trường (deps, version) rồi thoát")
    ap.add_argument("--verbose", "-v", action="store_true",
                    help="In chi tiết từng bước (giúp debug)")
    args = ap.parse_args()

    # --check: chỉ in info môi trường rồi thoát
    if args.check:
        _print_env_info()
        _check_dependencies(verbose=True)
        print("[OK] Môi trường có thể chạy tool.")
        return

    # Trong verbose mode in env info luôn
    if args.verbose:
        _print_env_info()

    # --dump-config
    if args.dump_config:
        try:
            out = dump_config(args.dump_config)
            print(f"[OK] Đã xuất config mặc định ra: {out}")
            print("Bạn có thể chỉnh sửa rồi truyền qua --config <file>")
        except Exception as e:
            print(f"[LỖI] Không xuất được config: {type(e).__name__}: {e}",
                  file=sys.stderr)
            raise
        return

    # Validate args
    if not args.inputs:
        ap.error("Cần ít nhất 1 file đầu vào (hoặc dùng --dump-config / --check)")
    if not args.output:
        ap.error("Cần truyền -o <output.xlsx>")

    process(
        args.inputs, args.output,
        rule_config_path=args.rule_config,
        table_config_path=args.table_config,
        system_name=args.system_name,
        mode=args.mode,
        verbose=args.verbose,
    )


# ============================================================================
# ENTRY POINT - wrap toan bo trong try/except de bat moi loi
# ============================================================================
if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        # argparse hoac sys.exit() - cho qua
        raise
    except KeyboardInterrupt:
        print("\n[CANCELLED] Đã dừng bởi người dùng (Ctrl+C)", file=sys.stderr)
        sys.exit(130)
    except Exception as e:
        print("\n" + "=" * 70, file=sys.stderr)
        print("[CRASH] Tool gặp lỗi. Vui lòng copy TOÀN BỘ output dưới đây",
              file=sys.stderr)
        print("        và gửi cho người support để debug:", file=sys.stderr)
        print("=" * 70, file=sys.stderr)
        print(f"\n[Loại lỗi]    {type(e).__name__}", file=sys.stderr)
        print(f"[Thông điệp]  {e}", file=sys.stderr)
        print(f"\n[Môi trường]", file=sys.stderr)
        print(f"  Python:     {sys.version.split(chr(10))[0]}", file=sys.stderr)
        print(f"  Platform:   {platform.platform()}", file=sys.stderr)
        print(f"  Encoding:   {sys.stdout.encoding}", file=sys.stderr)
        print(f"  Args:       {sys.argv}", file=sys.stderr)
        print(f"\n[Traceback đầy đủ]", file=sys.stderr)
        traceback.print_exc()
        print("\n" + "=" * 70, file=sys.stderr)
        sys.exit(2)
