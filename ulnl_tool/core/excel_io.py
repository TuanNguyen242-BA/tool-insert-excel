"""
core/excel_io.py - Đọc Excel danh sách CN + build ULNL output.

Build ULNL gồm các bước (clone từ skill):
1. Copy template_base.xlsx
2. Update Trang bia
3. Build sections JSON (Web, Mobile, UT) từ danh sách functions
4. Call add_functions.py
5. Call update_phi_cn_params.py
6. Call update_tong_hop.py
7. Apply formatting overrides
"""
from __future__ import annotations
import sys, os, json, subprocess, shutil
import re
import unicodedata
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from ulnl_tool.core.data_model import Project, Function, CellStyle, SheetFormatting
from ulnl_tool.core.classifier import Classifier, ClassificationRule


class PngImage(OpenpyxlImage):
    """PNG image wrapper that avoids requiring Pillow for workbook output."""
    _id = 1
    _path = "/xl/media/image{0}.{1}"

    def __init__(self, path: str | Path):
        self.ref = str(path)
        self.format = "png"
        self.anchor = "A1"
        self.width, self.height = _png_dimensions(path)

    def _data(self):
        return Path(self.ref).read_bytes()


def _png_dimensions(path: str | Path):
    data = Path(path).read_bytes()[:24]
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n":
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    return 300, 80


# Chuẩn hóa header để khớp nhiều biến thể
def _norm(s) -> str:
    if s is None:
        return ""
    return str(s).strip().lower()


def _norm_key(s) -> str:
    text = unicodedata.normalize("NFD", _norm(s))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", text).strip()


def import_functions_from_excel(xlsx_path: str | Path, sheet_name: Optional[str] = None) -> list[Function]:
    """Đọc danh sách chức năng từ file Excel.

    File phải có các cột: STT, Mã CN, Cấp, Tên chức năng, Mô tả ngắn (Mô tả cũng OK).
    Có thể có thêm cột phụ - ignore.

    Tìm header row tự động (kiếm dòng có cả 'STT', 'Mã' và 'Tên').
    """
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        # Tìm sheet có 'chuc_nang' hoặc 'chức năng' trong tên
        target = None
        for s in wb.sheetnames:
            if any(kw in s.lower() for kw in ["chuc_nang", "chuc nang", "chức năng", "chuc-nang"]):
                target = s
                break
        ws = wb[target] if target else wb.active

    # Tìm header row
    header_row = None
    col_map = {}  # 'stt' -> col_index, 'ma' -> ..., 'cap' -> ..., 'name' -> ..., 'desc' -> ...
    for r in range(1, min(ws.max_row + 1, 20)):
        row_vals = [_norm(ws.cell(r, c).value) for c in range(1, min(ws.max_column + 1, 15))]
        row_keys = [_norm_key(ws.cell(r, c).value) for c in range(1, min(ws.max_column + 1, 15))]
        if "stt" in row_keys and any(v in {"ma cn", "ma chuc nang"} for v in row_keys):
            header_row = r
            for c, (v, key) in enumerate(zip(row_vals, row_keys), 1):
                if key == "stt":
                    col_map["stt"] = c
                elif key in ("ma cn", "ma chuc nang"):
                    col_map["ma"] = c
                elif key == "cap":
                    col_map["cap"] = c
                elif "ten" in key or key == "name":
                    col_map["name"] = c
                elif "mo ta" in key or key == "description":
                    if "desc" not in col_map:
                        col_map["desc"] = c
            break

    if header_row is None:
        raise ValueError("Không tìm thấy header row trong sheet. Cần có cột 'STT' và 'Mã CN'.")

    if not all(k in col_map for k in ("ma", "name")):
        raise ValueError(f"Thiếu cột bắt buộc. Tìm thấy: {col_map}")

    functions = []
    for r in range(header_row + 1, ws.max_row + 1):
        ma = ws.cell(r, col_map["ma"]).value
        name = ws.cell(r, col_map["name"]).value
        if not ma and not name:
            continue
        if ma is None or name is None:
            continue
        stt = ws.cell(r, col_map["stt"]).value if "stt" in col_map else r
        cap = ws.cell(r, col_map["cap"]).value if "cap" in col_map else 3
        desc = ws.cell(r, col_map["desc"]).value if "desc" in col_map else ""

        try:
            stt_int = int(stt) if stt is not None else r
        except (ValueError, TypeError):
            stt_int = r
        try:
            cap_int = int(cap) if cap is not None else 3
        except (ValueError, TypeError):
            cap_int = 3

        functions.append(Function(
            stt=stt_int, ma=str(ma).strip(), cap=cap_int,
            name=str(name).strip(),
            desc=str(desc).strip() if desc else "",
        ))
    return functions


# ============== BUILD ULNL OUTPUT ==============

def build_sections_json(project: Project) -> dict:
    """Build sections JSON cho add_functions.py từ project.functions.

    Cấu trúc:
    - Section A: WEB - các functions có include_web=True
    - Section B: MOBILE - các functions có include_mobile=True (nếu enable_mobile)
    - Section C: Unit test (tự động)
    """
    sections = []

    if project.enable_web:
        web_groups = _build_groups_from_functions(
            [f for f in project.functions if f.include_web],
            platform="web"
        )
        sections.append({
            "code": "A",
            "name": "NỖ LỰC CHỨC NĂNG WEB" if project.enable_mobile else "NỖ LỰC CHỨC NĂNG",
            "groups": web_groups
        })

    if project.enable_mobile:
        mobile_groups = _build_groups_from_functions(
            [f for f in project.functions if f.include_mobile],
            platform="mobile"
        )
        # Mobile group thường ít hơn → flatten thành sections trực tiếp
        sections.append({
            "code": "B",
            "name": "NỖ LỰC CHỨC NĂNG MOBILE",
            "groups": mobile_groups
        })

    sections.append({
        "code": "C" if project.enable_mobile else "B",
        "name": "NỖ LỰC UNIT TEST",
        "unit_test": True
    })

    return {"sections": sections}


def _build_groups_from_functions(funcs: list[Function], platform: str) -> list[dict]:
    """Tổ chức list functions thành cấu trúc groups → subgroups → items.

    Dựa vào cột 'cap':
    - cap=1: group (I, II, III...)
    - cap=2: subgroup (I.1, I.2)
    - cap=3: leaf item

    Items rỗng (cap=2 không có cap=3 đi sau) → thêm generic item.
    """
    roman = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XII','XIII',
             'XIV','XV','XVI','XVII','XVIII','XIX','XX','XXI','XXII','XXIII','XXIV','XXV',
             'XXVI','XXVII','XXVIII','XXIX','XXX']

    groups = []
    current_group = None
    current_sub = None
    group_idx = 0

    for f in funcs:
        if f.cap == 1:
            group_idx += 1
            current_group = {
                "code": roman[group_idx - 1] if group_idx <= len(roman) else str(group_idx),
                "name": f"{f.ma} - {f.name}",
                "subgroups": []
            }
            groups.append(current_group)
            current_sub = None
        elif f.cap == 2:
            if current_group is None:
                # Auto-create group
                group_idx += 1
                current_group = {
                    "code": roman[group_idx - 1] if group_idx <= len(roman) else str(group_idx),
                    "name": "Nhóm chức năng",
                    "subgroups": []
                }
                groups.append(current_group)
            sub_code = f"{current_group['code']}.{len(current_group['subgroups']) + 1}"
            current_sub = {
                "code": sub_code,
                "name": f"{f.ma} - {f.name}",
                "items": []
            }
            current_group["subgroups"].append(current_sub)
        elif f.cap == 3:
            ma_cn = f.ma_cn_mobile if platform == "mobile" else f.ma_cn_web
            if not ma_cn:
                ma_cn = "CN_MOBILE Client2" if platform == "mobile" else "CN_WEB2"
            item = {
                "name": f.name,
                "ma_cn": ma_cn,
                "mota": f.desc if f.desc else None,
                "o": f.o, "p": f.p, "q": f.q,
                "thue_ngoai": {"gp": f.tn_gp, "pt": f.tn_pt, "kt": f.tn_kt},
                "ghi_chu": f.ghi_chu, "tsd_tu": f.tsd_tu,
            }
            if current_sub is not None:
                current_sub["items"].append(item)
            elif current_group is not None:
                # Group không có subgroups - thêm items trực tiếp
                if "items" not in current_group:
                    current_group["items"] = []
                    if "subgroups" in current_group:
                        del current_group["subgroups"]
                current_group["items"].append(item)
            else:
                # Orphan leaf - bỏ qua hoặc tạo group default
                group_idx += 1
                current_group = {
                    "code": roman[group_idx - 1] if group_idx <= len(roman) else str(group_idx),
                    "name": "Danh sách chức năng",
                    "items": [item],
                }
                groups.append(current_group)

    # Thêm generic items cho subgroups trống
    for g in groups:
        for s in g.get("subgroups", []):
            if not s.get("items"):
                ma_default = "CN_MOBILE Client2" if platform == "mobile" else "CN_WEB2"
                s["items"] = [{
                    "name": s["name"].split(" - ", 1)[-1] if " - " in s["name"] else s["name"],
                    "ma_cn": ma_default,
                    "mota": f"Triển khai chi tiết {s['name']}",
                    "o": 0, "p": 0, "q": 0,
                    "thue_ngoai": {"gp": "Có", "pt": "Có", "kt": "Có"},
                }]
    return groups


def enrich_with_classifier(project: Project, classifier: Classifier) -> int:
    """Gán ma_cn_web / ma_cn_mobile cho các functions chưa có mã.

    Trả về số functions được gán mã.
    """
    count = 0
    for f in project.functions:
        if f.cap != 3:
            continue
        group_code = f.ma.split(".")[0] if "." in f.ma else f.ma
        if f.include_web and not f.ma_cn_web:
            ma, _rid = classifier.classify(f.name, f.desc, "web", group_code)
            f.ma_cn_web = ma
            count += 1
        if f.include_mobile and not f.ma_cn_mobile:
            ma, _rid = classifier.classify(f.name, f.desc, "mobile", group_code)
            f.ma_cn_mobile = ma
            count += 1
    return count


def apply_cover_page(xlsx_path: str | Path, cover):
    """Update and style sheet 'Trang bia' to match the ULNL cover template."""
    wb = load_workbook(xlsx_path)
    ws = wb["Trang bia"]

    ws._images = []
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    for row in ws.iter_rows(min_row=1, max_row=36, max_col=9):
        for cell in row:
            cell.value = None
            cell.font = Font(name="Times New Roman", size=11, color="000000")
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    widths = {
        "A": 12, "B": 16, "C": 14, "D": 15, "E": 26,
        "F": 20, "G": 18, "H": 10, "I": 10,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row_idx in range(1, 37):
        ws.row_dimensions[row_idx].height = 20
    for row_idx in (1, 2, 7):
        ws.row_dimensions[row_idx].height = 28
    for row_idx in (4, 5, 8, 9, 10, 14, 15, 20, 21, 26, 27, 32, 33):
        ws.row_dimensions[row_idx].height = 18

    logo_path = Path(__file__).resolve().parents[1] / "resources" / "viettel_logo.png"
    if logo_path.exists():
        logo = PngImage(logo_path)
        logo.width = 210
        logo.height = 80
        ws.add_image(logo, "A1")

    ws.merge_cells("D1:G1")
    ws.merge_cells("D2:G2")
    ws["D1"] = "TẬP ĐOÀN CÔNG NGHIỆP - VIỄN THÔNG QUÂN ĐỘI"
    ws["D2"] = "TỔNG CÔNG TY GIẢI PHÁP DOANH NGHIỆP VIETTEL"
    for cell_ref in ("D1", "D2"):
        cell = ws[cell_ref]
        cell.font = Font(name="Times New Roman", size=14, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A7:I7")
    ws["A7"] = cover.project_name
    ws["A7"].font = Font(name="Times New Roman", size=18, bold=True)
    ws["A7"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    label_font = Font(name="Times New Roman", size=12)
    role_font = Font(name="Times New Roman", size=12, bold=True, italic=True)
    value_font = Font(name="Times New Roman", size=12)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    info_rows = [
        ("C11", "Mã dự án/ mã yêu cầu:", "E11", cover.project_code),
        ("C12", "Mã tài liệu:", "E12", cover.doc_code),
        ("C13", "Lần ước lượng", "E13", cover.version),
    ]
    for label_cell, label, value_cell, value in info_rows:
        ws[label_cell] = label
        ws[label_cell].font = label_font
        ws[label_cell].alignment = left
        ws[value_cell] = value
        ws[value_cell].font = value_font
        ws[value_cell].alignment = left

    sections = [
        (16, "Người lập:", cover.nguoi_lap_ten or "<Tên người lập>", cover.nguoi_lap_chucvu or "<Chức vụ>", cover.nguoi_lap_ngay or "<Ngày lập>"),
        (22, "Người kiểm tra:", cover.nguoi_ktra_ten or "<Tên người kiểm tra>", cover.nguoi_ktra_chucvu or "<Chức vụ>", cover.nguoi_ktra_ngay or "<Ngày kiểm tra>"),
        (28, "Người phê duyệt:", cover.nguoi_pduyet_ten or "<Tên người phê duyệt>", cover.nguoi_pduyet_chucvu or "<Chức vụ>", cover.nguoi_pduyet_ngay or "<Ngày phê duyệt>"),
    ]
    for row_idx, role, name, title, date in sections:
        ws[f"A{row_idx}"] = role
        ws[f"A{row_idx}"].font = role_font
        ws[f"A{row_idx}"].alignment = left
        ws[f"C{row_idx + 1}"] = name
        ws[f"C{row_idx + 2}"] = title
        ws[f"F{row_idx + 1}"] = date
        for cell_ref in (f"C{row_idx + 1}", f"C{row_idx + 2}", f"F{row_idx + 1}"):
            ws[cell_ref].font = value_font
            ws[cell_ref].alignment = center

    ws.sheet_view.showGridLines = True
    _apply_middle_align_and_autofit_workbook(wb)
    wb.save(xlsx_path)


def _make_font(style: CellStyle) -> Font:
    return Font(
        name=style.font_name, size=style.font_size,
        bold=style.bold, italic=style.italic,
        color=style.font_color
    )


def _make_fill(style: CellStyle):
    if not style.bg_color or style.bg_color.upper() in ("FFFFFF", "FFFFFFFF", ""):
        return None
    return PatternFill("solid", start_color=style.bg_color, end_color=style.bg_color)


def _make_align(style: CellStyle) -> Alignment:
    return Alignment(horizontal=style.horizontal, vertical=style.vertical, wrap_text=style.wrap_text)


def apply_formatting_overrides(xlsx_path: str | Path, formatting):
    """Apply formatting overrides cho các sheet.

    Logic: chỉ apply font/align/fill cho các vùng đã có content,
    KHÔNG xóa border / công thức / value gốc. Quét theo vùng:
    - Title: rows 1-3 (hoặc 1-5 với Trang bia)
    - Header: rows có style header
    - Total: rows có chứa 'Nỗ lực' hoặc 'Tổng' trong cột A/B
    - Data: phần còn lại
    """
    wb = load_workbook(xlsx_path)
    mappings = [
        ("Trang bia", formatting.trang_bia),
        ("Tong hop", formatting.tong_hop),
        ("NL Chuc nang", formatting.nl_chuc_nang),
        ("NL Phi CN", formatting.nl_phi_cn),
    ]
    for sheet_name, sheet_fmt in mappings:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        _apply_sheet_formatting(ws, sheet_fmt, sheet_name)
    wb.save(xlsx_path)


def apply_middle_align_and_autofit_rows(xlsx_path: str | Path):
    """Set all used cells to middle align and let Excel auto-fit row heights."""
    wb = load_workbook(xlsx_path)
    _apply_middle_align_and_autofit_workbook(wb)
    wb.save(xlsx_path)


def _apply_middle_align_and_autofit_workbook(wb):
    """Apply middle align/autofit to an in-memory workbook without dropping images."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                current = cell.alignment or Alignment()
                cell.alignment = current.copy(vertical="center")
        for row_idx in range(1, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = None


def _apply_sheet_formatting(ws, fmt: SheetFormatting, sheet_name: str):
    """Apply formatting cho 1 sheet."""
    max_row = ws.max_row
    max_col = ws.max_column

    # Heuristic: xác định row title / header / total
    title_rows = set()
    header_rows = set()
    total_rows = set()

    for r in range(1, max_row + 1):
        b_val = str(ws.cell(r, 2).value or "").strip().lower()
        a_val = str(ws.cell(r, 1).value or "").strip().lower()

        if r <= 5 and not b_val and not a_val:
            # Trang bia + tong hop có title rộng ở vùng đầu
            continue

        # Total / sub-total rows
        if "nỗ lực" in b_val or "tổng nỗ lực" in b_val or "tổng cộng" in b_val:
            total_rows.add(r)
            continue
        # Header rows - thường rows 6-7 trong NL Chuc nang, row 1-2 trong Tong hop
        if sheet_name == "NL Chuc nang" and r in (6, 7):
            header_rows.add(r)
            continue
        if sheet_name == "Tong hop" and r in (1, 2):
            title_rows.add(r)
            continue
        if sheet_name == "Tong hop" and ("stt" == a_val or "stt" == b_val):
            header_rows.add(r)
            continue
        if r in (1, 2) and sheet_name in ("Trang bia",):
            title_rows.add(r)

    # Apply
    for r in range(1, max_row + 1):
        if r in title_rows:
            style = fmt.title
        elif r in header_rows:
            style = fmt.header
        elif r in total_rows:
            style = fmt.total
        else:
            style = fmt.data

        font = _make_font(style)
        fill = _make_fill(style)
        align = _make_align(style)

        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            # Chỉ apply nếu cell có content (tránh format hàng triệu cell trống)
            if cell.value is None and not cell.has_style:
                continue
            cell.font = font
            if fill is not None:
                cell.fill = fill
            cell.alignment = align


# ============== MAIN PIPELINE ==============

def build_ulnl_file(
    project: Project,
    output_path: str | Path,
    template_path: str | Path,
    scripts_dir: str | Path,
    apply_format_overrides: bool = True,
) -> dict:
    """Build file ULNL từ project. Trả về dict {success, total_mh, total_mm, errors}."""
    output_path = Path(output_path)
    template_path = Path(template_path)
    scripts_dir = Path(scripts_dir)
    work_dir = output_path.parent / f"_tmp_{output_path.stem}"
    work_dir.mkdir(parents=True, exist_ok=True)

    step0 = work_dir / "step0.xlsx"
    step1 = work_dir / "step1.xlsx"
    step1b = work_dir / "step1b.xlsx"
    step2 = work_dir / "step2.xlsx"
    funcs_json = work_dir / "funcs.json"
    params_json = work_dir / "params.json"

    try:
        subprocess_env = os.environ.copy()
        subprocess_env.setdefault("PYTHONIOENCODING", "utf-8")

        # 1. Copy template
        shutil.copy(template_path, step0)

        # 2. Build sections JSON + run add_functions.py
        sections = build_sections_json(project)
        funcs_json.write_text(json.dumps(sections, ensure_ascii=False, indent=1), encoding="utf-8")

        cmd = [sys.executable, str(scripts_dir / "add_functions.py"),
               str(step0), str(step1), str(funcs_json)]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", env=subprocess_env)
        if result.returncode != 0:
            return {"success": False, "error": f"add_functions.py failed: {result.stderr}",
                    "stdout": result.stdout}

        # Parse output để lấy TOTAL_MH_ROW / UT_ROW
        total_mh_row = None
        ut_row = None
        for line in result.stdout.splitlines():
            if line.startswith("TOTAL_MH_ROW="):
                total_mh_row = int(line.split("=")[1])
            elif line.startswith("UT_ROW="):
                ut_row = int(line.split("=")[1])

        # 3. Apply cover page
        apply_cover_page(step1, project.cover)
        shutil.copy(step1, step1b)

        # 4. Update Phi CN params
        params_dict = {
            "n_module": project.phi_cn.n_module,
            "n_upcode": project.phi_cn.n_upcode,
            "n_bugs": project.phi_cn.n_bugs,
            "n_year": project.phi_cn.n_year,
            "n_hrs_bh_month": project.phi_cn.n_hrs_bh_month,
            "n_sys": project.phi_cn.n_sys,
            "pct_qtda": project.phi_cn.pct_qtda,
        }
        params_json.write_text(json.dumps(params_dict, indent=2), encoding="utf-8")
        cmd = [sys.executable, str(scripts_dir / "update_phi_cn_params.py"),
               str(step1b), str(step2), str(params_json)]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", env=subprocess_env)
        if result.returncode != 0:
            return {"success": False, "error": f"update_phi_cn_params.py failed: {result.stderr}",
                    "stdout": result.stdout}

        # 5. Update Tong hop
        if total_mh_row is None:
            total_mh_row = 50  # fallback
        if ut_row is None:
            ut_row = total_mh_row - 1

        cmd = [sys.executable, str(scripts_dir / "update_tong_hop.py"),
               str(step2), str(output_path),
               str(total_mh_row), "57", str(ut_row)]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", env=subprocess_env)
        if result.returncode != 0:
            return {"success": False, "error": f"update_tong_hop.py failed: {result.stderr}",
                    "stdout": result.stdout}

        # 6. Apply formatting overrides
        if apply_format_overrides:
            try:
                apply_formatting_overrides(output_path, project.formatting)
            except Exception as e:
                print(f"Warning: formatting apply failed: {e}")
        apply_middle_align_and_autofit_rows(output_path)
        apply_cover_page(output_path, project.cover)

        # 7. Read summary
        summary = _read_summary(output_path)

        return {
            "success": True,
            "output_path": str(output_path),
            "total_mh_row": total_mh_row,
            "ut_row": ut_row,
            **summary,
        }
    finally:
        # Cleanup tmp files
        for p in [step0, step1, step1b, step2, funcs_json, params_json]:
            try:
                p.unlink(missing_ok=True)
            except Exception:
                pass
        try:
            work_dir.rmdir()
        except Exception:
            pass


def _read_summary(xlsx_path: str | Path) -> dict:
    """Đọc tổng MH/MD/MM từ Tong hop."""
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb["Tong hop"]
        return {
            "tong_md_section1": ws["E14"].value,
            "tong_md_section2": ws["E25"].value,
            "tong_mm_section2": ws["E26"].value,
        }
    except Exception:
        return {}


# Forward-ref fix
from typing import Optional
