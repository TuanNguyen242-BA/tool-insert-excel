# -*- coding: utf-8 -*-
"""
DOCX Template Builder
=====================
Tool đọc thiết lập từ một file Word có sẵn rồi tạo ra một file template
theo cấu hình do người dùng tuỳ chỉnh.

Quy trình 7 bước (xem các tab trong giao diện):
  1) Chọn file + nhập trang bắt đầu/kết thúc cho phần mở đầu và phần nội dung
  2) Phân tích & sửa Document defaults và Section properties
  3) Phân tích & sửa nội dung phần mở đầu (text holders, ảnh)
  4) Phân tích & sửa Heading styles, Numbering, Header/Footer của phần nội dung
  5) Liệt kê & chỉnh sửa danh sách heading (auto-renumber)
  6) Chọn heading nào giữ nội dung gốc, heading nào để trống
  7) Sinh file template

Yêu cầu cài đặt:
    pip install python-docx lxml openpyxl
"""

import os
import re
import sys
import json
import copy
import shutil
import zipfile
import tempfile
from pathlib import Path
from collections import OrderedDict

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    tk = None
    ttk = None
    filedialog = None
    messagebox = None

from lxml import etree


# ============================================================================
# Namespace
# ============================================================================

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NSMAP = {'w': W_NS, 'r': R_NS, 'a': A_NS}

def w(tag):
    return f'{{{W_NS}}}{tag}'

def r(tag):
    return f'{{{R_NS}}}{tag}'


# ============================================================================
# Tiện ích chuyển đổi đơn vị
# ============================================================================

def twips_to_cm(t):
    """1 cm = 567 twips"""
    try:
        return round(int(t) / 567, 2)
    except (TypeError, ValueError):
        return 0.0

def cm_to_twips(c):
    try:
        return int(round(float(c) * 567))
    except (TypeError, ValueError):
        return 0

def half_pt_to_pt(hp):
    """sz dùng đơn vị half-point: 26 = 13pt"""
    try:
        return int(hp) / 2
    except (TypeError, ValueError):
        return 0

def pt_to_half_pt(p):
    return int(round(float(p) * 2))

def line_to_spacing_str(line, lineRule):
    """Convert (line, lineRule) → human-readable spacing"""
    if not line:
        return '1.0'
    line = int(line)
    if lineRule == 'auto':
        return f'{line/240:.2f}'  # 240 = 1.0
    elif lineRule == 'exact':
        return f'{line/20:.1f}pt (exact)'
    elif lineRule == 'atLeast':
        return f'{line/20:.1f}pt (atLeast)'
    return str(line)

def spacing_str_to_line(s):
    """'1.5' → 360 line, 'auto'"""
    try:
        f = float(s)
        return str(int(round(f * 240))), 'auto'
    except ValueError:
        return '240', 'auto'


# Schema element order theo OOXML — cần để chèn element mới đúng vị trí
PPR_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
    'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
    'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
    'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
    'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
    'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr',
    'pPrChange',
]

RPR_ORDER = [
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
    'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
    'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
    'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
    'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
    'eastAsianLayout', 'specVanish', 'oMath',
]


def ensure_child(parent, tag_name, schema_order):
    """
    Lấy hoặc tạo element con với tag_name. Nếu phải tạo mới, chèn ở đúng vị trí
    theo schema_order để file XML hợp lệ với schema OOXML.
    """
    existing = parent.find(w(tag_name))
    if existing is not None:
        return existing
    new_el = etree.Element(w(tag_name))
    if tag_name not in schema_order:
        parent.append(new_el)
        return new_el
    idx_of_tag = schema_order.index(tag_name)
    inserted = False
    for i, child in enumerate(parent):
        ctag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if ctag in schema_order and schema_order.index(ctag) > idx_of_tag:
            parent.insert(i, new_el)
            inserted = True
            break
    if not inserted:
        parent.append(new_el)
    return new_el


def parse_excel_column_spec(spec):
    """Parse 'A,B,D-F,H' hoặc 'A,C:E' -> list column indexes.

    Hỗ trợ tên cột Excel (A, B, AA), range bằng '-' hoặc ':'.
    Invalid token bị bỏ qua để UI preview không crash khi user đang gõ dở.
    """
    from openpyxl.utils import column_index_from_string

    result = []
    for part in (spec or '').split(','):
        part = part.strip().upper()
        if not part:
            continue

        sep = '-' if '-' in part else ':' if ':' in part else None
        if sep:
            a, b = part.split(sep, 1)
            try:
                start = column_index_from_string(a.strip())
                end = column_index_from_string(b.strip())
            except Exception:
                continue
            if start > end:
                start, end = end, start
            result.extend(range(start, end + 1))
        else:
            try:
                result.append(column_index_from_string(part))
            except Exception:
                continue

    seen = set()
    unique = []
    for i in result:
        if i not in seen:
            seen.add(i)
            unique.append(i)
    return unique



class DocxModel:
    """
    Lưu trạng thái file docx + cấu hình do người dùng đặt.
    Mọi tab trong wizard đọc/ghi vào instance này.
    """

    def __init__(self):
        self.path = None              # đường dẫn file gốc
        self.workdir = None           # thư mục unpack tạm
        self.doc_xml = None           # ElementTree document.xml
        self.styles_xml = None
        self.numbering_xml = None
        self.settings_xml = None

        # Body element list (paragraphs + tables) - tham chiếu nhanh
        self.body_elements = []

        # Mapping page number → list các index của body_elements
        self.page_to_elements = {}    # {1: [0,1,2,...], 2: [...]}

        # Số trang thật theo metadata Word ghi (app.xml)
        self.metadata_pages = None

        # Cấu hình do người dùng
        self.config = {
            # Element index bắt đầu phần nội dung. Mọi element TRƯỚC index này
            # là phần mở đầu (cover, changelog, signature, TOC).
            # Mặc định = None → tool dùng heuristic (Heading 1 đầu tiên không phải MỤC LỤC)
            'intro_end_idx': None,
            'doc_defaults': {},       # {font, size, lang, line_spacing, jc}
            'sections': [],           # [{pgSz, pgMar, headers, footers, ...}]
            'intro_replacements': {}, # {old_text: new_text}
            'intro_images': {},       # {rId: new_image_path}
            'heading_styles': {},     # {Heading1: {font, size, bold, ...}}
            'numbering_overrides': {},
            'header_footer_text': {}, # {header2: 'text', footer2: 'text'}
            # Mỗi entry: {level, text, original_idx, keep_content, insert_source, auto_number}
            #   keep_content=True               → giữ nguyên content gốc
            #   keep_content=False, source=None → xoá content, để 1 dòng trống
            #   keep_content=False, source=path → xoá rồi chèn từ file
            'headings_list': [],
            # Khi chèn từ file ngoài: có giữ bold/italic của từng từ không?
            'preserve_inline_formatting': True,
            # Font + size áp cho mọi cell khi chèn từ Excel.
            # Mặc định font theo template, size 11pt để bảng nhiều cột dễ đọc.
            'excel_font_name': '',
            'excel_font_size': 11,
            'output_path': None,
        }

    # -------------------------------------------------------------------
    # Load & unpack
    # -------------------------------------------------------------------

    def load(self, path):
        """Mở file docx, unpack ra workdir tạm, parse XML."""
        self.path = path
        if self.workdir:
            shutil.rmtree(self.workdir, ignore_errors=True)
        self.workdir = tempfile.mkdtemp(prefix='docx_template_')

        with zipfile.ZipFile(path) as z:
            z.extractall(self.workdir)

        parser = etree.XMLParser(remove_blank_text=False)
        self.doc_xml = etree.parse(os.path.join(self.workdir, 'word/document.xml'), parser)
        self.styles_xml = etree.parse(os.path.join(self.workdir, 'word/styles.xml'), parser)

        num_path = os.path.join(self.workdir, 'word/numbering.xml')
        self.numbering_xml = etree.parse(num_path, parser) if os.path.exists(num_path) else None

        sett_path = os.path.join(self.workdir, 'word/settings.xml')
        self.settings_xml = etree.parse(sett_path, parser) if os.path.exists(sett_path) else None

        # Build body elements list
        body = self.doc_xml.getroot().find(w('body'))
        self.body_elements = [el for el in body if el.tag in (w('p'), w('tbl'))]

        # Đọc số trang thật từ docProps/app.xml (Word ghi)
        self.metadata_pages = self._read_pages_from_app_xml()

        # Detect pages via lastRenderedPageBreak
        self._detect_pages()

        # Đọc các thiết lập mặc định và lưu vào config (lần đầu)
        self._load_defaults_into_config()

    def _read_pages_from_app_xml(self):
        """
        Đọc số trang từ docProps/app.xml. Thử theo thứ tự ưu tiên:
        1. <Pages> nếu giá trị > 1 (Word ghi)
        2. <Words> ước tính: words / 300 ≈ pages (fallback khi Pages bị sai)
        Trả về None nếu không đọc được.
        """
        app_path = os.path.join(self.workdir, 'docProps/app.xml')
        if not os.path.exists(app_path):
            return None
        try:
            tree = etree.parse(app_path)
            pages = None
            words = None
            for el in tree.iter():
                tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                if tag == 'Pages' and el.text and el.text.strip().isdigit():
                    p = int(el.text.strip())
                    if p > 0:
                        pages = p
                elif tag == 'Words' and el.text and el.text.strip().isdigit():
                    w_count = int(el.text.strip())
                    if w_count > 0:
                        words = w_count
            # Pages tin cậy nếu > 1 (=1 thường là default sai)
            if pages and pages > 1:
                return pages
            # Fallback: ước tính từ Words count
            if words and words > 300:
                return max(1, round(words / 300))
            # Pages = 1 cũng valid nếu chỉ có ít chữ
            if pages and words and words < 300:
                return pages
        except Exception:
            pass
        return None

    def get_word_count(self):
        """Trả về số chữ trong file (từ app.xml), None nếu không có."""
        app_path = os.path.join(self.workdir, 'docProps/app.xml')
        if not os.path.exists(app_path):
            return None
        try:
            tree = etree.parse(app_path)
            for el in tree.iter():
                tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                if tag == 'Words' and el.text and el.text.strip().isdigit():
                    return int(el.text.strip())
        except Exception:
            pass
        return None

    def _detect_pages(self):
        """
        Word lưu marker <w:lastRenderedPageBreak/> ở đầu run mà page break xảy ra
        khi render lần cuối. Đếm chúng để xác định ranh giới trang.
        Nếu không có, fallback sang <w:br w:type="page"/>.
        """
        self.page_to_elements = {1: []}
        current_page = 1

        for idx, el in enumerate(self.body_elements):
            # Xem xét cả w:lastRenderedPageBreak VÀ w:br type="page"
            page_breaks_in_el = 0

            # Đếm lastRenderedPageBreak (chỉ tính cái đứng đầu run)
            for lrpb in el.iter(w('lastRenderedPageBreak')):
                page_breaks_in_el += 1

            # Page break cứng <w:br w:type="page"/>
            for br in el.iter(w('br')):
                if br.get(w('type')) == 'page':
                    page_breaks_in_el += 1

            # Mỗi page break tăng số trang lên 1
            # Nhưng element đó vẫn thuộc trang TRƯỚC khi sang trang mới
            # Đơn giản hoá: thêm element vào trang hiện tại,
            # rồi tăng trang theo số page break tìm được
            self.page_to_elements.setdefault(current_page, []).append(idx)
            current_page += page_breaks_in_el

        # Nếu không có lastRenderedPageBreak nào (file chưa render trong Word),
        # ta coi mỗi sectPr cũng tạo trang mới
        if max(self.page_to_elements.keys()) == 1:
            current_page = 1
            self.page_to_elements = {1: []}
            for idx, el in enumerate(self.body_elements):
                self.page_to_elements.setdefault(current_page, []).append(idx)
                # Section break ngoại trừ continuous tạo page break
                pPr = el.find(w('pPr'))
                if pPr is not None:
                    sectPr = pPr.find(w('sectPr'))
                    if sectPr is not None:
                        type_el = sectPr.find(w('type'))
                        if type_el is None or type_el.get(w('val')) != 'continuous':
                            current_page += 1

    def get_total_pages(self):
        """
        Số trang ưu tiên metadata (app.xml). Nếu không có, fallback về
        page detection từ LRPB/sectPr.
        """
        if self.metadata_pages:
            return self.metadata_pages
        return max(self.page_to_elements.keys()) if self.page_to_elements else 1

    def get_total_pages_detected(self):
        """Số trang phát hiện được từ page break (có thể thấp hơn thực tế)"""
        return max(self.page_to_elements.keys()) if self.page_to_elements else 1

    def get_elements_in_pages(self, start, end):
        """Trả về list index trong body_elements thuộc khoảng trang [start, end]"""
        result = []
        for p in range(start, end + 1):
            result.extend(self.page_to_elements.get(p, []))
        return result

    def get_intro_end_idx(self):
        """
        Heuristic: tìm element index ngay sau khi kết thúc phần intro
        (trang bìa, changelog, signature, mục lục).
        Cách: tìm Heading 1 đầu tiên có text khác "MỤC LỤC".
        Mọi element trước nó là intro.

        Logic này độc lập với việc detect page (vốn không tin cậy
        khi file chưa được Word render đầy đủ).
        """
        for idx, el in enumerate(self.body_elements):
            if el.tag != w('p'):
                continue
            pPr = el.find(w('pPr'))
            if pPr is None:
                continue
            pStyle = pPr.find(w('pStyle'))
            if pStyle is None:
                continue
            style_id = pStyle.get(w('val')) or ''
            if style_id != 'Heading1':
                continue
            text = ''.join(t.text or '' for t in el.iter(w('t'))).strip().upper()
            # Skip MỤC LỤC (Heading 1 này thuộc intro)
            if not text:
                continue
            normalized = text.replace('Ụ', 'U').replace('Ụ', 'U')
            if 'MỤC LỤC' in text or 'MỤC LỤC' in normalized or 'MUC LUC' in normalized:
                continue
            return idx
        return 0  # Không tìm thấy → coi như không có intro

    def get_effective_intro_end_idx(self):
        """
        Trả về intro_end_idx user đã chọn ở Step 1.
        Nếu user chưa chọn (hoặc dùng auto), fallback heuristic.
        """
        cfg_val = self.config.get('intro_end_idx')
        if cfg_val is not None:
            return cfg_val
        return self.get_intro_end_idx()

    def get_all_headings(self):
        """
        Quét toàn bộ body, lấy mọi paragraph có pStyle="Heading1..9".
        KHÔNG filter theo page (vì page detection không tin cậy).
        """
        result = []
        for idx, el in enumerate(self.body_elements):
            if el.tag != w('p'):
                continue
            pPr = el.find(w('pPr'))
            if pPr is None:
                continue
            pStyle = pPr.find(w('pStyle'))
            if pStyle is None:
                continue
            style_id = pStyle.get(w('val')) or ''
            m = re.match(r'Heading(\d+)', style_id)
            if not m:
                continue
            level = int(m.group(1))
            text = ''.join(t.text or '' for t in el.iter(w('t')))
            result.append({
                'level': level,
                'text': text.strip(),
                'original_idx': idx,
                'keep_content': True,
                'auto_number': False,
            })
        return result

    # -------------------------------------------------------------------
    # Đọc thiết lập mặc định vào config
    # -------------------------------------------------------------------

    def _load_defaults_into_config(self):
        # Document defaults
        self.config['doc_defaults'] = self._read_doc_defaults()

        # Sections
        self.config['sections'] = self._read_sections()

        # Heading styles (1-9)
        self.config['heading_styles'] = self._read_heading_styles()

    def _read_doc_defaults(self):
        result = {
            'font': 'Times New Roman',
            'size_pt': 13.0,
            'lang': 'en-US',
            'line_spacing': '1.5',
            'jc': 'both',
            'space_before': 6.0,  # pt
            'space_after': 0.0,   # pt
            'indent_left_cm': 0.0,
        }
        root = self.styles_xml.getroot()
        dd = root.find(w('docDefaults'))
        if dd is None:
            return result

        # rPrDefault
        rpr = dd.find(f'{w("rPrDefault")}/{w("rPr")}')
        if rpr is not None:
            rfonts = rpr.find(w('rFonts'))
            if rfonts is not None:
                result['font'] = rfonts.get(w('ascii')) or result['font']
            sz = rpr.find(w('sz'))
            if sz is not None:
                result['size_pt'] = half_pt_to_pt(sz.get(w('val')))
            lang = rpr.find(w('lang'))
            if lang is not None:
                result['lang'] = lang.get(w('val')) or result['lang']

        # pPrDefault
        ppr = dd.find(f'{w("pPrDefault")}/{w("pPr")}')
        if ppr is not None:
            sp = ppr.find(w('spacing'))
            if sp is not None:
                line = sp.get(w('line'))
                rule = sp.get(w('lineRule'))
                if line:
                    result['line_spacing'] = line_to_spacing_str(line, rule)
                before = sp.get(w('before'))
                if before:
                    result['space_before'] = int(before) / 20
                after = sp.get(w('after'))
                if after:
                    result['space_after'] = int(after) / 20
            ind = ppr.find(w('ind'))
            if ind is not None:
                left = ind.get(w('left'))
                if left:
                    result['indent_left_cm'] = twips_to_cm(left)
            jc = ppr.find(w('jc'))
            if jc is not None:
                result['jc'] = jc.get(w('val')) or result['jc']

        return result

    def _read_sections(self):
        """Đọc tất cả sectPr trong document.xml"""
        sections = []
        root = self.doc_xml.getroot()
        body = root.find(w('body'))
        # sectPr trong body là default cho section cuối
        for sect in body.iter(w('sectPr')):
            s = {
                'pg_w_cm': 21.0, 'pg_h_cm': 29.7,
                'orient': 'portrait',
                'top_cm': 2.54, 'right_cm': 2.54, 'bottom_cm': 2.54, 'left_cm': 2.54,
                'header_cm': 1.27, 'footer_cm': 1.27, 'gutter_cm': 0.0,
                'titlePg': False,
                'header_refs': {}, 'footer_refs': {},
            }
            pg_sz = sect.find(w('pgSz'))
            if pg_sz is not None:
                s['pg_w_cm'] = twips_to_cm(pg_sz.get(w('w')))
                s['pg_h_cm'] = twips_to_cm(pg_sz.get(w('h')))
                s['orient'] = pg_sz.get(w('orient')) or 'portrait'
            pg_mar = sect.find(w('pgMar'))
            if pg_mar is not None:
                for key, attr in [('top_cm', 'top'), ('right_cm', 'right'),
                                   ('bottom_cm', 'bottom'), ('left_cm', 'left'),
                                   ('header_cm', 'header'), ('footer_cm', 'footer'),
                                   ('gutter_cm', 'gutter')]:
                    val = pg_mar.get(w(attr))
                    if val:
                        s[key] = twips_to_cm(val)
            s['titlePg'] = sect.find(w('titlePg')) is not None
            for hr in sect.findall(w('headerReference')):
                s['header_refs'][hr.get(w('type'))] = hr.get(r('id'))
            for fr in sect.findall(w('footerReference')):
                s['footer_refs'][fr.get(w('type'))] = fr.get(r('id'))
            sections.append(s)
        return sections

    def _read_heading_styles(self):
        """Đọc style Heading 1-9 (cả paragraph style và linked Char style)."""
        result = OrderedDict()
        root = self.styles_xml.getroot()
        for level in range(1, 10):
            sid = f'Heading{level}'
            char_id = f'Heading{level}Char'
            data = {
                'font': '', 'size_pt': 0,
                'bold': False, 'italic': False,
                'color': '', 'space_before_pt': 0, 'space_after_pt': 0,
            }
            # Paragraph style
            for s in root.findall(w('style')):
                if s.get(w('styleId')) == sid:
                    rpr = s.find(w('rPr'))
                    if rpr is not None:
                        self._fill_run_props(rpr, data)
                    ppr = s.find(w('pPr'))
                    if ppr is not None:
                        sp = ppr.find(w('spacing'))
                        if sp is not None:
                            if sp.get(w('before')):
                                data['space_before_pt'] = int(sp.get(w('before'))) / 20
                            if sp.get(w('after')):
                                data['space_after_pt'] = int(sp.get(w('after'))) / 20
                # Char style (override font/size khi paragraph style không có)
                if s.get(w('styleId')) == char_id:
                    rpr = s.find(w('rPr'))
                    if rpr is not None:
                        self._fill_run_props(rpr, data, only_if_empty=True)
            result[sid] = data
        return result

    def _fill_run_props(self, rpr, data, only_if_empty=False):
        rfonts = rpr.find(w('rFonts'))
        if rfonts is not None:
            f = rfonts.get(w('ascii'))
            if f and (not only_if_empty or not data.get('font')):
                data['font'] = f
        sz = rpr.find(w('sz'))
        if sz is not None:
            v = half_pt_to_pt(sz.get(w('val')))
            if v and (not only_if_empty or not data.get('size_pt')):
                data['size_pt'] = v
        if rpr.find(w('b')) is not None:
            data['bold'] = True
        if rpr.find(w('i')) is not None:
            data['italic'] = True
        color = rpr.find(w('color'))
        if color is not None:
            v = color.get(w('val'))
            if v and (not only_if_empty or not data.get('color')):
                data['color'] = v

    # -------------------------------------------------------------------
    # Đọc danh sách heading trong vùng nội dung
    # -------------------------------------------------------------------

    def get_headings_in_range(self, start_page, end_page):
        """Trả về [{level, text, original_idx}]"""
        result = []
        elem_indexes = set(self.get_elements_in_pages(start_page, end_page))
        for idx, el in enumerate(self.body_elements):
            if idx not in elem_indexes:
                continue
            if el.tag != w('p'):
                continue
            pPr = el.find(w('pPr'))
            if pPr is None:
                continue
            pStyle = pPr.find(w('pStyle'))
            if pStyle is None:
                continue
            style_id = pStyle.get(w('val')) or ''
            m = re.match(r'Heading(\d+)', style_id)
            if not m:
                continue
            level = int(m.group(1))
            # Lấy text
            text = ''.join(t.text or '' for t in el.iter(w('t')))
            result.append({
                'level': level,
                'text': text.strip(),
                'original_idx': idx,
                'keep_content': True,
            })
        return result

    # -------------------------------------------------------------------
    # Đọc nội dung phần mở đầu (text + ảnh) cho UI sửa
    # -------------------------------------------------------------------

    def get_intro_text_blocks(self):
        """
        Lấy danh sách các đoạn text trong phần mở đầu.
        Dùng intro_end_idx user đã chọn ở Step 1 (hoặc heuristic mặc định).
        Trả về [(idx, kind, text)] với kind ∈ {'p', 'tbl_cell'}.
        """
        result = []
        intro_end = self.get_effective_intro_end_idx()
        if intro_end == 0:
            # Không xác định được → lấy 100 element đầu cho an toàn
            intro_end = min(100, len(self.body_elements))
        for idx in range(intro_end):
            el = self.body_elements[idx]
            if el.tag == w('p'):
                text = ''.join(t.text or '' for t in el.iter(w('t')))
                if text.strip():
                    result.append((idx, 'p', text))
            elif el.tag == w('tbl'):
                for ri, row in enumerate(el.findall(w('tr'))):
                    for ci, cell in enumerate(row.findall(w('tc'))):
                        text = ''.join(t.text or '' for t in cell.iter(w('t')))
                        if text.strip():
                            result.append((f'{idx}.{ri}.{ci}', 'tbl_cell', text))
        return result

    def get_intro_images(self):
        """[(rId, target_path)] — ảnh trong phần mở đầu."""
        result = []
        rels_path = os.path.join(self.workdir, 'word/_rels/document.xml.rels')
        if not os.path.exists(rels_path):
            return result
        rels = etree.parse(rels_path)
        rid_to_target = {}
        for rel in rels.getroot():
            if 'image' in rel.get('Type', ''):
                rid_to_target[rel.get('Id')] = rel.get('Target')

        intro_end = self.get_effective_intro_end_idx()
        if intro_end == 0:
            intro_end = min(100, len(self.body_elements))
        seen = set()
        for idx in range(intro_end):
            el = self.body_elements[idx]
            for blip in el.iter(f'{{{A_NS}}}blip'):
                rid = blip.get(r('embed'))
                if rid and rid not in seen and rid in rid_to_target:
                    seen.add(rid)
                    result.append((rid, rid_to_target[rid]))
        return result

    # -------------------------------------------------------------------
    # Header & Footer text
    # -------------------------------------------------------------------

    def get_header_footer_files(self):
        """List các file header/footer + nội dung text"""
        result = {}
        for fname in os.listdir(os.path.join(self.workdir, 'word')):
            if fname.startswith(('header', 'footer')) and fname.endswith('.xml'):
                path = os.path.join(self.workdir, 'word', fname)
                tree = etree.parse(path)
                texts = [t.text or '' for t in tree.iter(w('t'))]
                result[fname] = ' '.join(texts).strip()
        return result

    # -------------------------------------------------------------------
    # Apply config & generate output
    # -------------------------------------------------------------------

    def generate(self, output_path, progress_callback=None):
        """Áp dụng toàn bộ config và xuất file mới.
        progress_callback(percent, message) - tuỳ chọn, để cập nhật progress bar.
        """
        def progress(p, msg):
            if progress_callback:
                progress_callback(p, msg)

        progress(2, 'Sao chép workdir...')
        # Chuẩn bị workdir mới
        out_dir = tempfile.mkdtemp(prefix='docx_out_')
        for item in os.listdir(self.workdir):
            src = os.path.join(self.workdir, item)
            dst = os.path.join(out_dir, item)
            if os.path.isdir(src):
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)

        progress(10, 'Parse XML trees...')
        parser = etree.XMLParser(remove_blank_text=False)
        doc_tree = etree.parse(os.path.join(out_dir, 'word/document.xml'), parser)
        styles_tree = etree.parse(os.path.join(out_dir, 'word/styles.xml'), parser)

        progress(20, 'Áp dụng document defaults...')
        self._apply_doc_defaults(styles_tree)

        progress(30, 'Áp dụng sections...')
        self._apply_sections(doc_tree)

        progress(40, 'Áp dụng heading styles...')
        self._apply_heading_styles(styles_tree)

        progress(50, 'Áp dụng text replacements...')
        self._apply_text_replacements(doc_tree)

        progress(60, 'Thay thế ảnh...')
        self._apply_image_replacements(out_dir)

        progress(70, 'Xử lý heading & content (xoá/chèn)...')
        self._apply_headings_and_content(doc_tree, styles_tree, progress_callback)

        progress(85, 'Cập nhật header/footer text...')
        self._apply_header_footer_text(out_dir)

        progress(90, 'Lưu XML trees...')
        doc_tree.write(os.path.join(out_dir, 'word/document.xml'),
                       xml_declaration=True, encoding='UTF-8', standalone=True)
        styles_tree.write(os.path.join(out_dir, 'word/styles.xml'),
                          xml_declaration=True, encoding='UTF-8', standalone=True)

        progress(95, 'Đóng gói file docx...')
        self._zip_dir(out_dir, output_path)
        shutil.rmtree(out_dir, ignore_errors=True)

        progress(100, 'Hoàn tất')

    def _zip_dir(self, src_dir, output_path):
        if os.path.exists(output_path):
            os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(src_dir):
                for f in files:
                    full = os.path.join(root_dir, f)
                    rel = os.path.relpath(full, src_dir)
                    rel = rel.replace(os.sep, '/')
                    zf.write(full, rel)

    def _apply_doc_defaults(self, styles_tree):
        cfg = self.config['doc_defaults']
        root = styles_tree.getroot()
        dd = root.find(w('docDefaults'))
        if dd is None:
            return

        rpr = dd.find(f'{w("rPrDefault")}/{w("rPr")}')
        if rpr is not None:
            rfonts = ensure_child(rpr, 'rFonts', RPR_ORDER)
            for attr in ['ascii', 'eastAsia', 'hAnsi', 'cs']:
                rfonts.set(w(attr), cfg['font'])
            sz = ensure_child(rpr, 'sz', RPR_ORDER)
            sz.set(w('val'), str(pt_to_half_pt(cfg['size_pt'])))
            szcs = ensure_child(rpr, 'szCs', RPR_ORDER)
            szcs.set(w('val'), str(pt_to_half_pt(cfg['size_pt'])))
            lang = ensure_child(rpr, 'lang', RPR_ORDER)
            lang.set(w('val'), cfg['lang'])

        ppr = dd.find(f'{w("pPrDefault")}/{w("pPr")}')
        if ppr is not None:
            sp = ensure_child(ppr, 'spacing', PPR_ORDER)
            line, rule = spacing_str_to_line(cfg['line_spacing'])
            sp.set(w('line'), line)
            sp.set(w('lineRule'), rule)
            sp.set(w('before'), str(int(cfg['space_before'] * 20)))
            sp.set(w('after'), str(int(cfg['space_after'] * 20)))
            jc = ensure_child(ppr, 'jc', PPR_ORDER)
            jc.set(w('val'), cfg['jc'])

    def _apply_sections(self, doc_tree):
        sections_cfg = self.config['sections']
        root = doc_tree.getroot()
        body = root.find(w('body'))
        sect_prs = list(body.iter(w('sectPr')))
        for i, sect in enumerate(sect_prs):
            if i >= len(sections_cfg):
                break
            cfg = sections_cfg[i]
            pg_sz = sect.find(w('pgSz'))
            if pg_sz is not None:
                pg_sz.set(w('w'), str(cm_to_twips(cfg['pg_w_cm'])))
                pg_sz.set(w('h'), str(cm_to_twips(cfg['pg_h_cm'])))
                if cfg['orient'] == 'landscape':
                    pg_sz.set(w('orient'), 'landscape')
                else:
                    if w('orient') in pg_sz.attrib:
                        del pg_sz.attrib[w('orient')]
            pg_mar = sect.find(w('pgMar'))
            if pg_mar is not None:
                pg_mar.set(w('top'), str(cm_to_twips(cfg['top_cm'])))
                pg_mar.set(w('right'), str(cm_to_twips(cfg['right_cm'])))
                pg_mar.set(w('bottom'), str(cm_to_twips(cfg['bottom_cm'])))
                pg_mar.set(w('left'), str(cm_to_twips(cfg['left_cm'])))
                pg_mar.set(w('header'), str(cm_to_twips(cfg['header_cm'])))
                pg_mar.set(w('footer'), str(cm_to_twips(cfg['footer_cm'])))
                pg_mar.set(w('gutter'), str(cm_to_twips(cfg['gutter_cm'])))

    def _apply_heading_styles(self, styles_tree):
        root = styles_tree.getroot()
        for sid, data in self.config['heading_styles'].items():
            for s in root.findall(w('style')):
                if s.get(w('styleId')) != sid:
                    continue
                # Style có thứ tự con: name, basedOn, next, link, ..., pPr, rPr
                # Schema: pPr trước rPr trong style element
                rpr = s.find(w('rPr'))
                if rpr is None:
                    rpr = etree.SubElement(s, w('rPr'))
                # font
                if data.get('font'):
                    rfonts = ensure_child(rpr, 'rFonts', RPR_ORDER)
                    for a in ['ascii', 'hAnsi', 'cs']:
                        rfonts.set(w(a), data['font'])
                # size
                if data.get('size_pt'):
                    sz = ensure_child(rpr, 'sz', RPR_ORDER)
                    sz.set(w('val'), str(pt_to_half_pt(data['size_pt'])))
                # bold/italic
                self._set_flag(rpr, 'b', data.get('bold', False))
                self._set_flag(rpr, 'i', data.get('italic', False))
                # color
                if data.get('color'):
                    color = ensure_child(rpr, 'color', RPR_ORDER)
                    color.set(w('val'), data['color'])
                # spacing trong pPr
                ppr = s.find(w('pPr'))
                if ppr is None:
                    ppr = etree.Element(w('pPr'))
                    # Chèn pPr trước rPr trong style element
                    rpr_in_style = s.find(w('rPr'))
                    if rpr_in_style is not None:
                        rpr_in_style.addprevious(ppr)
                    else:
                        s.append(ppr)
                sp = ensure_child(ppr, 'spacing', PPR_ORDER)
                if data.get('space_before_pt') is not None:
                    sp.set(w('before'), str(int(data['space_before_pt'] * 20)))
                if data.get('space_after_pt') is not None:
                    sp.set(w('after'), str(int(data['space_after_pt'] * 20)))

    def _set_flag(self, parent, tag_name, value):
        """Set/unset boolean flag tag, đảm bảo theo schema order nếu là rPr"""
        existing = parent.find(w(tag_name))
        if value and existing is None:
            # Chèn theo RPR_ORDER nếu parent là rPr
            ensure_child(parent, tag_name, RPR_ORDER)
        elif not value and existing is not None:
            parent.remove(existing)

    def _apply_text_replacements(self, doc_tree):
        replacements = self.config.get('intro_replacements', {})
        if not replacements:
            return
        for t in doc_tree.iter(w('t')):
            if t.text:
                for old, new in replacements.items():
                    if old in t.text:
                        t.text = t.text.replace(old, new)

    def _apply_image_replacements(self, out_dir):
        for rid, new_path in self.config.get('intro_images', {}).items():
            if not new_path or not os.path.exists(new_path):
                continue
            # Tìm target trong rels
            rels_path = os.path.join(out_dir, 'word/_rels/document.xml.rels')
            if not os.path.exists(rels_path):
                continue
            rels = etree.parse(rels_path)
            for rel in rels.getroot():
                if rel.get('Id') == rid:
                    target = rel.get('Target')
                    abs_target = os.path.join(out_dir, 'word', target)
                    os.makedirs(os.path.dirname(abs_target), exist_ok=True)
                    shutil.copy2(new_path, abs_target)
                    break

    def _apply_headings_and_content(self, doc_tree, styles_tree, progress_callback=None):
        """
        - Đổi text các heading theo config['headings_list']
        - Xoá nội dung dưới heading nếu keep_content=False
        - Sau khi xoá: chèn 1 paragraph trống hoặc content từ file ngoài (insert_source)
        - Auto-renumber: prepend số thứ tự nếu user yêu cầu
        """
        def progress(p, msg):
            if progress_callback:
                progress_callback(p, msg)

        headings_list = self.config.get('headings_list', [])
        if not headings_list:
            return

        root = doc_tree.getroot()
        body = root.find(w('body'))
        body_children = [el for el in body if el.tag in (w('p'), w('tbl'))]

        # Cache page content width — dùng khi build bảng từ Excel để fit page
        self._cached_page_width = self._compute_page_content_width(doc_tree)

        # Tính số thứ tự dạng "1.", "1.1.", "1.2.1." dựa trên thứ tự trong list
        counters = [0] * 10
        numbered = []
        for h in headings_list:
            lvl = h['level']
            counters[lvl] += 1
            for j in range(lvl + 1, 10):
                counters[j] = 0
            for k in range(1, lvl):
                if counters[k] == 0:
                    counters[k] = 1
            num_str = '.'.join(str(counters[k]) for k in range(1, lvl + 1))
            numbered.append(num_str)

        # Áp text mới + số (chỉ với entry đã có original_idx)
        for entry, num_str in zip(headings_list, numbered):
            if entry.get('original_idx') is None:
                continue
            idx = entry['original_idx']
            if idx >= len(body_children):
                continue
            el = body_children[idx]
            new_text = entry['text']
            if entry.get('auto_number', False):
                new_text = f'{num_str}. {new_text}'
            self._set_paragraph_text(el, new_text)

        # Chuẩn bị plan: với mỗi heading bỏ tick, lưu reference đến heading element
        # và xác định range cần xoá. Sau khi xoá xong sẽ chèn empty/external sau heading_el.
        plans = []
        for entry_idx, entry in enumerate(headings_list):
            if entry.get('keep_content', True):
                continue
            if entry.get('original_idx') is None:
                continue
            cur_idx = entry['original_idx']
            # Tìm heading kế tiếp BẤT KỲ (level nào cũng được)
            next_entry_orig_idx = None
            for j in range(entry_idx + 1, len(headings_list)):
                if headings_list[j].get('original_idx') is not None:
                    next_entry_orig_idx = headings_list[j]['original_idx']
                    break
            end_idx = next_entry_orig_idx if next_entry_orig_idx is not None else len(body_children)

            if cur_idx < len(body_children):
                heading_el = body_children[cur_idx]
                table_format_hint = None
                for k in range(cur_idx + 1, end_idx):
                    if k < len(body_children) and body_children[k].tag == w('tbl'):
                        table_format_hint = self._extract_table_format_hint(body_children[k])
                        break
                plans.append({
                    'heading_el': heading_el,
                    'start_idx': cur_idx + 1,
                    'end_idx': end_idx,
                    'insert_source': entry.get('insert_source') or None,
                    'insert_selection': entry.get('insert_selection') or None,
                    'heading_text': entry.get('text', ''),
                    'page_width': self._compute_page_content_width_for_index(doc_tree, cur_idx),
                    'table_format_hint': table_format_hint,
                })

        # Gom các index cần xoá thành một set duy nhất
        indexes_to_delete = set()
        for plan in plans:
            for k in range(plan['start_idx'], plan['end_idx']):
                indexes_to_delete.add(k)

        progress(72, f'Xoá content cũ ({len(indexes_to_delete)} elements)...')
        # Xoá theo thứ tự ngược để tránh ảnh hưởng index sau khi remove
        for k in sorted(indexes_to_delete, reverse=True):
            if k >= len(body_children):
                continue
            el = body_children[k]
            # Paragraph chứa sectPr cần được giữ để không làm mất section break,
            # nhưng text bên trong vẫn phải xoá nếu nó thuộc vùng content cần xoá
            # (PL01 có trường hợp "N/A" nằm trong paragraph mang sectPr).
            pPr = el.find(w('pPr')) if el.tag == w('p') else None
            if pPr is not None and pPr.find(w('sectPr')) is not None:
                self._clear_paragraph_keep_ppr(el)
                continue
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

        # Chèn empty paragraph hoặc external content sau mỗi heading
        preserve_inline = self.config.get('preserve_inline_formatting', True)
        n_plans = len(plans)
        for i, plan in enumerate(plans):
            heading_el = plan['heading_el']
            if heading_el.getparent() is None:
                continue  # heading bị xoá đâu đó (hiếm khi xảy ra)

            if plan['insert_source']:
                pct = 75 + int(8 * (i + 1) / max(n_plans, 1))
                progress(pct, f'Chèn từ file: {os.path.basename(plan["insert_source"])} '
                              f'(heading "{plan["heading_text"][:30]}")...')
                try:
                    # Mỗi heading có thể nằm ở section khác nhau (portrait/landscape).
                    # Build bảng Excel theo width của section chứa heading đó.
                    self._cached_page_width = plan.get('page_width') or self._cached_page_width
                    # Nếu heading gốc có table mẫu, Excel table kế thừa font/size
                    # từ table đó để bám template PL01 hơn doc defaults.
                    self._cached_table_format_hint = plan.get('table_format_hint') or {}
                    elements = self._read_external_content(
                        plan['insert_source'], preserve_inline,
                        selection=plan.get('insert_selection'),
                    )
                    if not elements:
                        elements = [self._create_empty_paragraph()]
                except Exception as e:
                    print(f'⚠ Lỗi đọc {plan["insert_source"]}: {e}')
                    import traceback; traceback.print_exc()
                    elements = [self._create_empty_paragraph()]
            else:
                # Chèn 1 paragraph trống → user paste vào sẽ format theo Normal style
                elements = [self._create_empty_paragraph()]

            # Chèn các element sau heading_el theo đúng thứ tự
            last = heading_el
            for new_el in elements:
                last.addnext(new_el)
                last = new_el

    def _compute_page_content_width(self, doc_tree):
        """Tính chiều rộng usable của trang (twips) từ first sectPr trong document.

        - A4 portrait: pgSz w=11906 twips (21cm)
        - Lề L 3.17cm + R 2.54cm = 5.71cm = 3239 twips
        - Usable = 11906 - 3239 = 8667 twips ~ 15.29cm

        Khi build bảng từ Excel, dùng giá trị này làm tổng width thay vì hardcode 9000.
        Tránh được tình trạng bảng tràn page → Word compress → cell wrap dọc.
        """
        if doc_tree is None:
            return 8500
        root = doc_tree.getroot()
        body = root.find(w('body'))
        if body is None:
            return 8500

        def extract(sectPr):
            pgSz = sectPr.find(w('pgSz'))
            pgMar = sectPr.find(w('pgMar'))
            try:
                page_w = int(pgSz.get(w('w'))) if pgSz is not None else 11906
                left = int(pgMar.get(w('left'))) if pgMar is not None else 1440
                right = int(pgMar.get(w('right'))) if pgMar is not None else 1440
                return page_w - left - right
            except (TypeError, ValueError):
                return None

        # Ưu tiên: first sectPr in body order (Section 1 thường là portrait)
        for p in body.findall(w('p')):
            pPr = p.find(w('pPr'))
            if pPr is not None:
                spr = pPr.find(w('sectPr'))
                if spr is not None:
                    cw = extract(spr)
                    if cw and cw >= 4000:
                        # Trừ thêm 100 twips safety margin
                        return cw - 100
        # Fallback: sectPr ở cuối body
        spr_body = body.find(w('sectPr'))
        if spr_body is not None:
            cw = extract(spr_body)
            if cw and cw >= 4000:
                return cw - 100
        return 8500

    def _sectPr_content_width(self, sectPr):
        """Tính usable width từ một sectPr cụ thể, trừ safety margin nhỏ."""
        if sectPr is None:
            return None
        pgSz = sectPr.find(w('pgSz'))
        pgMar = sectPr.find(w('pgMar'))
        try:
            page_w = int(pgSz.get(w('w'))) if pgSz is not None else 11906
            left = int(pgMar.get(w('left'))) if pgMar is not None else 1440
            right = int(pgMar.get(w('right'))) if pgMar is not None else 1440
            cw = page_w - left - right
        except (TypeError, ValueError):
            return None
        if cw < 4000:
            return None
        return cw - 100

    def _compute_page_content_width_for_index(self, doc_tree, body_index):
        """Tính usable width của section chứa body element tại body_index.

        Trong Word, sectPr nằm ở paragraph cuối của section và áp cho toàn bộ
        content phía trước nó. Nếu không gặp sectPr sau body_index, dùng sectPr
        cuối body. Điều này quan trọng với template PL01 vì document có cả
        section portrait và landscape.
        """
        if doc_tree is None:
            return self._compute_page_content_width(doc_tree)

        root = doc_tree.getroot()
        body = root.find(w('body'))
        if body is None:
            return 8500

        body_children = [el for el in body if el.tag in (w('p'), w('tbl'))]
        for idx in range(max(0, body_index), len(body_children)):
            el = body_children[idx]
            if el.tag != w('p'):
                continue
            pPr = el.find(w('pPr'))
            sectPr = pPr.find(w('sectPr')) if pPr is not None else None
            cw = self._sectPr_content_width(sectPr)
            if cw:
                return cw

        cw = self._sectPr_content_width(body.find(w('sectPr')))
        if cw:
            return cw
        return self._compute_page_content_width(doc_tree)

    def _extract_table_format_hint(self, tbl):
        """Lấy size đại diện từ table mẫu gốc dưới cùng heading.

        Font Excel vẫn theo doc defaults của template (PL01 = Times New Roman).
        Size từ table mẫu giúp giảm wrap trong bảng nhiều cột.
        """
        if tbl is None:
            return {}
        hint = {}
        for rpr in tbl.findall('.//' + w('rPr')):
            if not hint.get('size_pt'):
                sz = rpr.find(w('sz'))
                if sz is not None:
                    size_pt = half_pt_to_pt(sz.get(w('val')))
                    if size_pt:
                        hint['size_pt'] = size_pt
            if hint.get('size_pt'):
                break
        return hint

    def _create_empty_paragraph(self):
        """Tạo paragraph rỗng với pStyle="Normal" → kế thừa doc defaults."""
        p = etree.Element(w('p'))
        pPr = etree.SubElement(p, w('pPr'))
        pStyle = etree.SubElement(pPr, w('pStyle'))
        pStyle.set(w('val'), 'Normal')
        return p

    # ============================================================
    # Đọc nội dung từ file Word/Excel khác để chèn vào template
    # ============================================================

    def _read_external_content(self, source_path, preserve_inline=True, selection=None):
        """Đọc file Word/Excel ngoài, trả về list lxml elements (p, tbl)
        đã được clean để chèn vào template.

        selection: dict tuỳ chọn {mode, sheet, range, ...} để chọn phần cụ thể.
                   None hoặc {'mode': 'all'} → lấy toàn bộ.
        """
        ext = source_path.lower()
        if ext.endswith('.docx'):
            # Word selection sẽ làm phase sau, hiện tại vẫn lấy toàn bộ
            return self._read_word_content(source_path, preserve_inline)
        elif ext.endswith(('.xlsx', '.xlsm')):
            return self._read_excel_content(source_path, preserve_inline, selection)
        else:
            raise ValueError(f'Định dạng file không hỗ trợ: {source_path}')

    def _read_word_content(self, source_path, preserve_inline):
        """Đọc paragraphs/tables từ file .docx, clean format để khớp template."""
        with zipfile.ZipFile(source_path) as z:
            doc_xml = z.read('word/document.xml')
        src_tree = etree.fromstring(doc_xml)
        src_body = src_tree.find(w('body'))
        if src_body is None:
            return []

        result = []
        for el in src_body:
            if el.tag == w('p'):
                # Skip sectPr-only paragraph cuối
                pPr = el.find(w('pPr'))
                if pPr is not None and pPr.find(w('sectPr')) is not None:
                    has_text = any(t.text and t.text.strip() for t in el.iter(w('t')))
                    if not has_text:
                        continue
                clean_p = self._clean_paragraph_for_template(el, preserve_inline)
                if clean_p is not None:
                    result.append(clean_p)
            elif el.tag == w('tbl'):
                clean_tbl = self._clean_table_for_template(el, preserve_inline)
                if clean_tbl is not None:
                    result.append(clean_tbl)
        return result

    def _clean_paragraph_for_template(self, src_p, preserve_inline):
        """
        Tạo paragraph mới với:
        - pStyle = "Normal" → inherit doc defaults (font/size từ template)
        - Giữ alignment cơ bản (jc) nếu có
        - Mỗi run: giữ text + (tuỳ chọn) bold/italic/underline/strike, BỎ font/size/color
        """
        new_p = etree.Element(w('p'))
        pPr = etree.SubElement(new_p, w('pPr'))
        pStyle = etree.SubElement(pPr, w('pStyle'))
        pStyle.set(w('val'), 'Normal')

        # Giữ alignment nếu có (left/center/right/both)
        src_pPr = src_p.find(w('pPr'))
        if src_pPr is not None:
            jc = src_pPr.find(w('jc'))
            if jc is not None:
                new_jc = etree.SubElement(pPr, w('jc'))
                new_jc.set(w('val'), jc.get(w('val')) or 'left')

        # Process children (runs, hyperlinks)
        for child in src_p:
            if child.tag == w('pPr'):
                continue
            if child.tag == w('r'):
                new_run = self._clean_run_for_template(child, preserve_inline)
                if new_run is not None:
                    new_p.append(new_run)
            elif child.tag == w('hyperlink'):
                # Convert hyperlink → plain runs (đỡ phải giữ rels)
                for sub_run in child.findall(w('r')):
                    new_run = self._clean_run_for_template(sub_run, preserve_inline)
                    if new_run is not None:
                        new_p.append(new_run)

        return new_p

    def _clean_run_for_template(self, src_run, preserve_inline):
        """Clean 1 run: giữ text, (tuỳ chọn) giữ b/i/u/strike, bỏ rFonts/sz/color."""
        new_run = etree.Element(w('r'))

        if preserve_inline:
            src_rpr = src_run.find(w('rPr'))
            if src_rpr is not None:
                new_rpr = etree.Element(w('rPr'))
                added_any = False
                # Chỉ giữ các formatting flag inline cơ bản
                for tag in ['b', 'bCs', 'i', 'iCs', 'u', 'strike', 'dstrike',
                            'vertAlign', 'caps', 'smallCaps']:
                    src_el = src_rpr.find(w(tag))
                    if src_el is not None:
                        new_el = etree.Element(w(tag))
                        for k, v in src_el.attrib.items():
                            new_el.set(k, v)
                        new_rpr.append(new_el)
                        added_any = True
                if added_any:
                    new_run.append(new_rpr)

        # Copy text + tab + br
        for child in src_run:
            if child.tag == w('rPr'):
                continue
            if child.tag in (w('t'), w('tab'), w('br'), w('cr'), w('noBreakHyphen')):
                new_child = etree.Element(child.tag)
                if child.text:
                    new_child.text = child.text
                for k, v in child.attrib.items():
                    new_child.set(k, v)
                new_run.append(new_child)

        return new_run

    def _clean_table_for_template(self, src_tbl, preserve_inline):
        """Clone table, clean cell content."""
        new_tbl = etree.Element(w('tbl'))

        # Copy tblPr (deep clone)
        src_tblPr = src_tbl.find(w('tblPr'))
        if src_tblPr is not None:
            new_tbl.append(etree.fromstring(etree.tostring(src_tblPr)))
        else:
            # Fallback: thêm border đơn giản
            tblPr = etree.SubElement(new_tbl, w('tblPr'))
            tblBorders = etree.SubElement(tblPr, w('tblBorders'))
            for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                b = etree.SubElement(tblBorders, w(side))
                b.set(w('val'), 'single')
                b.set(w('sz'), '4')
                b.set(w('color'), 'auto')

        # Copy tblGrid (column widths)
        src_tblGrid = src_tbl.find(w('tblGrid'))
        if src_tblGrid is not None:
            new_tbl.append(etree.fromstring(etree.tostring(src_tblGrid)))

        # Process rows
        for src_tr in src_tbl.findall(w('tr')):
            new_tr = etree.SubElement(new_tbl, w('tr'))
            src_trPr = src_tr.find(w('trPr'))
            if src_trPr is not None:
                new_tr.append(etree.fromstring(etree.tostring(src_trPr)))

            for src_tc in src_tr.findall(w('tc')):
                new_tc = etree.SubElement(new_tr, w('tc'))
                src_tcPr = src_tc.find(w('tcPr'))
                if src_tcPr is not None:
                    new_tc.append(etree.fromstring(etree.tostring(src_tcPr)))

                # Process paragraphs in cell
                cell_has_p = False
                for el in src_tc:
                    if el.tag == w('tcPr'):
                        continue
                    if el.tag == w('p'):
                        cleaned = self._clean_paragraph_for_template(el, preserve_inline)
                        if cleaned is not None:
                            new_tc.append(cleaned)
                            cell_has_p = True
                # Word yêu cầu mỗi cell có ít nhất 1 paragraph
                if not cell_has_p:
                    new_tc.append(self._create_empty_paragraph())

        return new_tbl

    def _read_excel_content(self, source_path, preserve_inline, selection=None):
        """Đọc Excel sheets với selection mode.
        selection={
          'mode': 'custom'|'range'|'sheet'|'all',
          'sheet': name,
          # cho range:  'range': 'A1:D10'
          # cho custom: 'columns': 'A,B,D-F' (str), 'row_start': int, 'row_end': int|None
        }
        """
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise RuntimeError(
                'Tính năng chèn từ Excel cần openpyxl. Cài bằng: pip install openpyxl'
            )

        # Không có selection thường là dữ liệu legacy từ bản cũ chỉ lưu path.
        # Giữ hành vi an toàn: lấy toàn bộ workbook thay vì rơi vào custom thiếu sheet.
        selection = selection or {'mode': 'all'}
        mode = selection.get('mode', 'custom')

        # Bỏ read_only=True để max_row/max_column populate đúng
        wb = load_workbook(source_path, data_only=True)
        elements = []

        if mode == 'all':
            for ws in wb.worksheets:
                if len(wb.worksheets) > 1:
                    elements.append(self._make_subheading_paragraph(ws.title))
                rows = [list(r) for r in ws.iter_rows()]
                tbl = self._build_table_from_cell_rows(rows, preserve_inline)
                if tbl is not None:
                    elements.append(tbl)
                    elements.append(self._create_empty_paragraph())
        elif mode == 'sheet':
            sheet_name = selection.get('sheet')
            if not sheet_name or sheet_name not in wb.sheetnames:
                raise ValueError(f'Sheet "{sheet_name}" không tồn tại trong file. '
                                 f'Các sheet có sẵn: {wb.sheetnames}')
            ws = wb[sheet_name]
            rows = [list(r) for r in ws.iter_rows()]
            tbl = self._build_table_from_cell_rows(rows, preserve_inline)
            if tbl is not None:
                elements.append(tbl)
                elements.append(self._create_empty_paragraph())
        elif mode == 'range':
            sheet_name = selection.get('sheet')
            range_str = (selection.get('range') or '').strip()
            if not sheet_name or sheet_name not in wb.sheetnames:
                raise ValueError(f'Sheet "{sheet_name}" không tồn tại trong file.')
            if not range_str:
                raise ValueError('Range không được để trống khi mode=range')
            ws = wb[sheet_name]
            rows = self._extract_rows_from_range(ws, range_str)
            tbl = self._build_table_from_cell_rows(rows, preserve_inline)
            if tbl is not None:
                elements.append(tbl)
                elements.append(self._create_empty_paragraph())
        elif mode == 'custom':
            sheet_name = selection.get('sheet')
            cols_spec = (selection.get('columns') or '').strip()
            row_start = selection.get('row_start') or 1
            row_end = selection.get('row_end')  # None / 0 / '' → auto
            try:
                row_start = max(1, int(row_start))
            except (TypeError, ValueError):
                row_start = 1
            try:
                row_end = int(row_end) if row_end not in (None, '', 0, '0') else None
            except (TypeError, ValueError):
                row_end = None

            if not sheet_name or sheet_name not in wb.sheetnames:
                raise ValueError(f'Sheet "{sheet_name}" không tồn tại trong file.')
            ws = wb[sheet_name]

            max_col = ws.max_column or 1
            max_row_in_sheet = ws.max_row or 1

            # Parse columns spec
            if cols_spec:
                col_indexes = self._parse_column_spec(cols_spec)
            else:
                col_indexes = list(range(1, max_col + 1))
            if not col_indexes:
                raise ValueError(f'Danh sách cột custom không hợp lệ: "{cols_spec}"')

            if row_end is None:
                row_end = max_row_in_sheet
            row_end = max(row_start, row_end)

            # Build rows: lấy cell theo col_indexes × row range
            rows = []
            for r in range(row_start, row_end + 1):
                row_cells = [ws.cell(row=r, column=c) for c in col_indexes]
                rows.append(row_cells)

            tbl = self._build_table_from_cell_rows(rows, preserve_inline)
            if tbl is not None:
                elements.append(tbl)
                elements.append(self._create_empty_paragraph())
        else:
            raise ValueError(f'Mode không hợp lệ: {mode}')

        return elements

    def _parse_column_spec(self, spec):
        """Parse 'A,B,D-F,H' → list column indexes [1,2,4,5,6,8]."""
        return parse_excel_column_spec(spec)

    def _extract_rows_from_range(self, ws, range_str):
        """Trả về list of list of Cell từ ws[range_str]."""
        try:
            sel = ws[range_str]
        except Exception as e:
            raise ValueError(f'Range "{range_str}" không hợp lệ: {e}')
        rows = []
        if hasattr(sel, 'value'):
            rows = [[sel]]
        elif isinstance(sel, tuple):
            if not sel:
                return []
            if isinstance(sel[0], tuple):
                rows = [list(r) for r in sel]
            else:
                rows = [list(sel)]
        return rows

    def _make_subheading_paragraph(self, text):
        p = etree.Element(w('p'))
        pPr = etree.SubElement(p, w('pPr'))
        pStyle = etree.SubElement(pPr, w('pStyle'))
        pStyle.set(w('val'), 'Heading4')
        run = etree.SubElement(p, w('r'))
        t = etree.SubElement(run, w('t'))
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return p

    def _build_table_from_cell_rows(self, rows, preserve_inline):
        """Build <w:tbl> từ list of list of openpyxl Cell objects.

        Strategy: tính column widths dựa theo content length của từng cột,
        layout=fixed để Word không redistribute lại width khi render. Tổng width
        bám theo usable width của section chứa heading chèn bảng.
        """
        def _text_of(cell):
            if cell is None or cell.value is None:
                return ''
            value = cell.value
            if hasattr(value, 'isoformat'):
                value = value.isoformat()
            return str(value)

        def _excel_horizontal_to_word(value):
            mapping = {
                'center': 'center',
                'centerContinuous': 'center',
                'right': 'right',
                'left': 'left',
                'justify': 'both',
                'distributed': 'distribute',
                'fill': 'left',
                'general': None,
            }
            return mapping.get(value, None)

        def _excel_vertical_to_word(value):
            mapping = {
                'top': 'top',
                'center': 'center',
                'bottom': 'bottom',
                'justify': 'center',
                'distributed': 'center',
            }
            return mapping.get(value, None)

        def _cell_horizontal_alignment(cell, row_idx, col_idx):
            if row_idx == 0:
                return 'center'
            if cell is not None and getattr(cell, 'alignment', None) is not None:
                mapped = _excel_horizontal_to_word(cell.alignment.horizontal)
                if mapped:
                    # PL01 gốc dùng right cho STT; Excel mẫu dùng center.
                    # Tôn trọng Excel vì user yêu cầu căn theo file gốc insert.
                    return mapped
            value = cell.value if cell is not None else None
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                return 'center'
            if col_idx in no_wrap_cols:
                return 'center'
            return 'left'

        def _cell_vertical_alignment(cell, row_idx):
            if cell is not None and getattr(cell, 'alignment', None) is not None:
                mapped = _excel_vertical_to_word(cell.alignment.vertical)
                if mapped:
                    return mapped
            return 'center'

        if not rows:
            return None
        # Strip empty trailing rows
        while rows and all(c is None or c.value is None for c in rows[-1]):
            rows.pop()
        if not rows:
            return None
        ncols = max(len(row) for row in rows)
        if ncols == 0:
            return None

        # ---- Tính column widths theo header + content ----
        # Nếu chỉ dùng len(cell text), cột mô tả dài sẽ nuốt width, còn header
        # ngắn như STT/Key bị ép hẹp và Word bẻ thành S/T/T hoặc K/ey.
        def _token_lens(text):
            parts = re.split(r'[\s_\\/\-.,;:()\\[\\]{}]+', text or '')
            return [len(p) for p in parts if p]

        header_word_lens = [0] * ncols
        header_total_lens = [0] * ncols
        header_texts = [''] * ncols
        body_word_lens = [0] * ncols
        body_total_lens = [0] * ncols
        for r_idx, row in enumerate(rows):
            for i in range(ncols):
                cell = row[i] if i < len(row) else None
                text = _text_of(cell).strip()
                if not text:
                    continue
                token_len = max(_token_lens(text) or [0])
                if r_idx == 0:
                    header_texts[i] = text
                    header_word_lens[i] = max(header_word_lens[i], token_len)
                    header_total_lens[i] = max(header_total_lens[i], min(len(text), 32))
                else:
                    body_word_lens[i] = max(body_word_lens[i], token_len)
                    body_total_lens[i] = max(body_total_lens[i], min(len(text), 48))

        # Constants:
        # - HEADER_CHAR_W cao hơn body vì header thường bold và cần tránh bẻ chữ.
        # - FULL_TEXT_WEIGHT chỉ là trọng số phụ, không được nuốt width của header.
        HEADER_CHAR_W = 155
        BODY_WORD_CHAR_W = 100
        FULL_TEXT_WEIGHT = 45
        MIN_COL_W = 900
        MIN_COL_W_HARD = 400

        # Cell margin ngang giảm dần khi nhiều cột, để tăng usable text width.
        if ncols >= 13:
            cell_margin_h = 40
        elif ncols >= 9:
            cell_margin_h = 40
        else:
            cell_margin_h = 108
        CELL_OVERHEAD = cell_margin_h * 2 + 120

        # Lấy page content width đã cached, fallback 8500 (A4 portrait lề mặc định)
        total_target = int(getattr(self, '_cached_page_width', None) or 8500)

        # Adjust floor theo số cột để không có trường hợp tổng min width vượt page.
        # Với quá nhiều cột trên portrait (vd 23 cột), 400*ncols có thể > page.
        hard_floor = min(MIN_COL_W_HARD, max(120, total_target // ncols))
        if ncols <= 8:
            desired_min = MIN_COL_W
        elif ncols <= 12:
            desired_min = 700
        else:
            desired_min = MIN_COL_W_HARD
        effective_min = min(desired_min, max(hard_floor, total_target // ncols))

        if ncols <= 12:
            header_cap = 1350
            body_word_cap = 1100
        else:
            header_cap = 800
            body_word_cap = 750

        no_wrap_cols = set()
        min_widths = []
        col_floors = []
        weights = []
        for i in range(ncols):
            header_plain = re.sub(r'\s+', '', header_texts[i] or '').upper()
            is_short_code_col = (
                ncols <= 12 and
                header_plain in {'STT', 'KEY'} and
                body_word_lens[i] <= 4
            )
            if is_short_code_col:
                no_wrap_cols.add(i)
            col_floor = 950 if is_short_code_col else hard_floor
            col_floors.append(col_floor)

            header_req = min(
                header_cap,
                header_word_lens[i] * HEADER_CHAR_W + CELL_OVERHEAD
            )
            body_word_req = min(
                body_word_cap,
                body_word_lens[i] * BODY_WORD_CHAR_W + CELL_OVERHEAD
            )
            col_min = max(effective_min, header_req, body_word_req)
            if is_short_code_col:
                col_min = max(col_min, col_floor)
            min_widths.append(int(col_min))
            # Header được tính mạnh hơn body để không lặp lại lỗi header cao.
            weight = (
                header_total_lens[i] * 80 +
                header_word_lens[i] * 120 +
                body_total_lens[i] * FULL_TEXT_WEIGHT +
                body_word_lens[i] * 45
            )
            weights.append(max(weight, 1))

        def _norm_header(text):
            return re.sub(r'\s+', ' ', (text or '').strip().lower())

        semantic_col_widths = None
        semantic_table_hint = {}
        normalized_headers = [_norm_header(h) for h in header_texts]
        db_field_headers = [
            'stt',
            'tên bảng đích',
            'tên bảng đích (chuẩn hóa)',
            'trường đích',
            'trường đích (chuẩn hóa)',
            'ý nghĩa trường đích',
            'định dạng (chuẩn hóa)',
            'định dạng',
            'key',
            'hệ thống nguồn',
            'bảng nguồn',
        ]
        if ncols == 11 and normalized_headers[:11] == db_field_headers:
            # Preset riêng cho bảng danh sách trường CSDL trong PL01:
            # ưu tiên rộng cho cột tên bảng/trường chuẩn hóa và ý nghĩa,
            # giữ STT/Key đủ hẹp nhưng không wrap.
            ratios = [950, 1200, 1650, 1200, 1450, 1800,
                      1100, 1050, 950, 1150, 1050]
            ratio_total = sum(ratios)
            semantic_col_widths = [
                max(col_floors[i], int(total_target * ratios[i] / ratio_total))
                for i in range(ncols)
            ]
            semantic_table_hint = {'size_pt': 10}

        # Nếu tổng min đã vượt page, co tỷ lệ theo requirement từng cột.
        # Không bóp một vài cột xuống hard_floor vì sẽ tái hiện lỗi header cao.
        min_total = sum(min_widths)
        if semantic_col_widths:
            col_widths = semantic_col_widths
        elif min_total > total_target:
            base_total = sum(col_floors)
            remaining = max(0, total_target - base_total)
            req_weights = [max(min_widths[i] - col_floors[i], 1) for i in range(ncols)]
            req_total = sum(req_weights)
            col_widths = [
                col_floors[i] + int(remaining * req_weights[i] / req_total)
                for i in range(ncols)
            ]
        else:
            remaining = total_target - min_total
            weight_total = sum(weights)
            if weight_total > 0:
                col_widths = [
                    min_widths[i] + int(remaining * weights[i] / weight_total)
                    for i in range(ncols)
                ]
            else:
                col_widths = list(min_widths)

        # Final adjust để tổng = exactly target
        diff = total_target - sum(col_widths)
        if diff > 0:
            # Chia diff theo tỷ lệ vào các cột wide nhất (top 3)
            sorted_idx = sorted(range(len(col_widths)),
                                key=lambda i: -col_widths[i])
            top = sorted_idx[:min(3, len(sorted_idx))]
            per = diff // len(top)
            for j, idx in enumerate(top):
                add = per if j < len(top) - 1 else (diff - per * (len(top) - 1))
                col_widths[idx] += add
        elif diff < 0:
            need = -diff
            for idx in sorted(range(len(col_widths)), key=lambda i: -col_widths[i]):
                if need <= 0:
                    break
                reducible = max(0, col_widths[idx] - col_floors[idx])
                take = min(reducible, need)
                col_widths[idx] -= take
                need -= take

        # Final safety: tất cả phải >= floor riêng của cột và tổng phải bám target.
        col_widths = [max(col_floors[i], cw) for i, cw in enumerate(col_widths)]

        total_w = sum(col_widths)

        tbl = etree.Element(w('tbl'))

        # ---- tblPr ----
        tblPr = etree.SubElement(tbl, w('tblPr'))
        # tblW: total = computed sum of col widths, type=dxa (twips)
        tblW = etree.SubElement(tblPr, w('tblW'))
        tblW.set(w('w'), str(total_w))
        tblW.set(w('type'), 'dxa')

        # Borders
        tblBorders = etree.SubElement(tblPr, w('tblBorders'))
        for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            b = etree.SubElement(tblBorders, w(side))
            b.set(w('val'), 'single')
            b.set(w('sz'), '4')
            b.set(w('space'), '0')
            b.set(w('color'), 'auto')

        # tblLayout=FIXED — Word KHÔNG redistribute, tôn trọng widths của tôi
        tblLayout = etree.SubElement(tblPr, w('tblLayout'))
        tblLayout.set(w('type'), 'fixed')

        # tblLook
        tblLook = etree.SubElement(tblPr, w('tblLook'))
        tblLook.set(w('val'), '04A0')
        tblLook.set(w('firstRow'), '1')
        tblLook.set(w('lastRow'), '0')
        tblLook.set(w('firstColumn'), '0')
        tblLook.set(w('lastColumn'), '0')
        tblLook.set(w('noHBand'), '0')
        tblLook.set(w('noVBand'), '1')

        # ---- tblGrid theo widths đã tính ----
        tblGrid = etree.SubElement(tbl, w('tblGrid'))
        for w_val in col_widths:
            gc = etree.SubElement(tblGrid, w('gridCol'))
            gc.set(w('w'), str(w_val))

        # ---- Font/size override (default = doc defaults) ----
        ov_font = (self.config.get('excel_font_name') or '').strip()
        ov_size = self.config.get('excel_font_size') or 0
        try:
            ov_size = float(ov_size)
        except (TypeError, ValueError):
            ov_size = 0
        table_hint = getattr(self, '_cached_table_format_hint', {}) or {}
        # Nếu user không override font → luôn dùng doc defaults của template
        # (PL01 = Times New Roman), không lấy font từ table mẫu vì có table dùng Arial.
        # Size vẫn có thể lấy từ table mẫu để giảm wrap.
        if not ov_font:
            ov_font = (
                    self.config.get('doc_defaults', {}).get('font') or
                    'Times New Roman'
                )
        if ov_size <= 0:
            try:
                ov_size = float(
                    table_hint.get('size_pt') or
                    semantic_table_hint.get('size_pt') or
                    self.config.get('doc_defaults', {}).get('size_pt') or
                    13
                )
            except (TypeError, ValueError):
                ov_size = 13.0

        # ---- Rows ----
        for r_idx, row in enumerate(rows):
            tr = etree.SubElement(tbl, w('tr'))
            if r_idx == 0:
                # Lặp lại header nếu bảng bị tách sang trang sau.
                trPr = etree.SubElement(tr, w('trPr'))
                etree.SubElement(trPr, w('tblHeader'))
            for c_idx in range(ncols):
                cell = row[c_idx] if c_idx < len(row) else None
                tc = etree.SubElement(tr, w('tc'))

                tcPr = etree.SubElement(tc, w('tcPr'))
                tcW = etree.SubElement(tcPr, w('tcW'))
                tcW.set(w('w'), str(col_widths[c_idx]))
                tcW.set(w('type'), 'dxa')

                if c_idx in no_wrap_cols:
                    noWrap = etree.SubElement(tcPr, w('noWrap'))
                    noWrap.set(w('val'), '1')

                # tcMar explicit: giảm margin ngang khi bảng nhiều cột để
                # tránh Word bẻ header thành từng ký tự.
                # → tránh template inherit tcMar to gây squeeze content
                tcMar = etree.SubElement(tcPr, w('tcMar'))
                for side, val in [('top', '40'), ('left', str(cell_margin_h)),
                                   ('bottom', '40'), ('right', str(cell_margin_h))]:
                    m = etree.SubElement(tcMar, w(side))
                    m.set(w('w'), val)
                    m.set(w('type'), 'dxa')

                vAlign = etree.SubElement(tcPr, w('vAlign'))
                vAlign.set(w('val'), _cell_vertical_alignment(cell, r_idx))

                p = etree.SubElement(tc, w('p'))
                pPr = etree.SubElement(p, w('pPr'))
                pStyle = etree.SubElement(pPr, w('pStyle'))
                pStyle.set(w('val'), 'Normal')
                # Không kế thừa Normal 1.5 line + before 6pt trong table,
                # vì nó làm header và body cao bất thường.
                spacing = etree.SubElement(pPr, w('spacing'))
                spacing.set(w('before'), '0')
                spacing.set(w('after'), '0')
                spacing.set(w('line'), '240')
                spacing.set(w('lineRule'), 'auto')
                # Normal style trong PL01 có left indent (~0.48cm). Nếu không
                # reset, mọi cell bị trống một khoảng bên trái.
                ind = ensure_child(pPr, 'ind', PPR_ORDER)
                ind.set(w('left'), '0')
                ind.set(w('right'), '0')
                ind.set(w('firstLine'), '0')
                ind.set(w('hanging'), '0')
                jc = ensure_child(pPr, 'jc', PPR_ORDER)
                jc.set(w('val'), _cell_horizontal_alignment(cell, r_idx, c_idx))
                if c_idx in no_wrap_cols:
                    wordWrap = ensure_child(pPr, 'wordWrap', PPR_ORDER)
                    wordWrap.set(w('val'), '0')

                run = etree.SubElement(p, w('r'))

                # rPr: luôn có (ép font/size)
                rpr = etree.Element(w('rPr'))
                rfonts = ensure_child(rpr, 'rFonts', RPR_ORDER)
                rfonts.set(w('ascii'), ov_font)
                rfonts.set(w('hAnsi'), ov_font)
                rfonts.set(w('eastAsia'), ov_font)
                rfonts.set(w('cs'), ov_font)

                if preserve_inline and cell is not None and cell.font:
                    if cell.font.bold:
                        ensure_child(rpr, 'b', RPR_ORDER)
                    if cell.font.italic:
                        ensure_child(rpr, 'i', RPR_ORDER)

                sz = ensure_child(rpr, 'sz', RPR_ORDER)
                sz.set(w('val'), str(int(round(ov_size * 2))))
                szcs = ensure_child(rpr, 'szCs', RPR_ORDER)
                szcs.set(w('val'), str(int(round(ov_size * 2))))

                if preserve_inline and cell is not None and cell.font:
                    if cell.font.underline and cell.font.underline != 'none':
                        u = ensure_child(rpr, 'u', RPR_ORDER)
                        u.set(w('val'), 'single')

                run.append(rpr)

                value = cell.value if cell is not None and cell.value is not None else ''
                if hasattr(value, 'isoformat'):
                    value = value.isoformat()
                t = etree.SubElement(run, w('t'))
                t.text = str(value)
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        return tbl

    def _set_paragraph_text(self, p, text):
        """Xoá hết các run hiện tại trong paragraph rồi thêm 1 run mới với text mới.
        Giữ lại pPr."""
        pPr = p.find(w('pPr'))
        # Xoá tất cả children trừ pPr
        for child in list(p):
            if child.tag != w('pPr'):
                p.remove(child)
        # Thêm run mới
        run = etree.SubElement(p, w('r'))
        t = etree.SubElement(run, w('t'))
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text

    def _clear_paragraph_keep_ppr(self, p):
        """Xoá toàn bộ nội dung hiển thị của paragraph, giữ pPr/sectPr."""
        for child in list(p):
            if child.tag != w('pPr'):
                p.remove(child)

    def _apply_header_footer_text(self, out_dir):
        for fname, new_text in self.config.get('header_footer_text', {}).items():
            path = os.path.join(out_dir, 'word', fname)
            if not os.path.exists(path):
                continue
            tree = etree.parse(path)
            ts = list(tree.iter(w('t')))
            if ts:
                # Xoá text các <w:t> sau, gán toàn bộ text vào <w:t> đầu
                ts[0].text = new_text
                for t in ts[1:]:
                    t.text = ''
                tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)


# ============================================================================
# Excel Selection Dialog (Phase 1a)
# ============================================================================

class ExcelSelectionDialog:
    """
    Dialog cho phép user chọn nội dung từ file Excel:
    - Chọn 1 sheet
    - Hoặc chọn vùng cells trong 1 sheet (vd A1:D20)

    Dùng:
        result = ExcelSelectionDialog.show(parent, current_path, current_selection)
        if result:
            path = result['path']
            selection = result['selection']  # {'mode': ..., 'sheet': ..., 'range': ...}
    """

    def __init__(self, parent, current_path='', current_selection=None):
        self.result = None
        current_selection = current_selection or {}

        self.win = tk.Toplevel(parent)
        self.win.title('Chọn nội dung từ file Excel')
        self.win.geometry('960x860')
        self.win.transient(parent)
        self.win.grab_set()

        # ---- File picker ----
        file_frame = ttk.Frame(self.win)
        file_frame.pack(fill='x', padx=10, pady=10)
        ttk.Label(file_frame, text='File Excel:').pack(side='left')
        self.path_var = tk.StringVar(value=current_path)
        ttk.Entry(file_frame, textvariable=self.path_var, width=70).pack(side='left', padx=5)
        ttk.Button(file_frame, text='Browse...', command=self._browse_file).pack(side='left')

        self.file_info_label = ttk.Label(self.win, text='', foreground='blue', wraplength=850, justify='left')
        self.file_info_label.pack(anchor='w', padx=10, pady=2)

        ttk.Separator(self.win, orient='horizontal').pack(fill='x', padx=10, pady=8)

        # ---- Selection mode ----
        # Default = 'custom'
        self.mode_var = tk.StringVar(value=current_selection.get('mode') or 'custom')

        mode_frame = ttk.LabelFrame(self.win, text='Phạm vi nội dung cần lấy')
        mode_frame.pack(fill='x', padx=10, pady=5)

        ttk.Radiobutton(
            mode_frame,
            text='Tùy chỉnh — chọn sheet, cột, phạm vi dòng (mặc định, hay dùng nhất)',
            variable=self.mode_var, value='custom',
            command=self._on_mode_change,
        ).pack(anchor='w', padx=5, pady=2)
        ttk.Radiobutton(
            mode_frame, text='Chọn 1 vùng cells trong 1 sheet (vd A1:D20)',
            variable=self.mode_var, value='range',
            command=self._on_mode_change,
        ).pack(anchor='w', padx=5, pady=2)
        ttk.Radiobutton(
            mode_frame, text='Lấy 1 sheet (toàn bộ data của sheet đó)',
            variable=self.mode_var, value='sheet',
            command=self._on_mode_change,
        ).pack(anchor='w', padx=5, pady=2)
        ttk.Radiobutton(
            mode_frame, text='Lấy toàn bộ file (mọi sheet)',
            variable=self.mode_var, value='all',
            command=self._on_mode_change,
        ).pack(anchor='w', padx=5, pady=2)

        # ---- Params frame ----
        params_frame = ttk.LabelFrame(self.win, text='Tham số')
        params_frame.pack(fill='x', padx=10, pady=5)

        # Sheet (dùng cho custom/range/sheet)
        ttk.Label(params_frame, text='Sheet:').grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.sheet_var = tk.StringVar(value=current_selection.get('sheet') or '')
        self.sheet_combo = ttk.Combobox(params_frame, textvariable=self.sheet_var,
                                         width=40, state='readonly')
        self.sheet_combo.grid(row=0, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        self.sheet_combo.bind('<<ComboboxSelected>>', self._on_sheet_change)

        self.detected_columns = []  # [(col_idx, letter, header_text)]
        self._syncing_column_select = False

        # Cột (chỉ mode custom)
        ttk.Label(params_frame, text='Cột đã chọn:').grid(row=1, column=0, sticky='w', padx=5, pady=3)
        self.cols_var = tk.StringVar(value=current_selection.get('columns') or '')
        self.cols_entry = ttk.Entry(params_frame, textvariable=self.cols_var, width=30)
        self.cols_entry.grid(row=1, column=1, sticky='w', padx=5, pady=3)
        ttk.Label(params_frame, text='Tự sinh từ danh sách bên dưới; vẫn có thể gõ A,C-E,G',
                  foreground='gray').grid(row=1, column=2, columnspan=2, sticky='w', padx=5)

        ttk.Label(params_frame, text='Header/cột:').grid(row=2, column=0, sticky='nw', padx=5, pady=3)
        col_select_frame = ttk.Frame(params_frame)
        col_select_frame.grid(row=2, column=1, columnspan=3, sticky='we', padx=5, pady=3)
        self.column_listbox = tk.Listbox(
            col_select_frame,
            selectmode='extended',
            height=5,
            exportselection=False,
            width=72,
        )
        self.column_listbox.pack(side='left', fill='both', expand=True)
        self.column_listbox.bind('<<ListboxSelect>>', self._on_column_select)
        col_sb = ttk.Scrollbar(col_select_frame, orient='vertical', command=self.column_listbox.yview)
        col_sb.pack(side='left', fill='y')
        self.column_listbox.configure(yscrollcommand=col_sb.set)
        col_btn_frame = ttk.Frame(col_select_frame)
        col_btn_frame.pack(side='left', fill='y', padx=6)
        ttk.Button(
            col_btn_frame,
            text='Detect',
            command=lambda: self._detect_header_columns(auto_select=True, force_select_all=True)
        ).pack(fill='x', pady=1)
        ttk.Button(col_btn_frame, text='Chọn tất cả', command=self._select_all_detected_columns).pack(fill='x', pady=1)
        ttk.Button(col_btn_frame, text='Bỏ chọn', command=self._clear_detected_column_selection).pack(fill='x', pady=1)

        # Dòng (chỉ mode custom)
        ttk.Label(params_frame, text='Dòng (custom):').grid(row=3, column=0, sticky='w', padx=5, pady=3)
        row_range_frame = ttk.Frame(params_frame)
        row_range_frame.grid(row=3, column=1, columnspan=3, sticky='w', padx=5, pady=3)
        ttk.Label(row_range_frame, text='từ').pack(side='left')
        self.row_start_var = tk.StringVar(
            value=str(current_selection.get('row_start') or 1)
        )
        self.row_start_entry = ttk.Entry(row_range_frame, textvariable=self.row_start_var, width=8)
        self.row_start_entry.pack(side='left', padx=3)
        ttk.Label(row_range_frame, text='đến').pack(side='left')
        re_val = current_selection.get('row_end')
        self.row_end_var = tk.StringVar(value=str(re_val) if re_val else '')
        self.row_end_entry = ttk.Entry(row_range_frame, textvariable=self.row_end_var, width=8)
        self.row_end_entry.pack(side='left', padx=3)
        ttk.Label(row_range_frame, text='(rỗng = đến hết dòng có data)',
                  foreground='gray').pack(side='left', padx=5)

        # Range (chỉ mode range)
        ttk.Label(params_frame, text='Range (mode range):').grid(row=4, column=0, sticky='w', padx=5, pady=3)
        self.range_var = tk.StringVar(value=current_selection.get('range') or '')
        self.range_entry = ttk.Entry(params_frame, textvariable=self.range_var, width=20)
        self.range_entry.grid(row=4, column=1, sticky='w', padx=5, pady=3)
        ttk.Label(params_frame, text='VD: A1:D20', foreground='gray').grid(
            row=4, column=2, sticky='w', padx=5)
        ttk.Button(params_frame, text='Tự detect',
                   command=self._auto_detect_range).grid(row=4, column=3, padx=5, pady=3)

        # ---- Preview ----
        ttk.Separator(self.win, orient='horizontal').pack(fill='x', padx=10, pady=8)

        preview_header = ttk.Frame(self.win)
        preview_header.pack(fill='x', padx=10)
        ttk.Label(preview_header, text='Preview (10 row đầu của selection):',
                  font=('Segoe UI', 9, 'bold')).pack(side='left')
        ttk.Button(preview_header, text='Refresh',
                   command=self._refresh_preview).pack(side='left', padx=10)

        preview_frame = ttk.Frame(self.win)
        preview_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.preview_tree = ttk.Treeview(preview_frame, show='headings', height=10)
        self.preview_tree.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(preview_frame, orient='vertical', command=self.preview_tree.yview)
        sb.pack(side='right', fill='y')
        self.preview_tree.configure(yscrollcommand=sb.set)

        # ---- Action buttons ----
        btn_frame = ttk.Frame(self.win)
        btn_frame.pack(fill='x', padx=10, pady=10)
        ttk.Button(btn_frame, text='Cancel', command=self._cancel).pack(side='right', padx=5)
        ttk.Button(btn_frame, text='OK', command=self._ok).pack(side='right')

        # Initialize
        self._on_mode_change()
        if current_path and os.path.exists(current_path):
            self._load_file_metadata(current_path)
            if not self.sheet_var.get():
                sheets = list(self.sheet_combo['values'])
                if sheets:
                    self.sheet_var.set(sheets[0])
            if self.mode_var.get() == 'custom':
                self._detect_header_columns(auto_select=True)
            self._refresh_preview()

    # ----------------------------------------------------------------------

    def _browse_file(self):
        path = filedialog.askopenfilename(
            parent=self.win,
            title='Chọn file Excel',
            filetypes=[('Excel', '*.xlsx *.xlsm'), ('All files', '*.*')],
        )
        if path:
            self.path_var.set(path)
            self._load_file_metadata(path)
            # Reset sheet/range vì file mới
            self.sheet_var.set('')
            self.range_var.set('')
            self.cols_var.set('')
            sheets = list(self.sheet_combo['values'])
            if sheets:
                self.sheet_var.set(sheets[0])
            self._detect_header_columns(auto_select=True, force_select_all=True)
            self._refresh_preview()

    def _load_file_metadata(self, path):
        """Đọc danh sách sheet, fill vào combobox."""
        if not path.lower().endswith(('.xlsx', '.xlsm')):
            self.file_info_label.config(
                text='⚠ File không phải Excel. Phase 1 chỉ hỗ trợ .xlsx và .xlsm',
                foreground='orange'
            )
            self.sheet_combo['values'] = []
            return
        try:
            from openpyxl import load_workbook
            # read_only=True OK cho việc chỉ list sheet names (nhanh hơn với file 95 sheet)
            wb = load_workbook(path, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            self.sheet_combo['values'] = sheet_names
            display_names = ', '.join(sheet_names[:5])
            if len(sheet_names) > 5:
                display_names += f', ... (+{len(sheet_names)-5})'
            self.file_info_label.config(
                text=f'✓ Đã đọc file. Tìm thấy {len(sheet_names)} sheet: {display_names}',
                foreground='blue'
            )
        except Exception as e:
            self.file_info_label.config(text=f'✗ Lỗi đọc file: {e}', foreground='red')
            self.sheet_combo['values'] = []

    def _on_mode_change(self):
        mode = self.mode_var.get()
        # Sheet: enable cho custom/range/sheet
        if mode == 'all':
            self.sheet_combo.config(state='disabled')
        else:
            self.sheet_combo.config(state='readonly')
        # Cột + dòng: chỉ custom
        cols_state = 'normal' if mode == 'custom' else 'disabled'
        self.cols_entry.config(state=cols_state)
        self.column_listbox.config(state=cols_state)
        self.row_start_entry.config(state=cols_state)
        self.row_end_entry.config(state=cols_state)
        # Range: chỉ range
        range_state = 'normal' if mode == 'range' else 'disabled'
        self.range_entry.config(state=range_state)
        self._refresh_preview()

    def _on_sheet_change(self, _event=None):
        if self.mode_var.get() == 'custom':
            self._detect_header_columns(auto_select=True)
        if self.mode_var.get() == 'range' and not self.range_var.get():
            self._auto_detect_range()
        self._refresh_preview()

    def _columns_to_spec(self, col_indexes):
        """Compress [1,2,3,5] -> 'A:C,E'."""
        from openpyxl.utils import get_column_letter
        indexes = sorted(set(int(c) for c in col_indexes if c))
        if not indexes:
            return ''
        parts = []
        start = prev = indexes[0]
        for c in indexes[1:]:
            if c == prev + 1:
                prev = c
                continue
            parts.append(
                get_column_letter(start) if start == prev
                else f'{get_column_letter(start)}:{get_column_letter(prev)}'
            )
            start = prev = c
        parts.append(
            get_column_letter(start) if start == prev
            else f'{get_column_letter(start)}:{get_column_letter(prev)}'
        )
        return ','.join(parts)

    def _nonempty_bounds_in_scan(self, ws, max_scan=50):
        """Trả về (max_row_scan, max_col_with_data) không phụ thuộc dimension cache."""
        max_row_scan = min(ws.max_row or 1, max_scan)
        max_col = 1
        # Normal worksheet có _cells đầy đủ sau load_workbook(read_only=False).
        for (_r, _c), cell in getattr(ws, '_cells', {}).items():
            if _r > max_row_scan:
                continue
            v = cell.value
            if v is None or str(v).strip() == '':
                continue
            if _c > max_col:
                max_col = _c
        # Fallback nếu _cells không có đủ dữ liệu.
        if max_col <= 1:
            max_col = ws.max_column or 1
        return max_row_scan, max_col

    def _detect_header_row(self, ws):
        """Detect header row trong 50 dòng đầu, ưu tiên nhiều ô text không rỗng."""
        max_scan, max_col = self._nonempty_bounds_in_scan(ws)
        best_row = 1
        best_score = -1
        for r in range(1, max_scan + 1):
            nonempty = 0
            text_cells = 0
            numeric_cells = 0
            total_text_len = 0
            for c in range(1, max_col + 1):
                v = ws.cell(row=r, column=c).value
                if v is None or str(v).strip() == '':
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
        return best_row

    def _detect_header_columns(self, auto_select=True, force_select_all=False):
        """Detect header columns của sheet hiện tại và fill listbox chọn cột."""
        path = self.path_var.get()
        sheet = self.sheet_var.get()
        self.detected_columns = []
        self.column_listbox.delete(0, 'end')
        if not path or not sheet or not os.path.exists(path):
            return
        if not path.lower().endswith(('.xlsx', '.xlsm')):
            return
        try:
            from openpyxl import load_workbook
            from openpyxl.utils import get_column_letter
            # Không dùng read_only ở đây: nhiều file Excel thật có dimension cache
            # sai làm ws.max_column chỉ trả 1 dù sheet có nhiều cột.
            wb = load_workbook(path, read_only=False, data_only=True)
            if sheet not in wb.sheetnames:
                wb.close()
                return
            ws = wb[sheet]
            header_row = self._detect_header_row(ws)
            _max_scan, max_col = self._nonempty_bounds_in_scan(ws)
            columns = []
            for c in range(1, max_col + 1):
                header = ws.cell(row=header_row, column=c).value
                header_text = str(header).strip() if header is not None else ''
                if not header_text:
                    continue
                letter = get_column_letter(c)
                columns.append((c, letter, header_text))
            wb.close()
        except Exception as e:
            self.file_info_label.config(
                text=f'⚠ Không detect được header: {e}', foreground='orange')
            return

        self.detected_columns = columns
        self._syncing_column_select = True
        try:
            for c_idx, letter, header_text in columns:
                display = f'{letter}  —  {header_text}'
                self.column_listbox.insert('end', display)

            existing = parse_excel_column_spec(self.cols_var.get().strip())
            if existing and not force_select_all:
                selected_indexes = {
                    i for i, (c_idx, _letter, _text) in enumerate(columns)
                    if c_idx in existing
                }
            else:
                selected_indexes = set(range(len(columns))) if auto_select else set()
            for i in selected_indexes:
                self.column_listbox.selection_set(i)
            if auto_select and columns and (not existing or force_select_all):
                self.cols_var.set(self._columns_to_spec(c for c, _l, _h in columns))
            if auto_select:
                self.row_start_var.set(str(header_row))
        finally:
            self._syncing_column_select = False

    def _on_column_select(self, _event=None):
        if self._syncing_column_select:
            return
        selected = self.column_listbox.curselection()
        cols = [self.detected_columns[i][0] for i in selected]
        self.cols_var.set(self._columns_to_spec(cols))
        self._refresh_preview()

    def _selected_detected_column_indexes(self):
        return [
            self.detected_columns[i][0]
            for i in self.column_listbox.curselection()
            if i < len(self.detected_columns)
        ]

    def _select_all_detected_columns(self):
        self.column_listbox.selection_set(0, 'end')
        self._on_column_select()

    def _clear_detected_column_selection(self):
        self.column_listbox.selection_clear(0, 'end')
        self._on_column_select()

    def _auto_detect_range(self):
        """Tự detect range = từ ô đầu tiên có data đến ô cuối cùng có data.

        Strategy:
          1. Thử ws.calculate_dimension() (nhanh, chính xác nếu file đã lưu dim)
          2. Nếu kết quả 'A1:A1' nghi ngờ → fallback iterate cells để tìm extent
        """
        path = self.path_var.get()
        sheet = self.sheet_var.get()
        if not path or not sheet or not path.lower().endswith(('.xlsx', '.xlsm')):
            return
        try:
            from openpyxl import load_workbook
            from openpyxl.utils import get_column_letter, range_boundaries
            wb = load_workbook(path, data_only=True)
            ws = wb[sheet]

            # Thử dimension trước
            try:
                dim = ws.calculate_dimension()  # vd 'A1:Z100'
            except Exception:
                dim = None

            range_str = None
            if dim and dim != 'A1:A1':
                # Verify dim không phải fake (1 ô)
                try:
                    min_col, min_row, max_col, max_row = range_boundaries(dim)
                    if (max_col, max_row) != (min_col, min_row):
                        range_str = dim
                except Exception:
                    pass

            if not range_str:
                # Fallback: iterate qua cells để tìm extent
                min_row, max_row = None, None
                min_col, max_col = None, None
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue
                        r, c = cell.row, cell.column
                        if min_row is None or r < min_row: min_row = r
                        if max_row is None or r > max_row: max_row = r
                        if min_col is None or c < min_col: min_col = c
                        if max_col is None or c > max_col: max_col = c
                if min_row is not None:
                    range_str = (f'{get_column_letter(min_col)}{min_row}:'
                                 f'{get_column_letter(max_col)}{max_row}')
                else:
                    range_str = 'A1:A1'  # sheet rỗng

            wb.close()
            self.range_var.set(range_str)
            self._refresh_preview()
        except Exception as e:
            messagebox.showerror('Lỗi', f'Không detect được range: {e}', parent=self.win)

    def _refresh_preview(self):
        # Clear tree
        for col in self.preview_tree['columns']:
            self.preview_tree.heading(col, text='')
        self.preview_tree['columns'] = ()
        for row in self.preview_tree.get_children():
            self.preview_tree.delete(row)

        path = self.path_var.get()
        if not path or not path.lower().endswith(('.xlsx', '.xlsm')):
            return
        if not os.path.exists(path):
            return

        try:
            from openpyxl import load_workbook
            from openpyxl.utils import get_column_letter
            # Bỏ read_only để max_row/max_column populate đúng
            wb = load_workbook(path, data_only=True)
            mode = self.mode_var.get()

            rows = []
            col_headers = []  # Cột hiển thị (vd ['A','C','E'] cho custom)

            if mode == 'all':
                if not wb.worksheets:
                    wb.close()
                    return
                ws = wb.worksheets[0]
                rows_iter = ws.iter_rows(values_only=True, max_row=10)
                rows = list(rows_iter)
                if rows:
                    n = max(len(r) for r in rows)
                    col_headers = [get_column_letter(i+1) for i in range(n)]
            elif mode == 'sheet':
                sheet = self.sheet_var.get()
                if not sheet or sheet not in wb.sheetnames:
                    wb.close()
                    return
                ws = wb[sheet]
                rows = list(ws.iter_rows(values_only=True, max_row=10))
                if rows:
                    n = max(len(r) for r in rows)
                    col_headers = [get_column_letter(i+1) for i in range(n)]
            elif mode == 'range':
                sheet = self.sheet_var.get()
                range_str = self.range_var.get().strip()
                if not sheet or not range_str or sheet not in wb.sheetnames:
                    wb.close()
                    return
                ws = wb[sheet]
                try:
                    sel = ws[range_str]
                except Exception as e:
                    self.file_info_label.config(
                        text=f'⚠ Range không hợp lệ: {e}', foreground='orange')
                    wb.close()
                    return
                if hasattr(sel, 'value'):
                    rows = [(sel.value,)]
                    col_headers = [get_column_letter(sel.column)]
                elif isinstance(sel, tuple):
                    if not sel:
                        wb.close()
                        return
                    if isinstance(sel[0], tuple):
                        rows = [tuple(c.value for c in r) for r in sel[:10]]
                        if rows and sel[0]:
                            col_headers = [get_column_letter(c.column) for c in sel[0]]
                    else:
                        rows = [tuple(c.value for c in sel)]
                        col_headers = [get_column_letter(c.column) for c in sel]
            elif mode == 'custom':
                sheet = self.sheet_var.get()
                if not sheet or sheet not in wb.sheetnames:
                    wb.close()
                    return
                ws = wb[sheet]
                cols_spec = self.cols_var.get().strip()
                row_start_s = self.row_start_var.get().strip()
                row_end_s = self.row_end_var.get().strip()

                try:
                    rs = max(1, int(row_start_s)) if row_start_s else 1
                except ValueError:
                    rs = 1
                try:
                    re_ = int(row_end_s) if row_end_s else None
                except ValueError:
                    re_ = None
                max_col = ws.max_column or 1
                max_row = ws.max_row or 1
                if re_ is None:
                    re_ = max_row
                re_ = max(rs, re_)

                if cols_spec:
                    col_idxs = parse_excel_column_spec(cols_spec)
                elif self.detected_columns:
                    col_idxs = self._selected_detected_column_indexes()
                else:
                    col_idxs = list(range(1, max_col + 1))

                if not col_idxs:
                    wb.close()
                    return

                # Build preview (max 10 rows từ rs)
                preview_end = min(re_, rs + 9)
                for r_idx in range(rs, preview_end + 1):
                    row_vals = tuple(ws.cell(row=r_idx, column=c).value for c in col_idxs)
                    rows.append(row_vals)
                detected_map = {
                    c_idx: header_text
                    for c_idx, _letter, header_text in self.detected_columns
                }
                col_headers = [
                    f'{get_column_letter(c)}: {detected_map[c]}'
                    if detected_map.get(c) else get_column_letter(c)
                    for c in col_idxs
                ]

            wb.close()

            if not rows:
                return

            ncols = max(len(r) for r in rows) if rows else 0
            cols = tuple(f'c{i}' for i in range(ncols))
            self.preview_tree['columns'] = cols
            for i, c in enumerate(cols):
                hdr = col_headers[i] if i < len(col_headers) else str(i+1)
                self.preview_tree.heading(c, text=hdr)
                self.preview_tree.column(c, width=110, anchor='w')

            for row in rows:
                vals = [str(v) if v is not None else '' for v in row]
                while len(vals) < ncols:
                    vals.append('')
                vals = [v[:80] for v in vals]
                self.preview_tree.insert('', 'end', values=vals)
        except Exception as e:
            import traceback; traceback.print_exc()
            self.file_info_label.config(
                text=f'⚠ Lỗi preview: {e}', foreground='orange')

    def _ok(self):
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning('Thiếu file', 'Vui lòng chọn file Excel', parent=self.win)
            return
        if not os.path.exists(path):
            messagebox.showwarning('File không tồn tại', f'Không tìm thấy:\n{path}',
                                    parent=self.win)
            return
        mode = self.mode_var.get()
        sel = {'mode': mode}
        if mode in ('sheet', 'range', 'custom'):
            sheet = self.sheet_var.get()
            if not sheet:
                messagebox.showwarning('Thiếu sheet', 'Vui lòng chọn sheet', parent=self.win)
                return
            sel['sheet'] = sheet
        if mode == 'range':
            r = self.range_var.get().strip()
            if not r:
                messagebox.showwarning('Thiếu range',
                                        'Vui lòng nhập range (VD: A1:D20) '
                                        'hoặc bấm "Tự detect"', parent=self.win)
                return
            sel['range'] = r
        if mode == 'custom':
            cols_spec = self.cols_var.get().strip()
            if not cols_spec and self.detected_columns and self.column_listbox.curselection():
                cols_spec = self._columns_to_spec(self._selected_detected_column_indexes())
                self.cols_var.set(cols_spec)
            if self.detected_columns and not cols_spec and not self.column_listbox.curselection():
                messagebox.showwarning(
                    'Chưa chọn cột',
                    'Vui lòng chọn ít nhất một cột trong danh sách header '
                    'hoặc gõ cột dạng A,C-E,G.',
                    parent=self.win
                )
                return
            if cols_spec and not parse_excel_column_spec(cols_spec):
                messagebox.showwarning(
                    'Cột không hợp lệ',
                    'Vui lòng nhập cột dạng A,C-E,G hoặc A:C',
                    parent=self.win
                )
                return
            sel['columns'] = cols_spec
            try:
                sel['row_start'] = max(1, int(self.row_start_var.get().strip() or 1))
            except ValueError:
                sel['row_start'] = 1
            re_s = self.row_end_var.get().strip()
            try:
                sel['row_end'] = int(re_s) if re_s else None
            except ValueError:
                sel['row_end'] = None
        self.result = {'path': path, 'selection': sel}
        self.win.destroy()

    def _cancel(self):
        self.result = None
        self.win.destroy()

    @staticmethod
    def show(parent, current_path='', current_selection=None):
        dialog = ExcelSelectionDialog(parent, current_path, current_selection)
        dialog.win.wait_window()
        return dialog.result


# ============================================================================
# GUI Wizard (Tkinter)
# ============================================================================

class WizardApp:

    def __init__(self):
        self.model = DocxModel()

        self.root = tk.Tk()
        self.root.title('DOCX Template Builder')
        self.root.geometry('1100x780')

        # Notebook (tabs cho 7 bước)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        self.tab_step1 = ttk.Frame(self.notebook)
        self.tab_step2 = ttk.Frame(self.notebook)
        self.tab_step3 = ttk.Frame(self.notebook)
        self.tab_step4 = ttk.Frame(self.notebook)
        self.tab_step5 = ttk.Frame(self.notebook)
        self.tab_step6 = ttk.Frame(self.notebook)
        self.tab_step7 = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_step1, text='1. Chọn file & trang')
        self.notebook.add(self.tab_step2, text='2. Defaults & Sections')
        self.notebook.add(self.tab_step3, text='3. Phần mở đầu')
        self.notebook.add(self.tab_step4, text='4. Headings/Numbering/H&F')
        self.notebook.add(self.tab_step5, text='5. Danh sách heading')
        self.notebook.add(self.tab_step6, text='6. Giữ/xoá nội dung')
        self.notebook.add(self.tab_step7, text='7. Lưu template')

        # Disable các tab sau cho tới khi load file
        for i in range(1, 7):
            self.notebook.tab(i, state='disabled')

        self._build_step1()
        self._build_step2_placeholder()
        self._build_step3_placeholder()
        self._build_step4_placeholder()
        self._build_step5_placeholder()
        self._build_step6_placeholder()
        self._build_step7_placeholder()

    def run(self):
        self.root.mainloop()

    # -------------------------------------------------------------------
    # STEP 1: Chọn file + nhập trang
    # -------------------------------------------------------------------

    def _build_step1(self):
        f = self.tab_step1
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 1: Chọn file Word và xác định ranh giới',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)

        # File picker
        file_frame = ttk.Frame(f)
        file_frame.pack(fill='x', padx=10, pady=5)
        ttk.Label(file_frame, text='File gốc (.docx):').pack(side='left')
        self.file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_var, width=80).pack(side='left', padx=5)
        ttk.Button(file_frame, text='Chọn...', command=self._pick_file).pack(side='left')

        ttk.Button(f, text='Nạp file & phân tích',
                   command=self._load_file).pack(anchor='w', padx=10, pady=5)

        # Info
        self.info_label = ttk.Label(f, text='', foreground='blue', justify='left')
        self.info_label.pack(anchor='w', padx=10, pady=5)

        ttk.Separator(f, orient='horizontal').pack(fill='x', padx=10, pady=10)

        ttk.Label(f, text='Heading nào là bắt đầu phần nội dung?',
                  font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=10)
        ttk.Label(f,
                  text='Tool tự động phát hiện và tô đậm dòng đề xuất. Có thể click chọn '
                       'lại nếu thấy không đúng. Mọi heading TRƯỚC dòng được chọn được coi '
                       'là phần mở đầu (trang bìa, changelog, mục lục…). Heading được chọn '
                       'và các heading sau đó là phần nội dung.',
                  foreground='gray', wraplength=900, justify='left').pack(anchor='w', padx=10, pady=2)

        # Treeview heading
        list_frame = ttk.Frame(f)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)

        cols = ('idx', 'level', 'text')
        self.intro_tree = ttk.Treeview(list_frame, columns=cols,
                                        show='headings', height=14, selectmode='browse')
        self.intro_tree.heading('idx', text='Element idx')
        self.intro_tree.heading('level', text='Lvl')
        self.intro_tree.heading('text', text='Heading text')
        self.intro_tree.column('idx', width=80, anchor='center')
        self.intro_tree.column('level', width=50, anchor='center')
        self.intro_tree.column('text', width=900)
        self.intro_tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(list_frame, orient='vertical', command=self.intro_tree.yview)
        sb.pack(side='right', fill='y')
        self.intro_tree.configure(yscrollcommand=sb.set)

        # Tag style cho dòng đề xuất
        self.intro_tree.tag_configure('suggested', background='#fff7d6', font=('Segoe UI', 9, 'bold'))
        self.intro_tree.tag_configure('intro', foreground='#999999')

        ttk.Button(f, text='Lưu & sang bước 2 →',
                   command=self._save_step1_and_go).pack(anchor='e', padx=10, pady=15)

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title='Chọn file Word',
            filetypes=[('Word files', '*.docx'), ('All', '*.*')]
        )
        if path:
            self.file_var.set(path)

    def _load_file(self):
        path = self.file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror('Lỗi', 'Vui lòng chọn file hợp lệ')
            return
        try:
            self.model.load(path)
            detected = self.model.get_total_pages_detected()
            metadata = self.model.metadata_pages
            words = self.model.get_word_count()
            intro_end = self.model.get_intro_end_idx()
            n_headings = len(self.model.get_all_headings())

            info_lines = [
                f'✓ Đã nạp file thành công.',
                f'   • Tổng số element trong body: {len(self.model.body_elements)}',
                f'   • Số heading phát hiện: {n_headings}',
            ]
            if words:
                info_lines.append(f'   • Số chữ (theo metadata): {words:,}')
            if metadata:
                info_lines.append(f'   • Số trang ước tính: {metadata}')
            self.info_label.config(text='\n'.join(info_lines))

            # Populate heading tree với tất cả heading
            self._populate_intro_tree()

            # Enable các tab khác
            for i in range(1, 7):
                self.notebook.tab(i, state='normal')
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror('Lỗi nạp file', str(e))

    def _populate_intro_tree(self):
        """Đổ tất cả heading vào treeview Step 1, tự đề xuất dòng theo heuristic."""
        for row in self.intro_tree.get_children():
            self.intro_tree.delete(row)

        all_h = self.model.get_all_headings()
        suggested = self.model.get_intro_end_idx()  # từ heuristic

        suggested_iid = None
        for h in all_h:
            indent = '   ' * (h['level'] - 1)
            iid = str(h['original_idx'])
            tag = ()
            if h['original_idx'] == suggested:
                tag = ('suggested',)
                suggested_iid = iid
            elif h['original_idx'] < suggested:
                tag = ('intro',)
            self.intro_tree.insert(
                '', 'end', iid=iid,
                values=(h['original_idx'], h['level'], indent + h['text']),
                tags=tag
            )

        # Chọn sẵn dòng đề xuất
        if suggested_iid:
            self.intro_tree.selection_set(suggested_iid)
            self.intro_tree.see(suggested_iid)

    def _save_step1_and_go(self):
        if self.model.path is None or not self.model.config.get('doc_defaults'):
            messagebox.showerror(
                'Chưa nạp file',
                'Vui lòng click "Nạp file & phân tích" trước khi sang bước 2.'
            )
            return

        # Lấy heading được chọn = mốc bắt đầu của phần nội dung
        sel = self.intro_tree.selection()
        if not sel:
            # Không có heading nào hoặc user chưa chọn
            if not self.intro_tree.get_children():
                # File không có heading → coi toàn bộ là content
                self.model.config['intro_end_idx'] = 0
            else:
                messagebox.showwarning(
                    'Chưa chọn',
                    'Vui lòng chọn dòng heading bắt đầu phần nội dung trong danh sách.'
                )
                return
        else:
            self.model.config['intro_end_idx'] = int(sel[0])

        # Reset heading list để build lại theo intro_end_idx mới
        self.model.config['headings_list'] = []

        # Build các tab kế
        self._build_step2()
        self._build_step3()
        self._build_step4()
        self._build_step5()
        self._build_step6()
        self._build_step7()
        self.notebook.select(1)

    # -------------------------------------------------------------------
    # STEP 2: Document defaults + Sections
    # -------------------------------------------------------------------

    def _build_step2_placeholder(self):
        ttk.Label(self.tab_step2, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step2(self):
        f = self.tab_step2
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 2: Document defaults & Section properties',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)

        # --- Doc defaults form ---
        dd_frame = ttk.LabelFrame(f, text='Document defaults (mặc định toàn doc)')
        dd_frame.pack(fill='x', padx=10, pady=5)
        cfg = self.model.config.get('doc_defaults') or {}
        # Đảm bảo có đủ default keys phòng trường hợp file gốc thiếu
        defaults_fallback = {
            'font': 'Times New Roman', 'size_pt': 13.0, 'lang': 'en-US',
            'line_spacing': '1.5', 'jc': 'both',
            'space_before': 6.0, 'space_after': 0.0, 'indent_left_cm': 0.0,
        }
        for k, v in defaults_fallback.items():
            cfg.setdefault(k, v)
        self.model.config['doc_defaults'] = cfg

        self.dd_vars = {}
        rows = [
            ('font', 'Font:', cfg['font'], 25),
            ('size_pt', 'Cỡ chữ (pt):', cfg['size_pt'], 8),
            ('lang', 'Language:', cfg['lang'], 10),
            ('line_spacing', 'Line spacing (1.0/1.5/2.0…):', cfg['line_spacing'], 8),
            ('jc', 'Alignment (left/center/right/both):', cfg['jc'], 10),
            ('space_before', 'Space before (pt):', cfg['space_before'], 8),
            ('space_after', 'Space after (pt):', cfg['space_after'], 8),
            ('indent_left_cm', 'Indent left (cm):', cfg['indent_left_cm'], 8),
        ]
        for i, (key, label, val, width) in enumerate(rows):
            ttk.Label(dd_frame, text=label).grid(row=i, column=0, sticky='w', padx=5, pady=2)
            v = tk.StringVar(value=str(val))
            ttk.Entry(dd_frame, textvariable=v, width=width).grid(row=i, column=1, sticky='w', padx=5)
            self.dd_vars[key] = v

        # --- Sections form ---
        sect_frame = ttk.LabelFrame(f, text='Sections (mỗi section = một bộ thiết lập trang)')
        sect_frame.pack(fill='both', expand=True, padx=10, pady=5)

        self.section_widgets = []
        sections_cfg = self.model.config.get('sections') or []
        if not sections_cfg:
            ttk.Label(sect_frame, text='(Không tìm thấy section nào trong file)',
                      foreground='red').pack(padx=5, pady=5)
        for i, sect_cfg in enumerate(sections_cfg):
            self._build_section_form(sect_frame, i, sect_cfg)

        ttk.Button(f, text='Lưu & sang bước 3 →',
                   command=self._save_step2_and_go).pack(anchor='e', padx=10, pady=10)

    def _build_section_form(self, parent, i, cfg):
        # Đảm bảo đủ key (file có thể thiếu pgSz/pgMar)
        section_defaults = {
            'pg_w_cm': 21.0, 'pg_h_cm': 29.7, 'orient': 'portrait',
            'top_cm': 2.54, 'right_cm': 2.54, 'bottom_cm': 2.54, 'left_cm': 2.54,
            'header_cm': 1.27, 'footer_cm': 1.27, 'gutter_cm': 0.0,
            'titlePg': False, 'header_refs': {}, 'footer_refs': {},
        }
        for k, v in section_defaults.items():
            cfg.setdefault(k, v)

        sf = ttk.LabelFrame(parent, text=f'Section {i+1}')
        sf.pack(fill='x', padx=5, pady=5)
        widgets = {}
        rows = [
            ('pg_w_cm', 'Khổ rộng (cm):'), ('pg_h_cm', 'Khổ cao (cm):'),
            ('orient', 'Orientation (portrait/landscape):'),
            ('top_cm', 'Lề trên (cm):'), ('right_cm', 'Lề phải (cm):'),
            ('bottom_cm', 'Lề dưới (cm):'), ('left_cm', 'Lề trái (cm):'),
            ('header_cm', 'Header (cm):'), ('footer_cm', 'Footer (cm):'),
            ('gutter_cm', 'Gutter (cm):'),
        ]
        for r_idx, (key, label) in enumerate(rows):
            row, col = divmod(r_idx, 2)
            ttk.Label(sf, text=label).grid(row=row, column=col*2, sticky='w', padx=5, pady=2)
            v = tk.StringVar(value=str(cfg.get(key, '')))
            ttk.Entry(sf, textvariable=v, width=12).grid(row=row, column=col*2+1, sticky='w', padx=5)
            widgets[key] = v

        # titlePg
        v_tp = tk.BooleanVar(value=cfg.get('titlePg', False))
        ttk.Checkbutton(sf, text='titlePg (trang đầu có header/footer riêng)',
                        variable=v_tp).grid(row=99, column=0, columnspan=4, sticky='w', padx=5)
        widgets['titlePg'] = v_tp

        # Read-only: header/footer refs
        refs_text = f"Header refs: {cfg.get('header_refs', {})}  |  Footer refs: {cfg.get('footer_refs', {})}"
        ttk.Label(sf, text=refs_text, foreground='gray').grid(row=100, column=0, columnspan=4, sticky='w', padx=5, pady=2)

        self.section_widgets.append(widgets)

    def _save_step2_and_go(self):
        # Save doc defaults
        for key, v in self.dd_vars.items():
            try:
                if key in ('size_pt', 'space_before', 'space_after', 'indent_left_cm'):
                    self.model.config['doc_defaults'][key] = float(v.get())
                else:
                    self.model.config['doc_defaults'][key] = v.get().strip()
            except ValueError:
                messagebox.showerror('Lỗi', f'Giá trị không hợp lệ ở: {key}')
                return
        # Save sections
        for i, widgets in enumerate(self.section_widgets):
            cfg = self.model.config['sections'][i]
            try:
                for key, v in widgets.items():
                    if key == 'titlePg':
                        cfg[key] = v.get()
                    elif key == 'orient':
                        cfg[key] = v.get().strip()
                    else:
                        cfg[key] = float(v.get())
            except ValueError:
                messagebox.showerror('Lỗi', f'Giá trị không hợp lệ ở section {i+1}')
                return
        self.notebook.select(2)

    # -------------------------------------------------------------------
    # STEP 3: Phần mở đầu (text + ảnh)
    # -------------------------------------------------------------------

    def _build_step3_placeholder(self):
        ttk.Label(self.tab_step3, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step3(self):
        f = self.tab_step3
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 3: Sửa nội dung phần mở đầu (text holders + ảnh)',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)

        # Notebook nội bộ: text | images
        sub_nb = ttk.Notebook(f)
        sub_nb.pack(fill='both', expand=True, padx=10, pady=5)

        # --- Text tab ---
        text_tab = ttk.Frame(sub_nb)
        sub_nb.add(text_tab, text='Text holders')

        ttk.Label(text_tab, text='Click vào text rồi sửa ở ô bên dưới. '
                                 'Khi save, text gốc sẽ được thay bằng text mới.',
                  foreground='gray').pack(anchor='w', padx=5, pady=2)

        # Treeview các text block
        cols = ('idx', 'kind', 'text')
        tree = ttk.Treeview(text_tab, columns=cols, show='headings', height=15)
        tree.heading('idx', text='ID')
        tree.heading('kind', text='Loại')
        tree.heading('text', text='Nội dung gốc')
        tree.column('idx', width=80)
        tree.column('kind', width=80)
        tree.column('text', width=700)
        tree.pack(fill='both', expand=True, padx=5, pady=5)

        for idx, kind, text in self.model.get_intro_text_blocks():
            tree.insert('', 'end', values=(str(idx), kind, text))

        # Edit area
        edit_frame = ttk.LabelFrame(text_tab, text='Sửa text được chọn')
        edit_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(edit_frame, text='Text gốc:').grid(row=0, column=0, sticky='w', padx=5)
        self.intro_old_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.intro_old_var, width=80, state='readonly').grid(row=0, column=1, padx=5)
        ttk.Label(edit_frame, text='Text mới:').grid(row=1, column=0, sticky='w', padx=5)
        self.intro_new_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.intro_new_var, width=80).grid(row=1, column=1, padx=5)
        ttk.Button(edit_frame, text='Áp dụng thay thế',
                   command=lambda: self._apply_text_repl()).grid(row=2, column=1, sticky='w', padx=5, pady=5)

        # Khi click row
        def on_select(_):
            sel = tree.selection()
            if sel:
                vals = tree.item(sel[0])['values']
                self.intro_old_var.set(vals[2])
                self.intro_new_var.set(self.model.config['intro_replacements'].get(vals[2], vals[2]))
        tree.bind('<<TreeviewSelect>>', on_select)

        # Replacements list
        repl_frame = ttk.LabelFrame(text_tab, text='Danh sách thay thế đã thiết lập')
        repl_frame.pack(fill='x', padx=5, pady=5)
        self.repl_listbox = tk.Listbox(repl_frame, height=4)
        self.repl_listbox.pack(fill='x', padx=5, pady=2)
        self._refresh_repl_list()

        # --- Images tab ---
        img_tab = ttk.Frame(sub_nb)
        sub_nb.add(img_tab, text='Ảnh')
        ttk.Label(img_tab, text='Mỗi ảnh có thể chọn file PNG/JPG mới để thay thế.',
                  foreground='gray').pack(anchor='w', padx=5, pady=2)

        self.img_vars = {}
        for rid, target in self.model.get_intro_images():
            row = ttk.Frame(img_tab)
            row.pack(fill='x', padx=5, pady=3)
            ttk.Label(row, text=f'rId={rid}  →  {target}', width=50).pack(side='left')
            v = tk.StringVar(value=self.model.config['intro_images'].get(rid, ''))
            ttk.Entry(row, textvariable=v, width=40).pack(side='left', padx=5)
            ttk.Button(row, text='Chọn...',
                       command=lambda rr=rid, vv=v: self._pick_image(rr, vv)).pack(side='left')
            self.img_vars[rid] = v

        ttk.Button(f, text='Lưu & sang bước 4 →',
                   command=self._save_step3_and_go).pack(anchor='e', padx=10, pady=10)

    def _apply_text_repl(self):
        old = self.intro_old_var.get()
        new = self.intro_new_var.get()
        if not old:
            return
        if old == new:
            self.model.config['intro_replacements'].pop(old, None)
        else:
            self.model.config['intro_replacements'][old] = new
        self._refresh_repl_list()

    def _refresh_repl_list(self):
        self.repl_listbox.delete(0, 'end')
        for old, new in self.model.config['intro_replacements'].items():
            self.repl_listbox.insert('end', f'"{old}"  →  "{new}"')

    def _pick_image(self, rid, var):
        path = filedialog.askopenfilename(
            title='Chọn ảnh thay thế',
            filetypes=[('Image', '*.png *.jpg *.jpeg *.gif *.bmp'), ('All', '*.*')]
        )
        if path:
            var.set(path)

    def _save_step3_and_go(self):
        # Save image overrides
        for rid, v in self.img_vars.items():
            p = v.get().strip()
            if p:
                self.model.config['intro_images'][rid] = p
            else:
                self.model.config['intro_images'].pop(rid, None)
        self.notebook.select(3)

    # -------------------------------------------------------------------
    # STEP 4: Heading styles + Numbering + Header/Footer
    # -------------------------------------------------------------------

    def _build_step4_placeholder(self):
        ttk.Label(self.tab_step4, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step4(self):
        f = self.tab_step4
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 4: Heading styles, Numbering, Header & Footer',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)

        sub_nb = ttk.Notebook(f)
        sub_nb.pack(fill='both', expand=True, padx=10, pady=5)

        # --- Heading styles ---
        h_tab = ttk.Frame(sub_nb)
        sub_nb.add(h_tab, text='Heading styles')

        # Scroll frame
        canvas = tk.Canvas(h_tab)
        scrollbar = ttk.Scrollbar(h_tab, orient='vertical', command=canvas.yview)
        scrollable = ttk.Frame(canvas)
        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.heading_widgets = {}
        for sid, data in self.model.config['heading_styles'].items():
            self._build_heading_form(scrollable, sid, data)

        # --- Numbering (read-only summary) ---
        num_tab = ttk.Frame(sub_nb)
        sub_nb.add(num_tab, text='Numbering')
        ttk.Label(num_tab, text='Danh sách numbering definitions (numId) trong file. '
                                'Hiện tại tool chỉ hiển thị, không sửa được.',
                  foreground='gray').pack(anchor='w', padx=5, pady=2)
        if self.model.numbering_xml is not None:
            text = tk.Text(num_tab, height=20, wrap='none')
            text.pack(fill='both', expand=True, padx=5, pady=5)
            for an in self.model.numbering_xml.getroot().findall(w('abstractNum')):
                aid = an.get(w('abstractNumId'))
                mlt = an.find(w('multiLevelType'))
                mlt_v = mlt.get(w('val')) if mlt is not None else 'none'
                text.insert('end', f'abstractNumId={aid} multiLevelType={mlt_v}\n')
                for lvl in an.findall(w('lvl'))[:4]:
                    ilvl = lvl.get(w('ilvl'))
                    fmt = lvl.find(w('numFmt'))
                    fmt_v = fmt.get(w('val')) if fmt is not None else ''
                    txt_el = lvl.find(w('lvlText'))
                    txt_v = txt_el.get(w('val')) if txt_el is not None else ''
                    text.insert('end', f'  lvl{ilvl}: fmt={fmt_v} text="{txt_v}"\n')
                text.insert('end', '\n')

        # --- Header/Footer ---
        hf_tab = ttk.Frame(sub_nb)
        sub_nb.add(hf_tab, text='Header & Footer')
        ttk.Label(hf_tab, text='Sửa text trong các file header/footer. '
                               'Lưu ý: chỉ sửa được phần text đơn thuần, watermark giữ nguyên.',
                  foreground='gray').pack(anchor='w', padx=5, pady=2)

        self.hf_vars = {}
        for fname, content in self.model.get_header_footer_files().items():
            row = ttk.Frame(hf_tab)
            row.pack(fill='x', padx=5, pady=3)
            ttk.Label(row, text=f'{fname}:', width=18).pack(side='left')
            v = tk.StringVar(value=self.model.config['header_footer_text'].get(fname, content))
            ttk.Entry(row, textvariable=v, width=80).pack(side='left', padx=5)
            self.hf_vars[fname] = v

        ttk.Button(f, text='Lưu & sang bước 5 →',
                   command=self._save_step4_and_go).pack(anchor='e', padx=10, pady=10)

    def _build_heading_form(self, parent, sid, data):
        frame = ttk.LabelFrame(parent, text=sid)
        frame.pack(fill='x', padx=5, pady=3)
        widgets = {}
        rows = [
            ('font', 'Font:', 25),
            ('size_pt', 'Cỡ (pt):', 8),
            ('color', 'Màu (hex):', 10),
            ('space_before_pt', 'Space before (pt):', 8),
            ('space_after_pt', 'Space after (pt):', 8),
        ]
        for i, (key, label, width) in enumerate(rows):
            row, col = divmod(i, 3)
            ttk.Label(frame, text=label).grid(row=row, column=col*2, sticky='w', padx=5, pady=2)
            v = tk.StringVar(value=str(data.get(key, '')))
            ttk.Entry(frame, textvariable=v, width=width).grid(row=row, column=col*2+1, sticky='w', padx=5)
            widgets[key] = v

        v_b = tk.BooleanVar(value=data.get('bold', False))
        ttk.Checkbutton(frame, text='Bold', variable=v_b).grid(row=2, column=0, sticky='w', padx=5)
        widgets['bold'] = v_b
        v_i = tk.BooleanVar(value=data.get('italic', False))
        ttk.Checkbutton(frame, text='Italic', variable=v_i).grid(row=2, column=1, sticky='w', padx=5)
        widgets['italic'] = v_i

        self.heading_widgets[sid] = widgets

    def _save_step4_and_go(self):
        for sid, widgets in self.heading_widgets.items():
            cfg = self.model.config['heading_styles'][sid]
            try:
                for key, v in widgets.items():
                    if key in ('bold', 'italic'):
                        cfg[key] = v.get()
                    elif key in ('size_pt', 'space_before_pt', 'space_after_pt'):
                        s = v.get().strip()
                        cfg[key] = float(s) if s else 0.0
                    else:
                        cfg[key] = v.get().strip()
            except ValueError:
                messagebox.showerror('Lỗi', f'Số không hợp lệ ở {sid}')
                return
        for fname, v in self.hf_vars.items():
            self.model.config['header_footer_text'][fname] = v.get()
        self.notebook.select(4)

    # -------------------------------------------------------------------
    # STEP 5: Danh sách heading
    # -------------------------------------------------------------------

    def _build_step5_placeholder(self):
        ttk.Label(self.tab_step5, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step5(self):
        f = self.tab_step5
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 5: Danh sách heading của phần nội dung',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)
        ttk.Label(f, text='Số thứ tự (1, 1.1, 1.2.1...) sẽ được tự động cập nhật theo thứ tự và level.',
                  foreground='gray').pack(anchor='w', padx=10)

        # Đọc heading list từ file gốc nếu chưa có
        if not self.model.config['headings_list']:
            all_h = self.model.get_all_headings()
            intro_end = self.model.get_effective_intro_end_idx()
            # Lọc bỏ heading thuộc phần intro (trước mốc user chọn)
            self.model.config['headings_list'] = [
                h for h in all_h if h['original_idx'] >= intro_end
            ]

        # Treeview
        cols = ('lvl', 'num', 'text')
        self.h_tree = ttk.Treeview(f, columns=cols, show='headings', height=20)
        self.h_tree.heading('lvl', text='Lvl')
        self.h_tree.heading('num', text='Số')
        self.h_tree.heading('text', text='Nội dung heading')
        self.h_tree.column('lvl', width=50, anchor='center')
        self.h_tree.column('num', width=80)
        self.h_tree.column('text', width=800)
        self.h_tree.pack(fill='both', expand=True, padx=10, pady=5)

        self._refresh_heading_tree()

        # Edit panel
        edit_frame = ttk.LabelFrame(f, text='Sửa heading được chọn')
        edit_frame.pack(fill='x', padx=10, pady=5)
        ttk.Label(edit_frame, text='Level:').grid(row=0, column=0, sticky='w', padx=5)
        self.h_lvl_var = tk.StringVar()
        ttk.Combobox(edit_frame, textvariable=self.h_lvl_var,
                     values=['1', '2', '3', '4', '5', '6'], width=5).grid(row=0, column=1, padx=5)
        ttk.Label(edit_frame, text='Text:').grid(row=0, column=2, sticky='w', padx=5)
        self.h_text_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.h_text_var, width=80).grid(row=0, column=3, padx=5)
        ttk.Button(edit_frame, text='Cập nhật heading',
                   command=self._update_heading).grid(row=0, column=4, padx=5)

        btn_frame = ttk.Frame(f)
        btn_frame.pack(fill='x', padx=10, pady=5)
        ttk.Button(btn_frame, text='+ Thêm heading mới',
                   command=self._add_heading).pack(side='left', padx=2)
        ttk.Button(btn_frame, text='- Xoá heading',
                   command=self._delete_heading).pack(side='left', padx=2)
        ttk.Button(btn_frame, text='↑ Lên',
                   command=lambda: self._move_heading(-1)).pack(side='left', padx=2)
        ttk.Button(btn_frame, text='↓ Xuống',
                   command=lambda: self._move_heading(1)).pack(side='left', padx=2)

        v_auto = tk.BooleanVar(value=False)
        self.auto_number_var = v_auto
        ttk.Checkbutton(btn_frame, text='Tự động chèn số thứ tự (1., 1.1., …) vào text khi save',
                        variable=v_auto).pack(side='left', padx=20)

        # Khi click row
        def on_select(_):
            sel = self.h_tree.selection()
            if sel:
                idx = int(sel[0])
                h = self.model.config['headings_list'][idx]
                self.h_lvl_var.set(str(h['level']))
                self.h_text_var.set(h['text'])
        self.h_tree.bind('<<TreeviewSelect>>', on_select)

        ttk.Button(f, text='Lưu & sang bước 6 →',
                   command=lambda: (self._save_step5(), self.notebook.select(5))).pack(anchor='e', padx=10, pady=10)

    def _refresh_heading_tree(self):
        for row in self.h_tree.get_children():
            self.h_tree.delete(row)
        counters = [0] * 10
        for i, h in enumerate(self.model.config['headings_list']):
            lvl = h['level']
            counters[lvl] += 1
            for j in range(lvl + 1, 10):
                counters[j] = 0
            for k in range(1, lvl):
                if counters[k] == 0:
                    counters[k] = 1
            num_str = '.'.join(str(counters[k]) for k in range(1, lvl + 1))
            indent = '   ' * (lvl - 1)
            self.h_tree.insert('', 'end', iid=str(i),
                               values=(lvl, num_str, indent + h['text']))

    def _update_heading(self):
        sel = self.h_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        try:
            new_lvl = int(self.h_lvl_var.get())
        except ValueError:
            messagebox.showerror('Lỗi', 'Level phải là số')
            return
        self.model.config['headings_list'][idx]['level'] = new_lvl
        self.model.config['headings_list'][idx]['text'] = self.h_text_var.get()
        self._refresh_heading_tree()

    def _add_heading(self):
        sel = self.h_tree.selection()
        new_h = {
            'level': 1,
            'text': 'Heading mới',
            'original_idx': None,  # heading mới, chưa có vị trí gốc
            'keep_content': False,  # mặc định để trống
            'auto_number': False,
        }
        if sel:
            idx = int(sel[0]) + 1
            self.model.config['headings_list'].insert(idx, new_h)
        else:
            self.model.config['headings_list'].append(new_h)
        self._refresh_heading_tree()

    def _delete_heading(self):
        sel = self.h_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        del self.model.config['headings_list'][idx]
        self._refresh_heading_tree()

    def _move_heading(self, direction):
        sel = self.h_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        new_idx = idx + direction
        lst = self.model.config['headings_list']
        if 0 <= new_idx < len(lst):
            lst[idx], lst[new_idx] = lst[new_idx], lst[idx]
            self._refresh_heading_tree()
            self.h_tree.selection_set(str(new_idx))

    def _save_step5(self):
        # Cập nhật flag auto_number cho mọi entry
        auto = self.auto_number_var.get()
        for h in self.model.config['headings_list']:
            h['auto_number'] = auto

    # -------------------------------------------------------------------
    # STEP 6: Giữ/xoá nội dung
    # -------------------------------------------------------------------

    def _build_step6_placeholder(self):
        ttk.Label(self.tab_step6, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step6(self):
        f = self.tab_step6
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 6: Hành động cho nội dung dưới mỗi heading',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)
        ttk.Label(f, text=
                  '• Giữ nội dung gốc: copy nguyên content từ file gốc.\n'
                  '• Để trống: xoá content cũ, chừa 1 dòng trống (style Normal) để paste vào.\n'
                  '• Chèn từ file: xoá content cũ, chèn nội dung từ file Word (.docx) hoặc Excel (.xlsx). '
                  'Format sẽ ăn theo template (font, size, line spacing).',
                  foreground='gray', justify='left').pack(anchor='w', padx=10, pady=2)

        # Top buttons
        btn_frame = ttk.Frame(f)
        btn_frame.pack(fill='x', padx=10, pady=5)
        ttk.Button(btn_frame, text='Tất cả: Giữ nội dung',
                   command=lambda: self._set_all_action('keep')).pack(side='left', padx=2)
        ttk.Button(btn_frame, text='Tất cả: Để trống',
                   command=lambda: self._set_all_action('empty')).pack(side='left', padx=2)

        # Global setting
        self.preserve_inline_var = tk.BooleanVar(
            value=self.model.config.get('preserve_inline_formatting', True)
        )
        ttk.Checkbutton(
            btn_frame,
            text='Khi chèn từ file: giữ định dạng bold/italic/underline của TỪNG TỪ',
            variable=self.preserve_inline_var
        ).pack(side='left', padx=20)

        # Excel font override (row 2)
        font_frame = ttk.Frame(f)
        font_frame.pack(fill='x', padx=10, pady=2)
        ttk.Label(font_frame, text='Font khi chèn từ Excel:').pack(side='left')
        self.excel_font_name_var = tk.StringVar(
            value=self.model.config.get('excel_font_name', '') or ''
        )
        ttk.Entry(font_frame, textvariable=self.excel_font_name_var, width=20).pack(side='left', padx=3)
        ttk.Label(font_frame, text='Cỡ (pt):').pack(side='left', padx=(10, 3))
        self.excel_font_size_var = tk.StringVar(
            value=str(self.model.config.get('excel_font_size', 11) or 11)
        )
        ttk.Entry(font_frame, textvariable=self.excel_font_size_var, width=6).pack(side='left')
        ttk.Label(font_frame,
                  text='(mặc định 11; gõ giá trị khác để override)',
                  foreground='gray').pack(side='left', padx=8)

        # Scrollable list
        container = ttk.Frame(f)
        container.pack(fill='both', expand=True, padx=10, pady=5)
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        scrollable = ttk.Frame(canvas)
        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Mouse wheel scroll
        def _on_wheel(event):
            canvas.yview_scroll(int(-event.delta / 120), 'units')
        canvas.bind_all('<MouseWheel>', _on_wheel)

        self.heading_action_widgets = []  # list of widget references per heading
        for i, h in enumerate(self.model.config['headings_list']):
            row = ttk.Frame(scrollable)
            row.pack(fill='x', padx=2, pady=2)

            # Determine current action
            if h.get('keep_content', True):
                action = 'keep'
            elif h.get('insert_source'):
                action = 'insert'
            else:
                action = 'empty'

            action_var = tk.StringVar(value=action)
            action_combo = ttk.Combobox(
                row, textvariable=action_var, width=22, state='readonly',
                values=['keep — giữ nội dung gốc',
                        'empty — xoá để trống',
                        'insert — chèn từ file...']
            )
            display_map = {
                'keep': 'keep — giữ nội dung gốc',
                'empty': 'empty — xoá để trống',
                'insert': 'insert — chèn từ file...',
            }
            action_combo.set(display_map[action])
            action_combo.pack(side='left', padx=2)

            # Display label cho selection (ví dụ: "file.xlsx [Sheet1!A1:D20]")
            source_var = tk.StringVar(value=self._format_source_display(h))
            source_entry = ttk.Entry(row, textvariable=source_var, width=42, state='readonly')
            source_entry.pack(side='left', padx=2)

            browse_btn = ttk.Button(
                row, text='📁',
                command=lambda idx=i: self._open_selection_dialog(idx),
                width=3
            )
            browse_btn.pack(side='left', padx=2)

            indent = '    ' * (h['level'] - 1)
            ttk.Label(row, text=f'{indent}[H{h["level"]}] {h["text"]}',
                      foreground='#333').pack(side='left', padx=10)

            # Show/hide source widgets based on action
            def update_visibility(*_, ac=action_combo, se=source_entry, bb=browse_btn):
                txt = ac.get()
                if txt.startswith('insert'):
                    se.config(state='readonly')
                    bb.config(state='normal')
                else:
                    se.config(state='disabled')
                    bb.config(state='disabled')
            action_combo.bind('<<ComboboxSelected>>', update_visibility)
            update_visibility()

            self.heading_action_widgets.append({
                'index': i,
                'combo': action_combo,
                'source_var': source_var,
                'source_entry': source_entry,
            })

        ttk.Button(f, text='Lưu & sang bước 7 →',
                   command=self._save_step6_and_go).pack(anchor='e', padx=10, pady=10)

    def _format_source_display(self, h):
        """Tạo chuỗi hiển thị gọn cho source + selection của 1 heading."""
        path = h.get('insert_source', '') or ''
        sel = h.get('insert_selection') or {}
        if not path:
            return ''
        base = os.path.basename(path)
        mode = sel.get('mode', 'all')
        if mode == 'all' or not mode:
            return base
        if mode == 'sheet':
            return f'{base} [Sheet: {sel.get("sheet", "")}]'
        if mode == 'range':
            return f'{base} [{sel.get("sheet", "")}!{sel.get("range", "")}]'
        if mode == 'custom':
            sheet = sel.get('sheet', '')
            cols = sel.get('columns', '') or 'all'
            rs = sel.get('row_start', 1) or 1
            re_ = sel.get('row_end')
            re_disp = re_ if re_ else 'end'
            return f'{base} [{sheet}: cols={cols}, rows={rs}..{re_disp}]'
        return base

    def _open_selection_dialog(self, heading_idx):
        """Mở dialog chọn nội dung cho heading thứ heading_idx."""
        h = self.model.config['headings_list'][heading_idx]
        current_path = h.get('insert_source', '') or ''
        current_selection = h.get('insert_selection') or {}

        # Hiện chỉ Excel. Nếu file đang là docx, vẫn open Excel dialog
        # (user sẽ Browse để pick lại file Excel).
        result = ExcelSelectionDialog.show(self.root, current_path, current_selection)
        if result:
            h['insert_source'] = result['path']
            h['insert_selection'] = result['selection']
            # Update display
            for w in self.heading_action_widgets:
                if w['index'] == heading_idx:
                    w['source_var'].set(self._format_source_display(h))
                    break

    def _set_all_action(self, action):
        display_map = {
            'keep': 'keep — giữ nội dung gốc',
            'empty': 'empty — xoá để trống',
            'insert': 'insert — chèn từ file...',
        }
        for w in self.heading_action_widgets:
            w['combo'].set(display_map[action])
            w['combo'].event_generate('<<ComboboxSelected>>')

    def _pick_external_source(self, var):
        path = filedialog.askopenfilename(
            title='Chọn file Word hoặc Excel để chèn',
            filetypes=[
                ('Word/Excel', '*.docx *.xlsx *.xlsm'),
                ('Word', '*.docx'),
                ('Excel', '*.xlsx *.xlsm'),
                ('All', '*.*'),
            ]
        )
        if path:
            var.set(path)

    def _save_step6_and_go(self):
        for w in self.heading_action_widgets:
            i = w['index']
            txt = w['combo'].get()
            h = self.model.config['headings_list'][i]
            if txt.startswith('keep'):
                h['keep_content'] = True
            elif txt.startswith('empty'):
                h['keep_content'] = False
                h['insert_source'] = None
                h['insert_selection'] = None
            elif txt.startswith('insert'):
                src = h.get('insert_source')
                if not src:
                    messagebox.showwarning(
                        'Thiếu file',
                        f'Heading "{h["text"]}" chưa chọn file. '
                        f'Tool sẽ để trống thay vì chèn.'
                    )
                    h['keep_content'] = False
                    h['insert_source'] = None
                    h['insert_selection'] = None
                else:
                    h['keep_content'] = False

        self.model.config['preserve_inline_formatting'] = self.preserve_inline_var.get()

        # Excel font overrides
        self.model.config['excel_font_name'] = self.excel_font_name_var.get().strip()
        try:
            self.model.config['excel_font_size'] = float(
                self.excel_font_size_var.get().strip() or 11
            )
        except ValueError:
            self.model.config['excel_font_size'] = 11

        self._build_step7()
        self.notebook.select(6)

    def _pick_external_source(self, var):
        """Legacy method (chưa dùng) — giữ để tránh lỗi nếu có chỗ gọi."""
        path = filedialog.askopenfilename(
            title='Chọn file Word hoặc Excel',
            filetypes=[('Word/Excel', '*.docx *.xlsx *.xlsm'), ('All', '*.*')]
        )
        if path:
            var.set(path)

    # -------------------------------------------------------------------
    # STEP 7: Lưu template
    # -------------------------------------------------------------------

    def _build_step7_placeholder(self):
        ttk.Label(self.tab_step7, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step7(self):
        f = self.tab_step7
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 7: Tổng kết & lưu template',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)

        # Summary
        summary = tk.Text(f, height=18, wrap='word')
        summary.pack(fill='both', expand=True, padx=10, pady=5)
        s = self._build_summary()
        summary.insert('end', s)
        summary.config(state='disabled')

        # Output file picker
        out_frame = ttk.Frame(f)
        out_frame.pack(fill='x', padx=10, pady=5)
        ttk.Label(out_frame, text='Lưu thành:').pack(side='left')
        self.out_var = tk.StringVar()
        if self.model.path:
            base = os.path.basename(self.model.path)
            stem = os.path.splitext(base)[0]
            default_out = os.path.join(os.path.dirname(self.model.path), f'{stem}_template.docx')
            self.out_var.set(default_out)
        ttk.Entry(out_frame, textvariable=self.out_var, width=80).pack(side='left', padx=5)
        ttk.Button(out_frame, text='Chọn...', command=self._pick_output).pack(side='left')

        # Generate button
        self.gen_button = ttk.Button(f, text='🚀 SINH FILE TEMPLATE', command=self._generate)
        self.gen_button.pack(anchor='e', padx=10, pady=10)

        # Progress bar
        progress_frame = ttk.Frame(f)
        progress_frame.pack(fill='x', padx=10, pady=5)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame, mode='determinate', variable=self.progress_var,
            maximum=100, length=400
        )
        self.progress_bar.pack(side='left', fill='x', expand=True)
        self.progress_pct_label = ttk.Label(progress_frame, text='0%', width=6)
        self.progress_pct_label.pack(side='left', padx=5)

        self.gen_status = ttk.Label(f, text='', foreground='blue', wraplength=900, justify='left')
        self.gen_status.pack(anchor='w', padx=10, pady=2)

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title='Lưu template',
            defaultextension='.docx',
            filetypes=[('Word', '*.docx')]
        )
        if path:
            self.out_var.set(path)

    def _generate(self):
        out = self.out_var.get().strip()
        if not out:
            messagebox.showerror('Lỗi', 'Vui lòng chọn đường dẫn lưu')
            return

        # Disable button + reset progress
        self.gen_button.config(state='disabled')
        self.progress_var.set(0)
        self.progress_pct_label.config(text='0%')
        self.gen_status.config(text='Đang chuẩn bị...', foreground='blue')

        import threading

        def run_generation():
            try:
                def progress_cb(percent, message):
                    # Thread-safe UI update
                    self.root.after(0, lambda p=percent, m=message:
                                    self._update_progress(p, m))
                self.model.generate(out, progress_callback=progress_cb)
                self.root.after(0, lambda: self._gen_finished(out, success=True))
            except Exception as e:
                import traceback
                traceback.print_exc()
                err_msg = str(e)
                self.root.after(0, lambda: self._gen_finished(out, success=False, error=err_msg))

        thread = threading.Thread(target=run_generation, daemon=True)
        thread.start()

    def _update_progress(self, percent, message):
        self.progress_var.set(percent)
        self.progress_pct_label.config(text=f'{percent:.0f}%')
        self.gen_status.config(text=message, foreground='blue')

    def _gen_finished(self, out, success, error=None):
        self.gen_button.config(state='normal')
        if success:
            self.progress_var.set(100)
            self.progress_pct_label.config(text='100%')
            self.gen_status.config(text=f'✓ Đã tạo file template: {out}', foreground='green')
            messagebox.showinfo('Thành công', f'Template đã được tạo:\n{out}')
        else:
            self.gen_status.config(text=f'✗ Lỗi: {error}', foreground='red')
            messagebox.showerror('Lỗi sinh file', error)

    def _build_summary(self):
        c = self.model.config
        lines = []
        lines.append(f'File gốc: {self.model.path}')
        lines.append(f'Tổng số element trong body: {len(self.model.body_elements)}')
        intro_end = self.model.get_effective_intro_end_idx()
        lines.append(f'Mốc bắt đầu phần nội dung: element index {intro_end}')
        lines.append('')
        lines.append('=== Document defaults ===')
        for k, v in c['doc_defaults'].items():
            lines.append(f'  {k}: {v}')
        lines.append('')
        lines.append(f'=== Sections ({len(c["sections"])}) ===')
        for i, sect in enumerate(c['sections']):
            lines.append(f'  Section {i+1}: {sect["pg_w_cm"]}x{sect["pg_h_cm"]}cm '
                         f'{sect["orient"]}, lề T={sect["top_cm"]} R={sect["right_cm"]} '
                         f'B={sect["bottom_cm"]} L={sect["left_cm"]}')
        lines.append('')
        lines.append('=== Heading styles ===')
        for sid, data in c['heading_styles'].items():
            lines.append(f'  {sid}: font={data.get("font")} {data.get("size_pt")}pt '
                         f'bold={data.get("bold")} italic={data.get("italic")}')
        lines.append('')
        lines.append(f'=== Text replacements: {len(c["intro_replacements"])} ===')
        for old, new in list(c['intro_replacements'].items())[:5]:
            lines.append(f'  "{old[:40]}" → "{new[:40]}"')
        lines.append('')
        lines.append(f'=== Image replacements: {len(c["intro_images"])} ===')
        for rid, p in c['intro_images'].items():
            lines.append(f'  {rid}: {p}')
        lines.append('')
        lines.append(f'=== Headings list: {len(c["headings_list"])} ===')
        keep = sum(1 for h in c['headings_list'] if h.get('keep_content'))
        empty = sum(1 for h in c['headings_list']
                    if not h.get('keep_content') and not h.get('insert_source'))
        insert = sum(1 for h in c['headings_list']
                     if not h.get('keep_content') and h.get('insert_source'))
        lines.append(f'  Giữ nội dung gốc: {keep}')
        lines.append(f'  Để trống (1 dòng cho paste): {empty}')
        lines.append(f'  Chèn từ file ngoài: {insert}')
        if insert > 0:
            lines.append(f'  preserve_inline_formatting: {c.get("preserve_inline_formatting", True)}')
            ef_name = c.get('excel_font_name', '')
            ef_size = c.get('excel_font_size', 0)
            if ef_name or ef_size:
                lines.append(f'  excel_font_override: '
                             f'name="{ef_name or "auto"}" size={ef_size or "auto"}')
            for h in c['headings_list']:
                if h.get('insert_source'):
                    sel = h.get('insert_selection') or {}
                    mode = sel.get('mode', 'all')
                    if mode == 'all':
                        sel_str = '(toàn bộ)'
                    elif mode == 'sheet':
                        sel_str = f'(Sheet: {sel.get("sheet", "")})'
                    elif mode == 'range':
                        sel_str = f'({sel.get("sheet", "")}!{sel.get("range", "")})'
                    elif mode == 'custom':
                        cols = sel.get('columns', '') or 'all'
                        rs = sel.get('row_start', 1) or 1
                        re_ = sel.get('row_end') or 'end'
                        sel_str = f'({sel.get("sheet", "")}: cols={cols}, rows={rs}..{re_})'
                    else:
                        sel_str = ''
                    lines.append(f'    [H{h["level"]}] "{h["text"][:30]}" ← '
                                 f'{os.path.basename(h["insert_source"])} {sel_str}')
        return '\n'.join(lines)


# ============================================================================
# Main
# ============================================================================

def main():
    if tk is None:
        raise RuntimeError(
            'Không tìm thấy Tkinter. Dùng web_app.py cho bản web/server, '
            'hoặc cài Python có Tk để chạy GUI desktop.'
        )
    app = WizardApp()
    app.run()


if __name__ == '__main__':
    main()
