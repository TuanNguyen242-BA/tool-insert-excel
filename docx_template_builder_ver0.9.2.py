# -*- coding: utf-8 -*-
"""
DOCX Template Builder
=====================
Tool đọc thiết lập từ một file Word có sẵn rồi tạo ra một file template
theo cấu hình do người dùng tuỳ chỉnh.

Quy trình 8 bước (xem các tab trong giao diện):
  1) Chọn file + nhập trang bắt đầu/kết thúc cho phần mở đầu và phần nội dung
  2) Phân tích & sửa Document defaults và Section properties
  3) Phân tích & sửa nội dung phần mở đầu (text holders, ảnh)
  4) Phân tích & sửa Heading styles, Numbering, Header/Footer của phần nội dung
  5) Liệt kê & chỉnh sửa danh sách heading (auto-renumber)
  6) Chọn heading nào giữ nội dung gốc, heading nào để trống
  7) Cấu hình style mục lục (TOC)
  8) Sinh file template

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
import unicodedata
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
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

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
        return round(float(t) / 567, 2)
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
        return float(hp) / 2
    except (TypeError, ValueError):
        return 0

def pt_to_half_pt(p):
    return int(round(float(p) * 2))

def twips_to_pt(t):
    """Convert OOXML twips to points, accepting integer or decimal strings."""
    try:
        return float(t) / 20
    except (TypeError, ValueError):
        return 0.0

def line_to_spacing_str(line, lineRule):
    """Convert (line, lineRule) → human-readable spacing"""
    if not line:
        return '1.0'
    try:
        line = float(line)
    except (TypeError, ValueError):
        return str(line)
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


def normalize_match_text(text):
    text = str(text or '').strip().upper()
    text = ''.join(
        ch for ch in unicodedata.normalize('NFKD', text)
        if not unicodedata.combining(ch)
    )
    text = text.replace('Đ', 'D')
    return re.sub(r'\s+', ' ', text)


def default_toc_settings(doc_defaults=None):
    """Default gọn gàng cho Word TOC styles (TOC1..TOC9)."""
    doc_defaults = doc_defaults or {}
    font = doc_defaults.get('font') or 'Times New Roman'
    base_size = float(doc_defaults.get('size_pt') or 13.0)
    levels = OrderedDict()
    for level in range(1, 10):
        left = 0.0 if level == 1 else round(0.55 + (level - 2) * 0.5, 2)
        text_tab = round(left + 0.55 + max(level - 1, 0) * 0.12, 2)
        levels[f'TOC{level}'] = {
            'font': font,
            'size_pt': base_size,
            'bold': level <= 2,
            'italic': False,
            'color': '000000',
            'left_indent_cm': left,
            'text_tab_cm': text_tab,
            'space_before_pt': 0.0,
            'space_after_pt': 2.0 if level <= 3 else 1.0,
            'line_spacing': '1.15',
        }
    return {
        'enabled': True,
        'update_on_open': True,
        'levels': 4,
        'tab_leader': 'none',
        'right_tab_cm': 0.0,  # 0 = auto theo usable page width
        'styles': levels,
    }


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
            'toc_settings': default_toc_settings(),
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
            normalized = normalize_match_text(text)
            if normalized == 'MUC LUC':
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

        # TOC styles (TOC1-TOC9), mặc định bám theo document defaults.
        self.config['toc_settings'] = default_toc_settings(self.config.get('doc_defaults') or {})
        self._ensure_toc_settings_defaults()

    def _ensure_toc_settings_defaults(self):
        """Đảm bảo config TOC luôn đủ key khi load file cũ hoặc đổi template."""
        defaults = default_toc_settings(self.config.get('doc_defaults') or {})
        current = self.config.get('toc_settings') or {}
        merged = copy.deepcopy(defaults)
        for key, value in current.items():
            if key != 'styles':
                merged[key] = value
        current_styles = current.get('styles') or {}
        for sid, style_defaults in defaults['styles'].items():
            style = merged['styles'].setdefault(sid, copy.deepcopy(style_defaults))
            style.update(current_styles.get(sid, {}))
            for k, v in style_defaults.items():
                style.setdefault(k, v)
        self.config['toc_settings'] = merged
        return merged

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
                    result['space_before'] = twips_to_pt(before)
                after = sp.get(w('after'))
                if after:
                    result['space_after'] = twips_to_pt(after)
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
                                data['space_before_pt'] = twips_to_pt(sp.get(w('before')))
                            if sp.get(w('after')):
                                data['space_after_pt'] = twips_to_pt(sp.get(w('after')))
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
        numbering_path = os.path.join(out_dir, 'word/numbering.xml')
        numbering_tree = etree.parse(numbering_path, parser) if os.path.exists(numbering_path) else None

        progress(20, 'Áp dụng document defaults...')
        self._apply_doc_defaults(styles_tree)

        progress(30, 'Áp dụng sections...')
        self._apply_sections(doc_tree)

        progress(40, 'Áp dụng heading styles...')
        self._apply_heading_styles(styles_tree)

        progress(45, 'Áp dụng style mục lục...')
        self._apply_toc_styles(styles_tree, doc_tree)
        if numbering_tree is not None:
            self._ensure_numbering_suffix_spaces(numbering_tree, styles_tree)

        progress(50, 'Áp dụng text replacements...')
        self._apply_text_replacements(doc_tree)

        progress(60, 'Thay thế ảnh...')
        self._apply_image_replacements(out_dir)

        progress(70, 'Xử lý heading & content (xoá/chèn)...')
        self._apply_headings_and_content(doc_tree, styles_tree, progress_callback)

        progress(85, 'Cập nhật header/footer text...')
        self._apply_header_footer_text(out_dir)
        toc_cfg = self._ensure_toc_settings_defaults()
        should_update_toc = toc_cfg.get('enabled', True) and (
            toc_cfg.get('update_on_open', True) or
            str(os.getenv('DOCX_BUILDER_MARK_TOC_DIRTY', '')).strip().lower() in {'1', 'true', 'yes'}
        )
        if should_update_toc:
            progress(88, 'Chuẩn bị cập nhật mục lục khi mở file...')
            self._prepare_toc_fields_for_update(doc_tree)
            self._enable_field_update_on_open(out_dir)

        progress(90, 'Lưu XML trees...')
        doc_tree.write(os.path.join(out_dir, 'word/document.xml'),
                       xml_declaration=True, encoding='UTF-8', standalone=True)
        styles_tree.write(os.path.join(out_dir, 'word/styles.xml'),
                          xml_declaration=True, encoding='UTF-8', standalone=True)
        if numbering_tree is not None:
            numbering_tree.write(numbering_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        progress(95, 'Đóng gói file docx...')
        self._zip_dir(out_dir, output_path)
        shutil.rmtree(out_dir, ignore_errors=True)

        progress(100, 'Hoàn tất')

    def _enable_field_update_on_open(self, out_dir):
        settings_path = os.path.join(out_dir, 'word/settings.xml')
        parser = etree.XMLParser(remove_blank_text=False)
        if os.path.exists(settings_path):
            tree = etree.parse(settings_path, parser)
            root = tree.getroot()
        else:
            root = etree.Element(w('settings'), nsmap={'w': W_NS})
            tree = etree.ElementTree(root)
        update = root.find(w('updateFields'))
        if update is None:
            update = etree.SubElement(root, w('updateFields'))
        update.set(w('val'), 'true')
        tree.write(settings_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _mark_toc_fields_dirty(self, doc_tree):
        root = doc_tree.getroot()
        for fld in root.findall(f'.//{w("fldSimple")}'):
            instr = fld.get(w('instr')) or ''
            if 'TOC' in instr.upper():
                fld.set(w('dirty'), 'true')
                fld.attrib.pop(w('fldLock'), None)

        for p in root.findall(f'.//{w("p")}'):
            instr_text = ''.join(t.text or '' for t in p.findall(f'.//{w("instrText")}'))
            if 'TOC' not in instr_text.upper():
                continue
            for fld_char in p.findall(f'.//{w("fldChar")}'):
                fld_char.set(w('dirty'), 'true')
                fld_char.attrib.pop(w('fldLock'), None)

    def _prepare_toc_fields_for_update(self, doc_tree):
        """Chuẩn hóa field TOC để Word render lại số trang và format theo TOC styles."""
        toc_bookmark = self._ensure_content_toc_bookmark(doc_tree)
        self._ensure_content_starts_on_new_page_after_toc(doc_tree)
        self._normalize_toc_field_instructions(doc_tree, toc_bookmark)
        self._normalize_toc_result_number_spacing(doc_tree)
        self._clean_existing_toc_result_formatting(doc_tree)
        self._mark_toc_fields_dirty(doc_tree)

    def _normalize_toc_result_number_spacing(self, doc_tree):
        """Ensure cached TOC entries keep a visible space after manual outline numbers."""
        root = doc_tree.getroot()
        number_re = re.compile(r'^(\s*\d+(?:\.\d+)+)(?=\S)')
        for p in root.findall(f'.//{w("p")}'):
            ppr = p.find(w('pPr'))
            if ppr is None:
                continue
            pstyle = ppr.find(w('pStyle'))
            style_id = pstyle.get(w('val')) if pstyle is not None else ''
            if not re.match(r'TOC\d+$', style_id or ''):
                continue

            text_nodes = []
            for run in p.findall(f'.//{w("r")}'):
                if run.find(w('instrText')) is not None or run.find(w('fldChar')) is not None:
                    continue
                text_nodes.extend(run.findall(w('t')))
            text_nodes = [t for t in text_nodes if t.text]
            if not text_nodes:
                continue

            first = text_nodes[0]
            new_text = number_re.sub(r'\1 ', first.text or '', count=1)
            if new_text != first.text:
                first.text = new_text
                first.set(XML_SPACE, 'preserve')
                continue

            if re.fullmatch(r'\s*\d+(?:\.\d+)+', first.text or ''):
                for next_t in text_nodes[1:]:
                    if next_t.text and not next_t.text[0].isspace():
                        next_t.text = ' ' + next_t.text
                        next_t.set(XML_SPACE, 'preserve')
                        break

    def _ensure_content_toc_bookmark(self, doc_tree):
        """Tạo bookmark chỉ bao quanh phần nội dung để TOC không lấy heading intro."""
        root = doc_tree.getroot()
        body = root.find(w('body'))
        if body is None:
            return ''
        body_children = [el for el in body if el.tag in (w('p'), w('tbl'))]
        if not body_children:
            return ''

        first_idx = self._first_content_body_index(body_children)
        if first_idx is None or first_idx < 0 or first_idx >= len(body_children):
            return ''

        name = 'ContentTocRange'
        self._remove_bookmark(root, name)
        bookmark_id = self._next_bookmark_id(root)

        start_el = body_children[first_idx]
        if start_el.tag != w('p'):
            p = etree.Element(w('p'))
            start_el.addprevious(p)
            start_el = p
        self._insert_bookmark_start(start_el, bookmark_id, name)

        end_el = None
        for el in reversed(body_children):
            if el.tag == w('p'):
                end_el = el
                break
        if end_el is None:
            end_el = etree.Element(w('p'))
            sect_pr = body.find(w('sectPr'))
            if sect_pr is not None:
                sect_pr.addprevious(end_el)
            else:
                body.append(end_el)
        end_el.append(self._bookmark_end(bookmark_id))
        return name

    def _first_content_body_index(self, body_children=None):
        first_idx = self.config.get('intro_end_idx')
        for h in self.config.get('headings_list') or []:
            idx = h.get('original_idx')
            if idx is not None:
                first_idx = idx if first_idx is None else min(first_idx, idx)
                break
        if first_idx is None:
            first_idx = self.get_effective_intro_end_idx()
        try:
            first_idx = int(first_idx)
        except (TypeError, ValueError):
            return None
        if body_children is not None and not (0 <= first_idx < len(body_children)):
            return None
        return first_idx

    def _remove_bookmark(self, root, name):
        for start in list(root.findall(f'.//{w("bookmarkStart")}')):
            if start.get(w('name')) != name:
                continue
            bookmark_id = start.get(w('id'))
            parent = start.getparent()
            if parent is not None:
                parent.remove(start)
            for end in list(root.findall(f'.//{w("bookmarkEnd")}')):
                if end.get(w('id')) == bookmark_id:
                    end_parent = end.getparent()
                    if end_parent is not None:
                        end_parent.remove(end)

    def _next_bookmark_id(self, root):
        max_id = 0
        for el in root.findall(f'.//{w("bookmarkStart")}') + root.findall(f'.//{w("bookmarkEnd")}'):
            try:
                max_id = max(max_id, int(el.get(w('id')) or 0))
            except ValueError:
                pass
        return str(max_id + 1)

    def _insert_bookmark_start(self, paragraph, bookmark_id, name):
        bm = etree.Element(w('bookmarkStart'))
        bm.set(w('id'), str(bookmark_id))
        bm.set(w('name'), name)
        insert_at = 0
        if len(paragraph) and paragraph[0].tag == w('pPr'):
            insert_at = 1
        paragraph.insert(insert_at, bm)

    def _bookmark_end(self, bookmark_id):
        bm = etree.Element(w('bookmarkEnd'))
        bm.set(w('id'), str(bookmark_id))
        return bm

    def _allocate_bookmark_id(self, root):
        current = getattr(self, '_bookmark_id_counter', None)
        if current is None:
            current = int(self._next_bookmark_id(root))
        self._bookmark_id_counter = current + 1
        return str(current)

    def _safe_bookmark_name(self, prefix, text):
        raw = normalize_match_text(text)
        raw = re.sub(r'[^A-Z0-9_]+', '_', raw).strip('_')[:28] or 'ITEM'
        seed = abs(hash((prefix, raw))) % 1000000
        return f'{prefix}_{raw}_{seed}'

    def _ensure_paragraph_bookmark(self, doc_root, paragraph, name):
        if paragraph is None or paragraph.tag != w('p'):
            return ''
        for bm in paragraph.findall(w('bookmarkStart')):
            if bm.get(w('name')) == name:
                return name
        bookmark_id = self._allocate_bookmark_id(doc_root)
        self._insert_bookmark_start(paragraph, bookmark_id, name)
        paragraph.append(self._bookmark_end(bookmark_id))
        return name

    def _collect_child_heading_bookmarks(self, doc_tree, body_children, target_idx):
        if target_idx is None or target_idx < 0 or target_idx >= len(body_children):
            return {}
        root = doc_tree.getroot()
        target_el = body_children[target_idx]
        target_style = self._paragraph_style_id(target_el)
        match = re.match(r'Heading(\d+)$', target_style or '')
        base_level = int(match.group(1)) if match else 0
        result = {}
        for idx in range(target_idx + 1, len(body_children)):
            el = body_children[idx]
            if el.tag != w('p'):
                continue
            style_id = self._paragraph_style_id(el)
            m = re.match(r'Heading(\d+)$', style_id or '')
            if not m:
                continue
            level = int(m.group(1))
            if base_level and level <= base_level:
                break
            text = ''.join(t.text or '' for t in el.iter(w('t'))).strip()
            key = self._normalize_internal_link_key(text)
            if not key:
                continue
            bookmark = self._safe_bookmark_name('ILINK', text)
            result[key] = self._ensure_paragraph_bookmark(root, el, bookmark)
        return result

    def _ensure_content_starts_on_new_page_after_toc(self, doc_tree):
        """Đặt page break trước heading nội dung đầu tiên để TOC và nội dung không chung trang."""
        root = doc_tree.getroot()
        body = root.find(w('body'))
        if body is None:
            return
        body_children = [el for el in body if el.tag in (w('p'), w('tbl'))]
        first_idx = self._first_content_body_index(body_children)
        if first_idx is None:
            return
        if first_idx < 0 or first_idx >= len(body_children):
            return
        first_content = body_children[first_idx]
        if first_content.tag != w('p'):
            return
        ppr = first_content.find(w('pPr'))
        if ppr is None:
            ppr = etree.Element(w('pPr'))
            first_content.insert(0, ppr)
        ensure_child(ppr, 'pageBreakBefore', PPR_ORDER)

    def _normalize_toc_field_instructions(self, doc_tree, bookmark_name=''):
        cfg = self._ensure_toc_settings_defaults()
        levels = int(cfg.get('levels') or 4)
        levels = max(1, min(9, levels))
        bookmark_switch = f' \\b "{bookmark_name}"' if bookmark_name else ''
        instr = f'TOC \\o "1-{levels}" \\h \\z \\u{bookmark_switch}'

        root = doc_tree.getroot()
        for fld in root.findall(f'.//{w("fldSimple")}'):
            old_instr = fld.get(w('instr')) or ''
            if 'TOC' in old_instr.upper():
                fld.set(w('instr'), instr)

        for p in root.findall(f'.//{w("p")}'):
            instr_nodes = p.findall(f'.//{w("instrText")}')
            instr_text = ''.join(t.text or '' for t in instr_nodes)
            if 'TOC' not in instr_text.upper() or not instr_nodes:
                continue
            instr_nodes[0].text = ' ' + instr + ' '
            instr_nodes[0].set(XML_SPACE, 'preserve')
            for extra in instr_nodes[1:]:
                extra.text = ''

    def _clean_existing_toc_result_formatting(self, doc_tree):
        """
        Xóa direct format trong các paragraph TOC cũ để lần refresh kế tiếp dùng
        đúng TOC1..TOC9. Giữ pStyle và field code, chỉ bỏ override trình bày.
        """
        toc_styles = (self._ensure_toc_settings_defaults().get('styles') or {})
        root = doc_tree.getroot()
        for p in root.findall(f'.//{w("p")}'):
            ppr = p.find(w('pPr'))
            if ppr is None:
                continue
            pstyle = ppr.find(w('pStyle'))
            style_id = pstyle.get(w('val')) if pstyle is not None else ''
            if not re.match(r'TOC\d+$', style_id or ''):
                continue

            for child in list(ppr):
                if child.tag != w('pStyle'):
                    ppr.remove(child)

            for run in p.findall(w('r')):
                if run.find(w('instrText')) is not None or run.find(w('fldChar')) is not None:
                    continue
                rpr = run.find(w('rPr'))
                if rpr is not None:
                    run.remove(rpr)
                data = toc_styles.get(style_id) or {}
                if data:
                    self._apply_toc_result_run_format(run, data)

    def _apply_toc_result_run_format(self, run, data):
        """Apply TOC settings to cached TOC result text so Word shows the chosen size immediately."""
        rpr = run.find(w('rPr'))
        if rpr is None:
            rpr = etree.Element(w('rPr'))
            run.insert(0, rpr)
        if data.get('font'):
            rfonts = ensure_child(rpr, 'rFonts', RPR_ORDER)
            for a in ['ascii', 'eastAsia', 'hAnsi', 'cs']:
                rfonts.set(w(a), data['font'])
        if data.get('size_pt'):
            size_val = str(pt_to_half_pt(float(data.get('size_pt'))))
            ensure_child(rpr, 'sz', RPR_ORDER).set(w('val'), size_val)
            ensure_child(rpr, 'szCs', RPR_ORDER).set(w('val'), size_val)
        self._set_flag(rpr, 'b', bool(data.get('bold')))
        self._set_flag(rpr, 'bCs', bool(data.get('bold')))
        self._set_flag(rpr, 'i', bool(data.get('italic')))
        self._set_flag(rpr, 'iCs', bool(data.get('italic')))
        if data.get('color'):
            ensure_child(rpr, 'color', RPR_ORDER).set(w('val'), data.get('color'))

    def _apply_toc_styles(self, styles_tree, doc_tree):
        cfg = self._ensure_toc_settings_defaults()
        if not cfg.get('enabled', True):
            return

        root = styles_tree.getroot()
        page_width_twips = self._compute_page_content_width(doc_tree)
        if cfg.get('right_tab_cm'):
            page_width_twips = cm_to_twips(cfg.get('right_tab_cm'))

        leader = (cfg.get('tab_leader') or 'none').strip()
        leader_map = {
            'none': None,
            'dot': 'dot',
            'hyphen': 'hyphen',
            'underscore': 'underscore',
        }
        leader_val = leader_map.get(leader, None)

        for sid, data in (cfg.get('styles') or {}).items():
            if not re.match(r'TOC\d+$', sid):
                continue
            style_el = self._ensure_paragraph_style(root, sid, sid.replace('TOC', 'toc '))

            ppr = style_el.find(w('pPr'))
            rpr = style_el.find(w('rPr'))
            if ppr is None:
                ppr = etree.Element(w('pPr'))
                if rpr is not None:
                    rpr.addprevious(ppr)
                else:
                    style_el.append(ppr)
            if rpr is None:
                rpr = etree.SubElement(style_el, w('rPr'))

            ind = ensure_child(ppr, 'ind', PPR_ORDER)
            ind.set(w('left'), str(cm_to_twips(data.get('left_indent_cm', 0))))
            ind.attrib.pop(w('firstLine'), None)
            ind.attrib.pop(w('hanging'), None)

            old_tabs = ppr.find(w('tabs'))
            if old_tabs is not None:
                ppr.remove(old_tabs)
            tabs = ensure_child(ppr, 'tabs', PPR_ORDER)
            text_tab_cm = float(data.get('text_tab_cm') or 0)
            if text_tab_cm > 0:
                tab_text = etree.SubElement(tabs, w('tab'))
                tab_text.set(w('val'), 'left')
                tab_text.set(w('pos'), str(cm_to_twips(text_tab_cm)))
            tab_page = etree.SubElement(tabs, w('tab'))
            tab_page.set(w('val'), 'right')
            tab_page.set(w('pos'), str(page_width_twips))
            if leader_val:
                tab_page.set(w('leader'), leader_val)

            sp = ensure_child(ppr, 'spacing', PPR_ORDER)
            sp.set(w('before'), str(int(float(data.get('space_before_pt', 0)) * 20)))
            sp.set(w('after'), str(int(float(data.get('space_after_pt', 0)) * 20)))
            line, rule = spacing_str_to_line(str(data.get('line_spacing') or '1.15'))
            sp.set(w('line'), line)
            sp.set(w('lineRule'), rule)

            if data.get('font'):
                rfonts = ensure_child(rpr, 'rFonts', RPR_ORDER)
                for a in ['ascii', 'eastAsia', 'hAnsi', 'cs']:
                    rfonts.set(w(a), data['font'])
            if data.get('size_pt'):
                size_val = str(pt_to_half_pt(float(data.get('size_pt'))))
                ensure_child(rpr, 'sz', RPR_ORDER).set(w('val'), size_val)
                ensure_child(rpr, 'szCs', RPR_ORDER).set(w('val'), size_val)
            self._set_flag(rpr, 'b', bool(data.get('bold')))
            self._set_flag(rpr, 'bCs', bool(data.get('bold')))
            self._set_flag(rpr, 'i', bool(data.get('italic')))
            self._set_flag(rpr, 'iCs', bool(data.get('italic')))
            if data.get('color'):
                ensure_child(rpr, 'color', RPR_ORDER).set(w('val'), data.get('color'))

    def _ensure_numbering_suffix_spaces(self, numbering_tree, styles_tree):
        """Make Word's own TOC refresh keep a separator after heading numbers."""
        if numbering_tree is None:
            return
        root = numbering_tree.getroot()
        heading_abstract_ids = set()
        num_to_abs = {}
        for num in root.findall(w('num')):
            num_id = num.get(w('numId'))
            abs_el = num.find(w('abstractNumId'))
            if num_id and abs_el is not None:
                num_to_abs[str(num_id)] = str(abs_el.get(w('val')) or '')

        for level in range(1, 10):
            style_id = f'Heading{level}'
            num_pr = self._style_num_pr(styles_tree, style_id)
            num_id, _ilvl = self._num_pr_values(num_pr)
            if num_id and str(num_id) in num_to_abs:
                heading_abstract_ids.add(num_to_abs[str(num_id)])

        for abs_num in root.findall(w('abstractNum')):
            abs_id = str(abs_num.get(w('abstractNumId')) or '')
            has_heading_style = any(
                (lvl.find(w('pStyle')) is not None and
                 re.match(r'Heading\d+$', lvl.find(w('pStyle')).get(w('val')) or ''))
                for lvl in abs_num.findall(w('lvl'))
            )
            if not has_heading_style and abs_id not in heading_abstract_ids:
                continue
            for lvl in abs_num.findall(w('lvl')):
                suff = lvl.find(w('suff'))
                if suff is None:
                    suff = etree.SubElement(lvl, w('suff'))
                suff.set(w('val'), 'space')

    def _paragraph_style_id(self, p):
        ppr = p.find(w('pPr')) if p is not None else None
        pstyle = ppr.find(w('pStyle')) if ppr is not None else None
        return pstyle.get(w('val')) if pstyle is not None else ''

    def _style_num_pr(self, styles_tree, style_id):
        if styles_tree is None or not style_id:
            return None
        root = styles_tree.getroot()
        for style in root.findall(w('style')):
            if style.get(w('styleId')) != style_id:
                continue
            ppr = style.find(w('pPr'))
            return ppr.find(w('numPr')) if ppr is not None else None
        return None

    def _num_pr_values(self, num_pr):
        if num_pr is None:
            return None, None
        num_id_el = num_pr.find(w('numId'))
        ilvl_el = num_pr.find(w('ilvl'))
        num_id = num_id_el.get(w('val')) if num_id_el is not None else None
        ilvl = ilvl_el.get(w('val')) if ilvl_el is not None else None
        try:
            ilvl = int(ilvl) if ilvl is not None else 0
        except (TypeError, ValueError):
            ilvl = 0
        return num_id, ilvl

    def _paragraph_num_pr(self, p, styles_tree):
        ppr = p.find(w('pPr')) if p is not None else None
        num_pr = ppr.find(w('numPr')) if ppr is not None else None
        num_id, ilvl = self._num_pr_values(num_pr)
        if num_id:
            return num_id, ilvl
        return self._num_pr_values(self._style_num_pr(styles_tree, self._paragraph_style_id(p)))

    def _resolve_heading_number_map(self, body_children, styles_tree):
        """Resolve Word numbering for heading paragraphs in current body order."""
        if self.numbering_xml is None:
            return {}
        root = self.numbering_xml.getroot()
        num_to_abs = {}
        abs_levels = {}
        for num in root.findall(w('num')):
            num_id = num.get(w('numId'))
            abs_el = num.find(w('abstractNumId'))
            if num_id and abs_el is not None:
                num_to_abs[str(num_id)] = str(abs_el.get(w('val')) or '')
        for abs_num in root.findall(w('abstractNum')):
            abs_id = str(abs_num.get(w('abstractNumId')) or '')
            levels = {}
            for lvl in abs_num.findall(w('lvl')):
                try:
                    ilvl = int(lvl.get(w('ilvl')) or 0)
                except (TypeError, ValueError):
                    ilvl = 0
                lvl_text_el = lvl.find(w('lvlText'))
                start_el = lvl.find(w('start'))
                try:
                    start = int(start_el.get(w('val')) or 1) if start_el is not None else 1
                except (TypeError, ValueError):
                    start = 1
                levels[ilvl] = {
                    'text': lvl_text_el.get(w('val')) if lvl_text_el is not None else '',
                    'start': start,
                }
            abs_levels[abs_id] = levels

        counters_by_num = {}
        result = {}
        for idx, el in enumerate(body_children):
            if el.tag != w('p'):
                continue
            num_id, ilvl = self._paragraph_num_pr(el, styles_tree)
            if not num_id:
                continue
            num_id = str(num_id)
            abs_id = num_to_abs.get(num_id)
            if abs_id is None:
                continue
            counters = counters_by_num.setdefault(num_id, [0] * 10)
            level_data = (abs_levels.get(abs_id) or {}).get(ilvl, {})
            if counters[ilvl] == 0:
                counters[ilvl] = int(level_data.get('start') or 1)
            else:
                counters[ilvl] += 1
            for deeper in range(ilvl + 1, len(counters)):
                counters[deeper] = 0

            style_id = self._paragraph_style_id(el)
            if not re.match(r'Heading\d+$', style_id or ''):
                continue
            lvl_text = str(level_data.get('text') or '')
            if lvl_text:
                number = lvl_text
                for n in range(1, 10):
                    value = counters[n - 1] if counters[n - 1] else 1
                    number = number.replace(f'%{n}', str(value))
                number = re.sub(r'%\d+', '', number)
            else:
                number = '.'.join(str(counters[n]) for n in range(ilvl + 1) if counters[n])
            number = re.sub(r'\s+', ' ', number).strip()
            result[idx] = number.rstrip('.')
        return result

    def _ensure_paragraph_style(self, styles_root, style_id, style_name):
        for s in styles_root.findall(w('style')):
            if s.get(w('styleId')) == style_id:
                return s
        s = etree.SubElement(styles_root, w('style'))
        s.set(w('type'), 'paragraph')
        s.set(w('styleId'), style_id)
        name = etree.SubElement(s, w('name'))
        name.set(w('val'), style_name)
        based_on = etree.SubElement(s, w('basedOn'))
        based_on.set(w('val'), 'Normal')
        next_el = etree.SubElement(s, w('next'))
        next_el.set(w('val'), 'Normal')
        return s

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

    def _reorder_body_blocks_by_heading_order(self, body, body_children, headings_list):
        """Move original heading content blocks to match the configured heading order."""
        original_indexes = []
        seen = set()
        for h in headings_list:
            idx = h.get('original_idx')
            if idx is None:
                continue
            try:
                idx = int(idx)
            except (TypeError, ValueError):
                continue
            if idx in seen or not (0 <= idx < len(body_children)):
                continue
            seen.add(idx)
            original_indexes.append(idx)

        if not original_indexes:
            return {}

        sorted_indexes = sorted(original_indexes)
        next_by_original = {}
        for pos, idx in enumerate(sorted_indexes):
            next_by_original[idx] = (
                sorted_indexes[pos + 1] if pos + 1 < len(sorted_indexes) else len(body_children)
            )

        block_by_original = {
            idx: body_children[idx:next_by_original[idx]]
            for idx in sorted_indexes
        }
        ordered_indexes = original_indexes + [idx for idx in sorted_indexes if idx not in seen]
        if ordered_indexes != sorted_indexes:
            first_el = body_children[sorted_indexes[0]]
            anchor = first_el.getprevious()
            for idx in sorted_indexes:
                for el in block_by_original[idx]:
                    parent = el.getparent()
                    if parent is body:
                        parent.remove(el)

            last = anchor
            insert_at_start = last is None
            for idx in ordered_indexes:
                for el in block_by_original[idx]:
                    if insert_at_start:
                        body.insert(0, el)
                        last = el
                        insert_at_start = False
                    else:
                        last.addnext(el)
                        last = el

        refreshed = [el for el in body if el.tag in (w('p'), w('tbl'))]
        return {
            idx: refreshed.index(block_by_original[idx][0])
            for idx in sorted_indexes
            if block_by_original[idx] and block_by_original[idx][0] in refreshed
        }

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
        original_idx_to_current_idx = self._reorder_body_blocks_by_heading_order(
            body, body_children, headings_list
        )
        body_children = [el for el in body if el.tag in (w('p'), w('tbl'))]
        self._cached_page_width = self._compute_page_content_width(doc_tree)
        current_heading_indexes = sorted(original_idx_to_current_idx.values())
        resolved_heading_numbers = self._resolve_heading_number_map(body_children, styles_tree)

        def next_current_heading_index(cur_idx):
            for current_idx in current_heading_indexes:
                if current_idx > cur_idx:
                    return current_idx
            return len(body_children)

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
        number_by_original_idx = {
            int(entry['original_idx']): num_str
            for entry, num_str in zip(headings_list, numbered)
            if entry.get('original_idx') is not None
        }

        # Áp text mới + số (chỉ với entry đã có original_idx)
        for entry, num_str in zip(headings_list, numbered):
            if entry.get('original_idx') is None:
                continue
            idx = original_idx_to_current_idx.get(int(entry['original_idx']))
            if idx is None:
                continue
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
        for entry in headings_list:
            if entry.get('keep_content', True):
                continue
            if entry.get('original_idx') is None:
                continue
            cur_idx = original_idx_to_current_idx.get(int(entry['original_idx']))
            if cur_idx is None:
                continue
            # Tìm heading kế tiếp BẤT KỲ (level nào cũng được)
            end_idx = next_current_heading_index(cur_idx)

            if cur_idx < len(body_children):
                heading_el = body_children[cur_idx]
                table_format_hint = None
                for k in range(cur_idx + 1, end_idx):
                    if k < len(body_children) and body_children[k].tag == w('tbl'):
                        table_format_hint = self._extract_table_format_hint(body_children[k])
                        break
                plan_selection = entry.get('insert_selection') or None
                if isinstance(plan_selection, dict):
                    plan_selection = copy.deepcopy(plan_selection)
                    advanced = plan_selection.setdefault('advanced', {})
                    if isinstance(advanced, dict):
                        advanced['_parent_heading_level'] = entry.get('level') or 1
                        parent_number = resolved_heading_numbers.get(cur_idx)
                        if not parent_number and entry.get('auto_number', False):
                            parent_number = number_by_original_idx.get(int(entry['original_idx']))
                        advanced['_parent_heading_number'] = parent_number or ''
                        advanced['_parent_heading_auto_number'] = bool(entry.get('auto_number', False))
                        target_original_idx = advanced.get('internal_link_target_heading_original_idx')
                        try:
                            target_original_idx = int(target_original_idx)
                        except (TypeError, ValueError):
                            target_original_idx = None
                        if target_original_idx is not None:
                            target_current_idx = original_idx_to_current_idx.get(target_original_idx)
                            advanced['_internal_link_bookmarks'] = self._collect_child_heading_bookmarks(
                                doc_tree, body_children, target_current_idx
                            )
                plans.append({
                    'heading_el': heading_el,
                    'start_idx': cur_idx + 1,
                    'end_idx': end_idx,
                    'insert_source': entry.get('insert_source') or None,
                    'insert_selection': plan_selection,
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

    def _make_text_paragraph(self, text='', force_left=True, runs=None):
        p = self._create_empty_paragraph()
        pPr = p.find(w('pPr'))
        if force_left and pPr is not None:
            jc = ensure_child(pPr, 'jc', PPR_ORDER)
            jc.set(w('val'), 'left')
        run_specs = runs if isinstance(runs, list) else [{'text': text}]
        for spec in run_specs:
            run = etree.SubElement(p, w('r'))
            style = spec.get('style') if isinstance(spec, dict) else {}
            if style:
                rpr = etree.SubElement(run, w('rPr'))
                if style.get('bold'):
                    ensure_child(rpr, 'b', RPR_ORDER)
                    ensure_child(rpr, 'bCs', RPR_ORDER)
                if style.get('italic'):
                    ensure_child(rpr, 'i', RPR_ORDER)
                    ensure_child(rpr, 'iCs', RPR_ORDER)
                if style.get('underline'):
                    u = ensure_child(rpr, 'u', RPR_ORDER)
                    u.set(w('val'), 'single')
            t = etree.SubElement(run, w('t'))
            t.text = str(spec.get('text', '') if isinstance(spec, dict) else spec)
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return p

    def _make_text_paragraphs_from_excel_text(self, prefix, value, prefix_style=None, force_left=True):
        prefix_text = str(prefix or '').replace('\r\n', '\n').replace('\r', '\n')
        value_lines = str(value or '').replace('\r\n', '\n').replace('\r', '\n').split('\n') or ['']
        prefix_style = prefix_style if isinstance(prefix_style, dict) else {}
        paragraphs = []
        if not prefix_text:
            return [self._make_text_paragraph(line, force_left=force_left) for line in value_lines]

        prefix_lines = prefix_text.split('\n')
        prefix_ends_line = prefix_text.endswith('\n')
        standalone_prefix_lines = prefix_lines[:-1] if prefix_ends_line else prefix_lines[:-1]
        for line in standalone_prefix_lines:
            paragraphs.append(self._make_text_paragraph(
                force_left=force_left,
                runs=[{'text': line, 'style': prefix_style}],
            ))

        if prefix_ends_line:
            for line in value_lines:
                paragraphs.append(self._make_text_paragraph(line, force_left=force_left))
            return paragraphs

        last_prefix = prefix_lines[-1] if prefix_lines else ''
        first_value = value_lines[0] if value_lines else ''
        paragraphs.append(self._make_text_paragraph(
            force_left=force_left,
            runs=[
                {'text': last_prefix, 'style': prefix_style},
                {'text': first_value, 'style': {}},
            ],
        ))
        for line in value_lines[1:]:
            paragraphs.append(self._make_text_paragraph(line, force_left=force_left))
        return paragraphs

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

        class _ExcelCellView:
            """Cell view: formula result from data_only workbook, formatting from source workbook."""
            def __init__(self, value_cell, style_cell):
                self._value_cell = value_cell
                self._style_cell = style_cell
                self._override_value = None
                self._has_override = False
                self.row = getattr(style_cell, 'row', getattr(value_cell, 'row', None))
                self.column = getattr(style_cell, 'column', getattr(value_cell, 'column', None))
                self.parent = getattr(style_cell, 'parent', getattr(value_cell, 'parent', None))

            @property
            def value(self):
                if self._has_override:
                    return self._override_value
                return getattr(self._value_cell, 'value', None)

            @value.setter
            def value(self, value):
                self._override_value = value
                self._has_override = True

            def __getattr__(self, name):
                return getattr(self._style_cell, name)

        class _SyntheticExcelCellView:
            """Small cell-like object for values generated by advanced table transforms."""
            def __init__(self, value='', style_cell=None):
                self.value = value
                self._style_cell = style_cell
                self.row = getattr(style_cell, 'row', None)
                self.column = getattr(style_cell, 'column', None)
                self.parent = getattr(style_cell, 'parent', None)
                self.number_format = getattr(style_cell, 'number_format', 'General')
                self.font = copy.copy(getattr(style_cell, 'font', None)) if style_cell is not None else None
                self.alignment = copy.copy(getattr(style_cell, 'alignment', None)) if style_cell is not None else None
                self.hyperlink = None

        # Bỏ read_only=True để max_row/max_column populate đúng.
        # Mở 2 workbook: data_only lấy giá trị công thức đã cache, workbook gốc giữ style/hyperlink.
        wb_values = load_workbook(source_path, data_only=True)
        wb = load_workbook(source_path, data_only=False)

        def _wrap_cell(value_ws, style_cell):
            if style_cell is None:
                return None
            value_cell = value_ws.cell(row=style_cell.row, column=style_cell.column)
            return _ExcelCellView(value_cell, style_cell)

        def _wrap_iter_rows(value_ws, style_ws):
            return [
                [_wrap_cell(value_ws, cell) for cell in row]
                for row in style_ws.iter_rows()
            ]

        def _adv_bool_from(advanced_cfg, key, default=False):
            value = (advanced_cfg or {}).get(key, default)
            if isinstance(value, bool):
                return value
            return str(value).strip().lower() in {'1', 'true', 'yes', 'on', 'y'}

        def _cell_display_text(cell):
            if cell is None or cell.value is None:
                return ''
            value = cell.value
            if hasattr(value, 'isoformat'):
                value = value.isoformat()
            return str(value).strip()

        def _cell_has_real_value(cell):
            if cell is None:
                return False
            value = getattr(cell, 'value', None)
            if value is None:
                return False
            if isinstance(value, str):
                return bool(value.strip())
            return True

        def _effective_auto_row_end(style_ws, value_ws, row_start, col_indexes, extra_columns=None):
            """Find the real last data row when UI row_end is left empty.

            Excel files often keep formatting far below the actual table. In that
            case openpyxl's max_row points to the formatted area, and generating
            a DOCX table for every empty formatted row can look like a stuck job.
            """
            columns = {
                int(c) for c in (col_indexes or [])
                if str(c).strip().isdigit() and int(c) > 0
            }
            columns.update(
                int(c) for c in (extra_columns or [])
                if str(c).strip().isdigit() and int(c) > 0
            )
            if not columns:
                columns = set(range(1, (style_ws.max_column or 1) + 1))

            last_row = None
            for ws_candidate in (style_ws, value_ws):
                cells = getattr(ws_candidate, '_cells', {}) or {}
                for coord, cell in cells.items():
                    try:
                        row_idx, col_idx = coord
                    except (TypeError, ValueError):
                        continue
                    if row_idx < row_start or col_idx not in columns:
                        continue
                    if _cell_has_real_value(cell):
                        last_row = max(last_row or row_start, int(row_idx))

            if last_row is not None:
                return max(row_start, last_row)
            return row_start

        def _apply_auto_stt(rows, merge_positions, advanced_cfg):
            if not _adv_bool_from(advanced_cfg, 'auto_stt_enabled', False):
                return rows, merge_positions
            if not rows:
                return rows, merge_positions

            shifted_merge_positions = {int(pos) + 1 for pos in (merge_positions or set())}
            merge_stt = 1 in shifted_merge_positions
            new_rows = []
            header_style = rows[0][0] if rows[0] else None
            new_rows.append([_SyntheticExcelCellView('STT', header_style)] + rows[0])

            current_no = 0
            last_key = object()
            for row in rows[1:]:
                key = _cell_display_text(row[0] if row else None) if merge_stt else None
                if merge_stt:
                    if key != last_key:
                        current_no += 1
                        last_key = key
                    number = current_no
                else:
                    current_no += 1
                    number = current_no
                style_cell = row[0] if row else header_style
                new_rows.append([_SyntheticExcelCellView(number, style_cell)] + row)

            if merge_stt:
                shifted_merge_positions.add(0)
            return new_rows, shifted_merge_positions

        def _internal_link_positions(col_indexes, advanced_cfg):
            if not _adv_bool_from(advanced_cfg, 'internal_link_enabled', False):
                return {}
            try:
                link_col = int((advanced_cfg or {}).get('internal_link_column') or 0)
            except (TypeError, ValueError):
                link_col = 0
            if not link_col:
                return {}
            bookmarks = (advanced_cfg or {}).get('_internal_link_bookmarks') or {}
            if not isinstance(bookmarks, dict):
                bookmarks = {}
            positions = {
                pos for pos, col_idx in enumerate(col_indexes or [])
                if int(col_idx) == link_col
            }
            offset = 1 if _adv_bool_from(advanced_cfg, 'auto_stt_enabled', False) else 0
            return {pos + offset: bookmarks for pos in positions}

        def _normalize_lookup_key(value):
            text = unicodedata.normalize('NFD', str(value or '').strip())
            text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
            return re.sub(r'\s+', ' ', text).lower()

        def _merged_cell_display_text(ws, row, col):
            cell = ws.cell(row=row, column=col)
            value = cell.value
            if value is None:
                for merged_range in ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        value = ws.cell(merged_range.min_row, merged_range.min_col).value
                        break
            if value is None:
                return ''
            if hasattr(value, 'isoformat'):
                value = value.isoformat()
            return str(value).strip()

        lookup_rule_cache = {}
        lookup_rule_any_column_cache = {}

        def _lookup_values_for_rule(rule):
            src_path = str(rule.get('source_path') or '').strip()
            sheet_name = str(rule.get('sheet') or '').strip()
            try:
                key_col = int(rule.get('key_column') or 0)
                value_col = int(rule.get('value_column') or 0)
            except (TypeError, ValueError):
                return {}
            if not src_path or not sheet_name or not key_col or not value_col:
                return {}

            try:
                row_start = max(1, int(rule.get('row_start') or 1))
            except (TypeError, ValueError):
                row_start = 1
            row_end_raw = rule.get('row_end')
            signature = (src_path, sheet_name, key_col, value_col, row_start, str(row_end_raw or ''))
            if signature in lookup_rule_cache:
                return lookup_rule_cache[signature]

            values_by_key = {}
            lookup_wb = None
            try:
                lookup_wb = load_workbook(src_path, data_only=True)
                if sheet_name not in lookup_wb.sheetnames:
                    lookup_rule_cache[signature] = values_by_key
                    return values_by_key
                lookup_ws = lookup_wb[sheet_name]
                try:
                    row_end = int(row_end_raw) if row_end_raw not in (None, '', 0, '0') else lookup_ws.max_row
                except (TypeError, ValueError):
                    row_end = lookup_ws.max_row
                for row_idx in range(row_start, max(row_start, row_end) + 1):
                    key_text = _merged_cell_display_text(lookup_ws, row_idx, key_col)
                    normalized = _normalize_lookup_key(key_text)
                    if not normalized:
                        continue
                    value_text = _merged_cell_display_text(lookup_ws, row_idx, value_col)
                    if not value_text:
                        continue
                    bucket = values_by_key.setdefault(normalized, [])
                    if value_text not in bucket:
                        bucket.append(value_text)
            except Exception:
                values_by_key = {}
            finally:
                if lookup_wb is not None:
                    try:
                        lookup_wb.close()
                    except Exception:
                        pass
            lookup_rule_cache[signature] = values_by_key
            return values_by_key

        def _lookup_values_for_rule_any_column(rule):
            src_path = str(rule.get('source_path') or '').strip()
            sheet_name = str(rule.get('sheet') or '').strip()
            try:
                value_col = int(rule.get('value_column') or 0)
            except (TypeError, ValueError):
                return {}
            if not src_path or not sheet_name or not value_col:
                return {}

            try:
                row_start = max(1, int(rule.get('row_start') or 1))
            except (TypeError, ValueError):
                row_start = 1
            row_end_raw = rule.get('row_end')
            signature = (src_path, sheet_name, value_col, row_start, str(row_end_raw or ''))
            if signature in lookup_rule_any_column_cache:
                return lookup_rule_any_column_cache[signature]

            values_by_key = {}
            lookup_wb = None
            try:
                lookup_wb = load_workbook(src_path, data_only=True)
                if sheet_name not in lookup_wb.sheetnames:
                    lookup_rule_any_column_cache[signature] = values_by_key
                    return values_by_key
                lookup_ws = lookup_wb[sheet_name]
                try:
                    row_end = int(row_end_raw) if row_end_raw not in (None, '', 0, '0') else lookup_ws.max_row
                except (TypeError, ValueError):
                    row_end = lookup_ws.max_row

                for row_idx in range(row_start, max(row_start, row_end) + 1):
                    value_text = _merged_cell_display_text(lookup_ws, row_idx, value_col)
                    if not value_text:
                        continue
                    row_keys = set()
                    for col_idx in range(1, (lookup_ws.max_column or 1) + 1):
                        key_text = _merged_cell_display_text(lookup_ws, row_idx, col_idx)
                        normalized = _normalize_lookup_key(key_text)
                        if normalized:
                            row_keys.add(normalized)
                    for normalized in row_keys:
                        bucket = values_by_key.setdefault(normalized, [])
                        if value_text not in bucket:
                            bucket.append(value_text)
            except Exception:
                values_by_key = {}
            finally:
                if lookup_wb is not None:
                    try:
                        lookup_wb.close()
                    except Exception:
                        pass
            lookup_rule_any_column_cache[signature] = values_by_key
            return values_by_key

        def _lookup_text_paragraphs(group_key, advanced_cfg):
            rules = (advanced_cfg or {}).get('split_text_rules') or []
            if not isinstance(rules, list):
                return []
            out = []
            lookup_key = _normalize_lookup_key(group_key)
            for rule in rules:
                if not isinstance(rule, dict) or not _adv_bool_from(rule, 'enabled', True):
                    continue
                prefix_text = str(rule.get('prefix') or '')
                prefix_style = {
                    'bold': _adv_bool_from(rule, 'prefix_bold', False),
                    'italic': _adv_bool_from(rule, 'prefix_italic', False),
                    'underline': _adv_bool_from(rule, 'prefix_underline', False),
                }
                force_left = _adv_bool_from(rule, 'force_left', True)
                values = _lookup_values_for_rule(rule).get(lookup_key, [])
                if not values:
                    values = _lookup_values_for_rule_any_column(rule).get(lookup_key, [])
                for value_text in values:
                    out.extend(self._make_text_paragraphs_from_excel_text(
                        prefix_text, value_text, prefix_style, force_left
                    ))
            return out

        def _without_multi_sources(advanced_cfg):
            clone = copy.deepcopy(advanced_cfg or {})
            clone.pop('split_sources', None)
            clone.pop('multi_source_mode', None)
            return clone

        def _source_title(src_path, src_selection):
            sheet_title = str((src_selection or {}).get('sheet') or '').strip()
            file_title = Path(src_path).stem if src_path else ''
            return sheet_title or file_title or 'Nguồn'

        def _table_elements_from_rows(rows, col_indexes, merge_positions, advanced_cfg, group_key=None):
            link_columns = _internal_link_positions(col_indexes, advanced_cfg)
            table_rows, table_merge_positions = _apply_auto_stt(
                rows, merge_positions, advanced_cfg
            )
            tbl = self._build_table_from_cell_rows(
                table_rows, preserve_inline,
                auto_merge_columns=table_merge_positions,
                internal_link_columns=link_columns,
            )
            if tbl is None:
                return []
            extra = _lookup_text_paragraphs(group_key, advanced_cfg) if group_key is not None else []
            return [tbl, *extra, self._create_empty_paragraph()]

        def _prepare_custom_split_source(src_path, src_selection):
            """Return split groups for one custom Excel source used by interleave mode."""
            src_selection = src_selection or {}
            src_mode = src_selection.get('mode', 'custom')
            src_wb_values = wb_values if str(src_path) == str(source_path) else load_workbook(src_path, data_only=True)
            src_wb = wb if str(src_path) == str(source_path) else load_workbook(src_path, data_only=False)
            sheet_name = src_selection.get('sheet')
            if not sheet_name or sheet_name not in src_wb.sheetnames:
                raise ValueError(f'Sheet "{sheet_name}" không tồn tại trong file.')
            src_ws = src_wb[sheet_name]
            src_value_ws = src_wb_values[sheet_name]
            advanced_cfg = src_selection.get('advanced') or {}
            if not isinstance(advanced_cfg, dict):
                advanced_cfg = {}

            if src_mode == 'range':
                rows = self._extract_rows_from_range(src_ws, src_selection.get('range') or '', src_value_ws)
                col_indexes = list(range(1, max((len(row) for row in rows), default=0) + 1))
                row_start = 1
                row_end = len(rows)
            else:
                cols_spec = (src_selection.get('columns') or '').strip()
                if cols_spec:
                    col_indexes = self._parse_column_spec(cols_spec)
                else:
                    col_indexes = list(range(1, (src_ws.max_column or 1) + 1))
                if not col_indexes:
                    raise ValueError(f'Danh sách cột custom không hợp lệ: "{cols_spec}"')

                try:
                    row_start = max(1, int(src_selection.get('row_start') or 1))
                except (TypeError, ValueError):
                    row_start = 1
                try:
                    row_end = int(src_selection.get('row_end')) if src_selection.get('row_end') not in (None, '', 0, '0') else None
                except (TypeError, ValueError):
                    row_end = None
                if row_end is None:
                    extra_columns = [advanced_cfg.get('split_column')]
                    row_end = _effective_auto_row_end(
                        src_ws, src_value_ws, row_start, col_indexes, extra_columns
                    )
                row_end = max(row_start, row_end)

                rows = []
                for r in range(row_start, row_end + 1):
                    rows.append([
                        _ExcelCellView(
                            src_value_ws.cell(row=r, column=c),
                            src_ws.cell(row=r, column=c),
                        )
                        for c in col_indexes
                    ])

            header_overrides = src_selection.get('header_overrides') or {}
            if rows and isinstance(header_overrides, dict):
                for pos, col_idx in enumerate(col_indexes):
                    new_header = header_overrides.get(str(col_idx), header_overrides.get(col_idx))
                    if new_header is not None and pos < len(rows[0]):
                        rows[0][pos].value = str(new_header)

            merge_source_columns = advanced_cfg.get('merge_columns') or []
            if not isinstance(merge_source_columns, (list, tuple, set)):
                merge_source_columns = []
            merge_source_columns = {
                int(c) for c in merge_source_columns
                if str(c).strip().isdigit()
            }
            merge_positions = {
                pos for pos, col_idx in enumerate(col_indexes)
                if int(col_idx) in merge_source_columns
            }

            split_column = advanced_cfg.get('split_column')
            try:
                split_column = int(split_column)
            except (TypeError, ValueError):
                split_column = None

            header_row = rows[0] if rows else []
            grouped = []
            group_by_key = {}
            if _adv_bool_from(advanced_cfg, 'split_enabled', False) and split_column:
                for excel_row in range(row_start + 1, row_end + 1):
                    row_pos = excel_row - row_start
                    if row_pos < 0 or row_pos >= len(rows):
                        continue
                    group_cell = _ExcelCellView(
                        src_value_ws.cell(row=excel_row, column=split_column),
                        src_ws.cell(row=excel_row, column=split_column),
                    )
                    key = _cell_display_text(group_cell) or '(Trống)'
                    if key not in group_by_key:
                        group_by_key[key] = []
                        grouped.append((key, group_by_key[key]))
                    group_by_key[key].append(rows[row_pos])
            elif rows:
                grouped.append((_source_title(src_path, src_selection), rows[1:] if header_row else rows))

            return {
                'title': _source_title(src_path, src_selection),
                'header_row': header_row,
                'groups': grouped,
                'group_map': {key: data_rows for key, data_rows in grouped},
                'col_indexes': col_indexes,
                'merge_positions': merge_positions,
                'advanced': advanced_cfg,
            }

        def _build_interleaved_sources(sources, root_advanced):
            prepared = [
                item for item in (
                    _prepare_custom_split_source(src_path, src_selection)
                    for src_path, src_selection in sources
                )
                if item.get('groups')
            ]
            if not prepared:
                return []

            add_heading = _adv_bool_from(root_advanced, 'add_split_heading', False)
            prefix = str(root_advanced.get('heading_prefix', 'Bảng ') or '')
            suffix = str(root_advanced.get('heading_suffix', '') or '')
            add_group_heading = _adv_bool_from(root_advanced, 'interleave_group_heading_enabled', True)
            group_prefix = str(root_advanced.get('interleave_group_heading_prefix', 'Cụm ') or '')
            group_suffix = str(root_advanced.get('interleave_group_heading_suffix', '') or '')
            parent_level = root_advanced.get('_parent_heading_level') or 1
            try:
                group_heading_level = min(9, max(1, int(parent_level) + 1))
            except (TypeError, ValueError):
                group_heading_level = 4
            table_heading_level = min(9, group_heading_level + 1) if add_group_heading else group_heading_level
            parent_number = str(root_advanced.get('_parent_heading_number') or '').strip()
            heading_style = self._advanced_heading_style(root_advanced)

            out = []
            max_group_count = max(len(item['groups']) for item in prepared)
            for group_pos in range(max_group_count):
                group_items = []
                for item in prepared:
                    if group_pos >= len(item['groups']):
                        continue
                    group_key, data_rows = item['groups'][group_pos]
                    if not data_rows:
                        continue
                    group_items.append((item, group_key, data_rows))
                if not group_items:
                    continue

                cluster_idx = group_pos + 1
                cluster_number = f'{parent_number}.{cluster_idx}' if parent_number else ''
                if add_heading and add_group_heading:
                    cluster_key = group_items[0][1]
                    cluster_text = f'{group_prefix}{cluster_key}{group_suffix}'
                    if cluster_number:
                        cluster_text = f'{cluster_number} {cluster_text}'
                    out.append(self._make_subheading_paragraph(
                        cluster_text, group_heading_level, heading_style
                    ))

                for inner_idx, (item, group_key, data_rows) in enumerate(group_items, start=1):
                    bookmark_name = ''
                    if add_heading:
                        heading_text = f'{prefix}{group_key}{suffix}'
                        if add_group_heading and cluster_number:
                            heading_text = f'{cluster_number}.{inner_idx} {heading_text}'
                        elif parent_number:
                            flat_idx = group_pos * len(prepared) + inner_idx
                            heading_text = f'{parent_number}.{flat_idx} {heading_text}'
                        if _adv_bool_from(root_advanced, 'internal_link_enabled', False):
                            bookmark_name = self._safe_bookmark_name('ILINK', group_key)
                            item['advanced'].setdefault('_internal_link_bookmarks', {})[
                                self._normalize_internal_link_key(group_key)
                            ] = bookmark_name
                        out.append(self._make_subheading_paragraph(
                            heading_text, table_heading_level, heading_style, bookmark_name
                        ))
                    tbl_rows = [item['header_row']] + data_rows if item['header_row'] else data_rows
                    out.extend(_table_elements_from_rows(
                        tbl_rows,
                        item['col_indexes'],
                        item['merge_positions'],
                        item['advanced'],
                        group_key,
                    ))
            return out

        advanced_root = selection.get('advanced') or {}
        if isinstance(advanced_root, dict) and advanced_root.get('split_sources'):
            sources = []
            base_selection = copy.deepcopy(selection)
            base_selection['advanced'] = _without_multi_sources(advanced_root)
            sources.append((source_path, base_selection))
            for src in advanced_root.get('split_sources') or []:
                if not isinstance(src, dict):
                    continue
                src_path = str(src.get('source_path') or '').strip()
                if not src_path:
                    continue
                src_selection = copy.deepcopy(selection)
                src_selection.update({
                    'mode': 'custom',
                    'file_id': src.get('file_id') or selection.get('file_id'),
                    'sheet': src.get('sheet') or selection.get('sheet'),
                    'columns': src.get('columns') or selection.get('columns'),
                    'row_start': src.get('row_start') or selection.get('row_start') or 1,
                    'row_end': src.get('row_end') if src.get('row_end') not in ('', None) else selection.get('row_end'),
                })
                src_adv = _without_multi_sources(advanced_root)
                if src.get('split_column'):
                    src_adv['split_column'] = src.get('split_column')
                src_selection['advanced'] = src_adv
                sources.append((src_path, src_selection))
            if str(advanced_root.get('multi_source_mode') or '').strip().lower() == 'interleave':
                return _build_interleaved_sources(sources, _without_multi_sources(advanced_root))
            multi_elements = []
            next_heading_index = 1
            for src_path, src_selection in sources:
                src_selection = copy.deepcopy(src_selection)
                src_advanced = src_selection.get('advanced') if isinstance(src_selection.get('advanced'), dict) else {}
                src_advanced['_split_heading_start_index'] = next_heading_index
                src_selection['advanced'] = src_advanced
                multi_elements.extend(self._read_excel_content(src_path, preserve_inline, src_selection))
                try:
                    next_heading_index += len(_prepare_custom_split_source(src_path, src_selection).get('groups') or [])
                except Exception:
                    next_heading_index += 0
            return multi_elements

        elements = []

        if mode == 'all':
            for ws in wb.worksheets:
                if len(wb.worksheets) > 1:
                    elements.append(self._make_subheading_paragraph(ws.title))
                rows = _wrap_iter_rows(wb_values[ws.title], ws)
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
            rows = _wrap_iter_rows(wb_values[sheet_name], ws)
            col_indexes = list(range(1, max((len(row) for row in rows), default=0) + 1))
            row_start = 1
            row_end = len(rows)
            value_ws = wb_values[sheet_name]
            advanced = selection.get('advanced') or {}
            if not isinstance(advanced, dict):
                advanced = {}

            def _adv_bool(key, default=False):
                value = advanced.get(key, default)
                if isinstance(value, bool):
                    return value
                return str(value).strip().lower() in {'1', 'true', 'yes', 'on', 'y'}

            def _cell_display_text(cell):
                if cell is None or cell.value is None:
                    return ''
                value = cell.value
                if hasattr(value, 'isoformat'):
                    value = value.isoformat()
                return str(value).strip()

            merge_source_columns = advanced.get('merge_columns') or []
            if not isinstance(merge_source_columns, (list, tuple, set)):
                merge_source_columns = []
            merge_source_columns = {
                int(c) for c in merge_source_columns
                if str(c).strip().isdigit()
            }
            merge_positions = {
                pos for pos, col_idx in enumerate(col_indexes)
                if int(col_idx) in merge_source_columns
            }

            split_enabled = _adv_bool('split_enabled', False)
            split_column = advanced.get('split_column')
            try:
                split_column = int(split_column)
            except (TypeError, ValueError):
                split_column = None

            if split_enabled and split_column:
                header_row = rows[0] if rows else []
                grouped = []
                group_by_key = {}
                for excel_row in range(row_start + 1, row_end + 1):
                    group_cell = _ExcelCellView(
                        value_ws.cell(row=excel_row, column=split_column),
                        ws.cell(row=excel_row, column=split_column),
                    )
                    key = _cell_display_text(group_cell) or '(Trống)'
                    if key not in group_by_key:
                        group_by_key[key] = []
                        grouped.append((key, group_by_key[key]))
                    group_by_key[key].append(rows[excel_row - row_start])

                if not grouped:
                    link_columns = _internal_link_positions(col_indexes, advanced)
                    table_rows, table_merge_positions = _apply_auto_stt(
                        rows, merge_positions, advanced
                    )
                    tbl = self._build_table_from_cell_rows(
                        table_rows, preserve_inline,
                        auto_merge_columns=table_merge_positions,
                        internal_link_columns=link_columns,
                    )
                    if tbl is not None:
                        elements.append(tbl)
                        elements.append(self._create_empty_paragraph())
                    return elements

                add_heading = _adv_bool('add_split_heading', False)
                prefix = str(advanced.get('heading_prefix', 'Bảng ') or '')
                suffix = str(advanced.get('heading_suffix', '') or '')
                parent_level = advanced.get('_parent_heading_level') or 1
                try:
                    heading_level = min(9, max(1, int(parent_level) + 1))
                except (TypeError, ValueError):
                    heading_level = 4
                parent_number = str(advanced.get('_parent_heading_number') or '').strip()
                try:
                    split_heading_start_index = max(1, int(advanced.get('_split_heading_start_index') or 1))
                except (TypeError, ValueError):
                    split_heading_start_index = 1
                heading_style = self._advanced_heading_style(advanced)

                for group_idx, (group_key, data_rows) in enumerate(grouped, start=1):
                    heading_idx = split_heading_start_index + group_idx - 1
                    bookmark_name = ''
                    if add_heading:
                        heading_text = f'{prefix}{group_key}{suffix}'
                        if parent_number:
                            heading_text = f'{parent_number}.{heading_idx} {heading_text}'
                        if _adv_bool_from(advanced, 'internal_link_enabled', False):
                            bookmark_name = self._safe_bookmark_name('ILINK', group_key)
                            advanced.setdefault('_internal_link_bookmarks', {})[
                                self._normalize_internal_link_key(group_key)
                            ] = bookmark_name
                        elements.append(self._make_subheading_paragraph(
                            heading_text, heading_level, heading_style, bookmark_name
                        ))
                    tbl_rows = [header_row] + data_rows if header_row else data_rows
                    link_columns = _internal_link_positions(col_indexes, advanced)
                    tbl_rows, table_merge_positions = _apply_auto_stt(
                        tbl_rows, merge_positions, advanced
                    )
                    tbl = self._build_table_from_cell_rows(
                        tbl_rows, preserve_inline,
                        auto_merge_columns=table_merge_positions,
                        internal_link_columns=link_columns,
                    )
                    if tbl is not None:
                        elements.append(tbl)
                        elements.extend(_lookup_text_paragraphs(group_key, advanced))
                        elements.append(self._create_empty_paragraph())
            else:
                link_columns = _internal_link_positions(col_indexes, advanced)
                table_rows, table_merge_positions = _apply_auto_stt(
                    rows, merge_positions, advanced
                )
                tbl = self._build_table_from_cell_rows(
                    table_rows, preserve_inline,
                    auto_merge_columns=table_merge_positions,
                    internal_link_columns=link_columns,
                )
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
            rows = self._extract_rows_from_range(ws, range_str, wb_values[sheet_name])
            col_indexes = list(range(1, max((len(row) for row in rows), default=0) + 1))
            row_start = 1
            row_end = len(rows)
            value_ws = wb_values[sheet_name]
            advanced = selection.get('advanced') or {}
            if not isinstance(advanced, dict):
                advanced = {}

            def _adv_bool(key, default=False):
                value = advanced.get(key, default)
                if isinstance(value, bool):
                    return value
                return str(value).strip().lower() in {'1', 'true', 'yes', 'on', 'y'}

            def _cell_display_text(cell):
                if cell is None or cell.value is None:
                    return ''
                value = cell.value
                if hasattr(value, 'isoformat'):
                    value = value.isoformat()
                return str(value).strip()

            merge_source_columns = advanced.get('merge_columns') or []
            if not isinstance(merge_source_columns, (list, tuple, set)):
                merge_source_columns = []
            merge_source_columns = {
                int(c) for c in merge_source_columns
                if str(c).strip().isdigit()
            }
            merge_positions = {
                pos for pos, col_idx in enumerate(col_indexes)
                if int(col_idx) in merge_source_columns
            }

            split_enabled = _adv_bool('split_enabled', False)
            split_column = advanced.get('split_column')
            try:
                split_column = int(split_column)
            except (TypeError, ValueError):
                split_column = None

            if split_enabled and split_column:
                header_row = rows[0] if rows else []
                grouped = []
                group_by_key = {}
                for excel_row in range(row_start + 1, row_end + 1):
                    group_cell = _ExcelCellView(
                        value_ws.cell(row=excel_row, column=split_column),
                        ws.cell(row=excel_row, column=split_column),
                    )
                    key = _cell_display_text(group_cell) or '(Trống)'
                    if key not in group_by_key:
                        group_by_key[key] = []
                        grouped.append((key, group_by_key[key]))
                    group_by_key[key].append(rows[excel_row - row_start])

                if not grouped:
                    link_columns = _internal_link_positions(col_indexes, advanced)
                    table_rows, table_merge_positions = _apply_auto_stt(
                        rows, merge_positions, advanced
                    )
                    tbl = self._build_table_from_cell_rows(
                        table_rows, preserve_inline,
                        auto_merge_columns=table_merge_positions,
                        internal_link_columns=link_columns,
                    )
                    if tbl is not None:
                        elements.append(tbl)
                        elements.append(self._create_empty_paragraph())
                    return elements

                add_heading = _adv_bool('add_split_heading', False)
                prefix = str(advanced.get('heading_prefix', 'Bảng ') or '')
                suffix = str(advanced.get('heading_suffix', '') or '')
                parent_level = advanced.get('_parent_heading_level') or 1
                try:
                    heading_level = min(9, max(1, int(parent_level) + 1))
                except (TypeError, ValueError):
                    heading_level = 4
                parent_number = str(advanced.get('_parent_heading_number') or '').strip()
                try:
                    split_heading_start_index = max(1, int(advanced.get('_split_heading_start_index') or 1))
                except (TypeError, ValueError):
                    split_heading_start_index = 1
                heading_style = self._advanced_heading_style(advanced)

                for group_idx, (group_key, data_rows) in enumerate(grouped, start=1):
                    heading_idx = split_heading_start_index + group_idx - 1
                    bookmark_name = ''
                    if add_heading:
                        heading_text = f'{prefix}{group_key}{suffix}'
                        if parent_number:
                            heading_text = f'{parent_number}.{heading_idx} {heading_text}'
                        if _adv_bool_from(advanced, 'internal_link_enabled', False):
                            bookmark_name = self._safe_bookmark_name('ILINK', group_key)
                            advanced.setdefault('_internal_link_bookmarks', {})[
                                self._normalize_internal_link_key(group_key)
                            ] = bookmark_name
                        elements.append(self._make_subheading_paragraph(
                            heading_text, heading_level, heading_style, bookmark_name
                        ))
                    tbl_rows = [header_row] + data_rows if header_row else data_rows
                    link_columns = _internal_link_positions(col_indexes, advanced)
                    tbl_rows, table_merge_positions = _apply_auto_stt(
                        tbl_rows, merge_positions, advanced
                    )
                    tbl = self._build_table_from_cell_rows(
                        tbl_rows, preserve_inline,
                        auto_merge_columns=table_merge_positions,
                        internal_link_columns=link_columns,
                    )
                    if tbl is not None:
                        elements.append(tbl)
                        elements.extend(_lookup_text_paragraphs(group_key, advanced))
                        elements.append(self._create_empty_paragraph())
            else:
                link_columns = _internal_link_positions(col_indexes, advanced)
                table_rows, table_merge_positions = _apply_auto_stt(
                    rows, merge_positions, advanced
                )
                tbl = self._build_table_from_cell_rows(
                    table_rows, preserve_inline,
                    auto_merge_columns=table_merge_positions,
                    internal_link_columns=link_columns,
                )
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
            value_ws = wb_values[sheet_name]

            max_col = ws.max_column or 1

            # Parse columns spec
            if cols_spec:
                col_indexes = self._parse_column_spec(cols_spec)
            else:
                col_indexes = list(range(1, max_col + 1))
            if not col_indexes:
                raise ValueError(f'Danh sách cột custom không hợp lệ: "{cols_spec}"')

            if row_end is None:
                advanced_for_bounds = selection.get('advanced') or {}
                if not isinstance(advanced_for_bounds, dict):
                    advanced_for_bounds = {}
                row_end = _effective_auto_row_end(
                    ws, value_ws, row_start, col_indexes,
                    [advanced_for_bounds.get('split_column')]
                )
            row_end = max(row_start, row_end)

            # Build rows: lấy cell theo col_indexes × row range
            rows = []
            for r in range(row_start, row_end + 1):
                row_cells = [_ExcelCellView(
                    value_ws.cell(row=r, column=c),
                    ws.cell(row=r, column=c),
                ) for c in col_indexes]
                rows.append(row_cells)
            header_overrides = selection.get('header_overrides') or {}
            if rows and isinstance(header_overrides, dict):
                for pos, col_idx in enumerate(col_indexes):
                    new_header = header_overrides.get(str(col_idx), header_overrides.get(col_idx))
                    if new_header is None or pos >= len(rows[0]):
                        continue
                    try:
                        rows[0][pos].value = str(new_header)
                    except AttributeError:
                        pass

            advanced = selection.get('advanced') or {}
            if not isinstance(advanced, dict):
                advanced = {}

            def _adv_bool(key, default=False):
                value = advanced.get(key, default)
                if isinstance(value, bool):
                    return value
                return str(value).strip().lower() in {'1', 'true', 'yes', 'on', 'y'}

            def _cell_display_text(cell):
                if cell is None or cell.value is None:
                    return ''
                value = cell.value
                if hasattr(value, 'isoformat'):
                    value = value.isoformat()
                return str(value).strip()

            merge_source_columns = advanced.get('merge_columns') or []
            if not isinstance(merge_source_columns, (list, tuple, set)):
                merge_source_columns = []
            merge_source_columns = {
                int(c) for c in merge_source_columns
                if str(c).strip().isdigit()
            }
            merge_positions = {
                pos for pos, col_idx in enumerate(col_indexes)
                if int(col_idx) in merge_source_columns
            }

            split_enabled = _adv_bool('split_enabled', False)
            split_column = advanced.get('split_column')
            try:
                split_column = int(split_column)
            except (TypeError, ValueError):
                split_column = None

            if split_enabled and split_column:
                header_row = rows[0] if rows else []
                grouped = []
                group_by_key = {}
                for excel_row in range(row_start + 1, row_end + 1):
                    group_cell = _ExcelCellView(
                        value_ws.cell(row=excel_row, column=split_column),
                        ws.cell(row=excel_row, column=split_column),
                    )
                    key = _cell_display_text(group_cell) or '(Trống)'
                    if key not in group_by_key:
                        group_by_key[key] = []
                        grouped.append((key, group_by_key[key]))
                    group_by_key[key].append(rows[excel_row - row_start])

                if not grouped:
                    link_columns = _internal_link_positions(col_indexes, advanced)
                    table_rows, table_merge_positions = _apply_auto_stt(
                        rows, merge_positions, advanced
                    )
                    tbl = self._build_table_from_cell_rows(
                        table_rows, preserve_inline,
                        auto_merge_columns=table_merge_positions,
                        internal_link_columns=link_columns,
                    )
                    if tbl is not None:
                        elements.append(tbl)
                        elements.append(self._create_empty_paragraph())
                    return elements

                add_heading = _adv_bool('add_split_heading', False)
                prefix = str(advanced.get('heading_prefix', 'Bảng ') or '')
                suffix = str(advanced.get('heading_suffix', '') or '')
                parent_level = advanced.get('_parent_heading_level') or 1
                try:
                    heading_level = min(9, max(1, int(parent_level) + 1))
                except (TypeError, ValueError):
                    heading_level = 4
                parent_number = str(advanced.get('_parent_heading_number') or '').strip()
                try:
                    split_heading_start_index = max(1, int(advanced.get('_split_heading_start_index') or 1))
                except (TypeError, ValueError):
                    split_heading_start_index = 1
                heading_style = self._advanced_heading_style(advanced)

                for group_idx, (group_key, data_rows) in enumerate(grouped, start=1):
                    heading_idx = split_heading_start_index + group_idx - 1
                    bookmark_name = ''
                    if add_heading:
                        heading_text = f'{prefix}{group_key}{suffix}'
                        if parent_number:
                            heading_text = f'{parent_number}.{heading_idx} {heading_text}'
                        if _adv_bool_from(advanced, 'internal_link_enabled', False):
                            bookmark_name = self._safe_bookmark_name('ILINK', group_key)
                            advanced.setdefault('_internal_link_bookmarks', {})[
                                self._normalize_internal_link_key(group_key)
                            ] = bookmark_name
                        elements.append(self._make_subheading_paragraph(
                            heading_text, heading_level, heading_style, bookmark_name
                        ))
                    tbl_rows = [header_row] + data_rows if header_row else data_rows
                    link_columns = _internal_link_positions(col_indexes, advanced)
                    tbl_rows, table_merge_positions = _apply_auto_stt(
                        tbl_rows, merge_positions, advanced
                    )
                    tbl = self._build_table_from_cell_rows(
                        tbl_rows, preserve_inline,
                        auto_merge_columns=table_merge_positions,
                        internal_link_columns=link_columns,
                    )
                    if tbl is not None:
                        elements.append(tbl)
                        elements.extend(_lookup_text_paragraphs(group_key, advanced))
                        elements.append(self._create_empty_paragraph())
            else:
                link_columns = _internal_link_positions(col_indexes, advanced)
                table_rows, table_merge_positions = _apply_auto_stt(
                    rows, merge_positions, advanced
                )
                tbl = self._build_table_from_cell_rows(
                    table_rows, preserve_inline,
                    auto_merge_columns=table_merge_positions,
                    internal_link_columns=link_columns,
                )
                if tbl is not None:
                    elements.append(tbl)
                    elements.append(self._create_empty_paragraph())
        else:
            raise ValueError(f'Mode không hợp lệ: {mode}')

        return elements

    def _parse_column_spec(self, spec):
        """Parse 'A,B,D-F,H' → list column indexes [1,2,4,5,6,8]."""
        return parse_excel_column_spec(spec)

    def _extract_rows_from_range(self, ws, range_str, value_ws=None):
        """Trả về list of list of Cell từ ws[range_str]."""
        def wrap(cell):
            if cell is None or value_ws is None:
                return cell
            value_cell = value_ws.cell(row=cell.row, column=cell.column)

            class _RangeCellView:
                def __init__(self, value_cell, style_cell):
                    self._value_cell = value_cell
                    self._style_cell = style_cell
                    self.row = getattr(style_cell, 'row', getattr(value_cell, 'row', None))
                    self.column = getattr(style_cell, 'column', getattr(value_cell, 'column', None))
                    self.parent = getattr(style_cell, 'parent', getattr(value_cell, 'parent', None))

                @property
                def value(self):
                    return getattr(self._value_cell, 'value', None)

                @value.setter
                def value(self, value):
                    self._value_cell.value = value

                def __getattr__(self, name):
                    return getattr(self._style_cell, name)

            return _RangeCellView(value_cell, cell)

        try:
            sel = ws[range_str]
        except Exception as e:
            raise ValueError(f'Range "{range_str}" không hợp lệ: {e}')
        rows = []
        if hasattr(sel, 'value'):
            rows = [[wrap(sel)]]
        elif isinstance(sel, tuple):
            if not sel:
                return []
            if isinstance(sel[0], tuple):
                rows = [[wrap(cell) for cell in r] for r in sel]
            else:
                rows = [[wrap(cell) for cell in sel]]
        return rows

    def _make_subheading_paragraph(self, text, level=4, style_override=None, bookmark_name=''):
        p = etree.Element(w('p'))
        pPr = etree.SubElement(p, w('pPr'))
        pStyle = etree.SubElement(pPr, w('pStyle'))
        try:
            level = min(9, max(1, int(level)))
        except (TypeError, ValueError):
            level = 4
        pStyle.set(w('val'), f'Heading{level}')
        run = etree.SubElement(p, w('r'))
        style_override = style_override if isinstance(style_override, dict) else {}
        if style_override:
            rpr = etree.SubElement(run, w('rPr'))
            font = str(style_override.get('font') or '').strip()
            if font:
                rfonts = ensure_child(rpr, 'rFonts', RPR_ORDER)
                for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    rfonts.set(w(attr), font)
            size_pt = style_override.get('size_pt')
            try:
                size_pt = float(size_pt) if size_pt not in (None, '') else 0
            except (TypeError, ValueError):
                size_pt = 0
            if size_pt > 0:
                size_val = str(pt_to_half_pt(size_pt))
                ensure_child(rpr, 'sz', RPR_ORDER).set(w('val'), size_val)
                ensure_child(rpr, 'szCs', RPR_ORDER).set(w('val'), size_val)
            if style_override.get('bold'):
                ensure_child(rpr, 'b', RPR_ORDER)
                ensure_child(rpr, 'bCs', RPR_ORDER)
            if style_override.get('italic'):
                ensure_child(rpr, 'i', RPR_ORDER)
                ensure_child(rpr, 'iCs', RPR_ORDER)
            if style_override.get('underline'):
                u = ensure_child(rpr, 'u', RPR_ORDER)
                u.set(w('val'), 'single')
            color = str(style_override.get('color') or '').strip().lstrip('#')
            if re.fullmatch(r'[0-9A-Fa-f]{6}', color):
                ensure_child(rpr, 'color', RPR_ORDER).set(w('val'), color.upper())
        t = etree.SubElement(run, w('t'))
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        if bookmark_name:
            bookmark_id = str(abs(hash(bookmark_name)) % 2000000000)
            self._insert_bookmark_start(p, bookmark_id, bookmark_name)
            p.append(self._bookmark_end(bookmark_id))
        return p

    def _advanced_heading_style(self, advanced):
        if not isinstance(advanced, dict):
            return {}

        def as_bool(value, default=False):
            if value is None:
                return default
            if isinstance(value, bool):
                return value
            return str(value).strip().lower() in {'1', 'true', 'yes', 'on', 'y'}

        style = {
            'font': str(advanced.get('heading_font') or '').strip(),
            'size_pt': advanced.get('heading_size_pt') or '',
            'bold': as_bool(advanced.get('heading_bold'), False),
            'italic': as_bool(advanced.get('heading_italic'), False),
            'underline': as_bool(advanced.get('heading_underline'), False),
            'color': str(advanced.get('heading_color') or '').strip(),
        }
        return {k: v for k, v in style.items() if v not in ('', None)}

    def _build_table_from_cell_rows(
            self, rows, preserve_inline, auto_merge_columns=None,
            internal_link_columns=None):
        """Build <w:tbl> từ list of list of openpyxl Cell objects.

        Strategy: tính column widths dựa theo content length của từng cột,
        layout=fixed để Word không redistribute lại width khi render. Tổng width
        bám theo usable width của section chứa heading chèn bảng.
        """
        def _decimal_places_from_number_format(fmt):
            fmt = str(fmt or '').split(';')[0]
            if not fmt or fmt.lower() == 'general':
                return None
            fmt = re.sub(r'"[^"]*"|\\.', '', fmt)
            fmt = re.sub(r'\[[^\]]+\]', '', fmt)
            if '.' not in fmt:
                return 0
            frac = fmt.split('.', 1)[1]
            placeholders = re.match(r'[0#?]+', frac)
            return len(placeholders.group(0)) if placeholders else 0

        def _format_cell_value(cell):
            if cell is None or cell.value is None:
                return ''
            value = cell.value
            if hasattr(value, 'isoformat'):
                value = value.isoformat()
            if isinstance(value, bool):
                return str(value)
            if isinstance(value, int):
                return str(value)
            if isinstance(value, float):
                fmt = getattr(cell, 'number_format', '') or ''
                decimals = _decimal_places_from_number_format(fmt)
                percent_count = fmt.count('%')
                display_value = value * (100 ** percent_count)
                if decimals is not None:
                    text = f'{display_value:.{decimals}f}'
                elif display_value.is_integer():
                    text = str(int(display_value))
                else:
                    text = f'{display_value:.15g}'
                if percent_count:
                    text += '%' * percent_count
                return text
            return str(value)

        def _text_of(cell):
            return _format_cell_value(cell)

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
        merge_info = {}
        position_by_coord = {}
        worksheet = None
        for rr, row in enumerate(rows):
            for cc, cell in enumerate(row):
                if cell is None:
                    continue
                worksheet = getattr(cell, 'parent', worksheet)
                position_by_coord[(getattr(cell, 'row', None), getattr(cell, 'column', None))] = (rr, cc)

        if worksheet is not None and getattr(worksheet, 'merged_cells', None) is not None:
            for merged_range in worksheet.merged_cells.ranges:
                positions = []
                for rr, row in enumerate(rows):
                    for cc, cell in enumerate(row):
                        if cell is None:
                            continue
                        cell_row = getattr(cell, 'row', None)
                        cell_col = getattr(cell, 'column', None)
                        if (merged_range.min_row <= cell_row <= merged_range.max_row and
                                merged_range.min_col <= cell_col <= merged_range.max_col):
                            positions.append((rr, cc))
                if len(positions) <= 1:
                    continue
                row_indexes = sorted({p[0] for p in positions})
                col_indexes = sorted({p[1] for p in positions})
                if (row_indexes != list(range(row_indexes[0], row_indexes[-1] + 1)) or
                        col_indexes != list(range(col_indexes[0], col_indexes[-1] + 1))):
                    continue
                if len(positions) != len(row_indexes) * len(col_indexes):
                    continue

                master_pos = position_by_coord.get((merged_range.min_row, merged_range.min_col))
                master_cell = rows[master_pos[0]][master_pos[1]] if master_pos is not None else None
                top_row, left_col = row_indexes[0], col_indexes[0]
                rowspan, colspan = len(row_indexes), len(col_indexes)
                for rr in row_indexes:
                    for cc in col_indexes:
                        if rr == top_row and cc == left_col:
                            merge_info[(rr, cc)] = {
                                'role': 'restart',
                                'rowspan': rowspan,
                                'colspan': colspan,
                                'master_cell': master_cell,
                            }
                        elif cc == left_col:
                            merge_info[(rr, cc)] = {
                                'role': 'continue',
                                'rowspan': rowspan,
                                'colspan': colspan,
                                'master_cell': master_cell,
                            }
                        else:
                            merge_info[(rr, cc)] = {'role': 'skip'}

        auto_merge_columns = {
            int(c) for c in (auto_merge_columns or set())
            if str(c).strip().isdigit() and 0 <= int(c) < ncols
        }
        if auto_merge_columns and len(rows) > 2:
            for c_idx in sorted(auto_merge_columns):
                r_idx = 1
                while r_idx < len(rows):
                    cell = rows[r_idx][c_idx] if c_idx < len(rows[r_idx]) else None
                    value = _text_of(cell).strip()
                    if not value or (r_idx, c_idx) in merge_info:
                        r_idx += 1
                        continue
                    end_idx = r_idx + 1
                    while end_idx < len(rows):
                        next_cell = rows[end_idx][c_idx] if c_idx < len(rows[end_idx]) else None
                        if _text_of(next_cell).strip() != value or (end_idx, c_idx) in merge_info:
                            break
                        end_idx += 1
                    if end_idx - r_idx > 1:
                        group_positions = [(rr, c_idx) for rr in range(r_idx, end_idx)]
                        if not any(pos in merge_info for pos in group_positions):
                            merge_info[(r_idx, c_idx)] = {
                                'role': 'restart',
                                'rowspan': end_idx - r_idx,
                                'colspan': 1,
                                'master_cell': cell,
                            }
                            for rr in range(r_idx + 1, end_idx):
                                merge_info[(rr, c_idx)] = {
                                    'role': 'continue',
                                    'rowspan': end_idx - r_idx,
                                    'colspan': 1,
                                    'master_cell': cell,
                                }
                    r_idx = max(end_idx, r_idx + 1)

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
                merge = merge_info.get((r_idx, c_idx), {})
                if merge.get('role') == 'skip':
                    continue
                cell = row[c_idx] if c_idx < len(row) else None
                effective_cell = merge.get('master_cell') or cell
                colspan = int(merge.get('colspan') or 1)
                is_vmerge = int(merge.get('rowspan') or 1) > 1
                is_vmerge_continue = merge.get('role') == 'continue'
                tc = etree.SubElement(tr, w('tc'))

                tcPr = etree.SubElement(tc, w('tcPr'))
                tcW = etree.SubElement(tcPr, w('tcW'))
                tcW.set(w('w'), str(sum(col_widths[c_idx:c_idx + colspan])))
                tcW.set(w('type'), 'dxa')
                if colspan > 1:
                    grid_span = etree.SubElement(tcPr, w('gridSpan'))
                    grid_span.set(w('val'), str(colspan))
                if is_vmerge:
                    v_merge = etree.SubElement(tcPr, w('vMerge'))
                    if not is_vmerge_continue:
                        v_merge.set(w('val'), 'restart')

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
                vAlign.set(w('val'), _cell_vertical_alignment(effective_cell, r_idx))

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
                jc.set(w('val'), _cell_horizontal_alignment(effective_cell, r_idx, c_idx))
                if c_idx in no_wrap_cols:
                    wordWrap = ensure_child(pPr, 'wordWrap', PPR_ORDER)
                    wordWrap.set(w('val'), '0')

                # rPr: luôn có (ép font/size)
                rpr = etree.Element(w('rPr'))
                rfonts = ensure_child(rpr, 'rFonts', RPR_ORDER)
                rfonts.set(w('ascii'), ov_font)
                rfonts.set(w('hAnsi'), ov_font)
                rfonts.set(w('eastAsia'), ov_font)
                rfonts.set(w('cs'), ov_font)

                if preserve_inline and effective_cell is not None and effective_cell.font:
                    if effective_cell.font.bold:
                        ensure_child(rpr, 'b', RPR_ORDER)
                    if effective_cell.font.italic:
                        ensure_child(rpr, 'i', RPR_ORDER)

                sz = ensure_child(rpr, 'sz', RPR_ORDER)
                sz.set(w('val'), str(int(round(ov_size * 2))))
                szcs = ensure_child(rpr, 'szCs', RPR_ORDER)
                szcs.set(w('val'), str(int(round(ov_size * 2))))

                if preserve_inline and effective_cell is not None and effective_cell.font:
                    if effective_cell.font.underline and effective_cell.font.underline != 'none':
                        u = ensure_child(rpr, 'u', RPR_ORDER)
                        u.set(w('val'), 'single')

                text_value = '' if is_vmerge_continue else _format_cell_value(effective_cell)
                hyperlink = getattr(effective_cell, 'hyperlink', None) if effective_cell is not None else None
                hyperlink_target = ''
                if hyperlink is not None:
                    hyperlink_target = getattr(hyperlink, 'target', None) or ''
                    if not hyperlink_target and getattr(hyperlink, 'location', None):
                        hyperlink_target = '#' + str(hyperlink.location)

                internal_bookmark = ''
                link_map = (internal_link_columns or {}).get(c_idx) or {}
                if link_map and text_value:
                    internal_bookmark = link_map.get(self._normalize_internal_link_key(text_value), '')

                if internal_bookmark:
                    self._append_internal_hyperlink_run(p, text_value, internal_bookmark, rpr)
                elif self._is_external_web_hyperlink(hyperlink_target):
                    self._append_field_hyperlink_runs(p, text_value, hyperlink_target, rpr)
                else:
                    run = etree.SubElement(p, w('r'))
                    if hyperlink_target:
                        rpr = self._rpr_without_hyperlink_decoration(rpr)
                    run.append(rpr)
                    t = etree.SubElement(run, w('t'))
                    t.text = text_value
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        return tbl

    def _is_external_web_hyperlink(self, target):
        """Only keep real website hyperlinks; workbook-local links become plain text."""
        value = str(target or '').strip()
        if not value:
            return False
        lower = value.lower()
        if lower.startswith(('#', 'sheet', "'")):
            return False
        if lower.startswith(('http://', 'https://')):
            return True
        if lower.startswith('www.'):
            return True
        return False

    def _rpr_without_hyperlink_decoration(self, rpr):
        clean = etree.fromstring(etree.tostring(rpr))
        for tag in ('u', 'color'):
            child = clean.find(w(tag))
            if child is not None:
                clean.remove(child)
        return clean

    def _normalize_internal_link_key(self, value):
        text = normalize_match_text(value)
        text = re.sub(r'^\d+(?:\.\d+)*\.?\s+', '', text)
        text = re.sub(r'^(BANG|TABLE)\s+', '', text)
        return re.sub(r'\s+', ' ', text).strip()

    def _append_internal_hyperlink_run(self, paragraph, display_text, bookmark_name, base_rpr):
        hyperlink = etree.SubElement(paragraph, w('hyperlink'))
        hyperlink.set(w('anchor'), bookmark_name)
        hyperlink.set(w('history'), '1')
        run = etree.SubElement(hyperlink, w('r'))
        link_rpr = etree.fromstring(etree.tostring(base_rpr))
        color = ensure_child(link_rpr, 'color', RPR_ORDER)
        color.set(w('val'), '0563C1')
        underline = ensure_child(link_rpr, 'u', RPR_ORDER)
        underline.set(w('val'), 'single')
        run.append(link_rpr)
        t = etree.SubElement(run, w('t'))
        t.text = display_text
        t.set(XML_SPACE, 'preserve')

    def _append_field_hyperlink_runs(self, paragraph, display_text, target, base_rpr):
        """Append a Word HYPERLINK field without creating document relationships."""
        def add_fld_char(kind):
            run = etree.SubElement(paragraph, w('r'))
            fld = etree.SubElement(run, w('fldChar'))
            fld.set(w('fldCharType'), kind)

        add_fld_char('begin')

        instr_run = etree.SubElement(paragraph, w('r'))
        instr = etree.SubElement(instr_run, w('instrText'))
        instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        safe_target = str(target).replace('\\', '\\\\').replace('"', '\\"')
        instr.text = f' HYPERLINK "{safe_target}" '

        add_fld_char('separate')

        text_run = etree.SubElement(paragraph, w('r'))
        link_rpr = etree.fromstring(etree.tostring(base_rpr))
        color = ensure_child(link_rpr, 'color', RPR_ORDER)
        color.set(w('val'), '0563C1')
        underline = ensure_child(link_rpr, 'u', RPR_ORDER)
        underline.set(w('val'), 'single')
        text_run.append(link_rpr)
        t = etree.SubElement(text_run, w('t'))
        t.text = display_text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        add_fld_char('end')

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
        keywords = {
            'stt', 'ma', 'ten', 'mo ta', 'tan suat', 'loai', 'nguon', 'dich',
            'bang', 'cot', 'field', 'table', 'column', 'name', 'description',
            'key', 'type',
        }

        def normalize_header_text(value):
            text = str(value or '').strip().lower()
            text = ''.join(
                ch for ch in unicodedata.normalize('NFKD', text)
                if not unicodedata.combining(ch)
            )
            return re.sub(r'\s+', ' ', text)

        for r in range(1, max_scan + 1):
            nonempty = 0
            text_cells = 0
            numeric_cells = 0
            short_text_cells = 0
            header_keyword_hits = 0
            for c in range(1, max_col + 1):
                v = ws.cell(row=r, column=c).value
                if v is None or str(v).strip() == '':
                    continue
                nonempty += 1
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    numeric_cells += 1
                else:
                    text_cells += 1
                    text = normalize_header_text(v)
                    if len(text) <= 28:
                        short_text_cells += 1
                    if any(k in text for k in keywords):
                        header_keyword_hits += 1
            if nonempty < 2:
                continue
            score = (
                nonempty * 16
                + text_cells * 8
                + short_text_cells * 5
                + header_keyword_hits * 35
                - numeric_cells * 18
                - r * 2
            )
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

        # Notebook (tabs cho 8 bước)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        self.tab_step1 = ttk.Frame(self.notebook)
        self.tab_step2 = ttk.Frame(self.notebook)
        self.tab_step3 = ttk.Frame(self.notebook)
        self.tab_step4 = ttk.Frame(self.notebook)
        self.tab_step5 = ttk.Frame(self.notebook)
        self.tab_step6 = ttk.Frame(self.notebook)
        self.tab_step7 = ttk.Frame(self.notebook)
        self.tab_step8 = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_step1, text='1. Chọn file & trang')
        self.notebook.add(self.tab_step2, text='2. Defaults & Sections')
        self.notebook.add(self.tab_step3, text='3. Phần mở đầu')
        self.notebook.add(self.tab_step4, text='4. Headings/Numbering/H&F')
        self.notebook.add(self.tab_step5, text='5. Danh sách heading')
        self.notebook.add(self.tab_step6, text='6. Giữ/xoá nội dung')
        self.notebook.add(self.tab_step7, text='7. Mục lục')
        self.notebook.add(self.tab_step8, text='8. Lưu template')

        # Disable các tab sau cho tới khi load file
        for i in range(1, 8):
            self.notebook.tab(i, state='disabled')

        self._build_step1()
        self._build_step2_placeholder()
        self._build_step3_placeholder()
        self._build_step4_placeholder()
        self._build_step5_placeholder()
        self._build_step6_placeholder()
        self._build_step7_placeholder()
        self._build_step8_placeholder()

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
            for i in range(1, 8):
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
        self._build_step8()
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
        lst = self.model.config['headings_list']
        block_end = self._heading_block_end(idx)
        block = lst[idx:block_end]
        new_idx = idx
        if direction < 0:
            prev_start = self._previous_sibling_heading_block_start(idx)
            if prev_start < 0:
                return
            del lst[idx:block_end]
            lst[prev_start:prev_start] = block
            new_idx = prev_start
        else:
            if block_end >= len(lst):
                return
            base_level = int(lst[idx].get('level') or 1)
            next_level = int(lst[block_end].get('level') or 1)
            if next_level < base_level:
                return
            next_end = self._heading_block_end(block_end)
            del lst[idx:block_end]
            insert_at = next_end - len(block)
            lst[insert_at:insert_at] = block
            new_idx = insert_at
        self._refresh_heading_tree()
        self.h_tree.selection_set(str(new_idx))

    def _heading_block_end(self, start_idx):
        lst = self.model.config['headings_list']
        if not (0 <= start_idx < len(lst)):
            return start_idx
        base_level = int(lst[start_idx].get('level') or 1)
        end_idx = start_idx + 1
        while end_idx < len(lst) and int(lst[end_idx].get('level') or 1) > base_level:
            end_idx += 1
        return end_idx

    def _previous_sibling_heading_block_start(self, start_idx):
        lst = self.model.config['headings_list']
        if not (0 <= start_idx < len(lst)):
            return -1
        base_level = int(lst[start_idx].get('level') or 1)
        j = start_idx - 1
        while j >= 0 and int(lst[j].get('level') or 1) > base_level:
            j -= 1
        if j < 0 or int(lst[j].get('level') or 1) < base_level:
            return -1
        return j

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
            if action != 'insert':
                h['insert_source'] = None
                h['insert_selection'] = None

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
            def update_visibility(*_, ac=action_combo, se=source_entry, bb=browse_btn, sv=source_var, idx=i):
                txt = ac.get()
                if txt.startswith('insert'):
                    se.config(state='readonly')
                    bb.config(state='normal')
                else:
                    heading = self.model.config['headings_list'][idx]
                    heading['insert_source'] = None
                    heading['insert_selection'] = None
                    sv.set('')
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
                h['insert_source'] = None
                h['insert_selection'] = None
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
    # STEP 7: Style mục lục
    # -------------------------------------------------------------------

    def _build_step7_placeholder(self):
        ttk.Label(self.tab_step7, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step7(self):
        f = self.tab_step7
        for child in f.winfo_children():
            child.destroy()

        self.model._ensure_toc_settings_defaults()
        cfg = self.model.config['toc_settings']

        ttk.Label(f, text='Bước 7: Cấu hình style mục lục (TOC)',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(10, 5), padx=10)
        ttk.Label(
            f,
            text='Các setting này sẽ ghi vào style TOC1..TOC9. Khi sinh file, tool đánh dấu mục lục để Word cập nhật lại số trang và render theo style mới.',
            foreground='gray',
            wraplength=980,
            justify='left'
        ).pack(anchor='w', padx=10, pady=2)

        top = ttk.LabelFrame(f, text='Thiết lập chung')
        top.pack(fill='x', padx=10, pady=6)

        self.toc_enabled_var = tk.BooleanVar(value=cfg.get('enabled', True))
        self.toc_update_var = tk.BooleanVar(value=cfg.get('update_on_open', True))
        ttk.Checkbutton(top, text='Áp dụng style TOC khi sinh file', variable=self.toc_enabled_var)\
            .grid(row=0, column=0, sticky='w', padx=5, pady=3)
        ttk.Checkbutton(top, text='Cập nhật mục lục khi mở file bằng Word', variable=self.toc_update_var)\
            .grid(row=0, column=1, sticky='w', padx=5, pady=3)

        ttk.Label(top, text='Lấy heading level:').grid(row=1, column=0, sticky='w', padx=5, pady=3)
        self.toc_levels_var = tk.StringVar(value=str(cfg.get('levels', 4)))
        ttk.Combobox(top, textvariable=self.toc_levels_var, values=[str(i) for i in range(1, 10)],
                     width=6, state='readonly').grid(row=1, column=0, sticky='w', padx=(120, 5), pady=3)

        ttk.Label(top, text='Tab leader số trang:').grid(row=1, column=1, sticky='w', padx=5, pady=3)
        self.toc_leader_var = tk.StringVar(value=cfg.get('tab_leader', 'none'))
        ttk.Combobox(top, textvariable=self.toc_leader_var,
                     values=['none', 'dot', 'hyphen', 'underscore'],
                     width=12, state='readonly').grid(row=1, column=1, sticky='w', padx=(140, 5), pady=3)

        ttk.Label(top, text='Right tab (cm):').grid(row=1, column=2, sticky='w', padx=5, pady=3)
        self.toc_right_tab_var = tk.StringVar(value=str(cfg.get('right_tab_cm', 0) or 0))
        ttk.Entry(top, textvariable=self.toc_right_tab_var, width=8)\
            .grid(row=1, column=2, sticky='w', padx=(95, 5), pady=3)
        ttk.Label(top, text='0 = tự tính theo lề trang', foreground='gray')\
            .grid(row=1, column=3, sticky='w', padx=5, pady=3)

        style_frame = ttk.LabelFrame(f, text='TOC styles')
        style_frame.pack(fill='both', expand=True, padx=10, pady=5)

        header = ttk.Frame(style_frame)
        header.pack(fill='x', padx=5, pady=(4, 2))
        cols = [
            ('Style', 8), ('Font', 18), ('Size', 7), ('Bold', 6),
            ('Indent cm', 9), ('Text tab cm', 10),
            ('Before', 7), ('After', 7), ('Line', 7),
        ]
        for text, width in cols:
            ttk.Label(header, text=text, width=width, font=('Segoe UI', 9, 'bold')).pack(side='left', padx=2)

        container = ttk.Frame(style_frame)
        container.pack(fill='both', expand=True, padx=5, pady=2)
        canvas = tk.Canvas(container, height=320)
        scrollbar = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        scrollable = ttk.Frame(canvas)
        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.toc_style_widgets = []
        for sid, data in cfg.get('styles', {}).items():
            row = ttk.Frame(scrollable)
            row.pack(fill='x', padx=2, pady=2)
            ttk.Label(row, text=sid, width=8).pack(side='left', padx=2)

            font_var = tk.StringVar(value=str(data.get('font', 'Times New Roman')))
            size_var = tk.StringVar(value=str(data.get('size_pt', 13)))
            bold_var = tk.BooleanVar(value=bool(data.get('bold', False)))
            left_var = tk.StringVar(value=str(data.get('left_indent_cm', 0)))
            tab_var = tk.StringVar(value=str(data.get('text_tab_cm', 0)))
            before_var = tk.StringVar(value=str(data.get('space_before_pt', 0)))
            after_var = tk.StringVar(value=str(data.get('space_after_pt', 0)))
            line_var = tk.StringVar(value=str(data.get('line_spacing', '1.15')))

            ttk.Entry(row, textvariable=font_var, width=18).pack(side='left', padx=2)
            ttk.Entry(row, textvariable=size_var, width=7).pack(side='left', padx=2)
            ttk.Checkbutton(row, variable=bold_var, width=5).pack(side='left', padx=2)
            ttk.Entry(row, textvariable=left_var, width=9).pack(side='left', padx=2)
            ttk.Entry(row, textvariable=tab_var, width=10).pack(side='left', padx=2)
            ttk.Entry(row, textvariable=before_var, width=7).pack(side='left', padx=2)
            ttk.Entry(row, textvariable=after_var, width=7).pack(side='left', padx=2)
            ttk.Entry(row, textvariable=line_var, width=7).pack(side='left', padx=2)

            self.toc_style_widgets.append({
                'sid': sid,
                'font': font_var,
                'size_pt': size_var,
                'bold': bold_var,
                'left_indent_cm': left_var,
                'text_tab_cm': tab_var,
                'space_before_pt': before_var,
                'space_after_pt': after_var,
                'line_spacing': line_var,
            })

        ttk.Button(f, text='Lưu & sang bước 8 →',
                   command=self._save_step7_and_go).pack(anchor='e', padx=10, pady=10)

    def _save_step7_and_go(self):
        cfg = self.model._ensure_toc_settings_defaults()
        cfg['enabled'] = self.toc_enabled_var.get()
        cfg['update_on_open'] = self.toc_update_var.get()
        try:
            cfg['levels'] = max(1, min(9, int(self.toc_levels_var.get())))
            cfg['right_tab_cm'] = float(self.toc_right_tab_var.get().strip() or 0)
            for row in self.toc_style_widgets:
                sid = row['sid']
                cfg['styles'][sid]['font'] = row['font'].get().strip() or 'Times New Roman'
                cfg['styles'][sid]['size_pt'] = float(row['size_pt'].get())
                cfg['styles'][sid]['bold'] = row['bold'].get()
                cfg['styles'][sid]['left_indent_cm'] = float(row['left_indent_cm'].get())
                cfg['styles'][sid]['text_tab_cm'] = float(row['text_tab_cm'].get())
                cfg['styles'][sid]['space_before_pt'] = float(row['space_before_pt'].get())
                cfg['styles'][sid]['space_after_pt'] = float(row['space_after_pt'].get())
                cfg['styles'][sid]['line_spacing'] = row['line_spacing'].get().strip() or '1.15'
        except ValueError:
            messagebox.showerror('Lỗi', 'Một giá trị TOC không hợp lệ. Vui lòng kiểm tra số pt/cm/line spacing.')
            return
        cfg['tab_leader'] = self.toc_leader_var.get().strip() or 'none'
        self._build_step8()
        self.notebook.select(7)

    # -------------------------------------------------------------------
    # STEP 8: Lưu template
    # -------------------------------------------------------------------

    def _build_step8_placeholder(self):
        ttk.Label(self.tab_step8, text='Hoàn tất bước 1 trước').pack(pady=20)

    def _build_step8(self):
        f = self.tab_step8
        for child in f.winfo_children():
            child.destroy()

        ttk.Label(f, text='Bước 8: Tổng kết & lưu template',
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
        toc = c.get('toc_settings') or {}
        lines.append('=== TOC styles ===')
        lines.append(f'  enabled={toc.get("enabled", True)} '
                     f'update_on_open={toc.get("update_on_open", True)} '
                     f'levels=1-{toc.get("levels", 4)} '
                     f'leader={toc.get("tab_leader", "none")}')
        for sid, data in list((toc.get('styles') or {}).items())[:4]:
            lines.append(f'  {sid}: font={data.get("font")} {data.get("size_pt")}pt '
                         f'bold={data.get("bold")} indent={data.get("left_indent_cm")}cm '
                         f'text_tab={data.get("text_tab_cm")}cm line={data.get("line_spacing")}')
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

class ToolLauncherApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('Chon tool')
        self.root.geometry('720x360')
        self.root.minsize(640, 320)
        self._build()

    def run(self):
        self.root.mainloop()

    def _build(self):
        outer = ttk.Frame(self.root, padding=24)
        outer.pack(fill='both', expand=True)

        ttk.Label(
            outer,
            text='Chon tool can su dung',
            font=('Segoe UI', 16, 'bold')
        ).pack(anchor='w')
        ttk.Label(
            outer,
            text='Moi tool se mo giao dien va luong xu ly rieng.',
            font=('Segoe UI', 10)
        ).pack(anchor='w', pady=(4, 18))

        cards = ttk.Frame(outer)
        cards.pack(fill='both', expand=True)
        cards.columnconfigure(0, weight=1)
        cards.columnconfigure(1, weight=1)

        self._tool_card(
            cards,
            column=0,
            title='DOCX Template Builder',
            desc='Tool hien tai: doc file Word mau, chinh cau hinh va sinh file template .docx.',
            button='Mo tool template',
            command=self._open_template_tool,
        )
        self._tool_card(
            cards,
            column=1,
            title='CSDL -> Excel',
            desc='Mini tool moi: doc mo ta thiet ke CSDL tu file .docx va sinh Excel cau truc.',
            button='Mo tool CSDL',
            command=self._open_csdl_tool,
        )

    def _tool_card(self, parent, column, title, desc, button, command):
        frame = ttk.LabelFrame(parent, text=title, padding=16)
        frame.grid(row=0, column=column, sticky='nsew', padx=(0, 8) if column == 0 else (8, 0))
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text=title, font=('Segoe UI', 12, 'bold')).grid(row=0, column=0, sticky='w')
        ttk.Label(
            frame,
            text=desc,
            wraplength=280,
            justify='left'
        ).grid(row=1, column=0, sticky='nw', pady=(8, 18))
        ttk.Button(frame, text=button, command=command).grid(row=2, column=0, sticky='e')

    def _open_template_tool(self):
        self.root.destroy()
        app = WizardApp()
        app.run()

    def _open_csdl_tool(self):
        try:
            import tk_csdl_tool
        except Exception as exc:
            messagebox.showerror('Loi mo tool CSDL', str(exc))
            return

        self.root.destroy()
        tk_csdl_tool.launch_gui()


def main():
    if tk is None:
        raise RuntimeError(
            'Không tìm thấy Tkinter. Dùng web_app.py cho bản web/server, '
            'hoặc cài Python có Tk để chạy GUI desktop.'
        )
    app = ToolLauncherApp()
    app.run()


if __name__ == '__main__':
    main()
