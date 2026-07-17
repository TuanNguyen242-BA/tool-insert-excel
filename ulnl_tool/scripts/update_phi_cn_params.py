#!/usr/bin/env python3
"""
update_phi_cn_params.py - Cập nhật các tham số nhân trong NL Phi CN
(số module, số lần upcode, số năm bảo hành, số hệ thống tích hợp...).

Cách dùng:
  python3 update_phi_cn_params.py <input_xlsx> <output_xlsx> <params_json>

Format params_json:
  {
    "n_module": 6,         # Số module phần mềm
    "n_upcode": 10,        # Số lần upcode trong phát triển
    "n_bugs": 18,          # Số lỗi sau triển khai
    "n_year": 3,           # Số năm bảo hành
    "n_hrs_bh_month": 60,  # Giờ hỗ trợ phản ánh/tháng bảo hành
    "n_sys": 2,            # Số hệ thống tích hợp
    "pct_qtda": 0.1        # % nỗ lực QTDA
  }

Script sẽ phân tích công thức cũ trong cột E và thay thế các số hằng số
tương ứng. Tham chiếu pattern:
  - QTDA (row 11): ='NL Chuc nang'!E{TR}*{pct_qtda}
  - RV1 (rows 12-17): VLOOKUP * {n_module}*2  (Review tài liệu/module)
  - NT1 (row 18): VLOOKUP * {n_module}        (Nghiệm thu nội bộ)
  - NT2 (row 19): VLOOKUP * {n_module}*3      (Nghiệm thu KH x 3 người)
  - UC3 (row 20): VLOOKUP * {n_upcode}        (Merge code)
  - UC4 (row 21): VLOOKUP * {n_upcode}        (Test merge)
  - RV2 (row 22): VLOOKUP * {n_module}*2      (Review code)
  - VH6 (row 39): ={n_hrs_month_gd1*n_month_gd1*2}*0.5 (phản ánh GĐ1)
  - TH (row 40): =40*{n_sys}                  (test tích hợp)
  - DT6 (row 41): =2*{n_module}               (tài liệu hướng dẫn cài đặt)
  - Section III (rows 48-56): VLOOKUP * {n_bugs} hoặc * {n_year}
"""
import json, sys, re
from openpyxl import load_workbook


def update_params(input_xlsx, output_xlsx, params):
    n_module = params.get('n_module', 6)
    n_upcode = params.get('n_upcode', 10)
    n_bugs = params.get('n_bugs', 18)
    n_year = params.get('n_year', 3)
    n_hrs_bh = params.get('n_hrs_bh_month', 60)
    n_sys = params.get('n_sys', 2)
    pct_qtda = params.get('pct_qtda', 0.1)

    wb = load_workbook(input_xlsx)
    ws = wb['NL Phi CN']

    # Section I rows
    # Row 11: QTDA - ='NL Chuc nang'!E{TR}*{pct}
    v = ws.cell(11, 5).value
    if v:
        new_v = re.sub(r'\*[\d.]+\s*$', f'*{pct_qtda}', v)
        ws.cell(11, 5, new_v)

    # Pattern: VLOOKUP(...) * <số_module> * <số_lặp>  → cập nhật N_MODULE
    # Các row dùng N_MODULE: 12-17 (RV1), 18 (NT1), 19 (NT2 * 3), 22 (RV2 * 2)
    # Rows 13-16: *N_MODULE*2; Row 17: *N_MODULE; Row 18: *N_MODULE; Row 19: *N_MODULE*3; Row 22: *N_MODULE*2
    for r in [13, 14, 15, 16]:
        v = ws.cell(r, 5).value
        if v and 'VLOOKUP' in v:
            ws.cell(r, 5, re.sub(r'\)\*\d+\*2\s*$', f')*{n_module}*2', v))
    # Row 17 (RV1 HDSD): *N_MODULE chỉ
    v = ws.cell(17, 5).value
    if v and 'VLOOKUP' in v:
        ws.cell(17, 5, re.sub(r'\)\*\d+\s*$', f')*{n_module}', v))
    # Row 18 (NT1)
    v = ws.cell(18, 5).value
    if v and 'VLOOKUP' in v:
        ws.cell(18, 5, re.sub(r'\)\*\d+\s*$', f')*{n_module}', v))
    # Row 19 (NT2 * 3)
    v = ws.cell(19, 5).value
    if v and 'VLOOKUP' in v:
        ws.cell(19, 5, re.sub(r'\)\*\d+\*3\s*$', f')*{n_module}*3', v))
    # Row 20 (UC3 - merge code)
    v = ws.cell(20, 5).value
    if v and 'VLOOKUP' in v:
        ws.cell(20, 5, re.sub(r'\)\*\d+\s*$', f')*{n_upcode}', v))
    # Row 21 (UC4 - test merge)
    v = ws.cell(21, 5).value
    if v and 'VLOOKUP' in v:
        ws.cell(21, 5, re.sub(r'\)\*\d+\s*$', f')*{n_upcode}', v))
    # Row 22 (RV2 - review code * N_MODULE * 2)
    v = ws.cell(22, 5).value
    if v and 'VLOOKUP' in v:
        ws.cell(22, 5, re.sub(r'\)\*\d+\*2\s*$', f')*{n_module}*2', v))

    # Row 39 (VH6 GĐ1): =N_HRS_MONTH * N_MONTH * 2 * N_PCT_GD1
    # (đơn giản: giữ pattern *0.5)
    # Row 40 (TH): =40*N_SYS
    v = ws.cell(40, 5).value
    if v:
        ws.cell(40, 5, f'=40*{n_sys}')
    # Row 41 (DT6): =2*N_MODULE (template gốc dùng +1 module riêng cho DB)
    v = ws.cell(41, 5).value
    if v:
        ws.cell(41, 5, f'=2*{n_module + 1}')

    # Section III - rows 48-56
    # Row 47 (header): "Bảo hành và hỗ trợ vận hành <N_YEAR> năm"
    ws.cell(47, 3, f'Bảo hành và hỗ trợ vận hành {n_year} năm')

    # Row 48 (VH6 BH): ={n_hrs_bh}*12*{n_year}
    ws.cell(48, 5, f'={n_hrs_bh}*12*{n_year}')

    # Row 49 (VH7), 50 (VH8), 51 (UC3), 52 (UC4), 53 (VH1), 54 (VH2), 56 (UC6): VLOOKUP * N_BUGS
    for r in [49, 50, 51, 52, 53, 54, 56]:
        v = ws.cell(r, 5).value
        if v and 'VLOOKUP' in v:
            ws.cell(r, 5, re.sub(r'\)\*\d+\s*$', f')*{n_bugs}', v))
    # Row 55 (UC5): VLOOKUP * N_BUGS * 3
    v = ws.cell(55, 5).value
    if v and 'VLOOKUP' in v:
        ws.cell(55, 5, re.sub(r'\)\*\d+\*3\s*$', f')*{n_bugs}*3', v))

    wb.save(output_xlsx)
    print(f"OK: NL Phi CN đã update với n_module={n_module}, n_upcode={n_upcode}, "
          f"n_bugs={n_bugs}, n_year={n_year}, n_sys={n_sys}, n_hrs_bh={n_hrs_bh}, pct_qtda={pct_qtda}")


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[3], 'r', encoding='utf-8') as f:
        params = json.load(f)
    update_params(sys.argv[1], sys.argv[2], params)
