#!/usr/bin/env python3
"""
update_tong_hop.py - Cập nhật công thức tham chiếu trong sheet 'Tong hop'
sau khi xác định row tổng kết của 'NL Chuc nang' và 'NL Phi CN'.

Theo cấu trúc DTXD:
- Section I (rows 7-14): 7 dòng + row tổng
- Section II (rows 19-26): 6 data rows + row tổng MD + row tổng MM

Cách dùng:
  python3 update_tong_hop.py <input_xlsx> <output_xlsx> <total_row_nl_cn> <total_row_nl_pcn> [unit_test_row]

  total_row_nl_cn: row 'Nỗ lực (MH)' tổng general trong NL Chuc nang
  total_row_nl_pcn: row 'Nỗ lực (MH)' tổng trong NL Phi CN
  unit_test_row: row riêng cho Unit test (mặc định = total_row_nl_cn - 1)
"""
import sys, re
from openpyxl import load_workbook


def update_tong_hop(input_xlsx, output_xlsx, tr_cn, tr_pcn, ut_row=None):
    if ut_row is None:
        ut_row = tr_cn - 1

    wb = load_workbook(input_xlsx)
    ws = wb['Tong hop']

    # ==== Section I (rows 7-13) ====
    # Row 7: Tổng nỗ lực chức năng (general + UT)
    ws.cell(7, 5, f"='NL Chuc nang'!N{tr_cn}/8+'NL Chuc nang'!N{ut_row}/8")
    ws.cell(7, 6, f"=ROUNDUP('NL Chuc nang'!E{tr_cn}/8+'NL Chuc nang'!E{ut_row}/8,2)")
    # Row 8: Giải pháp
    ws.cell(8, 5, f"='NL Chuc nang'!K{tr_cn}/8")
    ws.cell(8, 6, f"=ROUNDUP('NL Chuc nang'!F{tr_cn}/8,2)")
    # Row 9: Phát triển
    ws.cell(9, 5, f"=ROUND('NL Chuc nang'!L{tr_cn}/8,2)")
    ws.cell(9, 6, f"='NL Chuc nang'!G{tr_cn}/8")
    # Row 10: Unit test
    ws.cell(10, 5, f"=ROUNDDOWN('NL Chuc nang'!L{ut_row}/8,2)")
    ws.cell(10, 6, f"=ROUND('NL Chuc nang'!G{ut_row}/8,2)")
    # Row 11: Kiểm thử
    ws.cell(11, 5, f"='NL Chuc nang'!M{tr_cn}/8")
    ws.cell(11, 6, f"=ROUNDDOWN('NL Chuc nang'!H{tr_cn}/8,2)")
    # Row 12: Phi chức năng
    ws.cell(12, 5, f"=ROUND('NL Phi CN'!E{tr_pcn}/8,2)")
    ws.cell(12, 6, f"=ROUNDUP('NL Phi CN'!E{tr_pcn}/8,2)")
    # Row 13: NL QTDL - giữ 0 (user cập nhật nếu có sheet QTDL)
    # Row 14: Tổng - giữ nguyên formula

    # ==== Section II (rows 19-26) ====
    # Row 19: Giải pháp
    ws.cell(19, 6, f"=ROUND('NL Chuc nang'!W{tr_cn}/8,2)")
    ws.cell(19, 7, f"=ROUND('NL Chuc nang'!Z{tr_cn}/8,2)")
    # Row 20: Phát triển
    ws.cell(20, 6, f"=ROUND('NL Chuc nang'!X{tr_cn}/8,2)")
    ws.cell(20, 7, f"='NL Chuc nang'!AA{tr_cn}/8")
    # Row 21: Unit test
    ws.cell(21, 6, f"=ROUND('NL Chuc nang'!X{ut_row}/8,2)")
    ws.cell(21, 7, f"=ROUND('NL Chuc nang'!AA{ut_row}/8,2)")
    # Row 22: Kiểm thử
    ws.cell(22, 6, f"=ROUNDDOWN('NL Chuc nang'!Y{tr_cn}/8,2)")
    ws.cell(22, 7, f"=ROUNDUP('NL Chuc nang'!AB{tr_cn}/8,2)")
    # Row 23: Phi chức năng
    ws.cell(23, 6, f"='NL Phi CN'!G{tr_pcn}/8")
    ws.cell(23, 7, f"=ROUND('NL Phi CN'!H{tr_pcn}/8,2)")
    # Row 24: NL QTDL - giữ 0
    # Row 25, 26: tổng - giữ nguyên SUM

    # ==== Cập nhật reference trong NL Phi CN dòng QTDA ====
    ws_pcn = wb['NL Phi CN']
    for r in range(11, 60):
        v = ws_pcn.cell(r, 5).value
        if v and isinstance(v, str) and "'NL Chuc nang'!E" in v:
            new_v = re.sub(r"'NL Chuc nang'!E\d+", f"'NL Chuc nang'!E{tr_cn}", v)
            ws_pcn.cell(r, 5, new_v)
            print(f"  Updated NL Phi CN!E{r}: {new_v}")

    # Update D6 trong NL Phi CN (Tổng MH)
    ws_pcn['D6'] = f'=E{tr_pcn}'

    wb.save(output_xlsx)
    print(f"OK: Tong hop tham chiếu NL Chuc nang!row={tr_cn} (UT={ut_row}), NL Phi CN!row={tr_pcn}")


if __name__ == '__main__':
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)
    ut_row = int(sys.argv[5]) if len(sys.argv) > 5 else None
    update_tong_hop(sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]), ut_row)
