# -*- coding: utf-8 -*-
"""Patch gen_xlsx.py: dua toan bo file Excel ve Times New Roman + bo sung
khoi the thuc 'dang bo / chi bo ban hanh' cho bieu 03 va bieu 08."""
import io

p = 'gen_xlsx.py'
s = io.open(p, encoding='utf-8').read()


def rep(old, new, n=1):
    global s
    assert s.count(old) == n, 'KHONG KHOP: ' + old[:70]
    s = s.replace(old, new, n)


# 1. Ham dat font Times New Roman cho ca workbook
rep('''OUT = sys.argv[1]''',
    '''OUT = sys.argv[1]
TNR = "Times New Roman"


def set_tnr(wb, bo_sheet_rong=True):
    """Dat font Times New Roman cho MOI o co du lieu + style Normal cua workbook.

    Giu nguyen bold/italic/size/mau da gan truoc do, chi doi ten font.
    """
    try:
        wb._named_styles["Normal"].font.name = TNR
    except Exception:
        pass
    for ws in list(wb.worksheets):
        if bo_sheet_rong and ws.max_row == 1 and ws.max_column == 1 \\
                and ws.cell(1, 1).value is None and len(wb.worksheets) > 1:
            wb.remove(ws)
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and c.font.name in (None, TNR):
                    continue
                f = c.font
                c.font = Font(name=TNR, size=f.size, bold=f.bold, italic=f.italic,
                              underline=f.underline, color=f.color,
                              vertAlign=f.vertAlign, strike=f.strike)
    return wb''')

# 2. Bieu 03: bo sung dong dang bo tren chi bo + goi set_tnr truoc khi luu
rep('''    ws["A1"] = CHI_BO; ws["A1"].font = Font(bold=True, size=11)
    ws["E1"] = "ĐẢNG CỘNG SẢN VIỆT NAM"; ws["E1"].font = Font(bold=True, size=11)''',
    '''    ws["A1"] = DANG_BO; ws["A1"].font = Font(bold=True, size=11)
    ws["A2"] = CHI_BO; ws["A2"].font = Font(bold=True, size=11)
    ws["E1"] = "ĐẢNG CỘNG SẢN VIỆT NAM"; ws["E1"].font = Font(bold=True, size=11)''')

rep('''    wb.save(os.path.join(OUT, "03_Bang_tong_hop_de_xuat.xlsx"))''',
    '''    set_tnr(wb)
    wb.save(os.path.join(OUT, "03_Bang_tong_hop_de_xuat.xlsx"))''')

# 3. Bieu 08: tach khoi the thuc thanh 2 dong chuan
rep('''    ws["A1"] = "CHI BỘ 13 - ĐẢNG BỘ THUẾ THÀNH PHỐ ĐỒNG NAI"; ws["A1"].font = Font(bold=True, size=11)''',
    '''    ws["A1"] = DANG_BO; ws["A1"].font = Font(bold=True, size=11)
    ws["A2"] = CHI_BO; ws["A2"].font = Font(bold=True, size=11)''')

rep('''    wb.save(os.path.join(OUT, "08_Thong_ke_cap_uy_vien.xlsx"))''',
    '''    set_tnr(wb)
    wb.save(os.path.join(OUT, "08_Thong_ke_cap_uy_vien.xlsx"))''')

# 4. Bieu 02 va 7 ban tu danh gia
rep('''    wb.save(os.path.join(OUT, "02_Danh_sach_trich_ngang.xlsx"))''',
    '''    set_tnr(wb)
    wb.save(os.path.join(OUT, "02_Danh_sach_trich_ngang.xlsx"))''')

rep('''        fn = "TDG_QIII-2026_%d_%s.xlsx" % (n["stt"], slugify(n["ten_hoa_thuong"]))
        wb.save(os.path.join(sub, fn))''',
    '''        fn = "TDG_QIII-2026_%d_%s.xlsx" % (n["stt"], slugify(n["ten_hoa_thuong"]))
        set_tnr(wb)
        wb.save(os.path.join(sub, fn))''')

io.open(p, 'w', encoding='utf-8').write(s)
print('gen_xlsx.py: Times New Roman + the thuc OK')
