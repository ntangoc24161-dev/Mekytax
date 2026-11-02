#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dien bien ban + ho so ve cho Cong ty TNHH Thuong mai San xuat Son Tien Hai (MST 3603859012).

Phoi: TONG HOP.docx tren Drive (clone tu mau DNTN Thuy Hong) - giu nguyen layout,
chi sua khoi nam (2017-2024 -> 2022-2026), do so lieu tu Bang tinh (sheet DTCHEP)
va sua cac truong con sot lai cua Thuy Hong.
"""
import copy
import shutil
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC = '/brains/Hồ sơ kiểm tra/exports/son-tien-hai-3603859012/TONG HOP.docx'

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# ---------------------------------------------------------------- so lieu vu
def vn(v, zero='0'):
    """Dinh dang so kieu Viet Nam; so am boc trong ngoac."""
    if v is None:
        return '......'
    if isinstance(v, str):
        return v
    v = int(round(v))
    if v == 0:
        return zero
    s = f'{abs(v):,}'.replace(',', '.')
    return f'({s})' if v < 0 else s


# BBXN thanh vien: Doanh thu ban hang va CCDV / Thue GTGT phai nop
BBXN_TV = {
    2022: [(1356457502, 1356457502, 0), (0, 0, 0)],
    2023: [(2550607100, 2550607100, 0), (0, 0, 0)],
    2024: [(4888008900, 4888008900, 0), (0, 0, 0)],
    2025: [(5832577595, 5832577595, 0), (40051890, 40051890, 0)],
    2026: [(997858720, 997858720, 0), (0, 0, 0)],
}

# BBXN truong doan: Gia von hang ban / Chi phi quan ly kinh doanh
BBXN_TD = {
    2022: [(1200449571, 1200449571, 0), (81544372, 81544372, 0)],
    2023: [(2701731176, 2701731176, 0), (203824754, 203824754, 0)],
    2024: [(4335139322, 4335139322, 0), (201343790, 201343790, 0)],
    2025: [(4910021268, 4850137853, -59883415), (320094772, 320094772, 0)],
    2026: [(1563849838, 737085429, -826764409), (101032554, 101032554, 0)],
}

# BBKT: thue GTGT phai nop, doanh thu, gia von, CP QLKD, TNCT, thue suat, thue TNDN
BBKT = {
    2022: dict(gtgt=(0, 0, 0), dt=(1356457502, 1356457502, 0),
               gv=(1200449571, 1200449571, 0), qlkd=(81544372, 81544372, 0),
               tnct=(-8345760, -8345760, 0), ts=('20%', '20%'), tndn=(0, 0, 0)),
    2023: dict(gtgt=(0, 0, 0), dt=(2550607100, 2550607100, 0),
               gv=(2701731176, 2701731176, 0), qlkd=(203824754, 203824754, 0),
               tnct=(-675305024, -675305024, 0), ts=('20%', '20%'), tndn=(0, 0, 0)),
    2024: dict(gtgt=(0, 0, 0), dt=(4888008900, 4888008900, 0),
               gv=(4335139322, 4335139322, 0), qlkd=(201343790, 201343790, 0),
               tnct=(741609928, 741609928, 0), ts=('20%', '20%'),
               tndn=(11591829, 11591829, 0)),
    2025: dict(gtgt=(40051890, 40051890, 0), dt=(5832577595, 5832577595, 0),
               gv=(4910021268, 4850137853, -59883415), qlkd=(320094772, 320094772, 0),
               tnct=(34203203, 94086618, 59883415), ts=('17%', '17%'),
               tndn=(5814545, 15994725, 10180180)),
    2026: dict(gtgt=(0, 0, 0), dt=(997858720, 997858720, 0),
               gv=(1563849838, 737085429, -826764409), qlkd=(101032554, 101032554, 0),
               tnct=(-809868861, 16895548, 826764409), ts=('17%', '17%'),
               tndn=(0, 2872243, 2872243)),
}

CAN_CU_2025 = 'Căn cứ Điều 10 Nghị định số 320/2025/NĐ-CP của Chính phủ.'

# Giai trinh BBXN truong doan cho nam co chenh lech
BBXN_TD_GT = {
    2025: ['Giải trình số liệu chênh lệch qua kiểm tra:',
           '*/ Giá vốn hàng bán giảm 59.883.415 đồng, do Công ty hạch toán giá vốn '
           'hàng bán không tương ứng doanh thu. ' + CAN_CU_2025],
    2026: ['Giải trình số liệu chênh lệch qua kiểm tra:',
           '*/ Giá vốn hàng bán giảm 826.764.409 đồng, do Công ty hạch toán giá vốn '
           'hàng bán không tương ứng doanh thu. ' + CAN_CU_2025],
}

# Giai trinh BBKT: danh sach dong thay cho "Khong chenh lech." duoi "+ Thue TNDN:"
BBKT_GT_TNDN = {
    2025: ['Công ty báo cáo thu nhập chịu thuế 34.203.203 đồng, Đoàn kiểm tra xác định '
           'thu nhập chịu thuế 94.086.618 đồng, tăng 59.883.415 đồng, nguyên nhân:',
           '*/ Giá vốn hàng bán giảm 59.883.415 đồng, do Công ty hạch toán giá vốn hàng '
           'bán không tương ứng doanh thu. ' + CAN_CU_2025,
           '=>Thuế TNDN truy thu qua kiểm tra: 10.180.180 đồng.'],
    2026: ['Công ty báo cáo lỗ 809.868.861 đồng, Đoàn kiểm tra xác định thu nhập chịu '
           'thuế 16.895.548 đồng, tăng 826.764.409 đồng, nguyên nhân:',
           '*/ Giá vốn hàng bán giảm 826.764.409 đồng, do Công ty hạch toán giá vốn hàng '
           'bán không tương ứng doanh thu. ' + CAN_CU_2025,
           '=>Thuế TNDN truy thu qua kiểm tra: 2.872.243 đồng.',
           'Giảm lỗ: 809.868.861 đồng.'],
}

# Muc III - chi dien o chac chan, o phu thuoc QHNS 970 de "......"
MUC_III_GTGT = {  # (3) ky truoc, (4) bao cao phat sinh, (5) xu ly qua KT, (6) da nop, (7) con phai nop
    2022: ('......', '0', '0', '......', '......'),
    2023: ('......', '0', '0', '......', '......'),
    2024: ('......', '0', '0', '......', '......'),
    2025: ('......', '40.051.890', '0', '......', '......'),
    2026: ('......', '0', '0', '......', '......'),
}
MUC_III_TNDN = {
    2022: ('......', '0', '0', '......', '......'),
    2023: ('......', '0', '0', '......', '......'),
    2024: ('......', '11.591.829', '0', '......', '......'),
    2025: ('......', '5.814.545', '10.180.180', '......', '......'),
    2026: ('......', '0', '2.872.243', '......', '......'),
}


# ---------------------------------------------------------------- helper docx
def set_text(p, text):
    """Dat text vao paragraph, giu dinh dang run dau tien."""
    runs = p.runs
    if not runs:
        p.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def set_cell(cell, text):
    p = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    set_text(p, text)


def clone_after(p, text):
    """Nhan ban paragraph p ngay sau p, doi text. Tra ve Paragraph moi."""
    new = copy.deepcopy(p._element)
    p._element.addnext(new)
    np = Paragraph(new, p._parent)
    set_text(np, text)
    return np


def drop(el):
    el.getparent().remove(el)


def main():
    shutil.copy2(SRC, SRC + '.bak-truoc-dien-bien-ban')
    doc = Document(SRC)
    body = doc.element.body
    kids = list(body.iterchildren())

    def rng(a, b):
        """Xoa cac child tu index a den b (inclusive) theo snapshot kids."""
        for i in range(a, b + 1):
            drop(kids[i])

    def par(i):
        return Paragraph(kids[i], doc)

    def tbl(i):
        return Table(kids[i], doc)

    # ---------------- 1. BBXN thanh vien: giu 5 khoi dau, xoa 3 khoi cuoi
    tv_blocks = [(218, 221), (222, 225), (226, 229), (230, 233), (234, 237)]
    for (y, (a, b)) in zip([2022, 2023, 2024, 2025, 2026], tv_blocks):
        set_text(par(a), f'+Năm {y}:')
        t = tbl(a + 2)
        for ri, vals in enumerate(BBXN_TV[y], start=1):
            for ci, v in enumerate(vals, start=2):
                set_cell(t.rows[ri].cells[ci], vn(v))
        set_text(par(b), '*/ Không chênh lệch.')
    rng(238, 251)

    # ---------------- 2. BBXN truong doan: giu 2017,2018,2019,2021,2024
    td_blocks = [(2022, 279, 282, None), (2023, 283, 286, None), (2024, 287, 290, None),
                 (2025, 295, 299, 'diff'), (2026, 308, 312, 'diff')]
    for (y, a, b, kind) in td_blocks:
        set_text(par(a), f'+Năm {y}:')
        t = tbl(a + 2)
        for ri, vals in enumerate(BBXN_TD[y], start=1):
            for ci, v in enumerate(vals, start=2):
                set_cell(t.rows[ri].cells[ci], vn(v))
        if kind is None:
            set_text(par(b), '*/ Không chênh lệch.')
        else:
            lines = BBXN_TD_GT[y]
            # khoi 2025: a+3 = 'Giai trinh...', a+4 = '*/ ...'
            # khoi 2026: a+3 = 'Giai trinh...', a+4 = '*/ ...'
            set_text(par(a + 3), lines[0])
            set_text(par(a + 4), lines[1])
            for extra in range(a + 5, b + 1):
                drop(kids[extra])
    rng(291, 294)
    rng(300, 307)

    # ---------------- 3. BBKT: giu 2017,2018,2019,2020,2022
    bbkt_blocks = [(2022, 348), (2023, 356), (2024, 364), (2025, 372), (2026, 391)]
    for (y, a) in bbkt_blocks:
        set_text(par(a), f'* NĂM {y}')
        t = tbl(a + 2)
        d = BBKT[y]
        rowmap = [(2, d['gtgt']), (4, d['dt']), (5, d['gv']), (6, d['qlkd']),
                  (7, d['tnct']), (9, d['tndn'])]
        for ri, vals in rowmap:
            for ci, v in enumerate(vals, start=2):
                set_cell(t.rows[ri].cells[ci], vn(v))
        set_cell(t.rows[8].cells[2], d['ts'][0])
        set_cell(t.rows[8].cells[3], d['ts'][1])
        # giai trinh: a+3 Giai trinh / a+4 '+ Thue GTGT:' / a+5 'Khong chenh lech.'
        #             a+6 '+ Thue TNDN:' / a+7 'Khong chenh lech.'
        set_text(par(a + 5), 'Không chênh lệch.')
        if y in BBKT_GT_TNDN:
            lines = BBKT_GT_TNDN[y]
            set_text(par(a + 7), lines[0])
            anchor = par(a + 7)
            for extra in lines[1:]:
                anchor = clone_after(anchor, extra)
        else:
            set_text(par(a + 7), 'Không chênh lệch.')
    rng(380, 390)
    rng(399, 406)
    rng(407, 419)

    # ---------------- 4. Muc III: dung lai khung 2022-2026
    t3 = tbl(421)
    # rows: 0,1,2 header | 3 'Thue GTGT' | 4..13 nam 2017-2026 | 14 'Thue TNDN' | 15..24 nam
    for y, vals in zip(range(2022, 2027), [MUC_III_GTGT[y] for y in range(2022, 2027)]):
        ri = 4 + (y - 2022)
        set_cell(t3.rows[ri].cells[1], str(y))
        for ci, v in enumerate(vals, start=2):
            set_cell(t3.rows[ri].cells[ci], v)
    for y in range(2022, 2027):
        ri = 15 + (y - 2022)
        set_cell(t3.rows[ri].cells[1], str(y))
        for ci, v in enumerate(MUC_III_TNDN[y], start=2):
            set_cell(t3.rows[ri].cells[ci], v)
    # xoa 5 dong du moi phan (dong 9..13 cua GTGT, 20..24 cua TNDN) - xoa tu duoi len
    for ri in sorted([20, 21, 22, 23, 24, 9, 10, 11, 12, 13], reverse=True):
        drop(t3.rows[ri]._element)

    doc.save(SRC)
    print('da luu', SRC)


if __name__ == '__main__':
    sys.exit(main())
