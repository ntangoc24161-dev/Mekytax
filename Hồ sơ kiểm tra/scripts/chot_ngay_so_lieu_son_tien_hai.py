#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pass 3 vu Son Tien Hai (MST 3603859012): chot moc ngay theo so
'Danh sach kiem tra quyet toan pro.xlsm' (ban Drive 29/10/2026), tinh lai tien
cham nop theo ngay ky BBKT 04/11/2026, bo phan Thue GTGT khoi Muc III,
va dong bo QD XPHC tach rieng.

Moc ngay: cong bo/bat dau 30/10/2026, ket thuc 03/11/2026,
ky BBKT 04/11/2026, bao cao 09/11/2026.
So QDKT/QDGS/QDXPHC va ngay ban hanh QD van de trong theo yeu cau.
"""
import shutil
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

BASE = '/brains/Hồ sơ kiểm tra/exports/son-tien-hai-3603859012'
TONG_HOP = f'{BASE}/TONG HOP.docx'
QD_XPHC = f'{BASE}/QD XPHC Cong ty Son Tien Hai.docx'

PH = '..../..../2026'
NGAY_CONG_BO = '30/10/2026'
NGAY_KET_THUC = '03/11/2026'
NGAY_BBKT = '04/11/2026'

# ---- so lieu tinh lai theo ngay ky BBKT 04/11/2026
# cham nop TNDN 2025 = 10.180.180 x (04/11/2026 - 01/04/2026 = 217 ngay) x 0,03%
CHAM_NOP_CU, CHAM_NOP = '387.865', '662.730'
TONG_PHAT_CU, TONG_PHAT = '2.998.350', '3.273.215'
TONG_CONG_CU, TONG_CONG = '16.050.773', '16.325.638'
TONG_CONG_CHU_CU = ('Mười sáu triệu không trăm năm mươi nghìn bảy trăm '
                    'bảy mươi ba đồng')
TONG_CONG_CHU = ('Mười sáu triệu ba trăm hai mươi lăm nghìn sáu trăm '
                 'ba mươi tám đồng')

# (chi so paragraph trong body, so thu tu lan xuat hien cua PH, ngay thay vao)
NGAY_THEO_INDEX = [
    (159, 1, NGAY_CONG_BO),    # BB cong bo QD kiem tra
    (207, 1, NGAY_KET_THUC),   # BBXN so lieu - thanh vien
    (254, 1, NGAY_KET_THUC),   # BBXN so lieu - truong doan
    (301, 1, NGAY_BBKT),       # BB kiem tra (giai the)
    (370, 1, NGAY_BBKT),       # Muc III: tinh den ngay
    (419, 1, NGAY_KET_THUC),   # BB cong khai du thao: ket thuc kiem tra
    (420, 1, NGAY_KET_THUC),   # BB cong khai du thao: ngay lap
    (446, 2, NGAY_CONG_BO),    # Bao cao KQ: tu ngay
    (446, 3, NGAY_KET_THUC),   # Bao cao KQ: den ngay
    (447, 1, NGAY_BBKT),       # Bao cao KQ: can cu BBKT ky ngay
    (459, 1, NGAY_CONG_BO),    # Bao cao KQ: cong bo QD kiem tra
    (460, 1, NGAY_CONG_BO),    # Bao cao KQ: thoi gian bat dau
    (461, 1, NGAY_KET_THUC),   # Bao cao KQ: thoi gian ket thuc
    (546, 2, NGAY_CONG_BO),    # Bao cao giam sat: tu ngay
    (546, 3, NGAY_KET_THUC),   # Bao cao giam sat: den ngay
    (626, 1, NGAY_CONG_BO),    # BB phan cong cong viec
    (658, 1, NGAY_BBKT),       # QD XPHC: can cu BBKT ky ngay
    (689, 1, NGAY_BBKT),       # QD XPHC: cham nop tinh den het ngay
    (689, 2, NGAY_BBKT),       # QD XPHC: tu tinh ke tu sau ngay
    (713, 1, NGAY_BBKT),       # To trinh QD XPHC: BBKT ngay
]

LUAT_CU = 'Căn cứ Luật Quản lý thuế ngày 13/6/2019 và các văn bản hướng dẫn thi hành;'
LUAT_MOI = ('Căn cứ Luật Quản lý thuế số 108/2025/QH15 ngày 10 tháng 12 năm 2025 '
            'và các văn bản hướng dẫn thi hành;')

SUBS_TONG_HOP = [
    # co quan cap DKKD: So Tai chinh Thanh pho Dong Nai
    ('do Sở Kế hoạch và Đầu tư Thành phố Đồng Nai cấp',
     'do Sở Tài chính Thành phố Đồng Nai cấp'),
    ('do Phòng đăng ký kinh doanh – Sở kế hoạch và đầu tư Thành phố Đồng Nai '
     'cấp ngày ..../..../........',
     'do Sở Tài chính Thành phố Đồng Nai cấp ngày 09/03/2022'),
    ('Doanh nghiệp Chúng tôi được thành lập từ ..../..../........,',
     'Doanh nghiệp Chúng tôi được thành lập từ ngày 09/03/2022,'),
    # can cu luat quan ly thue trong BB cong bo QD kiem tra
    (LUAT_CU, LUAT_MOI),
    # xu ly hoa don dien tu khi giai the
    ('và đã thực hiện ...... theo quy định.',
     'và đã thực hiện thủ tục ngừng sử dụng hóa đơn điện tử theo quy định.'),
    # tien cham nop tinh lai theo ngay ky BBKT
    (f'với số tiền: {CHAM_NOP_CU} đồng.', f'với số tiền: {CHAM_NOP} đồng.'),
    (f'Tổng cộng số tiền phạt: {TONG_PHAT_CU} đồng', f'Tổng cộng số tiền phạt: {TONG_PHAT} đồng'),
    (f'(TM 4918), số tiền: {CHAM_NOP_CU} đồng;', f'(TM 4918), số tiền: {CHAM_NOP} đồng;'),
    (f'Tổng số tiền truy thu, phạt: {TONG_CONG_CU} đồng (bằng chữ: {TONG_CONG_CHU_CU}).',
     f'Tổng số tiền truy thu, phạt: {TONG_CONG} đồng (bằng chữ: {TONG_CONG_CHU}).'),
]

SUBS_QD = [
    (f'(TM 4918), số tiền: {CHAM_NOP_CU} đồng;', f'(TM 4918), số tiền: {CHAM_NOP} đồng;'),
    (f'Tổng số tiền truy thu, phạt: {TONG_CONG_CU} đồng (bằng chữ: {TONG_CONG_CHU_CU}).',
     f'Tổng số tiền truy thu, phạt: {TONG_CONG} đồng (bằng chữ: {TONG_CONG_CHU}).'),
]

# QD XPHC tach rieng: (chi so paragraph, lan xuat hien, ngay)
NGAY_QD = [(13, 1, NGAY_BBKT), (44, 1, NGAY_BBKT), (44, 2, NGAY_BBKT)]


def rewrite(p, old, new, nth=None):
    """Thay old -> new trong paragraph. nth=None: thay het; nth=k: chi lan thu k."""
    full = p.text
    if old not in full:
        return 0
    if nth is None:
        out = full.replace(old, new)
        n = full.count(old)
    else:
        parts = full.split(old)
        if len(parts) <= nth:
            return 0
        out = old.join(parts[:nth]) + new + old.join(parts[nth:])
        n = 1
    runs = p.runs
    if not runs:
        p.add_run(out)
        return n
    runs[0].text = out
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    return n


def iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                yield p


def apply_subs(doc, subs):
    counts = {old: 0 for old, _ in subs}
    for p in list(iter_paragraphs(doc)):
        for old, new in subs:
            counts[old] += rewrite(p, old, new)
    return counts


def apply_ngay(doc, spec):
    kids = list(doc.element.body.iterchildren())
    # xu ly tu lan xuat hien lon ve nho de khong lech vi tri
    for idx, nth, ngay in sorted(spec, key=lambda x: (x[0], -x[1])):
        p = Paragraph(kids[idx], doc)
        n = rewrite(p, PH, ngay, nth=nth)
        flag = 'OK ' if n else '!! '
        print(f'{flag}p{idx} lan {nth} -> {ngay} | {p.text[:60]}')


def main():
    # ---------------- TONG HOP
    shutil.copy2(TONG_HOP, TONG_HOP + '.bak-truoc-chot-ngay')
    doc = Document(TONG_HOP)
    apply_ngay(doc, NGAY_THEO_INDEX)
    for old, n in apply_subs(doc, SUBS_TONG_HOP).items():
        print(f'{"OK " if n else "!! "}{n}x  {old[:62]}')
    # Muc III: chi the hien phan Thue TNDN -> bo nhan 'Thue GTGT' va 5 dong nam
    kids = list(doc.element.body.iterchildren())
    t3 = Table(kids[368], doc)
    assert t3.rows[3].cells[1].text.strip() == 'Thuế GTGT', t3.rows[3].cells[1].text
    assert t3.rows[9].cells[1].text.strip() == 'Thuế TNDN', t3.rows[9].cells[1].text
    for ri in range(8, 2, -1):
        t3.rows[ri]._element.getparent().remove(t3.rows[ri]._element)
    print('Muc III con', len(t3.rows), 'dong (chi Thue TNDN)')
    doc.save(TONG_HOP)
    print('da luu', TONG_HOP)

    # ---------------- QD XPHC tach rieng
    shutil.copy2(QD_XPHC, QD_XPHC + '.bak-truoc-chot-ngay')
    qd = Document(QD_XPHC)
    apply_ngay(qd, NGAY_QD)
    for old, n in apply_subs(qd, SUBS_QD).items():
        print(f'{"OK " if n else "!! "}{n}x  {old[:62]}')
    qd.save(QD_XPHC)
    print('da luu', QD_XPHC)


if __name__ == '__main__':
    sys.exit(main())
