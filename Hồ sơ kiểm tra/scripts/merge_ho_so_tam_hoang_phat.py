# -*- coding: utf-8 -*-
"""Tang 1 - Merge ho so Tam Hoang Phat (MST 3603721832) truc tiep tu sheet
01_DA_BAN_HANH (dong 15) vao phoi templates/MAUTONG HOP.docx.

Tien chu tinh bang scripts/doc_so.py, KHONG doc cot chu tren so (cot 53-58 bi
mat dau tieng Viet do ham VBA DocSo luu theo code page ANSI).
"""
import os
import sys
import io
import shutil
import datetime
import openpyxl
import docx

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from field_merger import merge_fields_in_element
from doc_so import doc_so

BRAIN = r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra'
TEMPLATE = os.path.join(BRAIN, 'templates', 'MAUTONG HOP.docx')
XLSM = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
EXPORT_DIR = os.path.join(BRAIN, 'exports', 'tam-hoang-phat-3603721832')
G_DIR = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\10. TAM HOANG PHAT\0. BIEN BAN'
ROW = 15
STAMP = datetime.datetime.now().strftime('%Y%m%d-%H%M')

CHU_MAP = {
    'Tong_truy_thu_chu': 'Tong_truy_thu',
    'Tong_cham_nop_chu': 'Tong_cham_nop',
    'Tong_tien_phat_chu': 'Tong_tien_phat',
    'Tong_phat_sinh_chu': 'Tong_phat_sinh',
    'Giam_lo_chu': 'Giam_lo',
    'Giam_khau_tru_chu': 'Giam_khau_tru',
}

log = io.open(os.path.join(HERE, 'log_merge_thp.txt'), 'w', encoding='utf-8')
def P(s):
    log.write(s + u"\n")

def main():
    os.makedirs(EXPORT_DIR, exist_ok=True)
    P(u"1. Doc so 01_DA_BAN_HANH dong %d" % ROW)
    wb = openpyxl.load_workbook(XLSM, data_only=True)
    ws = wb['01_DA_BAN_HANH']
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    vals = [ws.cell(row=ROW, column=c).value for c in range(1, ws.max_column + 1)]
    data = {}
    for h, v in zip(headers, vals):
        if h:
            data[h] = v
    # Tinh lai toan bo tien bang chu tu con so goc
    for k, src in CHU_MAP.items():
        data[k] = doc_so(data.get(src))
        P(u"   %s = %s" % (k, data[k]))
    P(u"   MST=%s | QDKT=%s | Ngay_bao_cao=%s | Ngay_ky_BBKT=%s | Thoi_han=%s"
      % (data.get('MST'), data.get('So_QDKT'), data.get('Ngay_bao_cao'),
         data.get('Ngay_ky_BBKT'), data.get('Thoi_han_KT')))
    P(u"   So truong nap: %d" % len(data))

    P(u"2. Nap phoi: %s" % TEMPLATE)
    doc = docx.Document(TEMPLATE)

    P(u"3. Tron du lieu")
    n = 0
    for p in doc.paragraphs:
        n += merge_fields_in_element(p._p, data)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    n += merge_fields_in_element(p._p, data)
    P(u"   Da merge %d truong" % n)

    left = []
    for instr in doc.element.body.iter():
        if instr.tag.endswith('instrText') and instr.text and 'MERGEFIELD' in instr.text:
            left.append(instr.text.strip())
    P(u"   MERGEFIELD con sot: %d %s" % (len(left), left[:10]))

    out_named = os.path.join(EXPORT_DIR, 'TONG HOP - Tam Hoang Phat.docx')
    out_main = os.path.join(EXPORT_DIR, 'TONG HOP.docx')
    for f in (out_named, out_main):
        if os.path.exists(f):
            shutil.copyfile(f, f.replace('.docx', '.bak-%s.docx' % STAMP))
    doc.save(out_named)
    doc.save(out_main)
    P(u"   Luu Brain: %s (%d bytes)" % (out_main, os.path.getsize(out_main)))

    if os.path.isdir(G_DIR):
        out_g = os.path.join(G_DIR, 'TONG HOP.docx')
        if os.path.exists(out_g):
            bak = os.path.join(G_DIR, 'TONG HOP.bak-truoc-merge-%s.docx' % STAMP)
            shutil.copyfile(out_g, bak)
            P(u"   Sao luu o G: %s" % bak)
        doc.save(out_g)
        P(u"   Luu o G: %s (%d bytes)" % (out_g, os.path.getsize(out_g)))
    P(u"=== XONG TANG 1 ===")
    log.close()

if __name__ == '__main__':
    main()
