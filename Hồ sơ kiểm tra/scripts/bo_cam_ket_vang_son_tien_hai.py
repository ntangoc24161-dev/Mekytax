#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bo van ban 'Cam ket hoat dong kinh doanh vang' con sot tu phoi DNTN Thuy Hong
khoi TONG HOP.docx cua vu Son Tien Hai (MST 3603859012).

Van ban nay khong ap duoc cho Son Tien Hai (ban buon vat lieu xay dung) va tu
mau thuan: muc 1 cam ket khong mua ban vang nhung muc 4 lai ghi chi mua vang
tu nguoi dan. Xoa tu paragraph ngat trang (833) den het van ban (859),
giu lai sectPr.
"""
import shutil
import sys
from docx import Document
from docx.text.paragraph import Paragraph

SRC = '/brains/Hồ sơ kiểm tra/exports/son-tien-hai-3603859012/TONG HOP.docx'
DAU, CUOI = 833, 859
MOC_DAU = 'V/v: Cam kết hoạt động kinh doanh vàng'
MOC_CUOI = 'Phan Thế Hanh'


def main():
    shutil.copy2(SRC, SRC + '.bak-truoc-bo-cam-ket-vang')
    doc = Document(SRC)
    kids = list(doc.element.body.iterchildren())
    # kiem tra dung khoi can xoa truoc khi cat
    assert MOC_DAU in Paragraph(kids[836], doc).text, Paragraph(kids[836], doc).text
    assert MOC_CUOI in Paragraph(kids[858], doc).text, Paragraph(kids[858], doc).text
    assert kids[860].tag.endswith('}sectPr'), kids[860].tag
    for i in range(DAU, CUOI + 1):
        kids[i].getparent().remove(kids[i])
    con = list(doc.element.body.iterchildren())
    print('da xoa', CUOI - DAU + 1, 'phan tu; body con', len(con))
    print('ket thuc o:', repr(Paragraph(con[-2], doc).text[:60]))
    doc.save(SRC)
    print('da luu', SRC)


if __name__ == '__main__':
    sys.exit(main())
