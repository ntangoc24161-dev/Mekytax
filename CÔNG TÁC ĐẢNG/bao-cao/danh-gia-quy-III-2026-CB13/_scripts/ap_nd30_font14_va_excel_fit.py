# -*- coding: utf-8 -*-
"""Word: TNR 14, le ND 30/2020, khong FitToPages.
Excel: in A4, Fit all columns on one page. Khong goi table.columns.
"""
from __future__ import print_function

import os
import shutil

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

VAULT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LAPTOP = (
    r"H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG"
    r"\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026"
)
TNR = "Times New Roman"
A4_W, A4_H = "11906", "16838"
# ND 30: top/bottom 20mm, left 30mm, right 15mm (twips)
MAR = dict(top="1134", bottom="1134", left="1701", right="851",
           header="709", footer="709", gutter="0")


def set_rfonts(rPr):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)


def set_run_size(r, pt):
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r.insert(0, rPr)
    set_rfonts(rPr)
    val = str(int(pt * 2))
    for tag in ("sz", "szCs"):
        el = rPr.find(qn("w:" + tag))
        if el is None:
            el = OxmlElement("w:" + tag)
            rPr.append(el)
        el.set(qn("w:val"), val)


def set_p_runs(p_el, pt):
    runs = p_el.findall(qn("w:r"))
    if not runs:
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        r.append(rPr)
        set_run_size(r, pt)
        p_el.append(r)
        return
    for r in runs:
        set_run_size(r, pt)


def set_p_spacing(p_el, after=120, before=60, line=276):
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    sp.set(qn("w:after"), str(after))
    sp.set(qn("w:before"), str(before))
    sp.set(qn("w:line"), str(line))
    sp.set(qn("w:lineRule"), "auto")


def set_a4(section, landscape):
    sectPr = section._sectPr
    pgSz = sectPr.find(qn("w:pgSz"))
    if pgSz is None:
        pgSz = OxmlElement("w:pgSz")
        sectPr.append(pgSz)
    if landscape:
        pgSz.set(qn("w:w"), A4_H)
        pgSz.set(qn("w:h"), A4_W)
        pgSz.set(qn("w:orient"), "landscape")
    else:
        pgSz.set(qn("w:w"), A4_W)
        pgSz.set(qn("w:h"), A4_H)
        if qn("w:orient") in pgSz.attrib:
            del pgSz.attrib[qn("w:orient")]
    pgMar = sectPr.find(qn("w:pgMar"))
    if pgMar is None:
        pgMar = OxmlElement("w:pgMar")
        sectPr.append(pgMar)
    for k, v in MAR.items():
        pgMar.set(qn("w:" + k), v)


def force_defaults(doc):
    st = doc.styles["Normal"]
    st.font.name = TNR
    st.font.size = Pt(14)
    rPr = st.element.get_or_add_rPr()
    set_rfonts(rPr)
    for tag, val in (("sz", "28"), ("szCs", "28")):
        el = rPr.find(qn("w:" + tag))
        if el is None:
            el = OxmlElement("w:" + tag)
            rPr.append(el)
        el.set(qn("w:val"), val)
    styles = doc.styles.element
    dd = styles.find(qn("w:docDefaults"))
    if dd is None:
        dd = OxmlElement("w:docDefaults")
        styles.insert(0, dd)
    rPrDef = dd.find(qn("w:rPrDefault"))
    if rPrDef is None:
        rPrDef = OxmlElement("w:rPrDefault")
        dd.insert(0, rPrDef)
    rPr = rPrDef.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rPrDef.append(rPr)
    set_rfonts(rPr)
    for tag, val in (("sz", "28"), ("szCs", "28")):
        el = rPr.find(qn("w:" + tag))
        if el is None:
            el = OxmlElement("w:" + tag)
            rPr.append(el)
        el.set(qn("w:val"), val)


def p_text(p_el):
    return "".join((t.text or "") for t in p_el.findall(".//" + qn("w:t")))


def ensure_sign_space(last_tbl, keep=5):
    trs = last_tbl.findall(qn("w:tr"))
    if not trs:
        return
    tcs = trs[0].findall(qn("w:tc"))
    if len(tcs) < 2:
        return
    cell = tcs[-1]
    paras = cell.findall(qn("w:p"))
    empty = 0
    for p in paras:
        if not p_text(p).strip():
            empty += 1
    need = keep - empty
    for _ in range(max(0, need)):
        p = OxmlElement("w:p")
        set_p_spacing(p, after=0, before=0, line=276)
        set_p_runs(p, 14)
        cell.append(p)


def patch_word(path, landscape):
    d = Document(path)
    force_defaults(d)
    set_a4(d.sections[0], landscape)
    body = d.element.body
    tbls = body.findall(qn("w:tbl"))
    header_tbl = tbls[0] if tbls else None
    last_tbl = tbls[-1] if tbls else None

    for p in body.findall(qn("w:p")):
        txt = p_text(p).strip()
        set_p_runs(p, 14)
        if not txt:
            set_p_spacing(p, after=0, before=0, line=276)
        else:
            set_p_spacing(p, after=120, before=60, line=276)

    for ti, tbl in enumerate(tbls):
        trs = tbl.findall(qn("w:tr"))
        for ri, tr in enumerate(trs):
            for ci, tc in enumerate(tr.findall(qn("w:tc"))):
                if tbl is header_tbl:
                    pt = 13 if ri < 2 else 14
                elif tbl is last_tbl and ci == 0:
                    pt = 12
                else:
                    pt = 14
                for p in tc.findall(qn("w:p")):
                    set_p_runs(p, pt)
                    set_p_spacing(p, after=0, before=0, line=240)

    if last_tbl is not None and last_tbl is not header_tbl:
        ensure_sign_space(last_tbl, keep=5)
    d.save(path)
    print("word", os.path.basename(path), os.path.getsize(path))


def used_range(ws):
    max_r, max_c = 1, 1
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 200),
                            max_col=min(ws.max_column or 1, 40)):
        for c in row:
            if c.value not in (None, ""):
                max_r = max(max_r, c.row)
                max_c = max(max_c, c.column)
    return max_r, max_c


def excel_fit(ws, landscape):
    from openpyxl.utils import get_column_letter
    max_r, max_c = used_range(ws)
    ws.print_area = "A1:%s%d" % (get_column_letter(max_c), max_r)
    ws.page_setup.paperSize = 9  # A4 trong openpyxl
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=1.181, right=0.591, top=0.787, bottom=0.787,
        header=0.3, footer=0.3,
    )
    ws.print_options.horizontalCentered = True
    ws.sheet_view.view = "pageBreakPreview"
    return max_r, max_c


def patch_excel(path, print_sheets):
    wb = load_workbook(path)
    info = []
    for ws in wb.worksheets:
        landscape = True
        if ws.title in print_sheets:
            landscape = print_sheets[ws.title]
        r, c = excel_fit(ws, landscape)
        info.append("%s %dx%d fitW=1" % (ws.title, r, c))
    wb.save(path)
    print("xlsx", os.path.basename(path), "; ".join(info))


def copy_laptop(rel):
    src = os.path.join(VAULT, rel)
    dst = os.path.join(LAPTOP, rel)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)


def main():
    words = [
        ("01_To_trinh_de_xuat_xep_loai.docx", False),
        ("04_Quyet_dinh_phan_cong_va_danh_muc_san_pham.docx", False),
        ("07_Bao_cao_ket_qua_QIII_2026.docx", False),
    ]
    kh = os.path.join(VAULT, "05_Ke_hoach_ca_nhan")
    for fn in sorted(os.listdir(kh)):
        if fn.lower().endswith(".docx"):
            words.append((os.path.join("05_Ke_hoach_ca_nhan", fn), True))
    for rel, land in words:
        patch_word(os.path.join(VAULT, rel), land)
        copy_laptop(rel)

    xls = [
        ("02_Danh_sach_trich_ngang.xlsx", {"Trich ngang": True}),
        ("03_Bang_tong_hop_de_xuat.xlsx", {"Tong hop de xuat": True, "Tinh diem KPI": True}),
        ("08_Thong_ke_cap_uy_vien.xlsx", {"Thong ke": True}),
    ]
    tdg = os.path.join(VAULT, "06_Ban_tu_danh_gia")
    for fn in sorted(os.listdir(tdg)):
        if fn.lower().endswith(".xlsx"):
            xls.append((os.path.join("06_Ban_tu_danh_gia", fn),
                        {"Khung đánh giá": True, "Tinh KPI 2A-2B": True}))
    for rel, sheets in xls:
        patch_excel(os.path.join(VAULT, rel), sheets)
        copy_laptop(rel)


if __name__ == "__main__":
    main()
