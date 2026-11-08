# -*- coding: utf-8 -*-
"""Gan header 3 dong (mau file 04) cho 9 file Word con lai, A4, in 1 trang.
Khong goi table.columns. So van ban de trong, ngay Tran Bien.
"""
from __future__ import print_function

import os
import shutil
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Twips

VAULT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LAPTOP = (
    r"H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG"
    r"\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026"
)
MAU_04 = os.path.join(VAULT, "04_Quyet_dinh_phan_cong_va_danh_muc_san_pham.docx")

# A4 twips
A4_W = "11906"
A4_H = "16838"

JOBS = [
    ("01_To_trinh_de_xuat_xep_loai.docx", False, "Số:        -TTr/CB"),
    ("07_Bao_cao_ket_qua_QIII_2026.docx", False, "Số:        -BC/CB"),
]


def kh_jobs():
    sub = os.path.join(VAULT, "05_Ke_hoach_ca_nhan")
    out = []
    for fn in sorted(os.listdir(sub)):
        if fn.lower().endswith(".docx"):
            out.append((os.path.join("05_Ke_hoach_ca_nhan", fn), True, "Số:"))
    return out


def set_t_keep_first(p_el, text):
    texts = p_el.findall(".//" + qn("w:t"))
    if not texts:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        p_el.append(r)
        return
    texts[0].text = text
    texts[0].set(qn("xml:space"), "preserve")
    for t in texts[1:]:
        t.text = ""


def clone_header_tbl(mau_doc, so_text):
    tbl = deepcopy(mau_doc.tables[0]._tbl)
    trs = tbl.findall(qn("w:tr"))
    if len(trs) < 3:
        raise SystemExit("mau header khong du 3 dong")
    tcs = trs[2].findall(qn("w:tc"))
    p0 = tcs[0].find(qn("w:p"))
    set_t_keep_first(p0, so_text)
    p1 = tcs[1].find(qn("w:p"))
    texts = p1.findall(".//" + qn("w:t"))
    joined = "".join((t.text or "") for t in texts)
    if "Trấn Biên" not in joined:
        set_t_keep_first(p1, "Trấn Biên, ngày ..... tháng ..... năm 2026")
    return tbl


def replace_first_table(doc, new_tbl):
    body = doc.element.body
    old = body.find(qn("w:tbl"))
    if old is None:
        body.insert(0, new_tbl)
        return
    old.addprevious(new_tbl)
    body.remove(old)


def scale_tbl_widths(tbl, new_total):
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        return
    cols = [c for c in list(grid) if c.tag == qn("w:gridCol")]
    if not cols:
        return
    olds = [int(c.get(qn("w:w")) or "0") for c in cols]
    s = sum(olds) or 1
    news = [max(200, int(round(w * new_total / float(s)))) for w in olds]
    news[-1] += new_total - sum(news)
    if news[-1] < 200:
        news[-1] = 200
    for c, w in zip(cols, news):
        c.set(qn("w:w"), str(w))
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(new_total))
    tblW.set(qn("w:type"), "dxa")
    for tr in tbl.findall(qn("w:tr")):
        tcs = tr.findall(qn("w:tc"))
        for tc, w in zip(tcs, news):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(w))
            tcW.set(qn("w:type"), "dxa")


def set_cell_margins(tbl, twips=40):
    for tc in tbl.iter(qn("w:tc")):
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        mar = tcPr.find(qn("w:tcMar"))
        if mar is None:
            mar = OxmlElement("w:tcMar")
            tcPr.append(mar)
        for edge in ("top", "left", "bottom", "right"):
            el = mar.find(qn("w:" + edge))
            if el is None:
                el = OxmlElement("w:" + edge)
                mar.append(el)
            el.set(qn("w:w"), str(twips))
            el.set(qn("w:type"), "dxa")


def set_a4(section, landscape, tight=False):
    sectPr = section._sectPr
    pgSz = sectPr.find(qn("w:pgSz"))
    if pgSz is None:
        pgSz = OxmlElement("w:pgSz")
        sectPr.append(pgSz)
    if landscape:
        pgSz.set(qn("w:w"), A4_H)
        pgSz.set(qn("w:h"), A4_W)
        pgSz.set(qn("w:orient"), "landscape")
        if tight:
            t, b, l, r = 567, 567, 851, 567  # 10/10/15/10 mm
        else:
            t, b, l, r = 680, 680, 1134, 680  # 12/12/20/12 mm
    else:
        pgSz.set(qn("w:w"), A4_W)
        pgSz.set(qn("w:h"), A4_H)
        if qn("w:orient") in pgSz.attrib:
            del pgSz.attrib[qn("w:orient")]
        if tight:
            t, b, l, r = 567, 567, 1134, 567  # 10/10/20/10 mm
        else:
            t, b, l, r = 850, 850, 1418, 850  # 15/15/25/15 mm
    pgMar = sectPr.find(qn("w:pgMar"))
    if pgMar is None:
        pgMar = OxmlElement("w:pgMar")
        sectPr.append(pgMar)
    pgMar.set(qn("w:top"), str(t))
    pgMar.set(qn("w:bottom"), str(b))
    pgMar.set(qn("w:left"), str(l))
    pgMar.set(qn("w:right"), str(r))
    pgMar.set(qn("w:header"), "284")
    pgMar.set(qn("w:footer"), "284")
    pgMar.set(qn("w:gutter"), "0")
    return int(pgSz.get(qn("w:w"))) - l - r


def tighten_paragraphs(doc, body_pt=None):
    for p in doc.paragraphs:
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        txt = (p.text or "").strip()
        if not txt:
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            continue
        pf.space_after = Pt(2)
        pf.line_spacing = 1.0
        if body_pt is None:
            continue
        for r in p.runs:
            r.font.size = Pt(body_pt)


def tighten_table_paras(table, font_pt=None):
    for row in table.rows:
        seen = set()
        for cell in row.cells:
            tc = cell._tc
            if id(tc) in seen:
                continue
            seen.add(id(tc))
            for p in cell.paragraphs:
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1.0
                if font_pt is None:
                    continue
                for r in p.runs:
                    r.font.size = Pt(font_pt)


def trim_empty_sign_paras(table, keep=2):
    """Giu toi da `keep` doan trong lien tiep o o chu ky."""
    for row in table.rows:
        seen = set()
        for cell in row.cells:
            tc = cell._tc
            if id(tc) in seen:
                continue
            seen.add(id(tc))
            paras = cell.paragraphs
            empty_run = 0
            to_clear = []
            for p in paras:
                if (p.text or "").strip():
                    empty_run = 0
                    continue
                empty_run += 1
                if empty_run > keep:
                    to_clear.append(p)
            for p in to_clear:
                p._element.getparent().remove(p._element)


def process_file(rel, landscape, so_text, mau, tight=False, body_pt=None, tbl_pt=None):
    path = os.path.join(VAULT, rel)
    d = Document(path)
    header = clone_header_tbl(mau, so_text)
    replace_first_table(d, header)
    content_w = set_a4(d.sections[0], landscape, tight=tight)
    scale_tbl_widths(d.tables[0]._tbl, content_w)
    tighten_paragraphs(d, body_pt=body_pt)
    for i, t in enumerate(d.tables):
        if i == 0:
            continue
        scale_tbl_widths(t._tbl, content_w)
        set_cell_margins(t._tbl, 40 if landscape else 50)
        tighten_table_paras(t, font_pt=tbl_pt)
        if i == len(d.tables) - 1:
            trim_empty_sign_paras(t, keep=2)
    d.save(path)
    return path


def word_fit_one_page(paths):
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    results = []
    try:
        for path in paths:
            doc = word.Documents.Open(path, ReadOnly=False)
            try:
                ps = doc.PageSetup
                ps.PaperSize = 7  # wdPaperA4 (9 la A5)
                if "Ke_hoach" in path.replace("/", "\\"):
                    ps.Orientation = 1
                    ps.PageWidth = 841.89
                    ps.PageHeight = 595.28
                else:
                    ps.Orientation = 0
                    ps.PageWidth = 595.28
                    ps.PageHeight = 841.89
                doc.Repaginate()
                before = int(doc.ComputeStatistics(2))
                n = before
                tries = 0
                while n > 1 and tries < 6:
                    try:
                        doc.FitToPages()
                    except Exception as e:
                        results.append((os.path.basename(path), before, n, "fit_err " + str(e)))
                        break
                    doc.Repaginate()
                    n = int(doc.ComputeStatistics(2))
                    tries += 1
                doc.Save()
                results.append((os.path.basename(path), before, n, "tries " + str(tries)))
            finally:
                doc.Close(False)
    finally:
        word.Quit()
    return results


def copy_laptop(rels):
    for rel in rels:
        src = os.path.join(VAULT, rel)
        dst = os.path.join(LAPTOP, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        print("laptop", rel, os.path.getsize(src), os.path.getsize(dst))


def dump_headers(rels):
    for rel in rels:
        path = os.path.join(VAULT, rel)
        d = Document(path)
        t0 = d.tables[0]
        print("==", rel, "T0 rows", len(t0.rows), "size", os.path.getsize(path))
        for ri, row in enumerate(t0.rows[:3]):
            cells = []
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                cells.append(" / ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())[:80])
            print("  R%d:" % ri, cells)


def main():
    jobs = JOBS + kh_jobs()
    mau = Document(MAU_04)
    rels = []
    for rel, landscape, so_text in jobs:
        body_pt = 11 if not landscape else 10
        tbl_pt = 9 if not landscape else 8
        process_file(rel, landscape, so_text, mau, tight=False, body_pt=body_pt, tbl_pt=tbl_pt)
        rels.append(rel)
        print("patched", rel)
    paths = [os.path.join(VAULT, r) for r in rels]
    print("WORD FIT...")
    results = word_fit_one_page(paths)
    leftover = [r for r in results if r[2] > 1]
    if leftover:
        print("STILL >1, tighter pass", leftover)
        mau = Document(MAU_04)
        left_rels = []
        for rel, landscape, so_text in jobs:
            bn = os.path.basename(rel)
            if not any(x[0] == bn and x[2] > 1 for x in leftover):
                continue
            process_file(
                rel, landscape, so_text, mau, tight=True,
                body_pt=10 if not landscape else 8,
                tbl_pt=8 if not landscape else 7,
            )
            left_rels.append(rel)
        if left_rels:
            more = word_fit_one_page([os.path.join(VAULT, r) for r in left_rels])
            results = [r for r in results if r[0] not in {x[0] for x in more}] + more
    dump_headers(rels)
    copy_laptop(rels)
    print("RESULTS")
    for r in results:
        print(" ", r)


if __name__ == "__main__":
    main()
