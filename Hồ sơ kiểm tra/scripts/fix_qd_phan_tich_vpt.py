# -*- coding: utf-8 -*-
"""Fix VPT: blank QD dates, proper-case company, phan tich layout, signature gaps."""
from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Twips

BB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
)
LIVE = BB / "TONG HOP.docx"
QDKT = BB / "QD Kiem tra tai cty Van Phat Thanh.docx"
QDGS = BB / "QDgiam sat tai cty Van Phat Thanh.docx"
ZIPP = BB / "hoso.zip"

OLD_CO = "CÔNG TY TNHH MỘT THÀNH VIÊN BAO BÌ GIẤY VẠN PHÁT THÀNH"
NEW_CO = "Công ty TNHH Một Thành Viên Bao Bì Giấy Vạn Phát Thành"
DATE_FILLED = re.compile(
    r"Trấn Biên,\s*ngày\s*\d{1,2}\s*tháng\s*\d{1,2}\s*năm\s*\d{4}",
    re.IGNORECASE,
)
DATE_BLANK = "Trấn Biên, ngày .... tháng .... năm 2026"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def set_run_font(run, size_pt=None, bold=None, italic=None):
    run.font.name = "Times New Roman"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(a), "Times New Roman")
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def replace_text_in_paragraph(p, old, new):
    if old not in (p.text or ""):
        return 0
    # simple path: single-run or rebuild
    full = p.text
    if old not in full:
        return 0
    new_full = full.replace(old, new)
    if len(p.runs) == 1:
        p.runs[0].text = new_full
        return 1
    # multi-run: replace across by rewriting first run and clearing others
    bold = p.runs[0].bold if p.runs else None
    size = p.runs[0].font.size if p.runs else None
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(new_full)
    set_run_font(run, size_pt=(size.pt if size else 13), bold=bold)
    return 1


def replace_in_doc_paragraphs(doc, old, new, predicate=None):
    n = 0
    for p in doc.paragraphs:
        if predicate and not predicate(p):
            continue
        n += replace_text_in_paragraph(p, old, new)
    return n


def replace_in_tables(doc, old, new, cell_predicate=None):
    n = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                if cell_predicate and not cell_predicate(cell):
                    continue
                for p in cell.paragraphs:
                    n += replace_text_in_paragraph(p, old, new)
    return n


def blank_qd_dates_in_tables(doc):
    """Blank issuance date only in QD header tables (have Số: .../QĐ-)."""
    n = 0
    for tbl in doc.tables:
        joined = "\n".join(c.text for row in tbl.rows for c in row.cells)
        if "QĐ-" not in joined and "QĐ–" not in joined:
            # also match QD header with Số:
            if not re.search(r"Số:\s*\d+/QĐ", joined):
                continue
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = p.text or ""
                    if DATE_FILLED.search(t) or (
                        t.strip().startswith("Trấn Biên") and "ngày" in t and "2026" in t
                    ):
                        new_t = DATE_BLANK
                        for r in list(p.runs):
                            r._element.getparent().remove(r._element)
                        run = p.add_run(new_t)
                        set_run_font(run, size_pt=13, italic=True, bold=False)
                        n += 1
    return n


def set_doc_lap_size_in_table(tbl, size_pt=14):
    n = 0
    for row in tbl.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if "Độc lập" not in (p.text or "") and "Độc" not in (p.text or ""):
                    continue
                for r in p.runs:
                    if any(k in (r.text or "") for k in ("Độc", "lập", "Tự", "do", "Hạnh", "phúc", "-")) or (r.text or "").strip() == "":
                        # set all runs in this para that are part of slogan
                        pass
                    set_run_font(r, size_pt=size_pt, bold=True)
                    n += 1
    return n


def find_table_before_phan_tich(doc):
    body = list(doc.element.body)
    para_i = 0
    last_tbl = None
    last_tbl_el = None
    for el in body:
        tag = el.tag.split("}")[-1]
        if tag == "p":
            texts = "".join(t.text or "" for t in el.findall(".//%st" % W))
            if texts.strip().startswith("PHÂN TÍCH THÔNG TIN"):
                return last_tbl_el, last_tbl, para_i
            para_i += 1
        elif tag == "tbl":
            last_tbl_el = el
            # map to python-docx table by xml element
            for ti, tbl in enumerate(doc.tables):
                if tbl._tbl is el:
                    last_tbl = tbl
                    break
    return None, None, None


def clear_spacing(p, ls=1.0, sb=0, sa=0):
    pf = p.paragraph_format
    pf.line_spacing = ls
    pf.space_before = Pt(sb)
    pf.space_after = Pt(sa)


def set_center_tab(p, pos="6663"):
    pPr = p._element.get_or_add_pPr()
    old = pPr.find(qn("w:tabs"))
    if old is not None:
        pPr.remove(old)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "center")
    tab.set(qn("w:pos"), str(pos))
    tabs.append(tab)
    pPr.append(tabs)


def style_right_date(p, text, bold=False):
    clear_spacing(p, 1.0, 0, 0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    set_center_tab(p)
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run("\t" + text.lstrip("\t"))
    set_run_font(run, size_pt=13, bold=bold, italic=True)


def ensure_empty_between(doc, title_text, name_text, keep=4, search_end=None):
    """Ensure exactly `keep` empty paras between title and name; insert/remove as needed."""
    ti = ni = None
    for i, p in enumerate(doc.paragraphs):
        if search_end is not None and i >= search_end:
            break
        if p.text.strip() == title_text:
            ti = i
        if ti is not None and p.text.strip() == name_text and i > ti:
            ni = i
            break
    if ti is None or ni is None:
        return "miss"
    # count empties
    empties = []
    el = doc.paragraphs[ti]._element.getnext()
    name_el = doc.paragraphs[ni]._element
    while el is not None and el is not name_el:
        if el.tag == qn("w:p"):
            texts = "".join(node.text or "" for node in el.iter(qn("w:t"))).strip()
            if texts == "":
                empties.append(el)
            else:
                break
        el = el.getnext()
    cur = len(empties)
    if cur > keep:
        for e in empties[keep:]:
            e.getparent().remove(e)
        return "removed %d (had %d)" % (cur - keep, cur)
    if cur < keep:
        for _ in range(keep - cur):
            empty = OxmlElement("w:p")
            pPr = OxmlElement("w:pPr")
            sp = OxmlElement("w:spacing")
            sp.set(qn("w:before"), "0")
            sp.set(qn("w:after"), "0")
            sp.set(qn("w:line"), "240")
            sp.set(qn("w:lineRule"), "auto")
            pPr.append(sp)
            empty.append(pPr)
            name_el.addprevious(empty)
        return "inserted %d (had %d)" % (keep - cur, cur)
    return "ok %d" % cur


def fix_phan_tich(doc):
    start = end = None
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if start is None and t.startswith("PHÂN TÍCH THÔNG TIN"):
            start = i
        if start is not None and (t.startswith("THƯ TRAO ĐỔI") or t.startswith("Kính gửi")):
            end = i
            break
    if start is None:
        print("phan tich not found")
        return
    end = end or start + 55
    print("phan tich range", start, end)

    # header slogan size 14
    _el, tbl, _pi = find_table_before_phan_tich(doc)
    if tbl is not None:
        n = set_doc_lap_size_in_table(tbl, 14)
        print("doc lap runs set to 14:", n)
    else:
        print("phan tich header table not found")

    for i in range(start, end):
        p = doc.paragraphs[i]
        t = (p.text or "").strip()
        # company name in this section
        replace_text_in_paragraph(p, OLD_CO, NEW_CO)

        # f) align with peers
        if t.startswith("f)"):
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            clear_spacing(p, 1.0, 0, 3)

        # III / IV body: tighten spacing, justify
        if i >= start:
            if t.startswith("III.") or t.startswith("IV."):
                clear_spacing(p, 1.0, 6, 3)
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=True)
            elif t.startswith(("1.", "2.", "3.", "a)", "b)", "c)", "d)", "đ)", "e)", "g)", "+/", "-")):
                clear_spacing(p, 1.0, 0, 3)
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.left_indent = Cm(0)
            elif t.startswith("I.") or t.startswith("II."):
                clear_spacing(p, 1.0, 6, 3)

        # IV date: not bold, right block
        if t.startswith("Trấn Biên") and i > start + 20:
            style_right_date(p, t.replace("Trấn Biên", "Trấn Biên"), bold=False)
            # ensure lowercase t in Trấn - already correct; unbold
            for r in p.runs:
                r.bold = False
                set_run_font(r, size_pt=13, bold=False, italic=True)

        # split multiline IV costs
        if t.startswith("- Chi phí quản lý") and "\n" in (p.text or ""):
            lines = [ln.strip() for ln in p.text.splitlines() if ln.strip()]
            if len(lines) > 1:
                for r in list(p.runs):
                    r._element.getparent().remove(r._element)
                run = p.add_run(lines[0])
                set_run_font(run, 13)
                clear_spacing(p, 1.0, 0, 3)
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                insert_after = p._element
                for line in lines[1:]:
                    new_p = OxmlElement("w:p")
                    pPr = OxmlElement("w:pPr")
                    jc = OxmlElement("w:jc")
                    jc.set(qn("w:val"), "both")
                    pPr.append(jc)
                    sp = OxmlElement("w:spacing")
                    sp.set(qn("w:before"), "0")
                    sp.set(qn("w:after"), "60")
                    sp.set(qn("w:line"), "240")
                    sp.set(qn("w:lineRule"), "auto")
                    pPr.append(sp)
                    new_p.append(pPr)
                    r = OxmlElement("w:r")
                    rPr = OxmlElement("w:rPr")
                    rFonts = OxmlElement("w:rFonts")
                    rFonts.set(qn("w:ascii"), "Times New Roman")
                    rFonts.set(qn("w:hAnsi"), "Times New Roman")
                    rFonts.set(qn("w:eastAsia"), "Times New Roman")
                    rPr.append(rFonts)
                    sz = OxmlElement("w:sz")
                    sz.set(qn("w:val"), "26")
                    rPr.append(sz)
                    r.append(rPr)
                    te = OxmlElement("w:t")
                    te.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    te.text = line
                    r.append(te)
                    new_p.append(r)
                    insert_after.addnext(new_p)
                    insert_after = new_p
                print("split IV cost lines", len(lines))


def fix_qd_doc(doc, *, is_standalone=False):
    # company name in paras
    n1 = replace_in_doc_paragraphs(doc, OLD_CO, NEW_CO)
    n2 = replace_in_tables(doc, OLD_CO, NEW_CO)
    nd = blank_qd_dates_in_tables(doc)
    print("  company paras/tables", n1, n2, "blank dates", nd)
    # title company line center + bold still ok with proper case
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == NEW_CO or t.startswith(NEW_CO):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True)


def fix_tong_hop_qd_sections(doc):
    # Replace company in QD blocks (from first QUYẾT ĐỊNH after to trinh through before PHÂN TÍCH)
    qd_start = None
    pt_start = None
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if qd_start is None and t.startswith("QUYẾT ĐỊNH") and i > 20:
            qd_start = i
        if t.startswith("PHÂN TÍCH THÔNG TIN"):
            pt_start = i
            break
    print("qd block", qd_start, pt_start)
    if qd_start is None:
        return
    end = pt_start or len(doc.paragraphs)
    for i in range(qd_start, end):
        replace_text_in_paragraph(doc.paragraphs[i], OLD_CO, NEW_CO)
    # blank dates only in QD header tables (Số QĐ)
    nd = blank_qd_dates_in_tables(doc)
    print("tong hop blank QD table dates", nd)


def fix_to_trinh_sign_gaps(doc):
    # Prefer 4 blanks; Word page check later may need reduce
    pairs = [
        ("CÁN BỘ ĐỀ XUẤT", "Nguyễn Thị Ánh Ngọc"),
        ("TỔ TRƯỞNG", "Nguyễn Trường Giang"),
        ("PHÓ TRƯỞNG THUẾ CƠ SỞ", "Phạm Đồng Xuân"),
    ]
    # find end of to trinh
    end = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("QUYẾT ĐỊNH") and i > 10:
            end = i
            break
    for title, name in pairs:
        print(" sign", title, ensure_empty_between(doc, title, name, keep=4, search_end=end))


def fix_phan_tich_sign_gap(doc):
    # NGƯỜI PHÂN TÍCH / name - already ~4; ensure 4 after title line
    # title is complex with tabs; ensure between P125-ish and name
    print(
        " phan tich sign",
        ensure_empty_between(doc, "TRƯỞNG NHÓM PHÂN TÍCH", "Nguyễn Thị Ánh Ngọc", keep=4),
    )


def remove_extra_pagebreaks_before_qdgs(doc):
    """Keep one page break before QDGS header; strip page breaks inside QDGS header table."""
    body = list(doc.element.body)
    para_i = 0
    zone = False
    for el in body:
        tag = el.tag.split("}")[-1]
        if tag == "p":
            texts = "".join(t.text or "" for t in el.findall(".//%st" % W))
            if "Điều 3. Quyết định này có hiệu lực" in texts:
                zone = True
            if zone and texts.strip().startswith("Về việc giám sát"):
                zone = False
            para_i += 1
        elif tag == "tbl" and zone:
            texts = " ".join(t.text or "" for t in el.findall(".//%st" % W))
            if "CỘNG HÒA" in texts and "THUẾ CƠ SỞ 3" in texts:
                brs = [
                    b
                    for b in el.findall(".//%sbr" % W)
                    if b.get("%stype" % W) == "page"
                ]
                for b in brs:
                    b.getparent().remove(b)
                print("stripped page brs in QDGS header tbl:", len(brs))


def main():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = BB / ("TONG HOP.bak-layout-%s.docx" % stamp)
    shutil.copy2(LIVE, bak)
    print("backup", bak.name)

    doc = Document(str(LIVE))
    fix_to_trinh_sign_gaps(doc)
    fix_tong_hop_qd_sections(doc)
    remove_extra_pagebreaks_before_qdgs(doc)
    fix_phan_tich(doc)
    # re-find after possible para inserts
    fix_phan_tich_sign_gap(doc)

    # also company name in to trinh body line
    replace_in_doc_paragraphs(doc, OLD_CO, NEW_CO)

    doc.save(str(LIVE))
    print("saved TONG HOP")

    for path, label in ((QDKT, "QDKT"), (QDGS, "QDGS")):
        d = Document(str(path))
        print("==", label, "==")
        fix_qd_doc(d, is_standalone=True)
        d.save(str(path))
        print("saved", path.name)

    with ZipFile(ZIPP, "w", ZIP_DEFLATED) as z:
        for f in (LIVE, QDKT, QDGS):
            z.write(f, f.name)
    print("zip ok")


if __name__ == "__main__":
    main()
