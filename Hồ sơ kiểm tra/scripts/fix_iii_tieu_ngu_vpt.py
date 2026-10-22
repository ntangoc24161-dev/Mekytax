# -*- coding: utf-8 -*-
"""Fix III bullets + remove gia von; sync tieu ngu = quoc hieu + 1pt everywhere."""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
)
LIVE = BB / "TONG HOP.docx"
QDKT = BB / "QD Kiem tra tai cty Van Phat Thanh.docx"
QDGS = BB / "QDgiam sat tai cty Van Phat Thanh.docx"
ZIPP = BB / "hoso.zip"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FI = Cm(1.25)


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
        # also szCs
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), str(int(size_pt * 2)))
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            rPr.append(szCs)
        szCs.set(qn("w:val"), str(int(size_pt * 2)))
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def make_para(text, *, fi=True, size=13, bold=False):
    new_p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "both")
    pPr.append(jc)
    ind = OxmlElement("w:ind")
    if fi:
        ind.set(qn("w:firstLine"), "708")  # 1.25cm ≈ 708 twips
    pPr.append(ind)
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
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    if bold:
        b = OxmlElement("w:b")
        rPr.append(b)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    new_p.append(r)
    return new_p


def clear_runs_set_text(p, text, size=13, bold=False):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(text)
    set_run_font(run, size_pt=size, bold=bold)
    return run


def fix_iii(doc):
    start = end = None
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if start is None and t.startswith("III. PHÂN TÍCH"):
            start = i
        if start is not None and t.startswith("IV. ĐỀ XUẤT"):
            end = i
            break
    print("III range", start, end)
    if start is None:
        return

    # Find and remove gia von para; fix bullets
    # Work from end to start when deleting
    to_delete = []
    for i in range(start, end or start + 30):
        p = doc.paragraphs[i]
        t = (p.text or "").strip()
        if "Tỷ lệ giá vốn hàng bán" in t:
            to_delete.append(p._element)
            print("mark delete gia von P", i)
            continue

        # Split multi-line cost ratio para into separate bullets
        if t.startswith("- Tỷ lệ chi phí quản lý") and ("\n" in (p.text or "") or p.text.count("- Tỷ lệ") > 1):
            raw = p.text
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            # also split if somehow joined with - Tỷ lệ without newline
            if len(lines) == 1 and raw.count("- Tỷ lệ") > 1:
                parts = raw.split("- Tỷ lệ")
                lines = []
                for j, part in enumerate(parts):
                    part = part.strip()
                    if not part:
                        continue
                    if j == 0 and part.startswith("-"):
                        lines.append(part)
                    else:
                        lines.append(("- Tỷ lệ " + part).strip())
            print("split ratio lines", len(lines))
            if lines:
                clear_runs_set_text(p, lines[0], 13)
                p.paragraph_format.first_line_indent = FI
                p.paragraph_format.left_indent = Cm(0)
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(3)
                insert_after = p._element
                for line in lines[1:]:
                    np = make_para(line, fi=True)
                    insert_after.addnext(np)
                    insert_after = np
            continue

        # Bullet lines under III.1: first-line indent 1.25cm
        if t.startswith("- ") and i < (end or i + 1):
            p.paragraph_format.first_line_indent = FI
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(3)
            for r in p.runs:
                set_run_font(r, size_pt=13)

        # +/ headers: no first-line indent (or small)
        if t.startswith("+/"):
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)

    for el in to_delete:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
            print("deleted gia von para")


def run_size_pt(rEl):
    rPr = rEl.find("%srPr" % W)
    if rPr is None:
        return None
    sz = rPr.find("%ssz" % W)
    if sz is None:
        return None
    try:
        return int(sz.get("%sval" % W)) / 2.0
    except Exception:
        return None


def set_run_size_xml(rEl, size_pt, bold=None, italic=None):
    rPr = rEl.find("%srPr" % W)
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rEl.insert(0, rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(a), "Times New Roman")
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), str(int(size_pt * 2)))
    if bold is True:
        if rPr.find(qn("w:b")) is None:
            rPr.append(OxmlElement("w:b"))
    if italic is True:
        if rPr.find(qn("w:i")) is None:
            rPr.append(OxmlElement("w:i"))


def para_full_text(pEl):
    return "".join(t.text or "" for t in pEl.findall(".//%st" % W))


def is_slogan_para(text):
    t = text.replace(" ", "")
    return "Độclập" in t or ("Độc" in text and "Tự do" in text and "Hạnh" in text)


def is_quoc_hieu_para(text):
    return "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in text or text.strip().startswith("CỘNG HÒA")


def sync_tieu_ngu_in_table(tbl):
    """In each table, find quoc hieu size and set slogan = that + 1."""
    # Collect paragraphs that are quoc hieu / slogan
    quoc_sizes = []
    slogan_paras = []
    for pEl in tbl.findall(".//%sp" % W):
        text = para_full_text(pEl)
        if is_quoc_hieu_para(text):
            for rEl in pEl.findall("%sr" % W):
                sz = run_size_pt(rEl)
                if sz:
                    quoc_sizes.append(sz)
        if is_slogan_para(text):
            slogan_paras.append(pEl)

    if not slogan_paras:
        return 0
    base = max(quoc_sizes) if quoc_sizes else 13.0
    # NĐ 30: quoc hieu 12-13, tieu ngu 13-14; tieu ngu = quoc hieu + 1
    target = base + 1.0
    if target > 14:
        target = 14.0
    if target < 13:
        target = 13.0
    n = 0
    for pEl in slogan_paras:
        for rEl in pEl.findall("%sr" % W):
            # only size runs that have text or are part of slogan
            texts = "".join(t.text or "" for t in rEl.findall("%st" % W))
            set_run_size_xml(rEl, target, bold=True, italic=True)
            n += 1
    return n, base, target


def sync_tieu_ngu_doc(doc, label):
    total = 0
    samples = []
    for ti, tbl in enumerate(doc.tables):
        # quick check
        joined = "\n".join(c.text for row in tbl.rows for c in row.cells)
        if "Độc lập" not in joined and "Độc" not in joined:
            continue
        if "CỘNG HÒA" not in joined:
            continue
        res = sync_tieu_ngu_in_table(tbl._tbl)
        if not res:
            continue
        n, base, target = res
        total += n
        if len(samples) < 8:
            samples.append("T%d CH=%.0f slogan->%.0f runs=%d" % (ti, base, target, n))
    # also body paragraphs (rare)
    for p in doc.paragraphs:
        if is_slogan_para(p.text or ""):
            # find nearby quoc hieu? default 13->14
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True, italic=True)
            total += 1
    print(label, "tieu ngu runs updated", total, samples)
    return total


def main():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = BB / ("TONG HOP.bak-layout-%s.docx" % stamp)
    shutil.copy2(LIVE, bak)
    print("backup", bak.name)

    doc = Document(str(LIVE))
    fix_iii(doc)
    sync_tieu_ngu_doc(doc, "TONG HOP")
    doc.save(str(LIVE))
    print("saved TONG HOP")

    for path, label in ((QDKT, "QDKT"), (QDGS, "QDGS")):
        d = Document(str(path))
        sync_tieu_ngu_doc(d, label)
        d.save(str(path))
        print("saved", path.name)

    with ZipFile(ZIPP, "w", ZIP_DEFLATED) as z:
        for f in (LIVE, QDKT, QDGS):
            z.write(f, f.name)
    print("zip ok")

    # verify
    v = Document(str(LIVE))
    print("\nVERIFY III:")
    for i, p in enumerate(v.paragraphs):
        t = (p.text or "").strip()
        if "III. PHÂN TÍCH" in t:
            start = i
        if "IV. ĐỀ XUẤT" in t:
            for j in range(start, i + 1):
                pp = v.paragraphs[j]
                tt = (pp.text or "").replace("\n", " | ")[:100]
                print("P%03d fi=%s | %s" % (j, pp.paragraph_format.first_line_indent, tt))
            break
    print("gia von left?", any("giá vốn" in (p.text or "") for p in v.paragraphs[100:130]))

    print("\nVERIFY tieu ngu samples:")
    hits = 0
    for ti, tbl in enumerate(v.tables):
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "CỘNG HÒA" in (p.text or "") or is_slogan_para(p.text or ""):
                        sizes = [(r.text[:15], r.font.size.pt if r.font.size else None) for r in p.runs if r.text and r.text.strip()]
                        if sizes:
                            print("T%d" % ti, sizes[:4])
                            hits += 1
        if hits >= 12:
            break


if __name__ == "__main__":
    main()
