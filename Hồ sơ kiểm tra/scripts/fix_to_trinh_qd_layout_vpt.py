# -*- coding: utf-8 -*-
"""Sua layout To trinh + QDKT + QDGS theo phoi MAUTONG HOP / skill soan-thao-van-ban.

- To trinh: trich yeu center; ngay/chuc danh/ten canh phai bang center-tab 6663; giu 1 trang
- QDKT: trich yeu center; can cu fi=1cm; gach VML + spacing TRUONG THUE; noi dung KT tach dong fi=1cm; 1 trang
- QDGS: giong the thuc QĐ nhung giu noi dung giam sat; bo sung gach neu thieu
"""
from __future__ import annotations

import shutil
from copy import deepcopy
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
TMPL = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\BIEN BAN\MAUTONG HOP.docx"
)
LIVE = BB / "TONG HOP.docx"
QDKT = BB / "QD Kiem tra tai cty Van Phat Thanh.docx"
QDGS = BB / "QDgiam sat tai cty Van Phat Thanh.docx"
ZIPP = BB / "hoso.zip"

CENTER_TAB_POS = "6663"  # twips - theo phoi
FI_CM = 1.0
FI_BODY_CM = 1.25


def set_run_font(run, size_pt=None, bold=None, italic=None):
    run.font.name = "Times New Roman"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "Times New Roman")
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def clear_spacing(p, ls=1.0):
    pf = p.paragraph_format
    pf.line_spacing = ls
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def set_center_tab(p, pos=CENTER_TAB_POS):
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


def ensure_leading_tab(p):
    text = p.text
    # rebuild runs with leading tab, keep visible text
    visible = text.lstrip("\t")
    # clear runs
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    # remove leftover text nodes in para
    new = p.add_run("\t" + visible)
    set_run_font(new, size_pt=13, bold=True if visible.isupper() and len(visible) < 40 else None)
    return new


def style_right_block_para(p, *, bold=False, size=13):
    """Ngay / chuc danh / ten: tab center nhu phoi."""
    clear_spacing(p, 1.0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    set_center_tab(p)
    visible = p.text.lstrip("\t").strip()
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run("\t" + visible)
    set_run_font(run, size_pt=size, bold=bold)


def find_range_by_title(doc, title_prefix, end_prefixes):
    start = None
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if start is None and t.startswith(title_prefix):
            start = i
            continue
        if start is not None and i > start:
            if any(t.startswith(ep) for ep in end_prefixes):
                return start, i
    if start is not None:
        return start, len(doc.paragraphs)
    return None, None


def copy_vml_line_from_template(target_p, tmpl_p):
    """Copy w:pict/v:line tu phoi neu target thieu."""
    has = target_p._element.find(".//{urn:schemas-microsoft-com:vml}line") is not None
    if has:
        return False
    pict = tmpl_p._element.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict")
    if pict is None:
        return False
    # append deepcopy pict into first run or paragraph
    target_p._element.append(deepcopy(pict))
    return True


def split_multiline_cost_para(p):
    """Neu para chua nhieu dong '- Chi phi...', tach thanh cac para rieng."""
    raw = p.text
    if "\n" not in raw and raw.count("- Chi phí") <= 1:
        return 0
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if len(lines) <= 1:
        return 0
    # set first line into p
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(lines[0])
    set_run_font(run, size_pt=13)
    clear_spacing(p, 1.0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(FI_CM)
    p.paragraph_format.left_indent = Cm(0)

    insert_after = p._element
    created = 0
    for line in lines[1:]:
        new_p = OxmlElement("w:p")
        # copy pPr baseline
        pPr = OxmlElement("w:pPr")
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "both")
        pPr.append(jc)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLine"), str(int(FI_CM * 360000 / 25.4 * 20 / 20)))  # wrong
        # firstLine in twips: 1cm ~= 567 twips
        ind.set(qn("w:firstLine"), "567")
        pPr.append(ind)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        pPr.append(spacing)
        new_p.append(pPr)
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), "Times New Roman")
        rFonts.set(qn("w:hAnsi"), "Times New Roman")
        rFonts.set(qn("w:eastAsia"), "Times New Roman")
        rPr.append(rFonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "26")  # 13pt
        rPr.append(sz)
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), "26")
        rPr.append(szCs)
        r.append(rPr)
        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = line
        r.append(t)
        new_p.append(r)
        insert_after.addnext(new_p)
        insert_after = new_p
        created += 1
    return created


def fix_to_trinh(doc, start, end):
    """Fix to trinh block [start:end)."""
    # Identify signature markers
    date_idxs = []
    title_idxs = []
    name_idxs = []
    for i in range(start, end):
        t = doc.paragraphs[i].text.strip()
        if t.startswith("Trấn Biên, ngày"):
            date_idxs.append(i)
        if t in ("CÁN BỘ ĐỀ XUẤT", "TỔ TRƯỞNG", "KT. TRƯỞNG THUẾ CƠ SỞ", "PHÓ TRƯỞNG THUẾ CƠ SỞ"):
            title_idxs.append(i)
        if t in ("Nguyễn Thị Ánh Ngọc", "Nguyễn Trường Giang", "Phạm Đồng Xuân"):
            name_idxs.append(i)

    for i in range(start, end):
        p = doc.paragraphs[i]
        t = p.text.strip()
        if not t and i not in range(start, end):
            continue

        # Title
        if i == start:
            clear_spacing(p, 1.0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True)
            continue

        # Trich yeu
        if t.startswith("V/v:") or (i == start + 2 and "giám sát" in t.lower()):
            clear_spacing(p, 1.0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=False)
            continue

        # Date / chuc danh / ten ky - right block via center tab
        if i in date_idxs or i in title_idxs or i in name_idxs:
            bold = i in title_idxs or i in name_idxs
            style_right_block_para(p, bold=bold, size=13)
            continue

        # Empty signature spacers: keep compact, no extra spacing props
        if not t:
            clear_spacing(p, 1.0)
            p.paragraph_format.first_line_indent = Cm(0)
            continue

        # Body labels
        if t.startswith("- ") and t.endswith(":") or t.startswith("- Ý kiến") or t.startswith("- Đề xuất") or t.startswith("- Nội dung") or t.startswith("- Căn cứ"):
            clear_spacing(p, 1.0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(FI_CM)
            for r in p.runs:
                set_run_font(r, size_pt=13)
            continue

        # Body content
        if t.startswith("Ý kiến"):
            clear_spacing(p, 1.0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(FI_CM)
            for r in p.runs:
                set_run_font(r, size_pt=13)
            continue

        clear_spacing(p, 1.0)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(FI_BODY_CM)
        for r in p.runs:
            set_run_font(r, size_pt=13)

    # Normalize blank lines between chuc danh and name: keep exactly 3 empties
    # Rebuild by editing existing empties around each title->name pair
    pairs = [
        ("CÁN BỘ ĐỀ XUẤT", "Nguyễn Thị Ánh Ngọc"),
        ("TỔ TRƯỞNG", "Nguyễn Trường Giang"),
        ("PHÓ TRƯỞNG THUẾ CƠ SỞ", "Phạm Đồng Xuân"),
    ]
    # Just report; blank count already roughly ok except can bo. Insert empties if needed.
    for title, name in pairs:
        ti = ni = None
        for i in range(start, end):
            if doc.paragraphs[i].text.strip() == title:
                ti = i
            if doc.paragraphs[i].text.strip() == name and ti is not None and i > ti:
                ni = i
                break
        if ti is None or ni is None:
            continue
        between = ni - ti - 1
        # target 3 blanks for 1-page balance (phoi ~3-4)
        if between < 3:
            # insert empty paragraphs before name
            name_el = doc.paragraphs[ni]._element
            for _ in range(3 - between):
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
        elif between > 4:
            # remove extras from the middle empties
            remove_n = between - 3
            # remove empty paras after title
            idx = ti + 1
            removed = 0
            while removed < remove_n and idx < ni:
                # indices shift - operate on XML siblings
                break
            # safer XML approach:
            el = doc.paragraphs[ti]._element.getnext()
            while el is not None and removed < remove_n:
                nxt = el.getnext()
                texts = "".join(node.text or "" for node in el.iter(qn("w:t"))).strip()
                if texts == "" and el.tag == qn("w:p"):
                    parent = el.getparent()
                    parent.remove(el)
                    removed += 1
                el = nxt


def fix_quyet_dinh(doc, start, end, *, is_gs: bool, tmpl_truong_p=None):
    for i in range(start, end):
        p = doc.paragraphs[i]
        t = p.text.strip()
        clear_spacing(p, 1.0)

        if t.startswith("QUYẾT ĐỊNH") and len(t) <= 12:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True)
            continue

        if t.startswith("Về việc"):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True)
            continue

        if t.startswith("CÔNG TY"):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True)
            continue

        if t.startswith("TRƯỞNG THUẾ CƠ SỞ"):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            # phoi QDKT dung left=-108 twips
            p.paragraph_format.left_indent = Twips(-108) if not is_gs else None
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(12)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True)
            if tmpl_truong_p is not None:
                copied = copy_vml_line_from_template(p, tmpl_truong_p)
                if copied:
                    print("  copied VML line into", "GS" if is_gs else "KT", "P", i)
            continue

        if t.startswith("Căn cứ ") or t.startswith("Theo đề nghị"):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(FI_CM)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, italic=True)
            continue

        if t == "QUYẾT ĐỊNH:":
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True)
            continue

        if t.startswith("Điều "):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(FI_CM)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=False)
            continue

        if t.startswith("- Chi phí") or t.startswith("- Ông") or t.startswith("Thời kỳ") or t.startswith("Thời hạn"):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(FI_CM)
            p.paragraph_format.left_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13)
            continue

        if t:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(FI_CM)
            for r in p.runs:
                set_run_font(r, size_pt=13)

    # Split cost lines if needed - find para containing multiple chi phi
    # Re-scan because indices may be used before split
    for i in range(start, min(end, len(doc.paragraphs))):
        t = doc.paragraphs[i].text
        if "- Chi phí quản lý" in t and "- Chi phí bán hàng" in t:
            n = split_multiline_cost_para(doc.paragraphs[i])
            print("  split cost lines:", n + 1)
            break


def tighten_qd_section_margins(doc):
    """QĐ trình ký: lề vẫn NĐ 30, siết dưới nhẹ để ưu tiên 1 trang."""
    for s in doc.sections:
        s.page_width = Cm(21.0)
        s.page_height = Cm(29.7)
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(2.0)
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(1.5)


def fix_standalone_qd(path: Path, *, is_gs: bool, tmpl_truong_p):
    doc = Document(str(path))
    tighten_qd_section_margins(doc)
    fix_quyet_dinh(doc, 0, len(doc.paragraphs), is_gs=is_gs, tmpl_truong_p=tmpl_truong_p)
    doc.save(str(path))
    print("saved standalone", path.name)


def main():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = BB / f"TONG HOP.bak-layout-{stamp}.docx"
    shutil.copy2(LIVE, bak)
    print("backup", bak)

    tmpl = Document(str(TMPL))
    # template TRUONG para for QDKT (has VML line)
    tmpl_truong = None
    for p in tmpl.paragraphs:
        if p.text.strip().startswith("TRƯỞNG THUẾ CƠ SỞ") and p._element.find(".//{urn:schemas-microsoft-com:vml}line") is not None:
            tmpl_truong = p
            break
    print("tmpl_truong_found", tmpl_truong is not None)

    doc = Document(str(LIVE))

    # To trinh dau file
    tt_start, tt_end = find_range_by_title(doc, "TỜ TRÌNH LÃNH ĐẠO THUẾ CƠ SỞ", ["QUYẾT ĐỊNH"])
    print("to_trinh", tt_start, tt_end)
    if tt_start is not None:
        fix_to_trinh(doc, tt_start, tt_end)

    # QDKT / QDGS trong TONG HOP - tim theo trich yeu
    qdkt_start = qdgs_start = None
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if t.startswith("QUYẾT ĐỊNH"):
            nxt = doc.paragraphs[i + 1].text.strip() if i + 1 < len(doc.paragraphs) else ""
            if "kiểm tra thuế" in nxt and qdkt_start is None:
                qdkt_start = i
            if "giám sát" in nxt.lower() and qdgs_start is None:
                qdgs_start = i
    # ends
    qdkt_end = qdgs_start if qdgs_start else None
    qdgs_end = None
    if qdgs_start is not None:
        for j in range(qdgs_start + 1, len(doc.paragraphs)):
            if doc.paragraphs[j].text.strip().startswith("PHÂN TÍCH THÔNG TIN"):
                qdgs_end = j
                break
        if qdgs_end is None:
            qdgs_end = qdgs_start + 25

    print("qdkt", qdkt_start, qdkt_end, "qdgs", qdgs_start, qdgs_end)
    if qdkt_start is not None:
        fix_quyet_dinh(doc, qdkt_start, qdkt_end or (qdkt_start + 30), is_gs=False, tmpl_truong_p=tmpl_truong)
    if qdgs_start is not None:
        fix_quyet_dinh(doc, qdgs_start, qdgs_end or (qdgs_start + 30), is_gs=True, tmpl_truong_p=tmpl_truong)

    # Section 0 (to trinh + QD) margins
    if doc.sections:
        s0 = doc.sections[0]
        s0.left_margin = Cm(3.0)
        s0.right_margin = Cm(2.0)
        s0.top_margin = Cm(2.0)
        s0.bottom_margin = Cm(1.5)

    doc.save(str(LIVE))
    print("saved", LIVE)

    fix_standalone_qd(QDKT, is_gs=False, tmpl_truong_p=tmpl_truong)
    fix_standalone_qd(QDGS, is_gs=True, tmpl_truong_p=tmpl_truong)

    with ZipFile(ZIPP, "w", ZIP_DEFLATED) as z:
        for f in (LIVE, QDKT, QDGS):
            if f.exists():
                z.write(f, f.name)
    print("zip", ZIPP)

    # verify snippets
    v = Document(str(LIVE))
    print("\nVERIFY to trinh:")
    for i in range(1, 36):
        p = v.paragraphs[i]
        t = p.text.replace("\t", "\\t")
        if t.strip() or i in (2, 3, 14, 15, 17):
            tabs = p._element.pPr.find(qn("w:tabs")) if p._element.pPr is not None else None
            print(f"P{i}|al={p.alignment}|fi={p.paragraph_format.first_line_indent}|tabs={tabs is not None}|{t[:70]}")

    print("\nVERIFY QDKT:")
    for i, p in enumerate(v.paragraphs):
        if i < 38 or i > 58:
            continue
        t = p.text.replace("\n", " | ")
        has_line = p._element.find(".//{urn:schemas-microsoft-com:vml}line") is not None
        print(f"P{i}|al={p.alignment}|fi={p.paragraph_format.first_line_indent}|line={has_line}|{t[:80]}")

    print("\nVERIFY QDGS:")
    for i, p in enumerate(v.paragraphs):
        if i < 62 or i > 78:
            continue
        t = p.text.replace("\n", " | ")
        has_line = p._element.find(".//{urn:schemas-microsoft-com:vml}line") is not None
        print(f"P{i}|al={p.alignment}|fi={p.paragraph_format.first_line_indent}|line={has_line}|{t[:80]}")

    qs = Document(str(QDGS))
    print("\nStandalone QDGS trich yeu:", qs.paragraphs[1].text[:60])
    print("Standalone QDGS dieu1:", [p.text[:50] for p in qs.paragraphs if p.text.startswith("Điều 1")][:1])


if __name__ == "__main__":
    main()
