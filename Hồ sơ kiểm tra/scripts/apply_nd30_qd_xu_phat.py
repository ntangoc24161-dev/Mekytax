# -*- coding: utf-8 -*-
"""
Chuẩn hóa chi tiết Quyết định xử phạt VPHC và Tờ trình xử phạt theo Nghị định 30/2020/NĐ-CP.
"""
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


def set_para_format(p, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Cm(1.0),
                     space_before=Pt(2), space_after=Pt(2), line_spacing=1.0):
    p.alignment = align
    pf = p.paragraph_format
    if first_indent is not None:
        pf.first_line_indent = first_indent
    else:
        pf.first_line_indent = None
    if space_before is not None:
        pf.space_before = space_before
    if space_after is not None:
        pf.space_after = space_after
    if line_spacing is not None:
        pf.line_spacing = line_spacing


def run_standardization():
    if not LIVE.exists():
        print(f"File {LIVE} does not exist.")
        return

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = BB / f"TONG HOP.bak-standardized-qdxphc-{ts}.docx"
    shutil.copy2(LIVE, bak)
    print(f"Backup created: {bak.name}")

    doc = Document(LIVE)

    # 1. Header Table 42 (QĐXPHC)
    t42 = doc.tables[42]
    # Row 0 Col 0
    c00 = t42.rows[0].cells[0]
    for p in c00.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            if "CƠ SỞ 3" in r.text or "THUẾ CƠ SỞ" in p.text:
                set_run_font(r, size_pt=12, bold=True, italic=False)
            else:
                set_run_font(r, size_pt=12, bold=False, italic=False)
    # Row 0 Col 1
    c01 = t42.rows[0].cells[1]
    for p in c01.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if "CỘNG HÒA" in p.text:
            for r in p.runs:
                set_run_font(r, size_pt=12, bold=True, italic=False)
        elif "Độc lập" in p.text:
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)

    # Row 2 (Số & Ngày)
    r_last = t42.rows[-1]
    for p in r_last.cells[0].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_run_font(r, size_pt=12, bold=False, italic=False)
    for p in r_last.cells[1].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_run_font(r, size_pt=13, bold=False, italic=True)

    # 2. Các đoạn văn của QĐXPHC
    qd_start = -1
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "QUYẾT ĐỊNH" and i >= 500:
            qd_start = i
            break

    if qd_start != -1:
        print(f"Standardizing QDXPHC from P{qd_start}")
        # P[qd_start]: QUYẾT ĐỊNH
        p_title = doc.paragraphs[qd_start]
        set_para_format(p_title, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None, space_before=Pt(6), space_after=Pt(2))
        for r in p_title.runs:
            set_run_font(r, size_pt=14, bold=True, italic=False)

        # P[qd_start+1]: Trích yếu
        p_trich_yeu = doc.paragraphs[qd_start + 1]
        p_trich_yeu.text = "Về việc xử phạt vi phạm hành chính về thuế"
        set_para_format(p_trich_yeu, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None, space_before=Pt(2), space_after=Pt(6))
        for r in p_trich_yeu.runs:
            set_run_font(r, size_pt=13, bold=True, italic=False)

        # Căn cứ pháp lý & Nội dung các Điều
        for i in range(qd_start + 2, min(len(doc.paragraphs), qd_start + 65)):
            p = doc.paragraphs[i]
            t = p.text.strip()
            if not t:
                continue
            if t.startswith("TỜ TRÌNH LÃNH ĐẠO"):
                break

            if t.startswith("Căn cứ"):
                set_para_format(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Cm(1.0), space_before=Pt(2), space_after=Pt(2), line_spacing=1.0)
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=False, italic=True)
            elif t == "TRƯỞNG THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI":
                set_para_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None, space_before=Pt(6), space_after=Pt(6))
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=True, italic=False)
            elif t.startswith("Theo đề nghị"):
                set_para_format(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Cm(1.0), space_before=Pt(2), space_after=Pt(4), line_spacing=1.0)
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=False, italic=False)
            elif t == "QUYẾT ĐỊNH:":
                set_para_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None, space_before=Pt(4), space_after=Pt(4), line_spacing=1.0)
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=True, italic=False)
            elif t.startswith("Điều "):
                set_para_format(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Cm(1.0), space_before=Pt(4), space_after=Pt(2), line_spacing=1.0)
                # Đậm phần Điều X.
                for r in p.runs:
                    if r.text.strip().startswith("Điều"):
                        set_run_font(r, size_pt=13, bold=True, italic=False)
                    else:
                        set_run_font(r, size_pt=13, bold=None, italic=False)
            else:
                set_para_format(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Cm(1.0), space_before=Pt(2), space_after=Pt(2), line_spacing=1.0)
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=None, italic=None)

    # 3. Footer Table 43
    t43 = doc.tables[43]
    for p in t43.rows[0].cells[0].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if "Nơi nhận" in p.text:
            for r in p.runs:
                set_run_font(r, size_pt=12, bold=True, italic=True)
        else:
            for r in p.runs:
                set_run_font(r, size_pt=11, bold=False, italic=False)

    for p in t43.rows[0].cells[1].paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            if any(k in r.text for k in ["KT. TRƯỞNG", "PHÓ TRƯỞNG", "Phạm Đồng Xuân"]):
                set_run_font(r, size_pt=13, bold=True, italic=False)
            else:
                set_run_font(r, size_pt=13, bold=False, italic=False)

    # 4. Tờ trình xử phạt (P624..)
    for i, p in enumerate(doc.paragraphs):
        if "TỜ TRÌNH LÃNH ĐẠO THUẾ CƠ SỞ 3" in p.text and i >= 600:
            set_para_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None, space_before=Pt(6), space_after=Pt(2))
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True, italic=False)
        elif "Về việc: ban hành Quyết Định xử phạt" in p.text and i >= 600:
            p.text = "Về việc: ban hành Quyết định xử phạt vi phạm trong lĩnh vực thuế."
            set_para_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=None, space_before=Pt(2), space_after=Pt(6))
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)
        elif "Nghị định số 125" in p.text and i >= 600:
            if "125/2022/NĐ-CP" in p.text:
                p.text = p.text.replace("125/2022/NĐ-CP", "125/2020/NĐ-CP")
            if "19/10/2022" in p.text:
                p.text = p.text.replace("19/10/2022", "19/10/2020")
            set_para_format(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Cm(1.0), space_before=Pt(3), space_after=Pt(3), line_spacing=1.0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=None, italic=None)
        elif "Trấn Biên, ngày 02 tháng 10 năm 2026" in p.text and i >= 630:
            p.text = p.text.replace("02 tháng 10", "14 tháng 10")
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=False, italic=True)

    doc.save(LIVE)
    print("Saved TONG HOP.docx successfully.")

    with ZipFile(ZIPP, "w", ZIP_DEFLATED) as z:
        z.write(LIVE, LIVE.name)
        if QDKT.exists():
            z.write(QDKT, QDKT.name)
        if QDGS.exists():
            z.write(QDGS, QDGS.name)
    print(f"Archive updated: {ZIPP}")


if __name__ == "__main__":
    run_standardization()
