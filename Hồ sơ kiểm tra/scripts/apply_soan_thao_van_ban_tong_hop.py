# -*- coding: utf-8 -*-
"""Ap dung skill shared-skills/soan-thao-van-ban (NĐ 30/2020) cho TONG HOP.docx.

- Tu nhan dien loai van ban
- Can le A4, TNR, gian dong 1.0-1.15, spacing 3-6pt, first-line indent body
- Giu floating table (tblpPr) cua phoi
"""
from __future__ import annotations

import shutil
from collections import Counter
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

KIND_RULES = [
    ("phan_tich", ["PHÂN TÍCH THÔNG TIN, XÁC ĐỊNH NỘI DUNG"]),
    ("thu_trao_doi", ["THƯ TRAO ĐỔI THÔNG TIN"]),
    ("ke_hoach_gs", ["KẾ HOẠCH GIÁM SÁT HOẠT ĐỘNG"]),
    ("nhat_ky", ["NHẬT KÝ THANH TRA KIỂM TRA"]),
    ("giay_xac_nhan", ["GIẤY XÁC NHẬN"]),
    ("bien_ban", ["BIÊN BẢN KIỂM TRA", "BIÊN BẢN LÀM VIỆC", "BIÊN BẢN"]),
    ("bao_cao", ["BÁO CÁO"]),
    ("to_trinh", ["TỜ TRÌNH LÃNH ĐẠO"]),
    ("quyet_dinh", ["QUYẾT ĐỊNH"]),
]

# title_pt, body_pt, line_spacing, space_before, space_after, first_indent_cm
# QĐ dùng 1.0 để ưu tiên gọn 1 trang trình ký (skill cho phép 1.0-1.15)
KIND_STYLE = {
    "to_trinh": (14, 13, 1.15, 3, 3, 1.0),
    "quyet_dinh": (14, 13, 1.0, 3, 3, 1.0),
    "phan_tich": (14, 13, 1.15, 3, 3, 1.0),
    "thu_trao_doi": (14, 13, 1.15, 3, 3, 1.0),
    "giay_xac_nhan": (14, 13, 1.15, 3, 3, 1.0),
    "bien_ban": (14, 13, 1.15, 3, 3, 1.0),
    "bao_cao": (14, 13, 1.15, 3, 3, 1.0),
    "ke_hoach_gs": (14, 13, 1.15, 3, 3, 1.0),
    "nhat_ky": (14, 12, 1.15, 3, 3, 1.0),
    "khac": (14, 13, 1.15, 3, 3, 1.0),
}


def detect_kind(text: str) -> str | None:
    t = text.strip()
    if not t or len(t) > 140:
        return None
    for kind, keys in KIND_RULES:
        for k in keys:
            if k in t:
                return kind
    return None


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


def is_section_heading(text: str) -> bool:
    t = text.strip()
    if not t or len(t) > 100:
        return False
    return t.startswith(("I.", "II.", "III.", "IV.", "V.", "Điều ", "QUYẾT ĐỊNH:", "Nơi nhận"))


def is_list_or_label(text: str) -> bool:
    t = text.strip()
    return (
        t.startswith(("-", "+/", "*/", "a)", "b)", "c)", "d)", "đ)", "e)", "g)", "1.", "2.", "3."))
        or t.startswith("V/v")
        or t.startswith("- ")
    )


def is_can_cu_line(text: str, kind: str) -> bool:
    if kind != "quyet_dinh":
        return False
    t = text.strip()
    return t.startswith("Căn cứ ") or t.startswith("Theo đề nghị")


def is_date_line(text: str) -> bool:
    t = text.strip()
    return ("ngày" in t.lower() and "tháng" in t.lower() and "năm" in t.lower()) and len(t) < 80


def build_ranges(doc: Document):
    starts = []
    for i, p in enumerate(doc.paragraphs):
        kind = detect_kind(p.text)
        if not kind:
            continue
        if p.text.strip() == "QUYẾT ĐỊNH:" and starts and starts[-1][2] == "quyet_dinh":
            continue
        if p.text.strip().startswith("II. NHẬT KÝ"):
            continue
        starts.append((i, p.text.strip()[:90], kind))
    ranges = []
    for idx, (start, title, kind) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(doc.paragraphs)
        ranges.append((start, end, kind, title))
    return ranges


def apply_section_margins(doc: Document):
    # NĐ 30: trên/dưới 20-25mm, trái 30-35mm, phải 15-20mm
    for s in doc.sections:
        s.page_width = Cm(21.0)
        s.page_height = Cm(29.7)
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(2.0)
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)


def apply_para_style(p, *, kind: str, is_title: bool, is_heading: bool, is_can_cu: bool):
    title_pt, body_pt, ls, sb, sa, indent_cm = KIND_STYLE.get(kind, KIND_STYLE["khac"])
    pf = p.paragraph_format
    text = p.text.strip()
    if not text:
        return

    if is_title:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.line_spacing = ls
        pf.space_before = Pt(6)
        pf.space_after = Pt(3)
        pf.first_line_indent = Cm(0)
        for r in p.runs:
            set_run_font(r, size_pt=title_pt, bold=True, italic=False)
        return

    if is_heading:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER if text == "QUYẾT ĐỊNH:" else WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = ls
        pf.space_before = Pt(sb)
        pf.space_after = Pt(sa)
        pf.first_line_indent = Cm(0)
        for r in p.runs:
            set_run_font(r, size_pt=body_pt, bold=True)
        return

    # body
    if text.isupper() and len(text) < 90 and not text.startswith("CỘNG HÒA"):
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Cm(0)
    else:
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if is_list_or_label(text) or is_can_cu or is_date_line(text):
            pf.first_line_indent = Cm(0)
        else:
            pf.first_line_indent = Cm(indent_cm)

    pf.line_spacing = ls
    pf.space_before = Pt(sb)
    pf.space_after = Pt(sa)

    italic = True if (is_can_cu or is_date_line(text)) else None
    for r in p.runs:
        set_run_font(r, size_pt=body_pt, italic=italic)


def apply_table_fonts(doc: Document):
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        set_run_font(r)


def count_floating(doc: Document) -> int:
    n = 0
    for t in doc.tables:
        tblPr = t._tbl.tblPr
        if tblPr is not None and tblPr.find(qn("w:tblpPr")) is not None:
            n += 1
    return n


def style_standalone_qd(path: Path):
    doc = Document(str(path))
    apply_section_margins(doc)
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        is_title = t == "QUYẾT ĐỊNH" or t.startswith("Về việc")
        is_heading = t.startswith("Điều ") or t == "QUYẾT ĐỊNH:"
        is_can_cu = t.startswith("Căn cứ ") or t.startswith("Theo đề nghị")
        apply_para_style(
            p,
            kind="quyet_dinh",
            is_title=is_title,
            is_heading=is_heading,
            is_can_cu=is_can_cu,
        )
        # ten cong ty / truong thue: center bold
        if ("CÔNG TY" in t or t.startswith("TRƯỞNG THUẾ")) and len(t) < 120:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True)
    apply_table_fonts(doc)
    doc.save(str(path))
    print("styled_qd", path.name)


def main():
    if not LIVE.exists():
        raise SystemExit(f"Missing {LIVE}")

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = BB / f"TONG HOP.bak-soan-thao-van-ban-{stamp}.docx"
    shutil.copy2(LIVE, bak)
    print("backup", bak)
    print("skill", "shared-skills/soan-thao-van-ban")

    doc = Document(str(LIVE))
    float_before = count_floating(doc)
    ranges = build_ranges(doc)
    print("detected_docs", len(ranges))
    for start, end, kind, title in ranges:
        print(f"  [{start}:{end}] {kind} | {title}")

    para_kind = ["khac"] * len(doc.paragraphs)
    para_is_title = [False] * len(doc.paragraphs)
    for start, end, kind, _title in ranges:
        for i in range(start, end):
            para_kind[i] = kind
        para_is_title[start] = True

    apply_section_margins(doc)
    changed = 0
    kind_counts = Counter()
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        kind = para_kind[i]
        kind_counts[kind] += 1
        if not text:
            continue
        title = para_is_title[i]
        heading = (not title) and is_section_heading(text)
        can_cu = is_can_cu_line(text, kind)
        apply_para_style(p, kind=kind, is_title=title, is_heading=heading, is_can_cu=can_cu)
        changed += 1

    apply_table_fonts(doc)
    float_after = count_floating(doc)
    if float_after < float_before:
        raise SystemExit(f"Floating tables lost: {float_before} -> {float_after}")

    doc.save(str(LIVE))
    print("saved", LIVE)
    print("paragraphs_styled", changed)
    print("kind_counts", dict(kind_counts))
    print("floating", float_after)

    for f in (QDKT, QDGS):
        if f.exists():
            style_standalone_qd(f)

    v = Document(str(LIVE))
    ls = Counter(str(p.paragraph_format.line_spacing) for p in v.paragraphs if p.text.strip())
    print("verify_ls_nonempty", dict(ls.most_common(6)))
    for s in v.sections:
        print(
            "margins_cm L/R/T/B",
            round(s.left_margin.cm, 2),
            round(s.right_margin.cm, 2),
            round(s.top_margin.cm, 2),
            round(s.bottom_margin.cm, 2),
        )

    with ZipFile(ZIPP, "w", ZIP_DEFLATED) as z:
        for f in (LIVE, QDKT, QDGS):
            if f.exists():
                z.write(f, f.name)
    print("zip", ZIPP, ZIPP.stat().st_size)


if __name__ == "__main__":
    main()
