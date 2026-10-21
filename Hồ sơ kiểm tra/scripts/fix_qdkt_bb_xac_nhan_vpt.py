# -*- coding: utf-8 -*-
"""Fix QDKT bold/align + biên bản xác nhận số liệu (TV/TD) for Vạn Phát Thành."""
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
from docx.shared import Pt

BB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
)
LIVE = BB / "TONG HOP.docx"
QDKT = BB / "QD Kiem tra tai cty Van Phat Thanh.docx"
QDGS = BB / "QDgiam sat tai cty Van Phat Thanh.docx"
ZIPP = BB / "hoso.zip"

CAN_CU = (
    "Căn cứ Điều 4 Thông tư 96/2015/TT-BTC ngày 22/06/2015 của Bộ Tài chính."
)
CLOSING_1_TMPL = (
    "Biên bản này được lập thành 03 bản, bao gồm {pages} trang, "
    "01 bản có nội dung và giá trị pháp lý như nhau, 01 bản được giao người nộp thuế."
)
CLOSING_2 = (
    "Sau khi đọc lại biên bản, những người có mặt đồng ý về nội dung biên bản, "
    "không có ý kiến gì khác và cùng ký xác nhận vào từng trang của biên bản "
    "hoặc có ý kiến khác như sau:"
)

CPTC_BY_YEAR = {
    2021: 31181930,
    2022: 33975531,
    2023: 0,
    2024: 0,
    2025: 17653341,
}


def fmt_num(n: int) -> str:
    return f"{int(n):,}".replace(",", ".")


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
        for tag in ("w:sz", "w:szCs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rPr.append(el)
            el.set(qn("w:val"), str(int(size_pt * 2)))
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def clear_runs(p):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)


def set_para_text(p, text, *, size=13, bold=False, italic=False, align=None, first_indent=None):
    clear_runs(p)
    run = p.add_run(text)
    set_run_font(run, size_pt=size, bold=bold, italic=italic)
    if align is not None:
        p.paragraph_format.alignment = align
    if first_indent is not None:
        p.paragraph_format.first_line_indent = first_indent
    return run


def set_dieu_para(p, text, size=13):
    """Only bold 'Điều N.' prefix; body regular; left aligned."""
    m = re.match(r"^(Điều\s+\d+\.)(\s*)(.*)$", text.strip(), flags=re.DOTALL)
    clear_runs(p)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if not m:
        run = p.add_run(text)
        set_run_font(run, size_pt=size, bold=False, italic=False)
        return
    label, sp, body = m.group(1), m.group(2) or " ", m.group(3)
    r1 = p.add_run(label)
    set_run_font(r1, size_pt=size, bold=True, italic=False)
    if body:
        r2 = p.add_run(sp + body)
        set_run_font(r2, size_pt=size, bold=False, italic=False)


def clone_pPr(src_p, new_p):
    src_pPr = src_p._element.find(qn("w:pPr"))
    if src_pPr is None:
        return
    new_pPr = new_p.find(qn("w:pPr"))
    if new_pPr is not None:
        new_p.remove(new_pPr)
    new_p.insert(0, src_pPr.__copy__())


def insert_para_after(after_p, text, *, size=13, bold=False, align=None):
    new_el = OxmlElement("w:p")
    clone_pPr(after_p, new_el)
    after_p._element.addnext(new_el)
    # wrap as paragraph via parent doc later by reloading indices; set text via XML
    # Build run
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
        rPr.append(OxmlElement("w:b"))
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    new_el.append(r)
    if align is not None:
        pPr = new_el.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            new_el.insert(0, pPr)
        jc = pPr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            pPr.append(jc)
        val = {
            WD_ALIGN_PARAGRAPH.LEFT: "left",
            WD_ALIGN_PARAGRAPH.CENTER: "center",
            WD_ALIGN_PARAGRAPH.RIGHT: "right",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "both",
        }.get(align, "left")
        jc.set(qn("w:val"), val)
    return new_el


def fix_qdkt_format(doc, label: str):
    n = 0
    in_qdkt = False
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        # enter QDKT block
        if t.startswith("Về việc kiểm tra thuế"):
            in_qdkt = True
        if t.startswith("Về việc giám sát"):
            in_qdkt = False
        if t.startswith("PHÂN TÍCH THÔNG TIN"):
            in_qdkt = False
        if not in_qdkt and not (label.startswith("QDKT") and t.startswith("Điều")):
            # standalone QDKT file: whole doc is QDKT
            if label != "QDKT-file":
                continue
        if label == "QDKT-file":
            in_qdkt = True

        if t.startswith("Điều "):
            set_dieu_para(p, t, size=13 if "TONG" in label else 12)
            n += 1
            continue
        # body under decision: left align
        if in_qdkt and (
            t.startswith("- ")
            or t.startswith("Thời hạn")
            or t.startswith("Thời kỳ")
            or t.startswith("Đoàn kiểm tra có nhiệm vụ")
        ):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # unbold accidental whole-line bold
            for r in p.runs:
                if r.bold:
                    r.bold = False
            n += 1
    print(label, "qdkt format fixes", n)
    return n


def cptc_line(year: int) -> str:
    giam = CPTC_BY_YEAR.get(year, 0)
    if not giam:
        return "*/ Không chênh lệch."
    return (
        f"*/ Chi phí tài chính giảm {fmt_num(giam)} đồng do Công ty hạch toán "
        f"chi phí lãi vay không đúng quy định. {CAN_CU}"
    )


def year_header_tv(year: int) -> str:
    return f"+ NĂM {year}: Đơn vị tính: Đồng"


def year_header_td(year: int) -> str:
    return f"Năm {year} Đơn vị tính: Đồng"


def find_section_ranges(doc):
    """Return (tv_start, tv_end_excl, td_start, td_end_excl) paragraph indices."""
    tv = td = None
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t == "Xác nhận số liệu kiểm tra":
            # look ahead for TV vs TD marker
            window = " ".join((doc.paragraphs[j].text or "") for j in range(i, min(i + 12, len(doc.paragraphs))))
            if "Thành viên." in window and tv is None:
                tv = i
            elif "Trưởng đoàn." in window and td is None:
                td = i
    if tv is None or td is None:
        raise RuntimeError(f"cannot find BB xac nhan sections tv={tv} td={td}")
    # ends
    tv_end = td
    td_end = None
    for i in range(td + 1, len(doc.paragraphs)):
        t = (doc.paragraphs[i].text or "").strip()
        if t.startswith("BIÊN BẢN KIỂM TRA") or t == "BIÊN BẢN KIỂM TRA":
            td_end = i
            break
        if t.startswith("BIÊN BẢN") and "Xác nhận" not in t and i > td + 5:
            td_end = i
            break
    if td_end is None:
        td_end = min(len(doc.paragraphs), td + 40)
    return tv, tv_end, td, td_end


def rewrite_result_block(doc, start, end, mode: str):
    """Rewrite II. Kết quả kiểm tra years 2021-2025 + closing lines before ý kiến."""
    paras = doc.paragraphs
    # find I / II / ý kiến inside range
    i_idx = ii_idx = yk_idx = None
    for i in range(start, end):
        t = (paras[i].text or "").strip()
        if t.startswith("I. Nội dung kiểm tra"):
            i_idx = i
        if t.startswith("II. Kết quả kiểm tra"):
            ii_idx = i
        if t.startswith("Ý kiến bổ sung"):
            yk_idx = i
            break
    if i_idx is None or ii_idx is None or yk_idx is None:
        raise RuntimeError(f"markers missing in {mode}: I={i_idx} II={ii_idx} YK={yk_idx}")

    # Fill I. nội dung
    # paragraph right after I.
    content_p = paras[i_idx + 1]
    if mode == "tv":
        set_para_text(
            content_p,
            "- Chi phí tài chính.",
            size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )
    else:
        set_para_text(
            content_p,
            "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.",
            size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )

    # Clear everything between II. and Ý kiến, then rebuild years + closing
    # Keep II. paragraph; remove following until Ý kiến
    to_remove = []
    el = paras[ii_idx]._element.getnext()
    yk_el = paras[yk_idx]._element
    while el is not None and el is not yk_el:
        nxt = el.getnext()
        if el.tag == qn("w:p"):
            to_remove.append(el)
        el = nxt
    for el in to_remove:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    # Insert year blocks after II.
    anchor = paras[ii_idx]
    years = [2021, 2022, 2023, 2024, 2025]
    last_el = anchor._element
    for y in years:
        if mode == "tv":
            hdr = year_header_tv(y)
            detail = cptc_line(y)
        else:
            hdr = year_header_td(y)
            # TD: QLDN/CPBH no diff all years per DTCHEP
            detail = "*/ Không chênh lệch."
        last_el = insert_para_after(type("P", (), {"_element": last_el, "paragraph_format": anchor.paragraph_format, "runs": []})(), hdr, size=13)
        # fix: insert_para_after expects a paragraph-like with _element; use simple wrapper
    # The above wrapper hack is messy because insert_para_after uses after_p._element and clone_pPr.
    # Redo cleanly below.
    return i_idx, ii_idx, yk_idx, mode


class _P:
    def __init__(self, el, style_src):
        self._element = el
        self.paragraph_format = style_src.paragraph_format
        self.runs = []


def rebuild_years_and_closing(doc, ii_idx: int, yk_idx: int, mode: str, pages: int):
    paras = doc.paragraphs
    ii_p = paras[ii_idx]
    yk_p = paras[yk_idx]

    # remove between II and Ý kiến again if needed (idempotent caller)
    el = ii_p._element.getnext()
    yk_el = yk_p._element
    while el is not None and el is not yk_el:
        nxt = el.getnext()
        if el.tag == qn("w:p"):
            el.getparent().remove(el)
        el = nxt

    last = _P(ii_p._element, ii_p)
    years = [2021, 2022, 2023, 2024, 2025]
    for y in years:
        if mode == "tv":
            hdr = year_header_tv(y)
            detail = cptc_line(y)
        else:
            hdr = year_header_td(y)
            detail = "*/ Không chênh lệch."
        hdr_el = insert_para_after(last, hdr, size=13, align=WD_ALIGN_PARAGRAPH.LEFT)
        last = _P(hdr_el, ii_p)
        det_el = insert_para_after(last, detail, size=13, align=WD_ALIGN_PARAGRAPH.LEFT)
        last = _P(det_el, ii_p)

    c1 = CLOSING_1_TMPL.format(pages=f"{pages:02d}" if pages < 10 else str(pages))
    # template uses "02 trang" with leading zero-ish; use plain number
    c1 = CLOSING_1_TMPL.format(pages=str(pages).zfill(2) if pages < 10 else str(pages))
    e1 = insert_para_after(last, c1, size=13, align=WD_ALIGN_PARAGRAPH.LEFT)
    last = _P(e1, ii_p)
    e2 = insert_para_after(last, CLOSING_2, size=13, align=WD_ALIGN_PARAGRAPH.LEFT)
    return e1, e2


def fill_section_i(doc, i_idx: int, mode: str):
    content_p = doc.paragraphs[i_idx + 1]
    # If next para is already II, need to insert
    t = (content_p.text or "").strip()
    if t.startswith("II."):
        el = insert_para_after(
            doc.paragraphs[i_idx],
            "- Chi phí tài chính." if mode == "tv" else "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.",
            size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )
        return
    if mode == "tv":
        set_para_text(content_p, "- Chi phí tài chính.", size=13, align=WD_ALIGN_PARAGRAPH.LEFT)
    else:
        set_para_text(
            content_p,
            "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.",
            size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )


def update_closing_pages_in_text(doc, marker_prefix: str, pages: int):
    n = 0
    for p in doc.paragraphs:
        t = p.text or ""
        if t.startswith("Biên bản này được lập thành 03 bản"):
            new_t = CLOSING_1_TMPL.format(pages=str(pages).zfill(2) if pages < 10 else str(pages))
            set_para_text(p, new_t, size=13, align=WD_ALIGN_PARAGRAPH.LEFT)
            n += 1
    return n


def count_section_pages(path: Path, start_text: str, end_text: str, occurrence: int = 1) -> int:
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        wdoc = word.Documents.Open(str(path), ReadOnly=True)
        wdoc.Repaginate()
        # find start occurrence
        rng = wdoc.Content
        f = rng.Find
        f.ClearFormatting()
        f.Text = start_text
        found_n = 0
        start_page = end_page = None
        # Word Find loops
        while f.Execute():
            found_n += 1
            if found_n == occurrence:
                start_page = int(rng.Information(3))
                break
        if start_page is None:
            wdoc.Close(False)
            return 2
        # find end after start
        rng2 = wdoc.Range(rng.End, wdoc.Content.End)
        f2 = rng2.Find
        f2.Text = end_text
        if f2.Execute():
            end_page = int(rng2.Information(3))
        else:
            end_page = start_page
        wdoc.Close(False)
        pages = max(1, end_page - start_page + 1)
        # if end marker is the next doc title on same page start, still 1
        return pages
    finally:
        try:
            word.Quit()
        except Exception:
            pass


def main():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = BB / f"TONG HOP.bak-qdkt-bbxn-{stamp}.docx"
    shutil.copy2(LIVE, bak)
    print("backup", bak.name)

    doc = Document(str(LIVE))
    fix_qdkt_format(doc, "TONG-HOP-QDKT")

    tv_s, tv_e, td_s, td_e = find_section_ranges(doc)
    print("sections", tv_s, tv_e, td_s, td_e)

    # markers inside each section
    def markers(start, end):
        i_idx = ii_idx = yk_idx = None
        for i in range(start, end):
            t = (doc.paragraphs[i].text or "").strip()
            if t.startswith("I. Nội dung kiểm tra"):
                i_idx = i
            elif t.startswith("II. Kết quả kiểm tra"):
                ii_idx = i
            elif t.startswith("Ý kiến bổ sung"):
                yk_idx = i
                break
        return i_idx, ii_idx, yk_idx

    # TV first (indices stable until we edit)
    i1, ii1, yk1 = markers(tv_s, tv_e)
    print("TV markers", i1, ii1, yk1)
    fill_section_i(doc, i1, "tv")
    # after possible insert, re-find
    tv_s, tv_e, td_s, td_e = find_section_ranges(doc)
    i1, ii1, yk1 = markers(tv_s, tv_e)
    rebuild_years_and_closing(doc, ii1, yk1, "tv", pages=2)

    # TD
    tv_s, tv_e, td_s, td_e = find_section_ranges(doc)
    i2, ii2, yk2 = markers(td_s, td_e)
    print("TD markers", i2, ii2, yk2)
    fill_section_i(doc, i2, "td")
    tv_s, tv_e, td_s, td_e = find_section_ranges(doc)
    i2, ii2, yk2 = markers(td_s, td_e)
    rebuild_years_and_closing(doc, ii2, yk2, "td", pages=2)

    doc.save(str(LIVE))
    print("saved preliminary TONG HOP")

    # Count real pages for each BB section
    # TV: from first "Xác nhận số liệu kiểm tra" to second one
    # TD: from second to "BIÊN BẢN KIỂM TRA"
    try:
        tv_pages = count_section_pages(LIVE, "Xác nhận số liệu kiểm tra", "Xác nhận số liệu kiểm tra", occurrence=1)
        # For TV end marker is the NEXT same title; Find from after first will get second.
        # Our function finds start occurrence then end_text from after start - for TV end_text same title works for 2nd.
        # But occurrence logic: start=1st, end=next "Xác nhận..." which is 2nd title = start of TD, so pages of TV ok.
        td_pages = count_section_pages(LIVE, "Xác nhận số liệu kiểm tra", "BIÊN BẢN KIỂM TRA", occurrence=2)
        print("page counts tv/td", tv_pages, td_pages)
    except Exception as e:
        print("page count failed", e)
        tv_pages, td_pages = 2, 2

    # Update page numbers in the two closing lines (in order)
    doc = Document(str(LIVE))
    closings = []
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t.startswith("Biên bản này được lập thành 03 bản"):
            closings.append(i)
    print("closing paras", closings)
    if len(closings) >= 1:
        set_para_text(
            doc.paragraphs[closings[0]],
            CLOSING_1_TMPL.format(pages=str(tv_pages).zfill(2)),
            size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )
    if len(closings) >= 2:
        set_para_text(
            doc.paragraphs[closings[1]],
            CLOSING_1_TMPL.format(pages=str(td_pages).zfill(2)),
            size=13,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )

    # ensure closing_2 exists before each ý kiến in those sections
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t.startswith("Ý kiến bổ sung"):
            prev = (doc.paragraphs[i - 1].text or "").strip() if i else ""
            if not prev.startswith("Sau khi đọc lại biên bản"):
                insert_para_after(doc.paragraphs[i - 1], CLOSING_2, size=13, align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.save(str(LIVE))
    print("saved TONG HOP with page counts")

    # standalone QDKT
    d = Document(str(QDKT))
    fix_qdkt_format(d, "QDKT-file")
    d.save(str(QDKT))
    print("saved QDKT")

    with ZipFile(ZIPP, "w", ZIP_DEFLATED) as z:
        for f in (LIVE, QDKT, QDGS):
            z.write(f, f.name)
    print("zip ok")

    # verify
    v = Document(str(LIVE))
    print("\n=== VERIFY QDKT Dieu 1 ===")
    for i, p in enumerate(v.paragraphs):
        t = (p.text or "").strip()
        if t.startswith("Điều 1. Kiểm tra"):
            print("align", p.alignment, "runs", [(r.text[:40], r.bold) for r in p.runs])
            break
    print("\n=== VERIFY BB TV/TD snippets ===")
    for i, p in enumerate(v.paragraphs):
        t = (p.text or "").strip()
        if "Xác nhận số liệu kiểm tra" == t or t.startswith("I. Nội dung") or t.startswith("+ NĂM") or t.startswith("Năm 202") or t.startswith("*/") or t.startswith("Biên bản này được lập") or t.startswith("Sau khi đọc lại"):
            if 190 <= i <= 290:
                print(f"P{i}: {t[:140]}")


if __name__ == "__main__":
    main()
