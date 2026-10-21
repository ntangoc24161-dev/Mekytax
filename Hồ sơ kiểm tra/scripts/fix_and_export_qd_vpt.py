# -*- coding: utf-8 -*-
"""Post-fix TONG HOP: thoi han, QĐGS blanks, double dots; re-export QĐs by markers."""
from __future__ import annotations

import copy
import os
import re
import sys
import zipfile

import docx
from docx import Document
from docx.shared import Cm, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_standardized_ho_so as base

TARGET = (
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
)
TONG = os.path.join(TARGET, "TONG HOP.docx")


def replace_paragraph_text(paragraph, new_text: str):
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.text = new_text


def fix_tong_hop(doc: Document):
    for p in doc.paragraphs:
        t = p.text
        if not t.strip():
            continue
        if "Thời hạn kiểm tra là 10 ngày làm việc" in t:
            replace_paragraph_text(p, t.replace("10 ngày", "20 ngày"))
            print("fix thoi han -> 20")
        if ".............../QĐ-TCS3" in t or "................./QĐ-TCS3" in t:
            nt = t.replace(".............../QĐ-TCS3", "6592/QĐ-TCS3")
            nt = nt.replace("................./QĐ-TCS3", "6592/QĐ-TCS3")
            nt = nt.replace("...../...../2026", "02/10/2026")
            replace_paragraph_text(p, nt)
            print("fix QĐGS blank:", p.text[:90])
        if ".." in t and "Chi phí" in t:
            nt = t
            while ".." in nt:
                nt = nt.replace("..", ".")
            replace_paragraph_text(p, nt)
            print("fix dots:", repr(nt[:90]))
        if t.strip().startswith("- - Chi phí tài chính"):
            replace_paragraph_text(p, "")
            print("cleared leftover phan cong")
        if "Trấn Biên, ngày" in t and ("….." in t or "....." in t):
            replace_paragraph_text(p, "Trấn Biên, ngày 02 tháng 10 năm 2026")
            print("fix Tran Bien ellipsis")

    for ti, table in enumerate(doc.tables[:6]):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pt = p.text.strip()
                    if pt.startswith("Trấn Biên") and not re.search(
                        r"ngày\s+\d{1,2}\s+tháng\s+\d{1,2}", pt
                    ):
                        replace_paragraph_text(p, "Trấn Biên, ngày 02 tháng 10 năm 2026")
                        print(f"T{ti} fix date cell")


def find_qd_ranges(master_doc):
    body = list(master_doc._element.body)
    qdkt_start = qdkt_end = qdgs_start = qdgs_end = None
    for i, el in enumerate(body):
        text = "".join(el.itertext())
        if qdkt_start is None and el.tag.endswith("tbl") and "THUẾ THÀNH PHỐ ĐỒNG NAI" in text and i > 20:
            nxt = "".join("".join(body[j].itertext()) for j in range(i, min(i + 5, len(body))))
            if "Về việc kiểm tra thuế tại" in nxt:
                qdkt_start = i
        if qdkt_start is not None and qdkt_end is None and el.tag.endswith("tbl") and "Nơi nhận" in text and "Điều 3" in text:
            qdkt_end = i
        if qdkt_end is not None and qdgs_start is None and el.tag.endswith("tbl") and "THUẾ THÀNH PHỐ ĐỒNG NAI" in text:
            nxt = "".join("".join(body[j].itertext()) for j in range(i, min(i + 5, len(body))))
            if "Về việc giám sát" in nxt:
                qdgs_start = i
        if qdgs_start is not None and qdgs_end is None and el.tag.endswith("tbl") and "Nơi nhận" in text and "Điều 4" in text:
            qdgs_end = i
            break
    print("ranges", qdkt_start, qdkt_end, qdgs_start, qdgs_end)
    return qdkt_start, qdkt_end, qdgs_start, qdgs_end


def export_qd(master_doc, start, end, output_path, kind):
    doc = docx.Document(base.TEMPLATE_PATH)
    body = doc._element.body
    sectPr = None
    for el in list(body):
        if el.tag.endswith("sectPr"):
            sectPr = copy.deepcopy(el)
            break
    all_elems = list(master_doc._element.body)
    segment = [copy.deepcopy(all_elems[i]) for i in range(start, end + 1)]
    body.clear()
    for el in segment:
        body.append(el)
    if sectPr is not None:
        body.append(sectPr)

    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2.7)
    sec.right_margin = Cm(1.8)

    for t in doc.tables:
        base.remove_floating_table_properties(t)
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Times New Roman"

    for p in list(doc.paragraphs):
        text = p.text.strip()
        if not text:
            if p._element.getparent() is not None:
                p._element.getparent().remove(p._element)
            continue
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        for r in p.runs:
            r.font.name = "Times New Roman"

    for p in doc.paragraphs:
        t = p.text
        if "10 ngày làm việc" in t and kind == "kt":
            replace_paragraph_text(p, t.replace("10 ngày", "20 ngày"))
        if ".............../QĐ-TCS3" in t or "................./QĐ-TCS3" in t:
            nt = t.replace(".............../QĐ-TCS3", "6592/QĐ-TCS3")
            nt = nt.replace("................./QĐ-TCS3", "6592/QĐ-TCS3")
            nt = nt.replace("...../...../2026", "02/10/2026")
            replace_paragraph_text(p, nt)
        if "Chi phí" in t and ".." in t:
            nt = t
            while ".." in nt:
                nt = nt.replace("..", ".")
            replace_paragraph_text(p, nt)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pt = p.text.strip()
                    if pt.startswith("Trấn Biên") and not re.search(
                        r"ngày\s+\d{1,2}\s+tháng\s+\d{1,2}", pt
                    ):
                        replace_paragraph_text(p, "Trấn Biên, ngày 02 tháng 10 năm 2026")
                if cell.text.strip() in ("Số:", "Số"):
                    num = "6592/QĐ-TCS3" if kind == "kt" else "6593/QĐ-TCS3"
                    if len(cell.paragraphs) == 1:
                        replace_paragraph_text(cell.paragraphs[0], f"Số: {num}")
                    else:
                        replace_paragraph_text(cell.paragraphs[0], "Số:")
                        replace_paragraph_text(cell.paragraphs[1], num)
                    print(kind, "filled Số", num)

    doc.save(output_path)
    print("exported", output_path)


def main():
    doc = Document(TONG)
    fix_tong_hop(doc)
    doc.save(TONG)
    print("saved TONG HOP")

    doc = Document(TONG)
    s1, e1, s2, e2 = find_qd_ranges(doc)
    if None in (s1, e1, s2, e2):
        raise SystemExit(f"missing ranges {(s1, e1, s2, e2)}")
    qd_kt = os.path.join(TARGET, "QD Kiem tra tai cty Van Phat Thanh.docx")
    qd_gs = os.path.join(TARGET, "QDgiam sat tai cty Van Phat Thanh.docx")
    export_qd(doc, s1, e1, qd_kt, "kt")
    export_qd(doc, s2, e2, qd_gs, "gs")

    zip_path = os.path.join(TARGET, "hoso.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in [
            "TONG HOP.docx",
            "QD Kiem tra tai cty Van Phat Thanh.docx",
            "QDgiam sat tai cty Van Phat Thanh.docx",
        ]:
            zf.write(os.path.join(TARGET, fname), fname)
    print("zip ok")

    d = Document(TONG)
    for p in d.paragraphs:
        if "Thời hạn kiểm tra là" in p.text:
            print("VERIFY thoi han:", p.text[:100])
            break
    for p in d.paragraphs:
        if "Căn cứ Quyết định kiểm tra số" in p.text:
            print("VERIFY GS:", p.text[:130])
            break
    for name in [
        "QD Kiem tra tai cty Van Phat Thanh.docx",
        "QDgiam sat tai cty Van Phat Thanh.docx",
    ]:
        dd = Document(os.path.join(TARGET, name))
        blob = "\n".join(p.text for p in dd.paragraphs)
        print(name, "paras", len(dd.paragraphs))
        print(
            "  6592",
            "6592" in blob,
            "| 10 ngày",
            "10 ngày" in blob,
            "| 20 ngày",
            "20 ngày" in blob,
        )
        print("  first:", next((p.text[:90] for p in dd.paragraphs if p.text.strip()), ""))
        print("  ve viec:", next((p.text[:90] for p in dd.paragraphs if "Về việc" in p.text), ""))


if __name__ == "__main__":
    main()
