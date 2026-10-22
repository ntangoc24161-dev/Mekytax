# -*- coding: utf-8 -*-
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
import os
import re
import zipfile

BASE = (
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
)
PATH = os.path.join(BASE, "TONG HOP.docx")


def replace_para(p, nt: str):
    if p.runs:
        p.runs[0].text = nt
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = nt


def insert_after(paragraph, texts):
    ref = paragraph._p
    for text in texts:
        new_p = deepcopy(paragraph._p)
        ts = new_p.findall(".//" + qn("w:t"))
        if ts:
            ts[0].text = text
            for t in ts[1:]:
                t.text = ""
        ref.addnext(new_p)
        ref = new_p


def ensure_year_2024(doc: Document, header_2023: str, header_2024: str, header_2025_prefix: str):
    paras = list(doc.paragraphs)
    for i, p in enumerate(paras):
        if p.text.strip() != header_2023:
            continue
        if i + 1 >= len(paras) or "Không chênh lệch" not in paras[i + 1].text:
            continue
        j = i + 2
        while j < len(paras) and not paras[j].text.strip():
            j += 1
        if j >= len(paras):
            continue
        nxt = paras[j].text.strip()
        if nxt.startswith(header_2025_prefix):
            # already has 2024?
            mid = " ".join(paras[k].text for k in range(i, j))
            if "2024" in mid:
                print("already has 2024 after", header_2023)
                return False
            insert_after(
                paras[i + 1],
                [header_2024, "*/ Không chênh lệch."],
            )
            print("inserted 2024 after", header_2023)
            return True
    return False


def main():
    doc = Document(PATH)
    ensure_year_2024(
        doc,
        "+ NĂM 2023: Đơn vị tính: Đồng",
        "+ NĂM 2024: Đơn vị tính: Đồng",
        "+ NĂM 2025",
    )
    doc.save(PATH)

    doc = Document(PATH)
    ensure_year_2024(
        doc,
        "Năm 2023 Đơn vị tính: Đồng",
        "Năm 2024 Đơn vị tính: Đồng",
        "Năm 2025",
    )
    doc.save(PATH)

    doc = Document(PATH)
    print("--- verify ---")
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if not t:
            continue
        if re.search(r"(NĂM|Năm)\s*202[1-5]", t) or t.startswith("*/"):
            if 190 <= i <= 310:
                print(f"P{i}: {t[:150]}")

    blob = "\n".join(p.text for p in doc.paragraphs)
    print("EBITDA", "EBITDA" in blob)
    print("TT96", "Thông tư số 96" in blob)
    print("lai vay count", blob.count("hạch toán chi phí lãi vay không đúng quy định"))
    print("thoi han 20", "Thời hạn kiểm tra là 20 ngày" in blob)

    for name in [
        "QD Kiem tra tai cty Van Phat Thanh.docx",
        "QDgiam sat tai cty Van Phat Thanh.docx",
    ]:
        dpath = os.path.join(BASE, name)
        d = Document(dpath)
        changed = False
        for p in d.paragraphs:
            if "Thời hạn kiểm tra là 10 ngày" in p.text:
                replace_para(p, p.text.replace("10 ngày", "20 ngày"))
                changed = True
        if changed:
            d.save(dpath)
            print("fixed thoi han in", name)

    with zipfile.ZipFile(os.path.join(BASE, "hoso.zip"), "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in [
            "TONG HOP.docx",
            "QD Kiem tra tai cty Van Phat Thanh.docx",
            "QDgiam sat tai cty Van Phat Thanh.docx",
        ]:
            zf.write(os.path.join(BASE, fname), fname)
    print("zip ok")


if __name__ == "__main__":
    main()
