# -*- coding: utf-8 -*-
"""Sua ten/MST cac dong lo 22-26/09 va OCR 6 ban scan bang PaddleOCR."""
from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

import openpyxl
import pymupdf

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from fast_batch_ocr import init_ocr, ocr_page_to_text  # noqa: E402
from update_so_theo_doi import DEFAULT_EXCEL_PATH  # noqa: E402

RECORDS_PATH = SCRIPTS / "records_batch_22_26.json"
OUT_DIR = Path(r"D:\HO SO CONG VAN CANH BAO")

SCAN_FILES = [
    r"D:\HO SO CONG VAN CANH BAO\[14036]_6621_TCS15-QLDN3_6621TCS15-QLDN3 14036.pdf",
    r"D:\HO SO CONG VAN CANH BAO\[14034]_612_TCS4-QLDN_612TCS4-QLDN 14034.pdf",
    r"D:\HO SO CONG VAN CANH BAO\[14031]_19138_TCS28-QLDN_19138TCS28-QLDN 14031.pdf",
    r"D:\HO SO CONG VAN CANH BAO\[14028]_613_TCS4-QLDN_613TCS4-QLDN 14028.pdf",
    r"D:\HO SO CONG VAN CANH BAO\[14860]_1168_TCS26-QLDN_5180.pdf",
    r"D:\HO SO CONG VAN CANH BAO\[14009]_608_TCS4-QLDN_608TCS4-QLDN 14009.pdf",
]

BAD_NAME = re.compile(
    r"dấu hiệu|rủi ro|không hoạt động|tại hệ thống|theo CV |và hóa đơn đối với",
    re.I,
)
LABEL_PREFIX = re.compile(
    r"^(?:Tên người nộp thuế|Tên NNT|Tên doanh nghiệp)\s*:\s*",
    re.I,
)


def log(msg: str) -> None:
    print(msg, flush=True)


def squash(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def extract_from_text(body: str, trich_yeu: str = "", filename: str = "") -> dict:
    blob = squash(body)
    ty = squash(trich_yeu)
    fn = squash(filename)

    ten = ""
    m = re.search(
        r"Tên(?:\s+người\s+nộp\s+thuế|\s+NNT|\s+doanh nghiệp)\s*:\s*"
        r"(.+?)(?:\s*\(sau|;\s*Mã số thuế|;\s*MST|;\s*Địa chỉ|$)",
        blob,
        re.I,
    )
    if m:
        ten = LABEL_PREFIX.sub("", m.group(1)).strip(" ;,.")
        ten = re.sub(r"\s*\(sau đây.*$", "", ten, flags=re.I).strip(" ;,.")

    if not ten or BAD_NAME.search(ten) or len(ten) < 8:
        m = re.search(
            r"((?:CÔNG TY|Công ty|CTY|Công Ty)\s+(?:TNHH|CP|Cổ phần|MTV|Trách Nhiệm)[^,;]{3,80})",
            blob + " " + ty + " " + fn,
            re.I,
        )
        if m:
            ten = m.group(1).strip(" ;,.")

    if not ten or BAD_NAME.search(ten):
        m = re.search(
            r"đối với\s+((?:CÔNG TY|Công ty|CTY)[^,;]{5,80})",
            ty,
            re.I,
        )
        if m:
            ten = m.group(1).strip(" ;,.")

    if ten:
        ten = re.sub(r"[\s:;.-]+(?:mst|mã số thuế).*$", "", ten, flags=re.I)
        ten = re.sub(r"[-–]\s*\d{10}(?:-\d{3})?\s*$", "", ten)
        ten = ten.strip(" ;,.")
        if BAD_NAME.search(ten) or ten.lower().startswith("doanh nghiệp có"):
            ten = ""

    mst = ""
    m = re.search(r"(?:Mã số thuế|MST)\s*[:.]?\s*(\d{10}(?:-\d{3})?)", blob + " " + ty, re.I)
    if m:
        mst = m.group(1)
    elif re.search(r"\b(\d{10})\b", ty):
        mst = re.search(r"\b(\d{10})\b", ty).group(1)

    dia = ""
    m = re.search(
        r"(?:Địa chỉ(?: kinh doanh| trụ sở)?|trụ sở)\s*:\s*([^;]{8,140})",
        blob,
        re.I,
    )
    if m:
        dia = m.group(1).strip(" ;,.")

    reasons = []
    if re.search(r"không hoạt động tại địa chỉ", blob, re.I):
        reasons.append("Không hoạt động tại địa chỉ đăng ký")
    if re.search(r"chưa nộp tờ khai|không nộp tờ khai", blob, re.I):
        reasons.append("Chưa nộp tờ khai thuế")
    if re.search(r"chênh lệch", blob, re.I):
        reasons.append("Chênh lệch số liệu tờ khai và HĐĐT")
    if re.search(r"cơ quan điều tra|chuyển tin báo", blob, re.I):
        reasons.append("Chuyển tin báo CQĐT")
    if re.search(r"tạm ngừng kinh doanh", blob, re.I):
        reasons.append("Tạm ngừng kinh doanh")
    if re.search(r"gian lận|mua bán hóa đơn", blob, re.I):
        reasons.append("Nghi vấn gian lận hóa đơn")
    noi = "; ".join(reasons) if reasons else (
        "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn"
    )
    return {"ten_ban": ten, "mst_ban": mst, "dia_chi": dia, "noi_dung": noi}


def ocr_scans() -> dict[str, str]:
    ocr = init_ocr()
    out = {}
    for pdf in SCAN_FILES:
        p = Path(pdf)
        if not p.exists():
            log(f"[OCR] thieu {p.name}")
            continue
        md_path = p.with_suffix(".md")
        log(f"[OCR] {p.name}")
        doc = pymupdf.open(str(p))
        pages_text = []
        n = min(doc.page_count, 3)
        for i in range(n):
            page = doc[i]
            # crop 70% phan tren trang dau
            if i == 0:
                r = page.rect
                page = page  # full page; ocr_page_to_text uses pixmap of page
            txt = ocr_page_to_text(ocr, page)
            pages_text.append(f"## Trang {i+1}\n{txt}")
            log(f"    trang {i+1}/{n}: {len(txt)} ky tu")
        doc.close()
        body = "\n\n".join(pages_text)
        md_path.write_text(f"# {p.stem}\n\n{body}\n", encoding="utf-8")
        out[str(p)] = body
        log(f"  [MD] {md_path.name} ({len(body)} ky tu)")
    return out


def main() -> int:
    records = json.loads(RECORDS_PATH.read_text(encoding="utf-8"))
    scan_text = ocr_scans()

    for rec in records:
        pdf = rec.get("file") or ""
        md = Path(pdf).with_suffix(".md")
        body = ""
        if pdf in scan_text:
            body = scan_text[pdf]
        elif md.exists():
            body = md.read_text(encoding="utf-8", errors="ignore")
        info = extract_from_text(body, rec.get("ghi_chu", ""), os.path.basename(pdf))
        if info["ten_ban"]:
            rec["ten_ban"] = info["ten_ban"]
        if info["mst_ban"]:
            rec["mst_ban"] = info["mst_ban"]
        if info["dia_chi"]:
            rec["dia_chi"] = info["dia_chi"]
        if info["noi_dung"] and not rec.get("noi_dung", "").startswith("Cảnh báo doanh nghiệp có dấu hiệu"):
            rec["noi_dung"] = info["noi_dung"]
        elif info["noi_dung"]:
            rec["noi_dung"] = info["noi_dung"]
        log(f"FIX {rec['so_cv']}: {rec.get('ten_ban','')[:70]} | {rec.get('mst_ban')}")

    RECORDS_PATH.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")

    by_cv = {re.sub(r"\s+", "", (r["so_cv"] or "")).upper(): r for r in records}
    wb = openpyxl.load_workbook(DEFAULT_EXCEL_PATH)
    ws = wb["THEO_DOI"]
    updated = 0
    for r in range(28336, ws.max_row + 1):
        socv = re.sub(r"\s+", "", str(ws.cell(r, 2).value or "")).upper()
        rec = by_cv.get(socv)
        if not rec:
            continue
        ws.cell(r, 5, rec.get("ten_ban"))
        ws.cell(r, 6, rec.get("mst_ban"))
        ws.cell(r, 7, rec.get("noi_dung"))
        if rec.get("dia_chi"):
            ws.cell(r, 16, rec.get("dia_chi"))
        updated += 1
        log(f"ROW {r} {socv} <- {rec.get('ten_ban')}")
    wb.save(DEFAULT_EXCEL_PATH)
    log(f"[DONE] cap nhat {updated} dong Excel")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
