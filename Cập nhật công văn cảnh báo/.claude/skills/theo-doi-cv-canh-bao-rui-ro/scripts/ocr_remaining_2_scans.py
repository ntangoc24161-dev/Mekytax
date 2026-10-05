# -*- coding: utf-8 -*-
"""OCR tiep 2 ban scan con stub (1168, 608), khong chay lai 4 ban da xong."""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import openpyxl
import pymupdf

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
from fast_batch_ocr import init_ocr, ocr_page_to_text  # noqa: E402
from update_so_theo_doi import DEFAULT_EXCEL_PATH  # noqa: E402
from repair_batch_22_26 import extract_from_text  # noqa: E402

RECORDS_PATH = SCRIPTS / "records_batch_22_26.json"
LOG_PATH = SCRIPTS / "run_ocr_remaining_2.log"
MAX_PAGES = 2

SCAN_FILES = [
    {
        "so_cv": "1168/TCS26-QLDN",
        "pdf": r"D:\HO SO CONG VAN CANH BAO\[14860]_1168_TCS26-QLDN_5180.pdf",
        "co_quan": "Thuế cơ sở 26 TP. Hồ Chí Minh",
        "ngay_vb": "23/09/2026",
    },
    {
        "so_cv": "608/TCS4-QLDN",
        "pdf": r"D:\HO SO CONG VAN CANH BAO\[14009]_608_TCS4-QLDN_608TCS4-QLDN 14009.pdf",
        "co_quan": "Thuế cơ sở 4 tỉnh Tây Ninh",
        "ngay_vb": "18/09/2026",
    },
]


def log(msg: str) -> None:
    print(msg, flush=True)
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(str(msg) + "\n")


def norm_socv(s: str) -> str:
    return re.sub(r"\s+", "", (s or "")).upper().replace("_", "/")


def write_md(path_md: Path, meta: dict, body: str) -> None:
    lines = [
        f"# {meta.get('so_cv', '')}",
        "",
        f"- Cơ quan: {meta.get('co_quan', '')}",
        f"- Ngày ban hành: {meta.get('ngay_vb', '')}",
        f"- Tên NNT: {meta.get('ten_ban', '')}",
        f"- MST: {meta.get('mst_ban', '')}",
        f"- Địa chỉ: {meta.get('dia_chi', '')}",
        f"- Nội dung: {meta.get('noi_dung', '')}",
        "",
        "## OCR (offline, bỏ bảng biểu)",
        "",
        body.strip(),
        "",
    ]
    path_md.write_text("\n".join(lines), encoding="utf-8")


def extract_noi_dung(body: str) -> str:
    reasons = []
    if re.search(r"không hoạt động tại địa chỉ", body, re.I):
        reasons.append("Không hoạt động tại địa chỉ đăng ký")
    if re.search(r"chưa nộp tờ khai|không nộp tờ khai", body, re.I):
        reasons.append("Chưa nộp tờ khai thuế")
    if re.search(r"chênh lệch", body, re.I):
        reasons.append("Chênh lệch số liệu tờ khai và HĐĐT")
    if re.search(r"cơ quan điều tra|chuyển tin báo", body, re.I):
        reasons.append("Chuyển tin báo CQĐT")
    if re.search(r"tạm ngừng kinh doanh", body, re.I):
        reasons.append("Tạm ngừng kinh doanh")
    if re.search(r"gian lận|mua bán hóa đơn", body, re.I):
        reasons.append("Nghi vấn gian lận hóa đơn")
    return "; ".join(reasons) if reasons else (
        "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn"
    )


def patch_excel(updates: dict[str, dict]) -> int:
    excel = DEFAULT_EXCEL_PATH
    log(f"[EXCEL] ghi de: {excel}")
    wb = openpyxl.load_workbook(excel)
    ws = wb["THEO_DOI"]
    want = {norm_socv(k): v for k, v in updates.items()}
    patched = 0
    for row_idx in range(8, ws.max_row + 1):
        socv = norm_socv(str(ws.cell(row_idx, 2).value or ""))
        if socv not in want:
            continue
        u = want[socv]
        if u.get("ten_ban"):
            ws.cell(row_idx, 5).value = u["ten_ban"]
        if u.get("mst_ban"):
            ws.cell(row_idx, 6).value = u["mst_ban"]
        if u.get("noi_dung"):
            ws.cell(row_idx, 7).value = u["noi_dung"]
        if u.get("dia_chi"):
            ws.cell(row_idx, 16).value = u["dia_chi"]
        patched += 1
        log(f"  [PATCH] row={row_idx} {socv} MST={u.get('mst_ban')} | {u.get('ten_ban')}")
    wb.save(excel)
    wb.close()
    return patched


def main() -> int:
    LOG_PATH.write_text("", encoding="utf-8")
    log("=== OCR REMAINING 2 SCANS (1168, 608) ===")
    ocr = init_ocr()
    records = json.loads(RECORDS_PATH.read_text(encoding="utf-8"))
    by_socv = {norm_socv(r.get("so_cv", "")): r for r in records}
    updates: dict[str, dict] = {}

    for i, item in enumerate(SCAN_FILES, 1):
        pdf = Path(item["pdf"])
        md = pdf.with_suffix(".md")
        log(f"[{i}/2] {item['so_cv']} {pdf.name}")
        if not pdf.exists():
            log(f"  [MISS] {pdf}")
            return 1
        doc = pymupdf.open(str(pdf))
        parts = []
        n = min(MAX_PAGES, doc.page_count)
        for pno in range(n):
            text = ocr_page_to_text(ocr, doc[pno])
            parts.append(f"### Trang {pno + 1}\n{text}")
            log(f"  [OCR] page {pno + 1}/{doc.page_count} chars={len(text)}")
        doc.close()
        body = "\n\n".join(parts)
        extracted = extract_from_text(body, filename=pdf.name)
        old = by_socv.get(norm_socv(item["so_cv"]), {})
        ten = re.sub(r"\s+", " ", (extracted.get("ten_ban") or old.get("ten_ban") or "")).strip(" ,.;:-")
        mst = extracted.get("mst_ban") or old.get("mst_ban") or ""
        dia = extracted.get("dia_chi") or old.get("dia_chi") or ""
        noi = extract_noi_dung(body) or old.get("noi_dung") or ""
        meta = {
            "so_cv": item["so_cv"],
            "co_quan": item.get("co_quan", old.get("don_vi_bh", "")),
            "ngay_vb": item.get("ngay_vb", old.get("ngay_cv", "")),
            "ten_ban": ten,
            "mst_ban": mst,
            "dia_chi": dia,
            "noi_dung": noi,
        }
        write_md(md, meta, body)
        log(f"  [MD] {md.name} size={md.stat().st_size}")
        updates[item["so_cv"]] = meta
        key = norm_socv(item["so_cv"])
        if key in by_socv:
            by_socv[key].update(
                {"ten_ban": ten, "mst_ban": mst, "dia_chi": dia, "noi_dung": noi}
            )

    RECORDS_PATH.write_text(
        json.dumps(list(by_socv.values()), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    n = patch_excel(updates)
    log(f"[DONE] ocr={len(updates)} excel_patched={n}")
    return 0 if len(updates) == 2 and n == 2 else 1


if __name__ == "__main__":
    raise SystemExit(main())
