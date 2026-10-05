# -*- coding: utf-8 -*-
"""
Xu ly lo cong van canh bao 22-26/09/2026:
- Loc trung cung noi dung (hash + so ky hieu)
- Bo CV da co tren so THEO_DOI
- Chuyen PDF sang Markdown
- Boc tach va nap Excel
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
import pymupdf

JAVIS_OS = Path(r"D:\SECOND BRAIN\javis-os")
BRAIN = Path(r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo")
SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

from update_so_theo_doi import DEFAULT_EXCEL_PATH, update_workbook  # noqa: E402

PROGRESS_PATH = BRAIN / "download_progress.json"
OUT_DIR = Path(r"D:\HO SO CONG VAN CANH BAO")
RECORDS_PATH = SCRIPTS / "records_batch_22_26.json"
PLAN_PATH = SCRIPTS / "plan_batch_22_26.json"
OCR_PY = JAVIS_OS / ".venv-paddleocr" / "Scripts" / "python.exe"

ALREADY_IN_BOOK = {"4482/TCS1-QLDN", "8692/TCS10-QLDN1"}

MAP_DON_VI = {
    "THUẾ CƠ SỞ 1": "Thuế cơ sở 1 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 2": "Thuế cơ sở 2 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 3": "Thuế cơ sở 3 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 4": "Thuế cơ sở 4 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 5": "Thuế cơ sở 5 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 6": "Thuế cơ sở 6 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 7": "Thuế cơ sở 7 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 8": "Thuế cơ sở 8 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 9": "Thuế cơ sở 9 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 10": "Thuế cơ sở 10 thành phố Đồng Nai",
    "THUẾ THÀNH PHỐ ĐỒNG NAI": "Thuế thành phố Đồng Nai",
    "THUẾ CS26 TP.HCM": "Thuế cơ sở 26 TP. Hồ Chí Minh",
    "THUẾ CƠ SỞ 3 TỈNH GIA LAI": "Thuế cơ sở 3 tỉnh Gia Lai",
    "THUẾ CƠ SỞ 4 TỈNH TÂY NINH": "Thuế cơ sở 4 tỉnh Tây Ninh",
    "THUẾ CS4 TỈNH TÂY NINH": "Thuế cơ sở 4 tỉnh Tây Ninh",
}


def log(msg: str) -> None:
    print(msg, flush=True)


def norm_socv(s: str) -> str:
    return re.sub(r"\s+", "", (s or "")).upper().replace("_", "/")


def file_sig(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read(1024 * 1024))
        f.seek(0, 2)
        h.update(str(f.tell()).encode())
    return h.hexdigest()[:16]


def pdf_stats(path: str) -> tuple[int, int]:
    if not path.lower().endswith(".pdf"):
        return 0, 0
    try:
        doc = pymupdf.open(path)
        pages = doc.page_count
        chars = 0
        for i, page in enumerate(doc):
            if i >= 2:
                break
            chars += len(page.get_text("text") or "")
        doc.close()
        return pages, chars
    except Exception:
        return 0, -1


def pdf_full_text(path: str, max_pages: int = 12) -> str:
    doc = pymupdf.open(path)
    parts = []
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        parts.append(page.get_text("text") or "")
    doc.close()
    return "\n".join(parts)


def write_md(path_md: Path, header: dict, body: str) -> None:
    lines = [
        f"# {header.get('so_ky_hieu', '')}",
        "",
        f"- Số đến: {header.get('so_den', '')}",
        f"- Ngày ban hành: {header.get('ngay_vb', '')}",
        f"- Ngày đến: {header.get('ngay_den', '')}",
        f"- Cơ quan: {header.get('co_quan', '')}",
        f"- Trích yếu: {header.get('trich_yeu', '')}",
        "",
        body.strip(),
        "",
    ]
    path_md.write_text("\n".join(lines), encoding="utf-8")


def convert_pdf_to_md(pdf_path: str, is_scan: bool) -> Path:
    md_path = Path(pdf_path).with_suffix(".md")
    if md_path.exists() and md_path.stat().st_size > 400:
        log(f"  [MD] da co {md_path.name}")
        return md_path
    if not is_scan:
        body = pdf_full_text(pdf_path)
        if len(body.strip()) >= 200:
            write_md(md_path, {}, body)
            log(f"  [MD] pymupdf {md_path.name} ({len(body)} ky tu)")
            return md_path
    backend = "paddle" if is_scan else "auto"
    cmd = [
        str(OCR_PY),
        "-m",
        "tools.doc2md",
        pdf_path,
        "-o",
        str(md_path),
        "--backend",
        backend,
        "--lang",
        "vi",
        "--force",
    ]
    log(f"  [OCR] doc2md --backend {backend} {Path(pdf_path).name}")
    proc = subprocess.run(
        cmd,
        cwd=str(JAVIS_OS),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.stdout:
        log(proc.stdout.strip()[-800:])
    if proc.returncode != 0:
        log(f"  [OCR-FAIL] code={proc.returncode} {proc.stderr[-500:]}")
    return md_path


def normalize_don_vi(dv_raw: str) -> str:
    dv = (dv_raw or "").strip()
    if not dv:
        return ""
    mapped = MAP_DON_VI.get(dv.upper())
    if mapped:
        return mapped
    dv2 = re.sub(r"THUẾ CƠ SỞ\s+(\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv2 = re.sub(r"Thuế CS\s*(\d+)", r"Thuế cơ sở \1", dv2, flags=re.IGNORECASE)
    dv2 = re.sub(r"TP\.?\s*HCM", "TP. Hồ Chí Minh", dv2, flags=re.IGNORECASE)
    return dv2


def clean_company_name(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s).strip(" ,.:;-\"'\n\r\t")
    s = re.sub(r"^(?:v/v|vv)\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(
        r"^(?:cảnh báo|công văn cảnh báo|cv cảnh báo|cbrr)[\s:.-]*",
        "",
        s,
        flags=re.IGNORECASE,
    )
    s = re.sub(
        r"^doanh nghiệp có dấu hiệu rủi ro cao về thuế[^\w]*",
        "",
        s,
        flags=re.IGNORECASE,
    )
    s = re.sub(r"^(?:đối với|về việc)\s+", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s:;.-]+(?:mst|mã số thuế).*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[-–]\s*\d{10}(?:-\d{3})?\s*$", "", s)
    return s.strip(" ,.:;-\"'")


def extract_mst(*texts: str) -> str:
    blob = "\n".join(t or "" for t in texts)
    m = re.search(
        r"(?:mã số thuế|mst|mã số doanh nghiệp)[:\s]*([0-9]{10}(?:-[0-9]{3})?)",
        blob,
        re.IGNORECASE,
    )
    if m:
        return m.group(1)
    m = re.search(r"\b(\d{10}(?:-\d{3})?)\b", blob)
    return m.group(1) if m else ""


def extract_ten_ban(trich_yeu: str, body: str, filename: str) -> str:
    sources = [trich_yeu or "", body or "", filename or ""]
    for src in sources:
        m = re.search(
            r"(?:Tên\s+người\s+nộp\s+thuế|Tên\s+doanh\s+nghiệp|Người\s+nộp\s+thuế)[:\s]*([^\n\r]{5,120})",
            src,
            re.IGNORECASE,
        )
        if m:
            cand = clean_company_name(m.group(1))
            if len(cand) > 6 and not any(
                w in cand.lower() for w in ["dấu hiệu", "rủi ro", "không hoạt động"]
            ):
                return cand
    ty = trich_yeu or ""
    m = re.search(
        r"((?:CÔNG TY|CTY|DOANH NGHIỆP|DNTN|HỘ KINH DOANH)[^,\n\r()]{3,90})",
        ty,
        re.IGNORECASE,
    )
    if m:
        cand = clean_company_name(m.group(1))
        if len(cand) > 6:
            return cand
    m = re.search(r"\(([^)]{8,90})\)", ty)
    if m:
        cand = clean_company_name(m.group(1))
        if len(cand) > 6:
            return cand
    m = re.search(
        r"((?:CÔNG TY|CTY|DOANH NGHIỆP)[^\n\r,.:;]{4,90})",
        body or "",
        re.IGNORECASE,
    )
    if m:
        cand = clean_company_name(m.group(1))
        skip = ["THÀNH PHỐ", "ĐỒNG NAI", "HẠNH PHÚC", "CỘNG HÒA"]
        if len(cand) > 6 and not any(x in cand.upper() for x in skip):
            return cand
    m = re.search(
        r"((?:CÔNG TY|CTY|DOANH NGHIỆP)[^.]+)",
        filename or "",
        re.IGNORECASE,
    )
    if m:
        cand = clean_company_name(m.group(1))
        if len(cand) > 6:
            return cand
    return ""


def extract_dia_chi(body: str) -> str:
    m = re.search(
        r"(?:địa chỉ|trụ sở)[:\s]*([^\n\r]{10,160})",
        body or "",
        re.IGNORECASE,
    )
    if not m:
        return ""
    dia = m.group(1).strip()
    dia = re.sub(
        r"^(?:đăng ký kinh doanh|trụ sở chính|trụ sở|đã đăng ký)[:\s]*",
        "",
        dia,
        flags=re.IGNORECASE,
    )
    return dia.strip(" ;,.")


def extract_noi_dung(body: str, trich_yeu: str) -> str:
    blob = f"{body}\n{trich_yeu}"
    reasons = []
    if re.search(r"bãi bỏ", blob, re.IGNORECASE):
        reasons.append("Bãi bỏ văn bản cảnh báo trước")
    if re.search(r"không hoạt động tại địa chỉ", blob, re.IGNORECASE):
        reasons.append("Không hoạt động tại địa chỉ đăng ký")
    if re.search(r"chưa nộp tờ khai|không nộp tờ khai", blob, re.IGNORECASE):
        reasons.append("Chưa nộp tờ khai thuế")
    if re.search(r"chênh lệch", blob, re.IGNORECASE):
        reasons.append("Chênh lệch số liệu tờ khai và HĐĐT")
    if re.search(r"cơ quan điều tra|chuyển tin báo", blob, re.IGNORECASE):
        reasons.append("Chuyển tin báo CQĐT")
    if re.search(r"tạm ngừng kinh doanh", blob, re.IGNORECASE):
        reasons.append("Tạm ngừng kinh doanh")
    if re.search(r"gian lận|mua bán hóa đơn", blob, re.IGNORECASE):
        reasons.append("Nghi vấn gian lận hóa đơn")
    if reasons:
        return "; ".join(reasons)
    return "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn"


def load_excel_keys(excel_path: str) -> tuple[set[str], set[str]]:
    log(f"[EXCEL] Doc khoa hien co: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb["THEO_DOI"]
    socv, mst = set(), set()
    for row in ws.iter_rows(min_row=8, max_col=6, values_only=True):
        if row[1]:
            socv.add(norm_socv(str(row[1])))
        if row[5]:
            mst.add(re.sub(r"\s+", "", str(row[5])))
    wb.close()
    log(f"[EXCEL] SoCV={len(socv)} MST={len(mst)}")
    return socv, mst


def pick_unique_items(raw: dict) -> tuple[list[dict], list[dict]]:
    groups: dict[str, list[dict]] = defaultdict(list)
    for did, item in raw.items():
        path = item["files"][0]
        pages, chars = pdf_stats(path) if os.path.exists(path) else (0, -1)
        rec = {
            **item,
            "doc_id": did,
            "path": path,
            "exists": os.path.exists(path),
            "sig": file_sig(path) if os.path.exists(path) else "MISSING",
            "pages": pages,
            "chars": chars,
            "is_xlsx": path.lower().endswith((".xls", ".xlsx")),
            "is_bangke": bool(
                re.search(r"bang\s*ke|danh sach hd|hđ bán ra|hdbr", path, re.I)
            ),
            "is_scan": (not path.lower().endswith((".xls", ".xlsx"))) and chars < 400,
        }
        groups[norm_socv(item["so_ky_hieu"])].append(rec)

    keep, skipped = [], []
    seen_hash: dict[str, str] = {}
    for kh, items in groups.items():
        items_ok = [x for x in items if x["exists"]]
        if not items_ok:
            skipped.extend({**x, "skip_reason": "file_thieu"} for x in items)
            continue

        # Uu tien: PDF cong van (khong bang ke) > digital > nhieu trang > file lon
        def score(x):
            return (
                0 if x["is_xlsx"] else 1,
                0 if x["is_bangke"] else 1,
                0 if x["is_scan"] else 1,
                x["pages"],
                os.path.getsize(x["path"]),
            )

        items_ok.sort(key=score, reverse=True)
        best = items_ok[0]
        if best["sig"] in seen_hash:
            skipped.append(
                {
                    **best,
                    "skip_reason": f"trung_hash_voi_{seen_hash[best['sig']]}",
                }
            )
            for extra in items_ok[1:]:
                skipped.append({**extra, "skip_reason": f"trung_so_ky_hieu_{kh}"})
            continue
        seen_hash[best["sig"]] = kh
        keep.append(best)
        for extra in items_ok[1:]:
            reason = (
                "trung_hash"
                if extra["sig"] == best["sig"]
                else f"trung_so_ky_hieu_{kh}"
            )
            skipped.append({**extra, "skip_reason": reason})
    return keep, skipped


def parse_item(item: dict, body: str) -> dict:
    trich = item.get("trich_yeu") or ""
    fname = os.path.basename(item["path"])
    ten = extract_ten_ban(trich, body, fname)
    mst = extract_mst(body, trich, fname)
    if not ten:
        ten = f"Doanh nghiệp theo CV {item['so_ky_hieu']}"
    return {
        "so_cv": item["so_ky_hieu"],
        "ngay_cv": item.get("ngay_vb") or "",
        "ngay_nhan": item.get("ngay_den") or "",
        "don_vi_bh": normalize_don_vi(item.get("co_quan") or ""),
        "ten_ban": ten,
        "mst_ban": mst,
        "dia_chi": extract_dia_chi(body),
        "noi_dung": extract_noi_dung(body, trich),
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "file": item["path"],
        "nam_cv": 2026,
        "ghi_chu": (
            f"Nguồn: CBRR-ben ban. Số đến {item.get('so_den')}. "
            f"Trích yếu: {trich}."
        ),
        "cv_dau": 1,
        "mst_dau": None,
    }


def xlsx_preview(path: str) -> str:
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = []
        for i, row in enumerate(ws.iter_rows(max_row=30, max_col=8, values_only=True), 1):
            vals = [str(c) for c in row if c is not None]
            if vals:
                rows.append(" | ".join(vals))
        wb.close()
        return "\n".join(rows)
    except Exception as e:
        return f"(khong doc duoc xlsx: {e})"


def main() -> int:
    raw = json.loads(PROGRESS_PATH.read_text(encoding="utf-8"))
    keep, skipped = pick_unique_items(raw)
    existing_socv, existing_mst = load_excel_keys(DEFAULT_EXCEL_PATH)

    to_process, skipped_excel = [], []
    for item in keep:
        kh = norm_socv(item["so_ky_hieu"])
        if kh in ALREADY_IN_BOOK or kh in existing_socv:
            skipped_excel.append({**item, "skip_reason": "da_co_tren_so"})
        else:
            to_process.append(item)

    plan = {
        "tong_tai": len(raw),
        "giu_doc_nhat": [
            {
                "so_ky_hieu": x["so_ky_hieu"],
                "so_den": x["so_den"],
                "file": os.path.basename(x["path"]),
                "scan": x["is_scan"],
            }
            for x in to_process
        ],
        "bo_qua": [
            {
                "so_ky_hieu": x["so_ky_hieu"],
                "so_den": x.get("so_den"),
                "file": os.path.basename(x["path"]),
                "ly_do": x.get("skip_reason"),
            }
            for x in skipped + skipped_excel
        ],
    }
    PLAN_PATH.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    log(
        f"[PLAN] goc={len(raw)} doc_nhat_xu_ly={len(to_process)} "
        f"bo_trung={len(skipped)} da_co_so={len(skipped_excel)}"
    )

    records = []
    seen_mst_batch: set[str] = set()
    for i, item in enumerate(to_process, 1):
        log(
            f"[{i}/{len(to_process)}] {item['so_ky_hieu']} so_den={item['so_den']} "
            f"{'SCAN' if item['is_scan'] else 'DIGITAL'}"
        )
        body = ""
        if item["is_xlsx"]:
            body = xlsx_preview(item["path"])
            md_path = Path(item["path"]).with_suffix(".md")
            write_md(md_path, item, body)
        else:
            md_path = convert_pdf_to_md(item["path"], item["is_scan"])
            if md_path.exists():
                body = md_path.read_text(encoding="utf-8", errors="ignore")
            if len(body.strip()) < 80:
                extra = pdf_full_text(item["path"])
                body = (body + "\n" + extra).strip()
        rec = parse_item(item, body)
        mst = rec["mst_ban"]
        if mst and mst not in existing_mst and mst not in seen_mst_batch:
            rec["mst_dau"] = 1
            seen_mst_batch.add(mst)
        else:
            rec["mst_dau"] = None
        records.append(rec)
        log(f"    -> {rec['ten_ban'][:70]} | MST={rec['mst_ban'] or '-'} | {rec['noi_dung'][:80]}")

    RECORDS_PATH.write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log(f"[JSON] {len(records)} ban ghi -> {RECORDS_PATH}")

    if not records:
        log("[STOP] Khong co ban ghi moi de nap Excel")
        return 0

    new_max = update_workbook(records, DEFAULT_EXCEL_PATH, backup=True)
    log(f"[DONE] nap {len(records)} dong, max_row={new_max}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
