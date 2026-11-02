# -*- coding: utf-8 -*-
"""Lam sach F0 tu sheet CBRR va F1 tu sheet TB DO KT1 MOI.

Nguon: INPUT/Tong hop danh sach moi rui ro (1).xlsb
- CBRR -> F0 (ben ban): A=MST, B=Ten, E=So CV, F=Ngay CV
  (tieu de that: C=Ngay CV den, D=Noi gui; khong dung C/D lam SoCV/NgayCV)
- TB DO KT1 MOI -> F1 (ben mua): B=MST, C=Ten, M=So CV, N=Ngay CV
Giu nguyen cac dong nguon khac (HDRR, 524, cong an 2026, F1 2026.09.16...).
"""
from __future__ import annotations

import re
import shutil
import sys
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, Protection
from pyxlsb import open_workbook

SRC_XLSB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\INPUT\Tong hop danh sach moi rui ro (1).xlsb"
)
EXCEL_PATH = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
SYNC_COPIES = [
    Path(r"D:\SECOND BRAIN\KIEM SOAT RUI RO\QUAN LY HD RUI RO\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"),
    Path(r"D:\Theo doi CVCBRR-ok\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"),
]

HEADER_ROWS = 7
NCOLS = 25
FONT = Font(name="Times New Roman", size=10)
MST_RE = re.compile(r"^(\d{10})(-\d{3})?$")
DATE_IN_TEXT_RE = re.compile(
    r"ng[aà]y\s+(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", re.IGNORECASE
)
SLASH_DATE_RE = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$")
XL_EPOCH = datetime(1899, 12, 30)


def log(msg: str) -> None:
    print(msg, flush=True)


def row_vals(row, n=20):
    vals = [None] * n
    for cell in row:
        idx = getattr(cell, "c", None)
        if idx is None:
            continue
        if 0 <= idx < n:
            vals[idx] = cell.v
    return vals


def as_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def norm_mst(v) -> str:
    s = as_text(v).replace(" ", "").replace(".", "")
    s = s.replace(";", "").replace(",", "")
    if s.endswith("-"):
        s = s[:-1]
    if MST_RE.fullmatch(s):
        return s
    return ""


def parse_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)):
        try:
            n = float(v)
            if 20000 < n < 80000:
                return XL_EPOCH + timedelta(days=n)
        except Exception:
            return None
        return None
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    m = SLASH_DATE_RE.match(s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime(y, mo, d)
        except ValueError:
            return None
    return None


def date_from_socv(socv: str):
    m = DATE_IN_TEXT_RE.search(socv or "")
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    try:
        return datetime(y, mo, d)
    except ValueError:
        return None


def looks_agency(s: str) -> bool:
    if not s:
        return False
    u = s.upper()
    return (
        u.startswith("CA ")
        or u.startswith("CA.")
        or u.startswith("CÔNG AN")
        or u.startswith("CONG AN")
        or "CSKT" in u
        or u.startswith("TCT")
        or "CỤC THUẾ" in u
        or "CUC THUE" in u
    )


def split_ngay_or_donvi(v):
    d = parse_date(v)
    if d:
        return d, ""
    s = as_text(v)
    if not s:
        return None, ""
    d2 = parse_date(s)
    if d2:
        return d2, ""
    return None, s


def is_cbrr_source(ghi: str) -> bool:
    return "CBRR" in (ghi or "").upper()


def is_tbkt1_source(ghi: str) -> bool:
    u = (ghi or "").upper()
    return "TB DO KT1" in u or "TB ĐO KT1" in u


def extract_cbrr(src: Path):
    recs = []
    skipped = Counter()
    seen = set()
    with open_workbook(str(src)) as wb:
        with wb.get_sheet("CBRR") as sh:
            for i, row in enumerate(sh.rows(), 1):
                vals = row_vals(row, 12)
                if i <= 2:
                    continue
                mst = norm_mst(vals[0])
                ten = as_text(vals[1])
                if not mst:
                    if as_text(vals[0]) or ten:
                        skipped["khong_mst"] += 1
                    continue
                if ten.upper() in {"TÊN ĐƠN VỊ", "TEN DON VI", "TÊN CÔNG TY"}:
                    skipped["header"] += 1
                    continue
                socv = as_text(vals[4])  # E So CV
                ngay = parse_date(vals[5])  # F Ngay CV
                ngay_nhan = parse_date(vals[2])  # C Ngay CV den
                if ngay_nhan is None:
                    ngay_nhan = as_text(vals[2]) or None
                donvi = as_text(vals[3])  # D Noi gui
                noidung = as_text(vals[6])
                dau_hieu = as_text(vals[7])
                key = (socv.upper(), mst)
                if key in seen:
                    skipped["trung"] += 1
                    continue
                seen.add(key)
                ghi = "Nguồn: CBRR."
                if dau_hieu:
                    ghi += " Dấu hiệu: " + dau_hieu[:180]
                recs.append(
                    {
                        "SoCV": socv or None,
                        "NgayCV": ngay,
                        "DonViBH": donvi or None,
                        "TenBan": ten or None,
                        "MSTBan": mst,
                        "NoiDung": noidung or "Cảnh báo rủi ro về thuế và hóa đơn",
                        "TenMua": None,
                        "MSTMua": None,
                        "DiaChi": None,
                        "NgayNhan": ngay_nhan if isinstance(ngay_nhan, datetime) else None,
                        "GhiChu": ghi,
                        "NhomND": dau_hieu or "Dấu hiệu rủi ro về thuế",
                        "File": "CBRR|" + SRC_XLSB.name,
                    }
                )
    return recs, skipped


def extract_tbkt1(src: Path):
    recs = []
    skipped = Counter()
    seen = set()
    current_section = ""
    with open_workbook(str(src)) as wb:
        with wb.get_sheet("TB DO KT1 MOI") as sh:
            for i, row in enumerate(sh.rows(), 1):
                vals = row_vals(row, 16)
                a = as_text(vals[0])
                mst = norm_mst(vals[1])
                if a and not mst and not isinstance(vals[0], (int, float)):
                    current_section = a
                    continue
                if not mst:
                    if as_text(vals[1]) or as_text(vals[2]):
                        skipped["khong_mst"] += 1
                    continue
                ten = as_text(vals[2])
                diachi = as_text(vals[3])
                phuong = as_text(vals[4])
                cbql = as_text(vals[5])
                mail = as_text(vals[10])
                sdt = as_text(vals[11])
                socv = as_text(vals[12])  # M
                ngay, donvi_n = split_ngay_or_donvi(vals[13])  # N
                if ngay is None:
                    ngay = date_from_socv(socv)
                donvi = donvi_n if looks_agency(donvi_n) else (donvi_n or None)
                key = (socv.upper(), mst, current_section)
                if key in seen:
                    skipped["trung"] += 1
                    continue
                seen.add(key)
                ghi_parts = ["Nguồn: TB DO KT1 MOI."]
                if current_section:
                    ghi_parts.append("Nhóm: " + current_section)
                if diachi:
                    ghi_parts.append("Địa chỉ: " + diachi)
                if phuong:
                    ghi_parts.append(phuong)
                if cbql:
                    ghi_parts.append("CBQL: " + cbql)
                if mail:
                    ghi_parts.append("Email: " + mail)
                if sdt:
                    ghi_parts.append("SĐT: " + sdt)
                recs.append(
                    {
                        "SoCV": socv or None,
                        "NgayCV": ngay,
                        "DonViBH": donvi,
                        "TenBan": None,
                        "MSTBan": None,
                        "NoiDung": current_section or "Bên mua rà soát",
                        "TenMua": ten or None,
                        "MSTMua": mst,
                        "DiaChi": " ".join(x for x in (diachi, phuong) if x) or None,
                        "NgayNhan": None,
                        "GhiChu": " ".join(ghi_parts),
                        "NhomND": current_section or "Bên mua rà soát",
                        "File": "TB DO KT1 MOI|" + SRC_XLSB.name,
                    }
                )
    return recs, skipped


def copy_style(src_cell, dst_cell):
    dst_cell.font = Font(
        name=src_cell.font.name or "Times New Roman",
        size=src_cell.font.size or 10,
        bold=src_cell.font.bold,
        italic=src_cell.font.italic,
        color=src_cell.font.color,
    )
    if src_cell.alignment:
        dst_cell.alignment = Alignment(
            horizontal=src_cell.alignment.horizontal,
            vertical=src_cell.alignment.vertical,
            wrap_text=src_cell.alignment.wrap_text,
        )
    dst_cell.number_format = src_cell.number_format
    if src_cell.fill and src_cell.fill.fill_type:
        dst_cell.fill = PatternFill(
            fill_type=src_cell.fill.fill_type,
            fgColor=src_cell.fill.fgColor,
            bgColor=src_cell.fill.bgColor,
        )
    if src_cell.border:
        dst_cell.border = Border(
            left=src_cell.border.left,
            right=src_cell.border.right,
            top=src_cell.border.top,
            bottom=src_cell.border.bottom,
        )
    if src_cell.protection:
        dst_cell.protection = Protection(
            locked=src_cell.protection.locked,
            hidden=src_cell.protection.hidden,
        )


def read_data_rows(ws):
    rows = []
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, NCOLS + 1)]
        if all(v is None or str(v).strip() == "" for v in vals):
            continue
        rows.append(vals)
    return rows


def rec_to_row(rec: dict) -> list:
    # Layout moi: STT MSTBan TenBan SoCV NgayCV DonViBH NoiDung TenMua MSTMua ...
    ngay = rec.get("NgayCV")
    nam = ngay.year if isinstance(ngay, datetime) else None
    socv = rec.get("SoCV") or ""
    khoa = f"{socv}_{nam}" if socv and nam else (socv or None)
    row = [None] * NCOLS
    row[1] = rec.get("MSTBan")
    row[2] = rec.get("TenBan")
    row[3] = rec.get("SoCV")
    row[4] = rec.get("NgayCV")
    row[5] = rec.get("DonViBH")
    row[6] = rec.get("NoiDung")
    row[7] = rec.get("TenMua")
    row[8] = rec.get("MSTMua")
    row[14] = rec.get("NhomND")
    row[15] = rec.get("DiaChi")
    row[16] = rec.get("NgayNhan")
    row[18] = rec.get("GhiChu")
    row[19] = rec.get("File")
    row[20] = nam
    row[22] = khoa
    return row


def recompute_flags(rows, mst_col_idx: int):
    seen_cv = set()
    seen_mst = set()
    out = []
    for i, row in enumerate(rows, 1):
        r = list(row)
        if len(r) < NCOLS:
            r.extend([None] * (NCOLS - len(r)))
        r[0] = i
        socv = as_text(r[3]).upper()
        ngay = r[4]
        nam = r[20]
        if not nam and isinstance(ngay, datetime):
            nam = ngay.year
            r[20] = nam
        if socv and nam and not r[22]:
            r[22] = f"{as_text(r[3])}_{nam}"
        if socv and socv not in seen_cv:
            seen_cv.add(socv)
            r[23] = 1
        else:
            r[23] = None
        mst = as_text(r[mst_col_idx])
        if mst and mst not in seen_mst:
            seen_mst.add(mst)
            r[24] = 1
        else:
            r[24] = None
        out.append(r)
    return out, len(seen_cv), len(seen_mst)


def write_rows(ws, rows):
    tmpl = [ws.cell(HEADER_ROWS + 1, c) for c in range(1, NCOLS + 1)]
    old_n = ws.max_row - HEADER_ROWS
    new_n = len(rows)
    date_cols = {5, 17}
    mst_cols = {2, 9}
    for i, row in enumerate(rows):
        r = HEADER_ROWS + 1 + i
        for c in range(1, NCOLS + 1):
            cell = ws.cell(r, c)
            val = row[c - 1] if c - 1 < len(row) else None
            cell.value = val
            cell.font = FONT
            if c in date_cols and isinstance(val, datetime):
                cell.number_format = "yyyy-mm-dd"
            elif c in mst_cols:
                cell.number_format = "@"
            elif tmpl[c - 1].number_format and tmpl[c - 1].number_format != "General":
                if c not in date_cols:
                    cell.number_format = tmpl[c - 1].number_format
    leftover = old_n - new_n
    if leftover > 0:
        ws.delete_rows(HEADER_ROWS + 1 + new_n, leftover)
    ws.freeze_panes = "A8"
    return old_n, new_n


def main():
    if not SRC_XLSB.is_file():
        log(f"[ERROR] Khong thay file nguon: {SRC_XLSB}")
        return 1
    if not EXCEL_PATH.is_file():
        log(f"[ERROR] Khong thay so theo doi: {EXCEL_PATH}")
        return 1

    log("[1] Trich xuat CBRR va TB DO KT1 MOI...")
    cbrr, skip_c = extract_cbrr(SRC_XLSB)
    tb, skip_t = extract_tbkt1(SRC_XLSB)
    log(f"    CBRR hop le: {len(cbrr)} | bo qua: {dict(skip_c)}")
    log(f"    TB DO KT1 MOI hop le: {len(tb)} | bo qua: {dict(skip_t)}")
    if len(cbrr) < 1000 or len(tb) < 1000:
        log("[ERROR] So dong trich xuat qua thap, dung lai.")
        return 1

    bak_dir = EXCEL_PATH.parent / "Sao lưu"
    bak_dir.mkdir(exist_ok=True)
    bak = bak_dir / f"{EXCEL_PATH.stem}.bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}{EXCEL_PATH.suffix}"
    log(f"[2] Sao luu -> {bak}")
    shutil.copy2(EXCEL_PATH, bak)

    log("[3] Mo so theo doi...")
    wb = load_workbook(EXCEL_PATH)
    ws0 = wb["F0"]
    ws1 = wb["F1"]

    old_f0 = read_data_rows(ws0)
    old_f1 = read_data_rows(ws1)
    keep_f0 = [r for r in old_f0 if not is_cbrr_source(as_text(r[18] if len(r) > 18 else ""))]
    drop_f0 = len(old_f0) - len(keep_f0)
    keep_f1 = [r for r in old_f1 if not is_tbkt1_source(as_text(r[18] if len(r) > 18 else ""))]
    drop_f1 = len(old_f1) - len(keep_f1)
    log(f"    F0 cu {len(old_f0)} | giu {len(keep_f0)} | go CBRR cu {drop_f0}")
    log(f"    F1 cu {len(old_f1)} | giu {len(keep_f1)} | go TB DO KT1 MOI cu {drop_f1}")

    new_f0 = keep_f0 + [rec_to_row(x) for x in cbrr]
    new_f1 = keep_f1 + [rec_to_row(x) for x in tb]
    new_f0, cv0, mst0 = recompute_flags(new_f0, 1)  # MSTBan
    new_f1, cv1, mst1 = recompute_flags(new_f1, 8)  # MSTMua

    log("[4] Ghi lai sheet F0 / F1...")
    o0, n0 = write_rows(ws0, new_f0)
    o1, n1 = write_rows(ws1, new_f1)
    log(f"    F0 {o0} -> {n0} | CVDau {cv0} | MSTDau {mst0}")
    log(f"    F1 {o1} -> {n1} | CVDau {cv1} | MSTDau {mst1}")

    log("[5] Luu file...")
    wb.save(EXCEL_PATH)
    wb.close()
    size = EXCEL_PATH.stat().st_size
    log(f"    Da luu {EXCEL_PATH} ({size} bytes)")

    for dest in SYNC_COPIES:
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(EXCEL_PATH, dest)
            log(f"[OK] Dong bo -> {dest}")
        except Exception as e:
            log(f"[WARN] Khong dong bo {dest}: {e}")

    log("[DONE]")
    log(
        f"SUMMARY|cbrr={len(cbrr)}|tb={len(tb)}|drop_f0={drop_f0}|drop_f1={drop_f1}"
        f"|f0={n0}|f1={n1}|cv0={cv0}|mst0={mst0}|cv1={cv1}|mst1={mst1}|bak={bak}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
