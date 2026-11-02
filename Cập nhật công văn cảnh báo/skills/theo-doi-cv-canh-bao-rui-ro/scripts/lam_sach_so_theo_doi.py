# -*- coding: utf-8 -*-
"""Lam sach / chuan hoa SoTheoDoi F0 + F1 tu file nguon .xlsb."""
from __future__ import annotations

import re
import shutil
from collections import Counter, OrderedDict
from copy import copy
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pyxlsb import open_workbook

XLSB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\INPUT\Tong hop danh sach moi rui ro (1).xlsb"
)
F524 = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\INPUT\524.xlsx"
)
SO_THEO_DOI = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
BACKUP_DIR = SO_THEO_DOI.parent / "Sao lưu"
SYNC_COPIES = [
    Path(r"D:\SECOND BRAIN\KIEM SOAT RUI RO\QUAN LY HD RUI RO\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"),
    Path(r"D:\Theo doi CVCBRR-ok\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"),
]

FONT_TNR = Font(name="Times New Roman", size=10)
FONT_TITLE = Font(name="Times New Roman", size=15, bold=True, color="FFFFFF")
FONT_SUB = Font(name="Times New Roman", size=10, color="FFFFFF")
FONT_HINT = Font(name="Times New Roman", size=9, color="595959")
FONT_GRP = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_H1 = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_H2 = Font(name="Times New Roman", size=8, bold=True, color="E0E0E0")
ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)

FILL_TITLE = PatternFill("solid", fgColor="1B3A4B")
FILL_SUB = PatternFill("solid", fgColor="24556E")
FILL_CV = PatternFill("solid", fgColor="1F4E79")
FILL_BAN = PatternFill("solid", fgColor="C65911")
FILL_ND = PatternFill("solid", fgColor="375623")
FILL_MUA = PatternFill("solid", fgColor="0F6E56")
FILL_TN = PatternFill("solid", fgColor="8A6D1B")
FILL_HD = PatternFill("solid", fgColor="2E75B6")
FILL_XL = PatternFill("solid", fgColor="41719C")
FILL_DS = PatternFill("solid", fgColor="5B9BD5")
FILL_CODE = PatternFill("solid", fgColor="333F48")
FILL_STT = PatternFill("solid", fgColor="1B3A4B")

XL_EPOCH = datetime(1899, 12, 30)
SOCV_RE = re.compile(r"(\d+[A-Z]?/[A-Za-z0-9Đđ.\-()]+)")
DATE_IN_TEXT_RE = re.compile(r"ngày\s+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", re.I)


def log(msg: str) -> None:
    print(msg, flush=True)


def rowdict(row) -> dict:
    return {c.c: c.v for c in row}


def xldate(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.replace(tzinfo=None)
    if isinstance(v, (int, float)) and 20000 < float(v) < 80000:
        try:
            return XL_EPOCH + timedelta(days=float(v))
        except Exception:
            return None
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%Y/%m/%d"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", s)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if y < 100:
                y += 2000
            try:
                return datetime(y, mo, d)
            except ValueError:
                return None
    return None


def norm_mst(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if abs(v - int(v)) < 1e-9:
            v = str(int(v))
        else:
            v = str(v)
    s = str(v).strip().replace(" ", "").replace("\u00a0", "")
    return s


def is_valid_mst(s: str) -> bool:
    if not s:
        return False
    core = s.replace("-", "")
    if not core.isdigit():
        return False
    return len(core) in (9, 10, 12, 13, 14)


def norm_socv(v) -> str:
    if v is None or v == "":
        return ""
    s = str(v).strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace("\\", "/")
    return s


def clean_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def parse_cv_blob(text: str):
    """Tach SoCV, NgayCV, DonViBH tu cot E sheet CV cong an."""
    s = clean_text(text)
    if not s:
        return "", None, ""
    su = s.upper()
    if "1798" in su and "TCT" in su:
        return "1798/TCT-TTKT", datetime(2023, 5, 16), "Tổng cục Thuế"

    ngay = None
    m = DATE_IN_TEXT_RE.search(s)
    if m:
        ngay = xldate(m.group(1))

    socv = ""
    m = SOCV_RE.search(s)
    if m:
        socv = m.group(1).strip()
        socv = socv.rstrip(".,;")

    donvi = ""
    m = re.search(r"của\s+(.+)$", s, re.I)
    if m:
        donvi = clean_text(m.group(1))
    elif re.search(r"Bộ Công an", s, re.I):
        donvi = "Bộ Công an"
    else:
        m = re.search(r"(Công an[^,]*)", s, re.I)
        if m:
            donvi = clean_text(m.group(1))
        elif re.search(r"Cơ quan ANĐT", s, re.I):
            donvi = "Cơ quan ANĐT - Bộ Công an"
    return socv, ngay, donvi


def rec_score(rec: dict) -> int:
    keys = ("so_cv", "ngay_cv", "don_vi_bh", "ten", "mst", "noi_dung", "dia_chi", "ghi_chu")
    return sum(1 for k in keys if rec.get(k) not in (None, "", 0))


def merge_rec(old: dict, new: dict) -> dict:
    out = dict(old)
    for k, v in new.items():
        if k in ("nguon",):
            continue
        if (out.get(k) in (None, "", 0)) and v not in (None, "", 0):
            out[k] = v
    srcs = []
    for r in (old, new):
        s = r.get("nguon") or ""
        if s and s not in srcs:
            srcs.append(s)
    out["nguon"] = " | ".join(srcs)
    if rec_score(new) > rec_score(old):
        for k in ("ten", "so_cv", "ngay_cv", "don_vi_bh", "noi_dung", "dia_chi"):
            if new.get(k) not in (None, "", 0):
                out[k] = new[k]
    return out


def add_unique(store: OrderedDict, rec: dict, stats: Counter, sheet_label: str):
    mst = rec.get("mst") or ""
    if not is_valid_mst(mst):
        stats["skip_mst"] += 1
        return
    socv = rec.get("so_cv") or ""
    key = (mst, socv.upper())
    rec["nguon"] = sheet_label
    if key in store:
        store[key] = merge_rec(store[key], rec)
        stats["dup_merged"] += 1
    else:
        store[key] = rec
        stats["kept"] += 1


def extract_cbrr(wb, store, stats):
    """CBRR: A MST, B ten, C ngay nhan, D noi gui (DonViBH), E SoCV, F NgayCV, G ND, H dau hieu.
    Nguoi dung ghi C=SoCV D=NgayCV nhung header that cua file khac.
    """
    with wb.get_sheet("CBRR") as sh:
        for i, row in enumerate(sh.rows(), 1):
            if i <= 2:
                continue
            d = rowdict(row)
            rec = {
                "mst": norm_mst(d.get(0)),
                "ten": clean_text(d.get(1)),
                "ngay_nhan": xldate(d.get(2)),
                "don_vi_bh": clean_text(d.get(3)),
                "so_cv": norm_socv(d.get(4)),
                "ngay_cv": xldate(d.get(5)),
                "noi_dung": clean_text(d.get(6)),
                "nhom_nd": clean_text(d.get(7)) or "Dấu hiệu rủi ro về thuế",
                "ghi_chu": "Nguồn: CBRR.",
                "sheet": "F0",
            }
            if rec["noi_dung"] and rec.get("nhom_nd") and rec["nhom_nd"] not in rec["ghi_chu"]:
                rec["ghi_chu"] = f"Nguồn: CBRR. Dấu hiệu: {rec['nhom_nd']}"
            if not rec["ten"] and not rec["mst"]:
                continue
            add_unique(store, rec, stats, "CBRR")


def fill_cv_from_note(rec: dict) -> None:
    """Khi cot SoCV/Ngay trong bang ke trong, boc tu Ghi chu."""
    note = rec.get("noi_dung") or rec.get("ghi_chu") or ""
    nu = note.upper()
    if not rec.get("so_cv"):
        if "524 DN" in nu or "CV 1798" in nu:
            rec["so_cv"] = "1798/TCT-TTKT"
            rec["ngay_cv"] = rec.get("ngay_cv") or datetime(2023, 5, 16)
            rec["don_vi_bh"] = rec.get("don_vi_bh") or "Tổng cục Thuế"
        else:
            m = SOCV_RE.search(note.replace(" ", ""))
            if not m:
                m = SOCV_RE.search(note)
            if m:
                rec["so_cv"] = norm_socv(m.group(1))
    if not rec.get("ngay_cv"):
        m = DATE_IN_TEXT_RE.search(note)
        if m:
            rec["ngay_cv"] = xldate(m.group(1))
        elif "16-05-23" in note or "16/05/2023" in note:
            rec["ngay_cv"] = datetime(2023, 5, 16)


def extract_hdrr_cuc(wb, store, stats):
    """HDRR -cuc (user goi HDTT-CUc): B ten, C MST, E DonViBH, F SoCV, G Ngay, H GhiChu."""
    with wb.get_sheet("HDRR -cuc") as sh:
        for i, row in enumerate(sh.rows(), 1):
            d = rowdict(row)
            ten = clean_text(d.get(1))
            dia = clean_text(d.get(3))
            if dia and ten and dia.upper() == ten.upper():
                dia = ""
            rec = {
                "mst": norm_mst(d.get(2)),
                "ten": ten,
                "dia_chi": dia,
                "don_vi_bh": clean_text(d.get(4)),
                "so_cv": norm_socv(d.get(5)),
                "ngay_cv": xldate(d.get(6)),
                "ghi_chu": clean_text(d.get(7)),
                "noi_dung": clean_text(d.get(7)),
                "nhom_nd": "Dấu hiệu rủi ro về thuế",
                "sheet": "F0",
            }
            fill_cv_from_note(rec)
            note = clean_text(d.get(7))
            if note:
                rec["ghi_chu"] = f"Nguồn: HDRR -cuc. {note}"
            else:
                rec["ghi_chu"] = "Nguồn: HDRR -cuc."
            if not rec["ten"] and not rec["mst"]:
                continue
            nu = (note + " " + (rec.get("so_cv") or "")).upper()
            if "DS 524 DN" in nu or "CV 1798" in nu or (rec.get("so_cv") or "").upper() in {
                "1798",
                "1798.0",
                "1798/TCT-TTKT",
                "1770/CAT-CSKT",
            }:
                stats["skip_524_xlsx"] += 1
                continue
            add_unique(store, rec, stats, "HDRR -cuc")


def extract_cv_cong_an(wb, store, stats):
    """CV cong an: B ten, C MST, D ghi chu/quy mo, E blob SoCV+Ngay+DonViBH."""
    with wb.get_sheet("CV cong an") as sh:
        for i, row in enumerate(sh.rows(), 1):
            d = rowdict(row)
            blob = clean_text(d.get(4))
            socv, ngay, donvi = parse_cv_blob(blob)
            rec = {
                "mst": norm_mst(d.get(2)),
                "ten": clean_text(d.get(1)),
                "so_cv": norm_socv(socv),
                "ngay_cv": ngay,
                "don_vi_bh": donvi,
                "ghi_chu": "",
                "noi_dung": blob,
                "nhom_nd": "Rủi ro do Công an chuyển",
                "sheet": "F0",
            }
            quy_mo = clean_text(d.get(3))
            extra = clean_text(d.get(6))
            parts = ["Nguồn: CV cong an."]
            if quy_mo:
                parts.append(f"Quy mô đợt: {quy_mo}")
            if blob and not socv:
                parts.append(f"Số ghi trong nguồn: {blob}")
            if extra:
                parts.append(extra)
            rec["ghi_chu"] = " ".join(parts)
            if not rec["ten"] and not rec["mst"]:
                continue
            qm = quy_mo.upper()
            su = (socv or "").upper()
            if (
                "524 CÔNG TY" in qm
                or "524 CONG TY" in qm
                or "637 CÔNG TY" in qm
                or "637 CONG TY" in qm
                or "1798/TCT" in su
                or "1770/CAT" in su
                or "05/BKL" in su
            ):
                stats["skip_524_xlsx"] += 1
                continue
            add_unique(store, rec, stats, "CV cong an")


def extract_524_xlsx(store, stats):
    """Danh sach 524/637 DN chuan. Khong lay lai tu HDRR hay CV cong an."""
    if not F524.is_file():
        stats["missing_file"] += 1
        return
    from openpyxl import load_workbook as _lw

    wb = _lw(F524, read_only=True, data_only=True)
    for row in wb.active.iter_rows(min_row=2, values_only=True):
        mst = norm_mst(row[1])
        ten = clean_text(row[2])
        ghi = clean_text(row[3])
        blob = clean_text(row[4])
        socv, ngay, donvi = parse_cv_blob(blob)
        if "ngày" in blob.lower() and not ngay:
            m = DATE_IN_TEXT_RE.search(blob)
            if m:
                ngay = xldate(m.group(1))
        if not socv:
            m = SOCV_RE.search(blob.replace(" ", ""))
            if m:
                socv = m.group(1)
        if "phú thọ" in (donvi or blob).lower() and "tỉnh" not in (donvi or "").lower():
            donvi = "Công an tỉnh Phú Thọ"
        if "637" in ghi:
            nd = "Danh sách 637 DN (524 mở rộng) - BKL 05/BKL-CSKT(Đ4)"
            note = (
                "Nguồn: 524.xlsx. Quy mô đợt: 637 Công ty (524 mở rộng). "
                "BKL 05/BKL-CSKT(Đ4) ngày 16/08/2023 của Công an tỉnh Phú Thọ."
            )
        else:
            nd = "Danh sách 524 DN bán rủi ro - CV 1770/CAT-CSKT gắn CV 1798/TCT-TTKT"
            note = (
                "Nguồn: 524.xlsx. Quy mô đợt: 524 Công ty. "
                "CV 1770/CAT-CSKT ngày 29/03/2023 của Công an tỉnh Phú Thọ; "
                "liên kết CV 1798/TCT-TTKT ngày 16/05/2023 của Tổng cục Thuế."
            )
        rec = {
            "mst": mst,
            "ten": ten,
            "so_cv": norm_socv(socv),
            "ngay_cv": ngay,
            "don_vi_bh": donvi or "Công an tỉnh Phú Thọ",
            "noi_dung": nd,
            "nhom_nd": "Rủi ro do Công an chuyển",
            "ghi_chu": note,
            "sheet": "F0",
        }
        add_unique(store, rec, stats, "524.xlsx")
    wb.close()


def extract_tb_do_kt1(wb, store, stats, review: set):
    """TB DO KT1 MOI -> F1: B MST, C ten, M SoCV, N NgayCV."""
    current_group = ""
    with wb.get_sheet("TB DO KT1 MOI") as sh:
        for i, row in enumerate(sh.rows(), 1):
            d = rowdict(row)
            if i == 1:
                continue
            mst = norm_mst(d.get(1))
            ten = clean_text(d.get(2))
            if not is_valid_mst(mst):
                title = clean_text(d.get(0))
                if title and not str(title).replace(".", "", 1).isdigit():
                    current_group = title
                continue
            review.add(mst)
            rec = {
                "mst": mst,
                "ten": ten,
                "dia_chi": clean_text(d.get(3)),
                "so_cv": norm_socv(d.get(12)),
                "ngay_cv": xldate(d.get(13)),
                "noi_dung": current_group or "Công văn cảnh báo",
                "nhom_nd": "Bên mua cần rà soát",
                "ghi_chu": f"Nguồn: TB DO KT1 MOI. Nhóm: {current_group}".strip(),
                "phuong": clean_text(d.get(4)),
                "cbql": clean_text(d.get(5)),
                "sheet": "F1",
            }
            extra = []
            if rec["phuong"]:
                extra.append(rec["phuong"])
            if rec["cbql"]:
                extra.append(f"CBQL: {rec['cbql']}")
            if extra:
                rec["ghi_chu"] += " | " + " | ".join(extra)
            add_unique(store, rec, stats, "TB DO KT1 MOI")


def extract_cv33(wb, store, stats, review: set):
    """CV 33 -> F1: B ten, C/D MST."""
    with wb.get_sheet("CV 33") as sh:
        for i, row in enumerate(sh.rows(), 1):
            if i == 1:
                continue
            d = rowdict(row)
            mst = norm_mst(d.get(3) or d.get(2))
            ten = clean_text(d.get(1))
            if not is_valid_mst(mst):
                continue
            review.add(mst)
            dia = " ".join(x for x in (clean_text(d.get(6)), clean_text(d.get(7)), clean_text(d.get(8))) if x)
            rec = {
                "mst": mst,
                "ten": ten,
                "dia_chi": dia,
                "so_cv": "33/PC03-Đ2",
                "ngay_cv": datetime(2025, 1, 20),
                "don_vi_bh": "Công an tỉnh Đồng Nai",
                "noi_dung": clean_text(d.get(12)) or "CV 33",
                "nhom_nd": "Bên mua cần rà soát",
                "tam_ngung": clean_text(d.get(12)),
                "ghi_chu": "Nguồn: CV 33.",
                "sheet": "F1",
            }
            add_unique(store, rec, stats, "CV 33")


def load_f0_extras(xlsx_path: Path, store: OrderedDict, stats: Counter):
    if not xlsx_path.is_file():
        return
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["F0"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        g = str(r[18] or "")
        gl = g.lower()
        if not any(
            x in gl
            for x in (
                "dot moi",
                "đợt mới",
                "đào minh",
                "dao minh",
                "b1152",
                "373a",
                "25.09.2026",
            )
        ):
            continue
        rec = {
            "mst": norm_mst(r[5]),
            "ten": clean_text(r[4]),
            "so_cv": norm_socv(r[1]),
            "ngay_cv": xldate(r[2]),
            "don_vi_bh": clean_text(r[3]),
            "noi_dung": clean_text(r[6]),
            "tam_ngung": clean_text(r[9]),
            "nhom_nd": clean_text(r[14]) or "Rủi ro do Công an chuyển",
            "dia_chi": clean_text(r[15]),
            "ngay_nhan": xldate(r[16]),
            "tinh_trang": clean_text(r[17]),
            "ghi_chu": g,
            "file": clean_text(r[19]),
            "sheet": "F0",
        }
        add_unique(store, rec, stats, "CV Cong an 2026")
    wb.close()


def load_f1_existing_in_review(xlsx_path: Path, store: OrderedDict, review: set, stats: Counter):
    if not xlsx_path.is_file():
        return
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["F1"]
    for r in ws.iter_rows(min_row=8, values_only=True):
        mst = norm_mst(r[8])
        if mst not in review:
            stats["f1_drop_not_review"] += 1
            continue
        g = str(r[18] or "")
        rec = {
            "mst": mst,
            "ten": clean_text(r[7]),
            "so_cv": norm_socv(r[1]),
            "ngay_cv": xldate(r[2]),
            "don_vi_bh": clean_text(r[3]),
            "ten_ban": clean_text(r[4]),
            "mst_ban": norm_mst(r[5]),
            "noi_dung": clean_text(r[6]),
            "tam_ngung": clean_text(r[9]),
            "nhom_nd": clean_text(r[14]) or "Bên mua cần rà soát",
            "dia_chi": clean_text(r[15]),
            "ngay_nhan": xldate(r[16]),
            "tinh_trang": clean_text(r[17]),
            "ghi_chu": g,
            "file": clean_text(r[19]),
            "sheet": "F1",
        }
        add_unique(store, rec, stats, "F1 hien huu trong DS ra soat")
    wb.close()


def sort_records(store: OrderedDict):
    recs = list(store.values())

    def key(r):
        ngay = r.get("ngay_cv") or datetime(1900, 1, 1)
        if not isinstance(ngay, datetime):
            ngay = datetime(1900, 1, 1)
        return (r.get("so_cv") or "", ngay, r.get("mst") or "")

    recs.sort(key=key)
    return recs


def stamp_flags(recs, mst_field="mst"):
    seen_cv = set()
    seen_mst = set()
    for r in recs:
        socv = (r.get("so_cv") or "").upper()
        mst = r.get(mst_field) or ""
        ngay = r.get("ngay_cv")
        nam = ngay.year if isinstance(ngay, datetime) else None
        r["nam_cv"] = nam
        r["khoa_cv"] = f"{socv}_{nam}" if socv and nam else (socv or "")
        r["cv_dau"] = 1 if socv and socv not in seen_cv else None
        if socv:
            seen_cv.add(socv)
        r["mst_dau"] = 1 if mst and mst not in seen_mst else None
        if mst:
            seen_mst.add(mst)
    return recs


def apply_fill(cell, fill, font, align=ALIGN_C):
    cell.fill = fill
    cell.font = font
    cell.alignment = align


def build_f0_sheet(wb: Workbook, recs):
    ws = wb.create_sheet("F0")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 26

    headers_vn = [
        "STT",
        "Số công văn",
        "Ngày công văn",
        "Đơn vị ban hành",
        "Tên công ty bán",
        "MST bên bán",
        "Nội dung chính",
        "Tạm ngưng kê khai",
        "Nhóm nội dung",
        "Địa chỉ bên bán",
        "Ngày nhận",
        "Tình trạng rà soát",
        "Ghi chú",
        "Đường dẫn file",
        "Năm công văn",
        "Dòng trùng",
        "Khóa công văn (ẩn)",
        "CV đầu tiên (ẩn)",
        "MST đầu tiên (ẩn)",
    ]
    headers_code = [
        "STT",
        "SoCV",
        "NgayCV",
        "DonViBH",
        "TenBan",
        "MSTBan",
        "NoiDung",
        "TamNgung",
        "NhomND",
        "DiaChi",
        "NgayNhan",
        "TinhTrang",
        "GhiChu",
        "File",
        "NamCV",
        "Trung",
        "KhoaCV",
        "CVDau",
        "MSTDau",
    ]
    last_col = len(headers_vn)
    last_letter = get_column_letter(last_col)

    ws.merge_cells(f"A1:{last_letter}1")
    ws.merge_cells(f"A2:{last_letter}2")
    ws.merge_cells(f"A4:{last_letter}4")
    ws.merge_cells("B5:D5")
    ws.merge_cells("E5:F5")
    ws.merge_cells("I5:N5")
    ws.merge_cells("O5:S5")

    c = ws.cell(1, 1, "SỔ THEO DÕI CÔNG VĂN CẢNH BÁO RỦI RO — DANH SÁCH BÊN BÁN (F0)")
    apply_fill(c, FILL_TITLE, FONT_TITLE, ALIGN_L)
    c = ws.cell(2, 1, "Danh sách doanh nghiệp rủi ro bán hàng (F0) do cơ quan thuế và công an cảnh báo.")
    apply_fill(c, FILL_SUB, FONT_SUB, ALIGN_L)
    ws.cell(3, 1, "Đơn vị sử dụng sổ: Chi cục Thuế").font = FONT_TNR
    ws.cell(3, 6, "Người lập:").font = FONT_TNR
    c = ws.cell(
        4,
        1,
        "Đọc từ trái sang phải: số công văn → ngày → đơn vị ban hành → bên bán → nội dung → xử lý. "
        "Đã bỏ cột thông tin bên mua và chi tiết hóa đơn. Phông Times New Roman.",
    )
    c.font = FONT_HINT

    groups = [
        (1, "STT", FILL_STT),
        (2, "THÔNG TIN CÔNG VĂN", FILL_CV),
        (5, "ĐƠN VỊ BÁN (F0)", FILL_BAN),
        (7, "NỘI DUNG", FILL_ND),
        (8, "TẠM NGƯNG KÊ KHAI", FILL_TN),
        (9, "THEO DÕI XỬ LÝ & LƯU TRỮ", FILL_XL),
        (15, "THÔNG TIN ĐỐI SOÁT", FILL_DS),
    ]
    for col, val, fill in groups:
        apply_fill(ws.cell(5, col, val), fill, FONT_GRP)

    fills_row6 = (
        [FILL_STT]
        + [FILL_CV] * 3
        + [FILL_BAN] * 2
        + [FILL_ND]
        + [FILL_TN]
        + [FILL_XL] * 6
        + [FILL_DS] * 5
    )
    for i, (vn, code) in enumerate(zip(headers_vn, headers_code), 1):
        apply_fill(ws.cell(6, i, vn), fills_row6[i - 1], FONT_H1)
        apply_fill(ws.cell(7, i, code), FILL_CODE, FONT_H2)

    widths = {
        1: 8,
        2: 23,
        3: 15,
        4: 29,
        5: 55,
        6: 17,
        7: 45,
        8: 21,
        9: 29,
        10: 36,
        11: 15,
        12: 23,
        13: 46,
        14: 18,
        15: 11,
        16: 13,
        17: 17,
        18: 11,
        19: 11,
    }
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = w

    for idx, r in enumerate(recs, 1):
        row = idx + 7
        ngay = r.get("ngay_cv")
        ngay_nhan = r.get("ngay_nhan")
        vals = [
            idx,
            r.get("so_cv") or None,
            ngay if isinstance(ngay, datetime) else None,
            r.get("don_vi_bh") or None,
            r.get("ten") or None,
            r.get("mst") or None,
            r.get("noi_dung") or None,
            r.get("tam_ngung") or None,
            r.get("nhom_nd") or None,
            r.get("dia_chi") or None,
            ngay_nhan if isinstance(ngay_nhan, datetime) else None,
            r.get("tinh_trang") or None,
            r.get("ghi_chu") or None,
            r.get("file") or None,
            r.get("nam_cv"),
            None,
            r.get("khoa_cv") or None,
            r.get("cv_dau"),
            r.get("mst_dau"),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row, c, v)
            cell.font = FONT_TNR
            if c in (3, 11) and isinstance(v, datetime):
                cell.number_format = "DD/MM/YYYY"
            if c in (1, 6, 15, 18, 19):
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if idx % 4000 == 0:
            log(f"  F0 wrote {idx}/{len(recs)}")

    last_row = 7 + len(recs)
    ws.auto_filter.ref = f"A6:{last_letter}{last_row}"
    return last_row


def build_f1_sheet(wb: Workbook, recs):
    ws = wb.create_sheet("F1")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 26

    headers_vn = [
        "STT",
        "Số công văn",
        "Ngày công văn",
        "Đơn vị ban hành",
        "Tên công ty bán",
        "MST bên bán",
        "Nội dung chính",
        "Tên công ty mua",
        "MST bên mua",
        "Tạm ngưng kê khai",
        "Kỳ kê khai",
        "Số hóa đơn",
        "Doanh thu (VNĐ)",
        "Tiền thuế GTGT (VNĐ)",
        "Nhóm nội dung",
        "Địa chỉ bên mua",
        "Ngày nhận",
        "Tình trạng rà soát",
        "Ghi chú",
        "Đường dẫn file",
        "Năm công văn",
        "Dòng trùng",
        "Khóa công văn (ẩn)",
        "CV đầu tiên (ẩn)",
        "MST đầu tiên (ẩn)",
    ]
    headers_code = [
        "STT",
        "SoCV",
        "NgayCV",
        "DonViBH",
        "TenBan",
        "MSTBan",
        "NoiDung",
        "TenMua",
        "MSTMua",
        "TamNgung",
        "KyKK",
        "SoHD",
        "DoanhThu",
        "ThueGTGT",
        "NhomND",
        "DiaChi",
        "NgayNhan",
        "TinhTrang",
        "GhiChu",
        "File",
        "NamCV",
        "Trung",
        "KhoaCV",
        "CVDau",
        "MSTDau",
    ]
    last_letter = "Y"
    ws.merge_cells("A1:Y1")
    ws.merge_cells("A2:Y2")
    ws.merge_cells("A4:Y4")
    ws.merge_cells("B5:D5")
    ws.merge_cells("E5:F5")
    ws.merge_cells("H5:I5")
    ws.merge_cells("K5:N5")
    ws.merge_cells("O5:T5")
    ws.merge_cells("U5:Y5")

    c = ws.cell(1, 1, "SỔ THEO DÕI CÔNG VĂN CẢNH BÁO RỦI RO — DANH SÁCH BÊN MUA (F1)")
    apply_fill(c, FILL_TITLE, FONT_TITLE, ALIGN_L)
    c = ws.cell(2, 1, "Danh sách doanh nghiệp mua hàng liên quan đến hóa đơn rủi ro (F1) cần rà soát, kiểm tra.")
    apply_fill(c, FILL_SUB, FONT_SUB, ALIGN_L)
    ws.cell(3, 1, "Đơn vị sử dụng sổ: Chi cục Thuế").font = FONT_TNR
    ws.cell(3, 6, "Người lập:").font = FONT_TNR
    c = ws.cell(
        4,
        1,
        "Đọc từ trái sang phải: số công văn → ngày → đơn vị ban hành → bên bán liên quan → bên mua rà soát → xử lý. "
        "Phông Times New Roman.",
    )
    c.font = FONT_HINT

    groups = [
        (1, "STT", FILL_STT),
        (2, "THÔNG TIN CÔNG VĂN", FILL_CV),
        (5, "ĐƠN VỊ BÁN LIÊN QUAN (F0)", FILL_BAN),
        (7, "NỘI DUNG", FILL_ND),
        (8, "ĐƠN VỊ MUA — ĐỐI TƯỢNG RÀ SOÁT (F1)", FILL_MUA),
        (10, "TẠM NGƯNG KÊ KHAI", FILL_TN),
        (11, "CHI TIẾT HÓA ĐƠN", FILL_HD),
        (15, "THEO DÕI XỬ LÝ & LƯU TRỮ", FILL_XL),
        (21, "THÔNG TIN ĐỐI SOÁT", FILL_DS),
    ]
    for col, val, fill in groups:
        apply_fill(ws.cell(5, col, val), fill, FONT_GRP)

    fills_row6 = (
        [FILL_STT]
        + [FILL_CV] * 3
        + [FILL_BAN] * 2
        + [FILL_ND]
        + [FILL_MUA] * 2
        + [FILL_TN]
        + [FILL_HD] * 4
        + [FILL_XL] * 6
        + [FILL_DS] * 5
    )
    for i, (vn, code) in enumerate(zip(headers_vn, headers_code), 1):
        apply_fill(ws.cell(6, i, vn), fills_row6[i - 1], FONT_H1)
        apply_fill(ws.cell(7, i, code), FILL_CODE, FONT_H2)

    widths = {
        1: 8,
        2: 23,
        3: 15,
        4: 29,
        5: 45,
        6: 17,
        7: 40,
        8: 46,
        9: 17,
        10: 21,
        11: 15,
        12: 19,
        13: 16,
        14: 18,
        15: 29,
        16: 36,
        17: 15,
        18: 23,
        19: 46,
        20: 18,
        21: 11,
        22: 13,
        23: 17,
        24: 11,
        25: 11,
    }
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = w

    for idx, r in enumerate(recs, 1):
        row = idx + 7
        ngay = r.get("ngay_cv")
        ngay_nhan = r.get("ngay_nhan")
        vals = [
            idx,
            r.get("so_cv") or None,
            ngay if isinstance(ngay, datetime) else None,
            r.get("don_vi_bh") or None,
            r.get("ten_ban") or None,
            r.get("mst_ban") or None,
            r.get("noi_dung") or None,
            r.get("ten") or None,
            r.get("mst") or None,
            r.get("tam_ngung") or None,
            None,
            None,
            None,
            None,
            r.get("nhom_nd") or None,
            r.get("dia_chi") or None,
            ngay_nhan if isinstance(ngay_nhan, datetime) else None,
            r.get("tinh_trang") or None,
            r.get("ghi_chu") or None,
            r.get("file") or None,
            r.get("nam_cv"),
            None,
            r.get("khoa_cv") or None,
            r.get("cv_dau"),
            r.get("mst_dau"),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row, c, v)
            cell.font = FONT_TNR
            if c in (3, 17) and isinstance(v, datetime):
                cell.number_format = "DD/MM/YYYY"
            if c in (1, 6, 9, 21, 24, 25):
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if idx % 3000 == 0:
            log(f"  F1 wrote {idx}/{len(recs)}")

    last_row = 7 + len(recs)
    ws.auto_filter.ref = f"A6:Y{last_row}"
    return last_row


def backup_current():
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BACKUP_DIR / f"SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.bak_{stamp}"
    shutil.copy2(SO_THEO_DOI, dest)
    return dest


def main():
    log(f"XLSB exists={XLSB.is_file()} size={XLSB.stat().st_size if XLSB.is_file() else 0}")
    log(f"SO exists={SO_THEO_DOI.is_file()} size={SO_THEO_DOI.stat().st_size if SO_THEO_DOI.is_file() else 0}")
    if not XLSB.is_file() or not SO_THEO_DOI.is_file():
        raise SystemExit("Thieu file nguon hoac so theo doi")

    inspect_copy = Path(r"C:\Users\NGOC\AppData\Local\Temp\SoTheoDoi_inspect.xlsx")
    extras_src = inspect_copy if inspect_copy.is_file() else SO_THEO_DOI

    f0_store: OrderedDict = OrderedDict()
    f1_store: OrderedDict = OrderedDict()
    review: set = set()
    st_cbrr, st_hdrr, st_ca = Counter(), Counter(), Counter()
    st_524, st_tb, st_cv33, st_ex, st_f1old = Counter(), Counter(), Counter(), Counter(), Counter()

    log("Mo xlsb...")
    xwb = open_workbook(str(XLSB))
    log("Trich CBRR -> F0")
    extract_cbrr(xwb, f0_store, st_cbrr)
    log(f"  CBRR {dict(st_cbrr)} F0 so far {len(f0_store)}")
    log("Trich HDRR -cuc -> F0")
    extract_hdrr_cuc(xwb, f0_store, st_hdrr)
    log(f"  HDRR {dict(st_hdrr)} F0 so far {len(f0_store)}")
    log("Trich CV cong an -> F0")
    extract_cv_cong_an(xwb, f0_store, st_ca)
    log(f"  CV cong an {dict(st_ca)} F0 so far {len(f0_store)}")
    log("Nap 524.xlsx chuan -> F0")
    extract_524_xlsx(f0_store, st_524)
    log(f"  524.xlsx {dict(st_524)} F0 so far {len(f0_store)}")
    log("Trich TB DO KT1 MOI -> F1")
    extract_tb_do_kt1(xwb, f1_store, st_tb, review)
    log(f"  TB {dict(st_tb)} F1 so far {len(f1_store)} review {len(review)}")
    log("Trich CV 33 -> F1")
    extract_cv33(xwb, f1_store, st_cv33, review)
    log(f"  CV33 {dict(st_cv33)} F1 so far {len(f1_store)} review {len(review)}")

    log("Giu dong F0 cong an 2026")
    load_f0_extras(extras_src, f0_store, st_ex)
    log(f"  extras {dict(st_ex)} F0 so far {len(f0_store)}")
    log("Giu dong F1 hien huu neu MST nam trong DS ra soat")
    load_f1_existing_in_review(extras_src, f1_store, review, st_f1old)
    log(f"  F1 old {dict(st_f1old)} F1 so far {len(f1_store)}")

    f0_recs = stamp_flags(sort_records(f0_store))
    f1_recs = stamp_flags(sort_records(f1_store))
    log(f"FINAL F0={len(f0_recs)} unique MST={len({r['mst'] for r in f0_recs})} unique SoCV={len({r.get('so_cv') for r in f0_recs})}")
    log(f"FINAL F1={len(f1_recs)} unique MST={len({r['mst'] for r in f1_recs})} unique SoCV={len({r.get('so_cv') for r in f1_recs})}")
    log(f"F0 co SoCV={sum(1 for r in f0_recs if r.get('so_cv'))} co Ngay={sum(1 for r in f0_recs if r.get('ngay_cv'))}")
    log(f"F1 co SoCV={sum(1 for r in f1_recs if r.get('so_cv'))} co Ngay={sum(1 for r in f1_recs if r.get('ngay_cv'))}")

    log("Sao luu so hien tai...")
    bak = backup_current()
    log(f"Backup: {bak}")

    out_tmp = Path(r"C:\Users\NGOC\AppData\Local\Temp\SoTheoDoi_cleaned.xlsx")
    log("Ghi workbook moi...")
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    n0 = build_f0_sheet(wb, f0_recs)
    n1 = build_f1_sheet(wb, f1_recs)
    wb.save(out_tmp)
    log(f"Saved tmp {out_tmp} F0_last={n0} F1_last={n1} size={out_tmp.stat().st_size}")

    shutil.copy2(out_tmp, SO_THEO_DOI)
    log(f"Copied to {SO_THEO_DOI} size={SO_THEO_DOI.stat().st_size}")
    for p in SYNC_COPIES:
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(out_tmp, p)
            log(f"Synced {p}")
        except Exception as e:
            log(f"Skip sync {p}: {e}")

    stats_path = Path(r"C:\Users\NGOC\AppData\Local\Temp\so_theo_doi_clean_stats.txt")
    stats_path.write_text(
        "\n".join(
            [
                f"backup={bak}",
                f"CBRR={dict(st_cbrr)}",
                f"HDRR={dict(st_hdrr)}",
                f"CVCA={dict(st_ca)}",
                f"F524={dict(st_524)}",
                f"TB={dict(st_tb)}",
                f"CV33={dict(st_cv33)}",
                f"F0EX={dict(st_ex)}",
                f"F1OLD={dict(st_f1old)}",
                f"F0={len(f0_recs)} MST={len({r['mst'] for r in f0_recs})}",
                f"F1={len(f1_recs)} MST={len({r['mst'] for r in f1_recs})}",
                f"review={len(review)}",
                f"out={SO_THEO_DOI} size={SO_THEO_DOI.stat().st_size}",
            ]
        ),
        encoding="utf-8",
    )
    log(f"STATS {stats_path.read_text(encoding='utf-8')}")
    log("DONE")


if __name__ == "__main__":
    main()
