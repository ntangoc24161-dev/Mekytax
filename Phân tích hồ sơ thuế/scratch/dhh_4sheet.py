# -*- coding: utf-8 -*-
"""Skill tong-hop-hddt-bang-ke-4-sheet cho CTY TNHH DAI HUY HOANG MST 3602179712, ky 2021-2025.

File 1: Tong_hop_HDDT_BangKe_3602179712_2021-2025.xlsx
  HDDT_Mua_vao, HDDT_Ban_ra, Bang_ke_Mua_vao, Bang_ke_Ban_ra,
  Tong_hop_doi_tac, Tong_hop_trang_thai_rui_ro
File 2: Tong_hop_NCC_mua_vao_CV_trang_thai_3602179712_2021-2025.xlsx
  Phan_1_F1_NNT, Bang_2_1_NCC_F0, Bang_2_2_Toan_bo_NCC, Bang_2_3_Ky_ke_khai_CT, Phan_3_Nhan_xet
NCC mua vao: chi gan co F0. NNT: tra F0, F1, Cong an.
"""
from __future__ import annotations

import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST_NNT = "3602179712"
TEN_NNT = "CÔNG TY TNHH ĐẠI HUY HOÀNG"
KY = "2021-2025"
YEARS = {2021, 2022, 2023, 2024, 2025}
TODAY = date(2026, 11, 25)

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\DAI_HUY_HOANG_3602179712"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
DIR_HDDT = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT.xlsx"
DIR_BK = ROOT / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
DIR_DANHBA = ROOT / "3. DU LIEU BO SUNG & TRA CUU QLT" / "3.2. Danh ba trang thai NNT" / "Danh ba NNT.xls"
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
OUT_CT_G = ROOT / "4. KET QUA PHAN TICH" / "cau-thanh"
OUT_CT_V = VAULT / "ket-qua" / "cau-thanh"
OUT_FILE1 = f"Tong_hop_HDDT_BangKe_{MST_NNT}_{KY}.xlsx"
OUT_FILE2 = f"Tong_hop_NCC_mua_vao_CV_trang_thai_{MST_NNT}_{KY}.xlsx"

FONT_TITLE = Font(name="Times New Roman", size=12, bold=True, color="1F497D")
FONT_SUBTITLE = Font(name="Times New Roman", size=11, bold=True)
FONT_HEADER = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_REGULAR = Font(name="Times New Roman", size=10)
FONT_BOLD = Font(name="Times New Roman", size=10, bold=True)
FONT_RED_BOLD = Font(name="Times New Roman", size=10, bold=True, color="C00000")
FONT_ORG_BOLD = Font(name="Times New Roman", size=10, bold=True, color="B25900")
FILL_HEADER = PatternFill("solid", fgColor="1F497D")
FILL_TOTAL = PatternFill("solid", fgColor="DCE6F1")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
TOTAL_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="double", color="000000"),
)
NUM_FMT = "#,##0"
DATE_FMT = "dd/mm/yyyy"
WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)


def lp(p) -> str:
    s = str(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s


def chuan_hoa_mst(val) -> str:
    if val is None:
        return ""
    s = str(val).strip().replace(" ", "").replace("\xa0", "")
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    s = re.sub(r"[^\d\-]", "", s)
    if not s:
        return ""
    digits = s.replace("-", "")
    if len(s) == 14 and "-" in s:
        return s
    if len(digits) == 10:
        return digits
    if len(digits) == 9:
        return "0" + digits
    if len(digits) == 12:
        return "0" + digits[:9] + "-" + digits[9:]
    if len(digits) == 13:
        return digits[:10] + "-" + digits[10:]
    return s


def mst_clean(s) -> str:
    if not s:
        return ""
    return re.sub(r"\D", "", str(s).split("-")[0].strip())


def mst_keys(mst_str, parent10=True):
    s = str(mst_str or "").strip()
    keys = set()
    if s:
        keys.add(s)
    d = re.sub(r"\D", "", s)
    if d:
        keys.add(d)
        if parent10 and len(d) >= 10:
            keys.add(d[:10])
    c = chuan_hoa_mst(s)
    if c:
        keys.add(c)
        keys.add(c.replace("-", ""))
        if parent10:
            keys.add(c.split("-")[0])
    m = mst_clean(s)
    if m:
        keys.add(m)
    return {k for k in keys if k}


def danhba_keys(mst_str):
    """Exact MST only. Do not index a 13-digit branch onto the 10-digit parent."""
    return mst_keys(mst_str, parent10=False)


def sanitize_val(val):
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip()
        if s.startswith("="):
            return "'" + s
        return s
    return val


def nnum(v, default=None):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\xa0", "").replace("%", "")
    if s.count(".") > 1 and s.count(",") == 0:
        s = s.replace(".", "")
    elif s.count(",") == 1 and s.count(".") > 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


def as_date(v, datemode=0):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            return datetime(*xlrd.xldate_as_tuple(float(v), datemode)).date()
        except Exception:
            try:
                return datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(v)).date()
            except Exception:
                return None
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def so_hd_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s


def is_label_row(so, ten) -> bool:
    blob = " ".join([str(so or ""), str(ten or "")]).casefold().strip()
    if not blob:
        return True
    if any(k in blob for k in ["công ty", "cong ty", "tổng công ty", "tong cong ty"]):
        return False
    if blob in {"tổng", "tong", "tổng cộng", "tong cong"}:
        return True
    if blob.startswith("tổng cộng") or blob.startswith("tong cong"):
        return True
    if "nhóm hhdv" in blob or blob.startswith("nhóm"):
        return True
    return False


def is_huy(st) -> bool:
    t = str(st or "").casefold()
    return any(
        k in t
        for k in [
            "bị xóa bỏ",
            "bi xoa bo",
            "bị hủy",
            "bi huy",
            "bị thay thế",
            "bi thay the",
            "hủy bỏ",
            "huy bo",
        ]
    )


def fmt_money_vn(n) -> str:
    try:
        v = int(round(float(n or 0)))
    except (TypeError, ValueError):
        return "0"
    s = f"{v:,}".replace(",", ".")
    return s


def autowidth(ws, min_w=10, cap=55):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in list(col)[:250]:
            if cell.row in (1, 2):
                continue
            v = str(cell.value or "")
            if v.startswith("="):
                continue
            max_len = max(max_len, min(len(v), cap))
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, cap), min_w)


def style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row, c)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = WRAP
        cell.border = THIN_BORDER


def write_total_row(ws, row, n_cols, money_cols, data_start=4):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row, c)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        cell.font = FONT_BOLD
    ws.cell(row, 1, value="TỔNG CỘNG")
    ws.cell(row, 1).alignment = Alignment(horizontal="center", vertical="center")
    if row <= data_start:
        return
    for c in money_cols:
        letter = get_column_letter(c)
        cell = ws.cell(row, c, value=f"=SUM({letter}{data_start}:{letter}{row-1})")
        cell.font = FONT_BOLD
        cell.number_format = NUM_FMT
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        cell.alignment = Alignment(horizontal="right", vertical="center")


def finish_sheet(ws, n_cols):
    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_title_rows = "1:3"
    ws.row_dimensions[3].height = 32
    autowidth(ws)


def lookup_map(d, mst_str):
    for k in mst_keys(mst_str):
        if k in d:
            return d[k]
    return None


def load_danhba_map(p_danhba: Path):
    if not p_danhba.exists():
        print(f"WARNING Danh ba missing: {p_danhba}")
        return {}
    print(f"Loading Danh ba NNT from {p_danhba}...")
    tree = ET.parse(lp(p_danhba) if os.path.exists(lp(p_danhba)) else p_danhba)
    root = tree.getroot()

    def ss_index(cell, default):
        for key, val in cell.attrib.items():
            if key == "Index" or key.endswith("}Index"):
                return int(val)
        return default

    def ss_text(cell):
        for child in list(cell):
            if child.tag.endswith("}Data") or child.tag == "Data":
                return (child.text or "").strip()
        return ""

    danhba = {}
    n_row = 0
    for row in root.iter():
        if not (row.tag.endswith("}Row") or row.tag == "Row"):
            continue
        cells = {}
        idx = 1
        for cell in list(row):
            if not (cell.tag.endswith("}Cell") or cell.tag == "Cell"):
                continue
            idx = ss_index(cell, idx)
            cells[idx] = ss_text(cell)
            idx += 1
        mst = cells.get(4, "").strip()
        if mst and mst != "Mã số thuế":
            rec = {
                "tt": cells.get(45, "").strip() or cells.get(44, "").strip() or cells.get(53, "").strip(),
                "cqt": cells.get(2, "").strip(),
                "ten": cells.get(5, "").strip(),
            }
            n_row += 1
            for k in danhba_keys(mst):
                danhba[k] = rec
    print(f"Loaded Danh ba unique rows {n_row}, keys {len(danhba)}")
    return danhba


def get_danhba_info(danhba_map, mst_str):
    for k in danhba_keys(mst_str):
        if k in danhba_map:
            hit = danhba_map[k]
            return hit.get("tt", ""), hit.get("cqt", "")
    return "", ""


def load_cvcbrr_and_ca():
    print("Loading CVCBRR...")
    f0_map = {}
    f1_by_seller = {}
    f1_by_buyer = defaultdict(list)
    if CVCBRR.exists():
        wb_cv = load_workbook(lp(CVCBRR) if os.path.exists(lp(CVCBRR)) else CVCBRR, read_only=True, data_only=True)
        if "F0" in wb_cv.sheetnames:
            for r in wb_cv["F0"].iter_rows(min_row=8, values_only=True):
                raw = r[1] if r and len(r) > 1 else ""
                m = mst_clean(raw)
                if not m:
                    continue
                rec = {
                    "mst": chuan_hoa_mst(raw),
                    "ten": str(r[2] or "").strip() if len(r) > 2 else "",
                    "cv": str(r[3] or "").strip() if len(r) > 3 else "",
                    "ngay": str(r[4])[:10] if len(r) > 4 and r[4] else "",
                    "co_quan": str(r[5] or "").strip() if len(r) > 5 else "",
                    "nd": str(r[6] or "").strip() if len(r) > 6 else "",
                }
                for k in mst_keys(raw):
                    f0_map[k] = rec
        if "F1" in wb_cv.sheetnames:
            for r in wb_cv["F1"].iter_rows(min_row=8, values_only=True):
                m_ban = mst_clean(r[1] if r and len(r) > 1 else "")
                m_mua = mst_clean(r[8] if r and len(r) > 8 else "")
                rec = {
                    "mst_ban": m_ban,
                    "ten_ban": str(r[2] or "").strip() if len(r) > 2 else "",
                    "cv": str(r[3] or "").strip() if len(r) > 3 else "",
                    "ngay": str(r[4])[:10] if len(r) > 4 and r[4] else "",
                    "co_quan": str(r[5] or "").strip() if len(r) > 5 else "",
                    "m_mua": m_mua,
                    "ten_mua": str(r[9] or "").strip() if len(r) > 9 else "",
                }
                if m_ban:
                    f1_by_seller[m_ban] = rec
                if m_mua:
                    f1_by_buyer[m_mua].append(rec)
        wb_cv.close()
    print(f"Loaded F0 keys {len(f0_map)}, F1 sellers {len(f1_by_seller)}, F1 buyers {len(f1_by_buyer)}")

    print("Loading CA...")
    ca_map = {}
    if CA.exists():
        wb_ca = load_workbook(lp(CA) if os.path.exists(lp(CA)) else CA, read_only=True, data_only=True)
        if "NHAP_HO_SO" in wb_ca.sheetnames:
            ws = wb_ca["NHAP_HO_SO"]
            last_cv, last_ngay, last_cq, last_nd = "", "", "", ""
            for r in list(ws.iter_rows(values_only=True))[4:]:
                if len(r) > 2 and r[2] is not None and str(r[2]).strip():
                    last_cv = str(r[2]).strip()
                    last_ngay = str(r[3])[:10] if len(r) > 3 and r[3] is not None else ""
                    last_cq = str(r[4] or "").strip() if len(r) > 4 else ""
                    last_nd = str(r[7] or "").strip() if len(r) > 7 else ""
                mst = mst_clean(r[9]) if len(r) > 9 and r[9] else ""
                if mst:
                    rec = {
                        "mst": mst,
                        "ten": str(r[10] or "").strip() if len(r) > 10 else "",
                        "cv": last_cv,
                        "ngay": last_ngay,
                        "co_quan": last_cq,
                        "nd": last_nd,
                    }
                    for k in mst_keys(mst):
                        ca_map[k] = rec
        for sname in ["TAM_NGUNG", "DANH_SACH_LOC"]:
            if sname in wb_ca.sheetnames:
                for r in list(wb_ca[sname].iter_rows(values_only=True))[2:]:
                    for val in r:
                        m = mst_clean(val)
                        if len(m) in (10, 13) and (m.startswith("0") or m.startswith("3")):
                            if m not in ca_map:
                                rec = {
                                    "mst": m,
                                    "ten": "",
                                    "cv": f"Theo dõi Công an ({sname})",
                                    "ngay": "",
                                    "co_quan": "Công an",
                                    "nd": f"Nằm trong danh sách {sname}",
                                }
                                for k in mst_keys(m):
                                    ca_map.setdefault(k, rec)
        wb_ca.close()
    print(f"Loaded CA keys {len(ca_map)}")
    return f0_map, f1_by_seller, f1_by_buyer, ca_map


def fmt_cv(info, prefix=""):
    if not info:
        return ""
    cv = info.get("cv") or ""
    ngay = info.get("ngay") or ""
    cq = info.get("co_quan") or ""
    parts = []
    if prefix:
        parts.append(prefix)
    if cv:
        parts.append(f"CV {cv}")
    if ngay:
        parts.append(f"ngày {ngay}")
    if cq:
        parts.append(f"- {cq}")
    return " ".join(parts).strip()


def get_ncc_f0(mst_str, f0_map):
    info = lookup_map(f0_map, mst_str)
    if not info:
        return "", ""
    return "F0", fmt_cv(info, "F0:")


def nnt_lookup(mst, f0_map, f1_by_buyer, ca_map):
    f0 = lookup_map(f0_map, mst)
    f1s = []
    for k in mst_keys(mst):
        if k in f1_by_buyer:
            f1s = f1_by_buyer[k]
            break
    ca = lookup_map(ca_map, mst)
    return f0, f1s, ca


def load_hddt():
    print(f"Reading HDDT {DIR_HDDT}")
    wb = load_workbook(lp(DIR_HDDT) if os.path.exists(lp(DIR_HDDT)) else DIR_HDDT, data_only=True)
    out = {}
    skip = {}
    for sheet, chieu in (("Mua vao", "MUA"), ("Ban ra", "BAN")):
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        header = list(rows[0])
        data = []
        n_skip = 0
        for r in rows[1:]:
            if not r or not any(r):
                continue
            st = str(r[20] if len(r) > 20 else "" or "")
            if is_huy(st):
                n_skip += 1
                continue
            data.append(list(r))
        out[chieu] = (header, data)
        skip[chieu] = n_skip
        print(f"  HDDT {sheet}: keep {len(data)} skip_huy {n_skip}")
    wb.close()
    return out, skip


def parse_bang_ke():
    files = sorted(
        p
        for p in DIR_BK.rglob("*")
        if p.is_file()
        and p.suffix.lower() == ".xls"
        and not p.name.startswith("Bang_ke_mua_vao_ban_ra_tong_hop")
        and not p.name.startswith("Tong_hop")
        and not p.name.startswith("Danh")
    )
    print(f"BK source files {len(files)}")
    all_rows = []
    for p in files:
        m = re.search(r"(20\d{2})\s*Q([1-4])", p.name, re.I)
        nam_hint = int(m.group(1)) if m else None
        quy_hint = int(m.group(2)) if m else None
        try:
            book = xlrd.open_workbook(lp(p) if os.path.exists(lp(p)) else str(p))
        except Exception as e:
            print(f"  ERR {p.name}: {e}")
            continue
        n_file = 0
        for sname in book.sheet_names():
            su = sname.strip().upper()
            if su in {"VAO", "MUA"} or "vao" in sname.lower() or "mua" in sname.lower():
                chieu = "MUA"
            elif su in {"RA", "BAN"} or "ban" in sname.lower() or su == "RA":
                chieu = "BAN"
            else:
                continue
            sh = book.sheet_by_name(sname)
            hdr = None
            start = 1
            for i in range(min(6, sh.nrows)):
                vals = [str(sh.cell_value(i, j)).strip().upper() for j in range(min(12, sh.ncols))]
                joined = " ".join(vals)
                if "SO_HD" in vals or "SỐ HĐ" in joined or "SO HD" in joined:
                    hdr = {vals[j]: j for j in range(len(vals)) if vals[j]}
                    start = i + 1
                    break
            if hdr is None:
                continue

            def g(row, *keys):
                for k in keys:
                    ku = k.upper()
                    if ku in hdr:
                        return sh.cell_value(row, hdr[ku])
                return None

            for r in range(start, sh.nrows):
                so = g(r, "SO_HD", "SO HD")
                ten = str(g(r, "TENKH") or "").strip()
                if is_label_row(so, ten):
                    continue
                so_s = so_hd_str(so)
                if not so_s:
                    continue
                ngay = as_date(g(r, "NGAY_HD"), book.datemode)
                y = ngay.year if ngay else nam_hint
                if y not in YEARS:
                    continue
                mst_goc = g(r, "MS_DN", "MST")
                mst_c = chuan_hoa_mst(mst_goc)
                ds = nnum(g(r, "TTVND"), 0) or 0
                thue = nnum(g(r, "THUEVND", "THUE"), 0) or 0
                ts = g(r, "TS_GTGT", "TS")
                ts_n = nnum(ts)
                if ts_n is not None and ts_n > 1:
                    ts_n = ts_n / 100.0
                quy = ((ngay.month - 1) // 3 + 1) if ngay else quy_hint
                all_rows.append(
                    {
                        "Chieu": chieu,
                        "Nam": y,
                        "Quy": quy,
                        "Thang": ngay.month if ngay else None,
                        "So_HD": so_s,
                        "Ngay_HD": ngay,
                        "MST_goc": str(mst_goc or "").strip(),
                        "MST_chuan": mst_c,
                        "Ten_DoiTac": ten,
                        "Mat_hang": str(g(r, "MATHANG") or "").strip(),
                        "GT_chua_thue": ds,
                        "Thue_suat": ts_n if ts_n is not None else (str(ts).strip() if ts not in (None, "") else ""),
                        "Thue_GTGT": thue,
                        "Tong_tien": ds + thue,
                        "Ghi_chu": str(g(r, "GHICHU", "GC") or "").strip(),
                        "File_nguon": p.name,
                    }
                )
                n_file += 1
        print(f"  {p.name} -> {n_file}")
    n_m = sum(1 for r in all_rows if r["Chieu"] == "MUA")
    n_b = sum(1 for r in all_rows if r["Chieu"] == "BAN")
    print(f"BK TOTAL {len(all_rows)} MUA {n_m} BAN {n_b}")
    return all_rows


def empty_ncc(mst, ten, tt, cqt, risk_flag, risk_cv):
    return {
        "mst": mst,
        "ten": ten or "",
        "tt": tt or "",
        "cqt": cqt or "",
        "risk_flag": risk_flag or "",
        "risk_cv": risk_cv or "",
        "bk_count": 0,
        "bk_ds": 0.0,
        "bk_thue": 0.0,
        "hddt_count": 0,
        "hddt_ds": 0.0,
        "hddt_thue": 0.0,
    }


def ensure_ncc(all_ncc, mst_c, ten, tt, cqt, risk_flag, risk_cv):
    if not mst_c:
        return None
    rec = all_ncc.get(mst_c)
    if rec is None:
        rec = empty_ncc(mst_c, ten, tt, cqt, risk_flag, risk_cv)
        all_ncc[mst_c] = rec
    else:
        if ten and len(str(ten)) > len(rec["ten"]):
            rec["ten"] = str(ten)
        if not rec["tt"] and tt:
            rec["tt"] = tt
        if not rec["cqt"] and cqt:
            rec["cqt"] = cqt
        if not rec["risk_flag"] and risk_flag:
            rec["risk_flag"] = risk_flag
            rec["risk_cv"] = risk_cv
    return rec


def src_ncc(ncc):
    if ncc["hddt_count"] > 0 and ncc["bk_count"] > 0:
        return "Chung"
    if ncc["hddt_count"] > 0:
        return "HĐĐT"
    return "Bảng kê"


def abnormal_status(tt_s: str) -> bool:
    t = (tt_s or "").lower()
    if not t:
        return False
    return "đang hoạt động" not in t and "dang hoat dong" not in t


def put_cell(ws, r, c, val, money=False, date_v=False, center=False, risk=""):
    cell = ws.cell(r, c, value=sanitize_val(val))
    cell.font = FONT_REGULAR
    cell.border = THIN_BORDER
    if money:
        if isinstance(val, (int, float)):
            cell.value = float(val)
        cell.number_format = NUM_FMT
        cell.alignment = Alignment(horizontal="right", vertical="center")
    elif date_v and isinstance(val, date):
        cell.number_format = DATE_FMT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    elif center:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    if risk == "F0":
        cell.fill = FILL_RED
        if c in (2, 3, 6) or True:
            pass
    return cell


def apply_risk_row(ws, row, n_cols, flag, flag_col):
    if flag == "F0":
        for c in range(1, n_cols + 1):
            cell = ws.cell(row, c)
            cell.fill = FILL_RED
        ws.cell(row, flag_col).font = FONT_RED_BOLD


def write_title(ws, text, n_cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(n_cols, 8))
    c = ws.cell(1, 1, value=text)
    c.font = FONT_TITLE


def run():
    danhba_map = load_danhba_map(DIR_DANHBA)
    f0_map, f1_by_seller, f1_by_buyer, ca_map = load_cvcbrr_and_ca()
    hddt, skip_huy = load_hddt()
    bk_rows = parse_bang_ke()
    bk_mua = sorted([r for r in bk_rows if r["Chieu"] == "MUA"], key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))
    bk_ban = sorted([r for r in bk_rows if r["Chieu"] == "BAN"], key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))

    all_ncc = {}
    all_invoices = []
    partners = {}

    def add_partner(mst, ten, src):
        m = chuan_hoa_mst(mst)
        if not m:
            return
        rec = partners.setdefault(m, {"mst": m, "ten": str(ten or ""), "src": set()})
        if ten and len(str(ten)) > len(rec["ten"]):
            rec["ten"] = str(ten)
        rec["src"].add(src)

    # ----- File 1 -----
    wb1 = Workbook()
    default1 = wb1.active

    header_m, data_m = hddt["MUA"]
    new_h_m = list(header_m[:6]) + ["MST_chuan"] + list(header_m[6:]) + [
        "Tình trạng hoạt động",
        "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro",
        "Thông tin công văn CA / CVCBRR",
    ]
    ws1 = wb1.create_sheet("HDDT_Mua_vao")
    write_title(ws1, f"TỔNG HỢP HÓA ĐƠN ĐIỆN TỬ MUA VÀO - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}", len(new_h_m))
    for i, h in enumerate(new_h_m, 1):
        ws1.cell(3, i, value=h)
    style_header_row(ws1, 3, len(new_h_m))
    money_m = [15, 16, 17, 18, 19]
    rr = 4
    stt = 1
    n_hddt_m_2021 = 0
    years_m = defaultdict(int)
    for r in data_m:
        mst_goc = r[5] if len(r) > 5 else ""
        mst_c = chuan_hoa_mst(mst_goc)
        ten_nb = str(r[6] or "").strip() if len(r) > 6 else ""
        ngay = as_date(r[4] if len(r) > 4 else None)
        if ngay and ngay.year == 2021:
            n_hddt_m_2021 += 1
        if ngay:
            years_m[ngay.year] += 1
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)
        risk_flag, risk_cv = get_ncc_f0(mst_c, f0_map)
        new_row = list(r[:6]) + [mst_c] + list(r[6:]) + [tt_hd, cqt_ql, risk_flag, risk_cv]
        new_row[0] = stt
        ds_val = nnum(r[13] if len(r) > 13 else 0, 0) or 0
        thue_val = nnum(r[14] if len(r) > 14 else 0, 0) or 0
        rec = ensure_ncc(all_ncc, mst_c, ten_nb, tt_hd, cqt_ql, risk_flag, risk_cv)
        if rec:
            rec["hddt_count"] += 1
            rec["hddt_ds"] += ds_val
            rec["hddt_thue"] += thue_val
            all_invoices.append(
                {
                    "mst": mst_c,
                    "ten": ten_nb,
                    "nguon": "HĐĐT",
                    "so_hd": so_hd_str(r[3] if len(r) > 3 else ""),
                    "ngay_hd": ngay.strftime("%d/%m/%Y") if ngay else "",
                    "mat_hang": "",
                    "ds": ds_val,
                    "thue": thue_val,
                    "risk_flag": risk_flag,
                    "risk_cv": risk_cv,
                    "tt": tt_hd,
                    "cqt": cqt_ql,
                }
            )
        add_partner(mst_c, ten_nb, "HDDT_Mua")
        for c, val in enumerate(new_row, 1):
            cell = ws1.cell(rr, c, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if c in money_m:
                nv = nnum(val)
                if nv is not None:
                    cell.value = nv
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (1, 2, 5, 6, 7):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if c == 5 and ngay:
                cell.value = ngay
                cell.number_format = DATE_FMT
            if c == 26:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if risk_flag == "F0":
                    cell.font = FONT_RED_BOLD
                    cell.fill = FILL_RED
        if risk_flag == "F0":
            apply_risk_row(ws1, rr, len(new_h_m), "F0", 26)
        stt += 1
        rr += 1
    write_total_row(ws1, rr, len(new_h_m), money_m)
    finish_sheet(ws1, len(new_h_m))
    n1 = stt - 1

    header_b, data_b = hddt["BAN"]
    new_h_b = list(header_b[:8]) + ["MST_chuan"] + list(header_b[8:]) + [
        "Tình trạng hoạt động",
        "Cơ quan thuế quản lý",
    ]
    ws2 = wb1.create_sheet("HDDT_Ban_ra")
    write_title(ws2, f"TỔNG HỢP HÓA ĐƠN ĐIỆN TỬ BÁN RA - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}", len(new_h_b))
    for i, h in enumerate(new_h_b, 1):
        ws2.cell(3, i, value=h)
    style_header_row(ws2, 3, len(new_h_b))
    money_b = [15, 16, 17, 18, 19]
    rr = 4
    stt = 1
    years_b = defaultdict(int)
    for r in data_b:
        kq = str(r[21] if len(r) > 21 else "" or "").casefold()
        if "không đủ điều kiện" in kq:
            continue
        mst_goc = r[7] if len(r) > 7 else ""
        mst_c = chuan_hoa_mst(mst_goc)
        ten_nm = str(r[8] or "").strip() if len(r) > 8 else ""
        ngay = as_date(r[4] if len(r) > 4 else None)
        if ngay:
            years_b[ngay.year] += 1
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)
        new_row = list(r[:8]) + [mst_c] + list(r[8:]) + [tt_hd, cqt_ql]
        new_row[0] = stt
        add_partner(mst_c, ten_nm, "HDDT_Ban")
        for c, val in enumerate(new_row, 1):
            cell = ws2.cell(rr, c, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if c in money_b:
                nv = nnum(val)
                if nv is not None:
                    cell.value = nv
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (1, 2, 5, 8, 9):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if c == 5 and ngay:
                cell.value = ngay
                cell.number_format = DATE_FMT
        stt += 1
        rr += 1
    write_total_row(ws2, rr, len(new_h_b), money_b)
    finish_sheet(ws2, len(new_h_b))
    n2 = stt - 1

    h_bk_m = [
        "STT",
        "Nam",
        "Quy",
        "Thang",
        "So_HD",
        "Ngay_HD",
        "MST_goc",
        "MST_chuan",
        "Ten_DoiTac",
        "Mat_hang",
        "GT_chua_thue",
        "Thue_suat",
        "Thue_GTGT",
        "Tong_tien",
        "Ghi_chu",
        "File_nguon",
        "Tình trạng hoạt động",
        "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro",
        "Thông tin công văn CA / CVCBRR",
    ]
    ws3 = wb1.create_sheet("Bang_ke_Mua_vao")
    write_title(ws3, f"TỔNG HỢP BẢNG KÊ MUA VÀO (NNT NỘP) - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}", len(h_bk_m))
    for i, h in enumerate(h_bk_m, 1):
        ws3.cell(3, i, value=h)
    style_header_row(ws3, 3, len(h_bk_m))
    money_bk = [11, 13, 14]
    rr = 4
    stt = 1
    for rec in bk_mua:
        mst_c = rec["MST_chuan"]
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)
        risk_flag, risk_cv = get_ncc_f0(mst_c, f0_map)
        ncc = ensure_ncc(all_ncc, mst_c, rec["Ten_DoiTac"], tt_hd, cqt_ql, risk_flag, risk_cv)
        if ncc:
            ncc["bk_count"] += 1
            ncc["bk_ds"] += rec["GT_chua_thue"] or 0
            ncc["bk_thue"] += rec["Thue_GTGT"] or 0
            all_invoices.append(
                {
                    "mst": mst_c,
                    "ten": rec["Ten_DoiTac"],
                    "nguon": "Bảng kê",
                    "so_hd": rec["So_HD"],
                    "ngay_hd": rec["Ngay_HD"].strftime("%d/%m/%Y") if rec["Ngay_HD"] else "",
                    "mat_hang": rec["Mat_hang"],
                    "ds": rec["GT_chua_thue"] or 0,
                    "thue": rec["Thue_GTGT"] or 0,
                    "risk_flag": risk_flag,
                    "risk_cv": risk_cv,
                    "tt": tt_hd,
                    "cqt": cqt_ql,
                }
            )
        add_partner(mst_c, rec["Ten_DoiTac"], "Bang_ke_Mua")
        vals = [
            stt,
            rec["Nam"],
            rec["Quy"],
            rec["Thang"],
            rec["So_HD"],
            rec["Ngay_HD"],
            rec["MST_goc"],
            mst_c,
            rec["Ten_DoiTac"],
            rec["Mat_hang"],
            rec["GT_chua_thue"],
            rec["Thue_suat"],
            rec["Thue_GTGT"],
            rec["Tong_tien"],
            rec["Ghi_chu"],
            rec["File_nguon"],
            tt_hd,
            cqt_ql,
            risk_flag,
            risk_cv,
        ]
        for c, val in enumerate(vals, 1):
            cell = ws3.cell(rr, c, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if c in money_bk:
                nv = nnum(val)
                if nv is not None:
                    cell.value = nv
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (1, 2, 3, 4, 5, 7, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if c == 6 and isinstance(val, date):
                cell.number_format = DATE_FMT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if c == 19 and risk_flag == "F0":
                cell.font = FONT_RED_BOLD
                cell.fill = FILL_RED
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if risk_flag == "F0":
            apply_risk_row(ws3, rr, len(h_bk_m), "F0", 19)
        stt += 1
        rr += 1
    write_total_row(ws3, rr, len(h_bk_m), money_bk)
    finish_sheet(ws3, len(h_bk_m))
    n3 = stt - 1

    h_bk_b = [
        "STT",
        "Nam",
        "Quy",
        "Thang",
        "So_HD",
        "Ngay_HD",
        "MST_goc",
        "MST_chuan",
        "Ten_DoiTac",
        "Mat_hang",
        "GT_chua_thue",
        "Thue_suat",
        "Thue_GTGT",
        "Tong_tien",
        "Ghi_chu",
        "File_nguon",
        "Tình trạng hoạt động",
        "Cơ quan thuế quản lý",
    ]
    ws4 = wb1.create_sheet("Bang_ke_Ban_ra")
    write_title(ws4, f"TỔNG HỢP BẢNG KÊ BÁN RA (NNT NỘP) - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}", len(h_bk_b))
    for i, h in enumerate(h_bk_b, 1):
        ws4.cell(3, i, value=h)
    style_header_row(ws4, 3, len(h_bk_b))
    rr = 4
    stt = 1
    for rec in bk_ban:
        mst_c = rec["MST_chuan"]
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)
        add_partner(mst_c, rec["Ten_DoiTac"], "Bang_ke_Ban")
        vals = [
            stt,
            rec["Nam"],
            rec["Quy"],
            rec["Thang"],
            rec["So_HD"],
            rec["Ngay_HD"],
            rec["MST_goc"],
            mst_c,
            rec["Ten_DoiTac"],
            rec["Mat_hang"],
            rec["GT_chua_thue"],
            rec["Thue_suat"],
            rec["Thue_GTGT"],
            rec["Tong_tien"],
            rec["Ghi_chu"],
            rec["File_nguon"],
            tt_hd,
            cqt_ql,
        ]
        for c, val in enumerate(vals, 1):
            cell = ws4.cell(rr, c, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if c in money_bk:
                nv = nnum(val)
                if nv is not None:
                    cell.value = nv
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (1, 2, 3, 4, 5, 7, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if c == 6 and isinstance(val, date):
                cell.number_format = DATE_FMT
                cell.alignment = Alignment(horizontal="center", vertical="center")
        stt += 1
        rr += 1
    write_total_row(ws4, rr, len(h_bk_b), money_bk)
    finish_sheet(ws4, len(h_bk_b))
    n4 = stt - 1

    def src_label(s):
        labels = []
        if "HDDT_Mua" in s and "Bang_ke_Mua" in s:
            labels.append("Chung_Mua")
        elif "HDDT_Mua" in s:
            labels.append("HDDT_Mua")
        elif "Bang_ke_Mua" in s:
            labels.append("Bang_ke_Mua")
        if "HDDT_Ban" in s and "Bang_ke_Ban" in s:
            labels.append("Chung_Ban")
        elif "HDDT_Ban" in s:
            labels.append("HDDT_Ban")
        elif "Bang_ke_Ban" in s:
            labels.append("Bang_ke_Ban")
        return "; ".join(labels)

    h5 = [
        "STT",
        "MST_chuan",
        "Ten_doi_tac",
        "Ghi_chu_nguon",
        "Tình trạng hoạt động",
        "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro",
        "Thông tin công văn",
    ]
    ws_dt = wb1.create_sheet("Tong_hop_doi_tac")
    write_title(ws_dt, f"TỔNG HỢP ĐỐI TÁC - {TEN_NNT} (MST {MST_NNT}) - KỲ {KY}", len(h5))
    for i, h in enumerate(h5, 1):
        ws_dt.cell(3, i, value=h)
    style_header_row(ws_dt, 3, len(h5))
    rr = 4
    stt = 1
    n_chung_m = n_hddt_only_m = n_bk_only_m = 0
    for rec in sorted(partners.values(), key=lambda x: x["mst"]):
        tt_hd, cqt_ql = get_danhba_info(danhba_map, rec["mst"])
        rf, rcv = "", ""
        if "HDDT_Mua" in rec["src"] or "Bang_ke_Mua" in rec["src"]:
            rf, rcv = get_ncc_f0(rec["mst"], f0_map)
        lab = src_label(rec["src"])
        if "Chung_Mua" in lab:
            n_chung_m += 1
        elif "HDDT_Mua" in lab:
            n_hddt_only_m += 1
        elif "Bang_ke_Mua" in lab:
            n_bk_only_m += 1
        vals = [stt, rec["mst"], rec["ten"], lab, tt_hd, cqt_ql, rf, rcv]
        for c, val in enumerate(vals, 1):
            cell = ws_dt.cell(rr, c, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if c in (1, 2, 4, 7):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if rf == "F0" and c == 7:
                cell.font = FONT_RED_BOLD
                cell.fill = FILL_RED
        if rf == "F0":
            apply_risk_row(ws_dt, rr, len(h5), "F0", 7)
        stt += 1
        rr += 1
    write_total_row(ws_dt, rr, len(h5), [])
    finish_sheet(ws_dt, len(h5))
    n_dt = stt - 1

    # Sheet Tong_hop_trang_thai_rui_ro
    ws5 = wb1.create_sheet("Tong_hop_trang_thai_rui_ro")
    ws5.cell(1, 1, value="TỔNG HỢP NHÀ CUNG CẤP THEO TRẠNG THÁI HOẠT ĐỘNG VÀ CẢNH BÁO RỦI RO").font = FONT_TITLE
    ws5.cell(2, 1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ phân tích: {KY}").font = FONT_SUBTITLE

    status_summary = {}
    for ncc in all_ncc.values():
        st = ncc["tt"] or "Không rõ / Chưa tra cứu"
        d = status_summary.setdefault(
            st,
            {"count_ncc": 0, "hd_count": 0, "hd_ds": 0.0, "hd_thue": 0.0, "bk_count": 0, "bk_ds": 0.0, "bk_thue": 0.0},
        )
        d["count_ncc"] += 1
        d["hd_count"] += ncc["hddt_count"]
        d["hd_ds"] += ncc["hddt_ds"]
        d["hd_thue"] += ncc["hddt_thue"]
        d["bk_count"] += ncc["bk_count"]
        d["bk_ds"] += ncc["bk_ds"]
        d["bk_thue"] += ncc["bk_thue"]

    ws5.cell(4, 1, value="PHẦN I. TỔNG HỢP THEO TÌNH TRẠNG HOẠT ĐỘNG (ĐKKD - DANH BẠ NNT)").font = FONT_SUBTITLE
    hdr1 = [
        "STT",
        "Tình trạng hoạt động (ĐKKD)",
        "Số lượng NCC",
        "Số HĐ (HĐĐT)",
        "Doanh số chưa thuế (HĐĐT)",
        "Thuế GTGT (HĐĐT)",
        "Số HĐ (Bảng kê)",
        "Doanh số chưa thuế (Bảng kê)",
        "Thuế GTGT (Bảng kê)",
        "Chênh lệch DS (HĐĐT - BK)",
        "Chênh lệch Thuế (HĐĐT - BK)",
    ]
    for i, h in enumerate(hdr1, 1):
        ws5.cell(5, i, value=h)
    style_header_row(ws5, 5, len(hdr1))
    r_idx = 6
    stt_s = 1
    for st, d in sorted(status_summary.items(), key=lambda x: x[1]["count_ncc"], reverse=True):
        row_vals = [
            stt_s,
            st,
            d["count_ncc"],
            d["hd_count"],
            d["hd_ds"],
            d["hd_thue"],
            d["bk_count"],
            d["bk_ds"],
            d["bk_thue"],
            d["hd_ds"] - d["bk_ds"],
            d["hd_thue"] - d["bk_thue"],
        ]
        for col_i, val in enumerate(row_vals, 1):
            cell = ws5.cell(r_idx, col_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if col_i in (1, 3, 4, 7):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i in (5, 6, 8, 9, 10, 11):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        stt_s += 1
        r_idx += 1
    t1_tot = r_idx
    write_total_row(ws5, t1_tot, len(hdr1), [3, 4, 5, 6, 7, 8, 9, 10, 11], data_start=6)
    r_idx += 2

    risk_summary = {
        "F0": {
            "name": "NCC dính CVCBRR F0 (Bán hóa đơn)",
            "ncc": 0,
            "hd_cnt": 0,
            "hd_ds": 0.0,
            "hd_thue": 0.0,
            "bk_cnt": 0,
            "bk_ds": 0.0,
            "bk_thue": 0.0,
            "ghi_chu": "Rủi ro cao (HD01), xác minh hàng hóa và thanh toán",
        },
        "Bình thường": {
            "name": "NCC bình thường (Không thuộc F0)",
            "ncc": 0,
            "hd_cnt": 0,
            "hd_ds": 0.0,
            "hd_thue": 0.0,
            "bk_cnt": 0,
            "bk_ds": 0.0,
            "bk_thue": 0.0,
            "ghi_chu": "Đối chiếu theo kế hoạch thông thường",
        },
    }
    for ncc in all_ncc.values():
        cat = "F0" if ncc["risk_flag"] == "F0" else "Bình thường"
        risk_summary[cat]["ncc"] += 1
        risk_summary[cat]["hd_cnt"] += ncc["hddt_count"]
        risk_summary[cat]["hd_ds"] += ncc["hddt_ds"]
        risk_summary[cat]["hd_thue"] += ncc["hddt_thue"]
        risk_summary[cat]["bk_cnt"] += ncc["bk_count"]
        risk_summary[cat]["bk_ds"] += ncc["bk_ds"]
        risk_summary[cat]["bk_thue"] += ncc["bk_thue"]

    ws5.cell(r_idx, 1, value="PHẦN II. TỔNG HỢP THEO CẢNH BÁO RỦI RO (NCC chỉ đối soát F0)").font = FONT_SUBTITLE
    r_idx += 1
    hdr2 = [
        "STT",
        "Phân loại rủi ro",
        "Số lượng NCC",
        "Số HĐ (HĐĐT)",
        "Doanh số chưa thuế (HĐĐT)",
        "Thuế GTGT (HĐĐT)",
        "Số HĐ (Bảng kê)",
        "Doanh số chưa thuế (Bảng kê)",
        "Thuế GTGT (Bảng kê)",
        "Đánh giá & Định hướng xử lý",
    ]
    for i, h in enumerate(hdr2, 1):
        ws5.cell(r_idx, i, value=h)
    style_header_row(ws5, r_idx, len(hdr2))
    r_idx += 1
    t2_start = r_idx
    stt_r = 1
    for k in ["F0", "Bình thường"]:
        d = risk_summary[k]
        row_vals = [stt_r, d["name"], d["ncc"], d["hd_cnt"], d["hd_ds"], d["hd_thue"], d["bk_cnt"], d["bk_ds"], d["bk_thue"], d["ghi_chu"]]
        for col_i, val in enumerate(row_vals, 1):
            cell = ws5.cell(r_idx, col_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if k == "F0":
                cell.fill = FILL_RED
                if col_i in (1, 2):
                    cell.font = FONT_RED_BOLD
            if col_i in (1, 3, 4, 7):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i in (5, 6, 8, 9):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        stt_r += 1
        r_idx += 1
    t2_tot = r_idx
    write_total_row(ws5, t2_tot, len(hdr2), [3, 4, 5, 6, 7, 8, 9], data_start=t2_start)
    r_idx += 2

    ws5.cell(r_idx, 1, value="PHẦN III. DANH SÁCH CHI TIẾT NCC F0 HOẶC TRẠNG THÁI ĐKKD BẤT THƯỜNG").font = FONT_SUBTITLE
    r_idx += 1
    hdr3 = [
        "STT",
        "MST_chuan",
        "Tên người bán",
        "Tình trạng hoạt động (ĐKKD)",
        "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro",
        "Thông tin công văn CA / CVCBRR",
        "Nguồn dữ liệu",
        "Số HĐ (HĐĐT)",
        "Doanh số (HĐĐT)",
        "Thuế (HĐĐT)",
        "Số HĐ (Bảng kê)",
        "Doanh số (Bảng kê)",
        "Thuế (Bảng kê)",
        "Đề xuất kiểm tra & xác minh",
    ]
    for i, h in enumerate(hdr3, 1):
        ws5.cell(r_idx, i, value=h)
    style_header_row(ws5, r_idx, len(hdr3))
    r_idx += 1
    t3_start = r_idx
    stt_ab = 1
    n_abn = 0
    for mst_c, ncc in sorted(all_ncc.items(), key=lambda x: (0 if x[1]["risk_flag"] == "F0" else 1, -x[1]["hddt_ds"])):
        rf = ncc["risk_flag"]
        tt_s = ncc["tt"] or ""
        is_abn = abnormal_status(tt_s)
        if rf != "F0" and not is_abn:
            continue
        n_abn += 1
        dexuat = (
            "Rà soát tính có thật của hàng hóa, hồ sơ thanh toán, loại trừ chi phí nếu HĐ bất hợp pháp"
            if rf == "F0"
            else "Kiểm tra thời điểm lập hóa đơn so với ngày cơ quan thuế thông báo ngừng hoạt động"
        )
        row_vals = [
            stt_ab,
            mst_c,
            ncc["ten"],
            tt_s,
            ncc["cqt"],
            rf if rf else "ĐKKD bất thường",
            ncc["risk_cv"],
            src_ncc(ncc),
            ncc["hddt_count"],
            ncc["hddt_ds"],
            ncc["hddt_thue"],
            ncc["bk_count"],
            ncc["bk_ds"],
            ncc["bk_thue"],
            dexuat,
        ]
        for col_i, val in enumerate(row_vals, 1):
            cell = ws5.cell(r_idx, col_i, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if rf == "F0":
                cell.fill = FILL_RED
                if col_i in (2, 3, 6):
                    cell.font = FONT_RED_BOLD
            elif is_abn and col_i == 4:
                cell.font = FONT_RED_BOLD
            if col_i in (1, 2, 6, 8, 9, 12):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i in (10, 11, 13, 14):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        stt_ab += 1
        r_idx += 1
    t3_tot = r_idx
    write_total_row(ws5, t3_tot, len(hdr3), [9, 10, 11, 12, 13, 14], data_start=t3_start)
    autowidth(ws5)

    wb1.remove(default1)
    OUT_CT_G.mkdir(parents=True, exist_ok=True)
    OUT_CT_V.mkdir(parents=True, exist_ok=True)
    p1g = OUT_CT_G / OUT_FILE1
    p1v = OUT_CT_V / OUT_FILE1
    wb1.save(p1v)
    print(f"Saved File 1 vault {p1v} size {p1v.stat().st_size}")
    try:
        wb1.save(p1g)
        print(f"Saved File 1 ho so {p1g} size {p1g.stat().st_size}")
    except Exception as e:
        print(f"WARN save G: {e}; retry long path")
        wb1.save(lp(p1g))
        print(f"Saved File 1 ho so via lp")

    # ----- File 2 -----
    print("Processing File 2 NCC...")
    wb2 = Workbook()
    default2 = wb2.active
    nnt_tt, nnt_cqt = get_danhba_info(danhba_map, MST_NNT)
    nnt_f0, nnt_f1s, nnt_ca = nnt_lookup(MST_NNT, f0_map, f1_by_buyer, ca_map)

    ws_p1 = wb2.create_sheet("Phan_1_F1_NNT")
    ws_p1.cell(1, 1, value="PHẦN 1. RÀ SOÁT RỦI RO TỰ THÂN CỦA NGƯỜI NỘP THUẾ").font = FONT_TITLE
    ws_p1.cell(2, 1, value=f"Doanh nghiệp kiểm tra: {TEN_NNT} - MST: {MST_NNT} - Kỳ phân tích: {KY}").font = FONT_SUBTITLE
    hdr_p1 = ["STT", "Nội dung kiểm tra rủi ro tự thân", "Cơ sở dữ liệu đối soát", "Kết quả tra cứu", "Đánh giá mức rủi ro", "Ghi chú & Căn cứ"]
    for i, h in enumerate(hdr_p1, 1):
        ws_p1.cell(4, i, value=h)
    style_header_row(ws_p1, 4, len(hdr_p1))

    def nnt_row_f0():
        if nnt_f0:
            return [
                1,
                "Kiểm tra rủi ro NNT là F0 (Doanh nghiệp bị cảnh báo bán hóa đơn)",
                "Sổ CVCBRR CQT - Sheet F0",
                "CÓ TRONG DANH SÁCH F0",
                "Cao",
                fmt_cv(nnt_f0, "F0:"),
            ]
        return [
            1,
            "Kiểm tra rủi ro NNT là F0 (Doanh nghiệp bị cảnh báo bán hóa đơn)",
            "Sổ CVCBRR CQT - Sheet F0",
            "KHÔNG CÓ TRONG DANH SÁCH",
            "An toàn (Không kích hoạt rủi ro)",
            "NNT không có công văn cảnh báo bán hóa đơn bất hợp pháp",
        ]

    def nnt_row_f1():
        if nnt_f1s:
            txt = "; ".join(fmt_cv(x) for x in nnt_f1s[:5])
            return [
                2,
                "Kiểm tra rủi ro NNT là F1 (Doanh nghiệp mua HĐ từ doanh nghiệp F0)",
                "Sổ CVCBRR CQT - Sheet F1",
                f"CÓ TRONG DANH SÁCH F1 ({len(nnt_f1s)} dòng)",
                "Cao (HD01)",
                txt,
            ]
        return [
            2,
            "Kiểm tra rủi ro NNT là F1 (Doanh nghiệp mua HĐ từ doanh nghiệp F0)",
            "Sổ CVCBRR CQT - Sheet F1",
            "KHÔNG CÓ TRONG DANH SÁCH",
            "An toàn (Không kích hoạt rủi ro HD01)",
            "NNT không nằm trong danh mục người mua hóa đơn của các doanh nghiệp F0",
        ]

    def nnt_row_ca():
        if nnt_ca:
            return [
                3,
                "Kiểm tra NNT trong danh sách phối hợp xác minh của Cơ quan Công an",
                "Sổ theo dõi hồ sơ Công an (Bao_cao_cong_van.xlsx)",
                "CÓ TRONG DANH SÁCH CÔNG AN",
                "Cao",
                fmt_cv(nnt_ca, "CA:"),
            ]
        return [
            3,
            "Kiểm tra NNT trong danh sách phối hợp xác minh của Cơ quan Công an",
            "Sổ theo dõi hồ sơ Công an (Bao_cao_cong_van.xlsx)",
            "KHÔNG CÓ TRONG DANH SÁCH",
            "An toàn (Bình thường)",
            "Cơ quan Công an không có công văn yêu cầu cung cấp hồ sơ/điều tra đối với NNT",
        ]

    tt_txt = nnt_tt or "Chưa tra cứu được trên Danh bạ NNT"
    p1_rows = [
        nnt_row_f0(),
        nnt_row_f1(),
        nnt_row_ca(),
        [
            4,
            "Kiểm tra trạng thái hoạt động ĐKKD của NNT",
            "Danh bạ NNT - Hệ thống TMS",
            tt_txt,
            "Bình thường" if (nnt_tt and "đang hoạt động" in nnt_tt.lower()) else ("Cần xác minh" if nnt_tt else "Chưa tra cứu"),
            f"CQT quản lý: {nnt_cqt}" if nnt_cqt else "Chưa có CQT trên danh bạ",
        ],
    ]
    for r_i, r_data in enumerate(p1_rows, start=5):
        for c_i, val in enumerate(r_data, 1):
            cell = ws_p1.cell(r_i, c_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if c_i in (1, 4, 5):
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if c_i in (4, 5):
                    cell.font = FONT_BOLD
                    if "CÓ TRONG" in str(r_data[3]):
                        cell.fill = FILL_RED
                        cell.font = FONT_RED_BOLD
    autowidth(ws_p1)
    ws_p1.column_dimensions["B"].width = 45
    ws_p1.column_dimensions["F"].width = 50

    ws_21 = wb2.create_sheet("Bang_2_1_NCC_F0")
    ws_21.cell(1, 1, value="BẢNG 2.1. DANH SÁCH NHÀ CUNG CẤP THUỘC DIỆN CVCBRR F0").font = FONT_TITLE
    ws_21.cell(2, 1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}. NCC mua vào chỉ đối soát F0.").font = FONT_SUBTITLE
    hdr_21 = [
        "STT",
        "MST_chuan",
        "Tên người bán / Nhà cung cấp",
        "Nhóm rủi ro",
        "Thông tin công văn cảnh báo",
        "Tình trạng hoạt động (ĐKKD)",
        "Cơ quan thuế quản lý",
        "Nguồn dữ liệu",
        "Số HĐ (BK)",
        "Doanh số (BK)",
        "Thuế (BK)",
        "Số HĐ (HĐĐT)",
        "Doanh số (HĐĐT)",
        "Thuế (HĐĐT)",
        "Mức độ rủi ro",
        "Đề xuất xử lý khi thanh tra",
    ]
    for i, h in enumerate(hdr_21, 1):
        ws_21.cell(4, i, value=h)
    style_header_row(ws_21, 4, len(hdr_21))
    r21 = 5
    stt_21 = 1
    f0_nccs = [n for n in all_ncc.values() if n["risk_flag"] == "F0"]
    for ncc in sorted(f0_nccs, key=lambda x: -x["hddt_ds"]):
        r_vals = [
            stt_21,
            ncc["mst"],
            ncc["ten"],
            "F0",
            ncc["risk_cv"],
            ncc["tt"],
            ncc["cqt"],
            src_ncc(ncc),
            ncc["bk_count"],
            ncc["bk_ds"],
            ncc["bk_thue"],
            ncc["hddt_count"],
            ncc["hddt_ds"],
            ncc["hddt_thue"],
            "Cao",
            "Rà soát hóa đơn mua vào, tính có thật của hàng hóa, loại trừ chi phí nếu sử dụng HĐ bất hợp pháp",
        ]
        for c_i, val in enumerate(r_vals, 1):
            cell = ws_21.cell(r21, c_i, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            cell.fill = FILL_RED
            if c_i in (2, 3, 4, 15):
                cell.font = FONT_RED_BOLD
            if c_i in (1, 2, 4, 8, 9, 12, 15):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i in (10, 11, 13, 14):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        stt_21 += 1
        r21 += 1
    write_total_row(ws_21, r21, len(hdr_21), [9, 10, 11, 12, 13, 14], data_start=5)
    autowidth(ws_21)

    ws_22 = wb2.create_sheet("Bang_2_2_Toan_bo_NCC")
    ws_22.cell(1, 1, value="BẢNG 2.2. TỔNG HỢP TOÀN BỘ 100% NHÀ CUNG CẤP MUA VÀO (UNION BẢNG KÊ & HĐĐT)").font = FONT_TITLE
    ws_22.cell(2, 1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}").font = FONT_SUBTITLE
    hdr_22 = [
        "STT",
        "MST_chuan",
        "Tên người bán / Nhà cung cấp",
        "Nguồn dữ liệu",
        "Số HĐ (BK)",
        "Doanh số (BK)",
        "Thuế (BK)",
        "Số HĐ (HĐĐT)",
        "Doanh số (HĐĐT)",
        "Thuế (HĐĐT)",
        "Chênh lệch DS (HĐĐT - BK)",
        "Chênh lệch Thuế (HĐĐT - BK)",
        "Tình trạng hoạt động (ĐKKD)",
        "Cơ quan thuế quản lý",
        "Nhóm rủi ro (F0)",
        "Thông tin công văn",
    ]
    for i, h in enumerate(hdr_22, 1):
        ws_22.cell(4, i, value=h)
    style_header_row(ws_22, 4, len(hdr_22))
    r22 = 5
    stt_22 = 1
    for ncc in sorted(all_ncc.values(), key=lambda x: (0 if x["risk_flag"] == "F0" else 1, -x["hddt_ds"])):
        rf = ncc["risk_flag"]
        is_abn = abnormal_status(ncc["tt"])
        r_vals = [
            stt_22,
            ncc["mst"],
            ncc["ten"],
            src_ncc(ncc),
            ncc["bk_count"],
            ncc["bk_ds"],
            ncc["bk_thue"],
            ncc["hddt_count"],
            ncc["hddt_ds"],
            ncc["hddt_thue"],
            ncc["hddt_ds"] - ncc["bk_ds"],
            ncc["hddt_thue"] - ncc["bk_thue"],
            ncc["tt"],
            ncc["cqt"],
            rf,
            ncc["risk_cv"],
        ]
        for c_i, val in enumerate(r_vals, 1):
            cell = ws_22.cell(r22, c_i, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if rf == "F0":
                cell.fill = FILL_RED
                if c_i in (2, 3, 15):
                    cell.font = FONT_RED_BOLD
            elif is_abn and c_i == 13:
                cell.font = FONT_RED_BOLD
            if c_i in (1, 2, 4, 5, 8, 15):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i in (6, 7, 9, 10, 11, 12):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        stt_22 += 1
        r22 += 1
    write_total_row(ws_22, r22, len(hdr_22), [5, 6, 7, 8, 9, 10, 11, 12], data_start=5)
    autowidth(ws_22)

    ws_23 = wb2.create_sheet("Bang_2_3_Ky_ke_khai_CT")
    ws_23.cell(1, 1, value="BẢNG 2.3. CHI TIẾT HÓA ĐƠN THEO KỲ KÊ KHAI CỦA CÁC NHÀ CUNG CẤP RỦI RO / ĐKKD BẤT THƯỜNG").font = FONT_TITLE
    ws_23.cell(2, 1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}").font = FONT_SUBTITLE
    hdr_23 = [
        "STT",
        "Nguồn",
        "Kỳ/Năm",
        "Số HĐ",
        "Ngày HĐ",
        "MST_chuan",
        "Tên người bán",
        "Mặt hàng",
        "Doanh số chưa thuế",
        "Thuế GTGT",
        "Cảnh báo rủi ro",
        "Thông tin công văn",
        "Tình trạng hoạt động (ĐKKD)",
        "Cơ quan thuế quản lý",
    ]
    for i, h in enumerate(hdr_23, 1):
        ws_23.cell(4, i, value=h)
    style_header_row(ws_23, 4, len(hdr_23))
    risky_msts = {
        m for m, n in all_ncc.items() if n["risk_flag"] == "F0" or abnormal_status(n["tt"])
    }
    r23 = 5
    stt_23 = 1
    for inv in sorted([x for x in all_invoices if x["mst"] in risky_msts], key=lambda x: (x["ngay_hd"], x["mst"])):
        ky = ""
        if inv["ngay_hd"] and len(inv["ngay_hd"]) >= 10:
            ky = inv["ngay_hd"][6:10] + "-" + inv["ngay_hd"][3:5]
        r_vals = [
            stt_23,
            inv["nguon"],
            ky,
            inv["so_hd"],
            inv["ngay_hd"],
            inv["mst"],
            inv["ten"],
            inv["mat_hang"],
            inv["ds"],
            inv["thue"],
            inv["risk_flag"] if inv["risk_flag"] else "ĐKKD bất thường",
            inv["risk_cv"],
            inv["tt"],
            inv["cqt"],
        ]
        for c_i, val in enumerate(r_vals, 1):
            cell = ws_23.cell(r23, c_i, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if inv["risk_flag"] == "F0":
                cell.fill = FILL_RED
                if c_i in (4, 6, 7, 11):
                    cell.font = FONT_RED_BOLD
            if c_i in (1, 2, 3, 4, 5, 6, 11):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i in (9, 10):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        stt_23 += 1
        r23 += 1
    write_total_row(ws_23, r23, len(hdr_23), [9, 10], data_start=5)
    autowidth(ws_23)
    n_inv_risk = stt_23 - 1

    ws_p3 = wb2.create_sheet("Phan_3_Nhan_xet")
    ws_p3.cell(1, 1, value="PHẦN 3. ĐÁNH GIÁ RỦI RO VÀ ĐỀ XUẤT TRỌNG TÂM THANH TRA KIỂM TRA").font = FONT_TITLE
    ws_p3.cell(2, 1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}").font = FONT_SUBTITLE
    hdr_p3 = ["STT", "Chuyên đề / Nhóm nội dung", "Kết quả rà soát thực tế tại hồ sơ", "Đề cương trọng tâm kiểm tra xác minh tại NNT"]
    for i, h in enumerate(hdr_p3, 1):
        ws_p3.cell(4, i, value=h)
    style_header_row(ws_p3, 4, len(hdr_p3))

    if nnt_f0 or nnt_f1s or nnt_ca:
        bits = []
        if nnt_f0:
            bits.append("F0 " + fmt_cv(nnt_f0))
        if nnt_f1s:
            bits.append(f"F1 {len(nnt_f1s)} dòng")
        if nnt_ca:
            bits.append("Công an " + fmt_cv(nnt_ca))
        nnt_kq = "NNT dính rủi ro tự thân: " + "; ".join(bits)
        nnt_dx = "Ưu tiên xác minh hồ sơ tự thân NNT trước khi mở rộng chuỗi NCC."
    else:
        nnt_kq = (
            f"NNT không nằm trong danh sách F0 (CVCBRR), không nằm trong danh sách F1 (mua từ F0), "
            f"và không có trong Sổ theo dõi phối hợp Công an. Trạng thái ĐKKD: {tt_txt}."
        )
        nnt_dx = "Kết luận: NNT tự thân an toàn. Hồ sơ không thuộc diện trọng điểm điều tra mua bán hóa đơn tự thân."

    if f0_nccs:
        parts = []
        for n in sorted(f0_nccs, key=lambda x: -max(x["hddt_ds"], x["bk_ds"])):
            parts.append(
                f"{n['ten']} (MST {n['mst']}, {n['risk_cv']}, HĐĐT {n['hddt_count']} HĐ DS {fmt_money_vn(n['hddt_ds'])} đ, "
                f"BK {n['bk_count']} HĐ DS {fmt_money_vn(n['bk_ds'])} đ)"
            )
        f0_kq = f"Phát hiện {len(f0_nccs)} NCC thuộc F0: " + "; ".join(parts) + "."
        f0_dx = (
            "Trọng tâm 1: Yêu cầu NNT giải trình và cung cấp hợp đồng, hóa đơn, biên bản giao nhận, "
            "chứng từ vận chuyển, UNC/sao kê cho 100% hóa đơn phát sinh với các NCC F0."
        )
    else:
        f0_kq = "Không phát hiện NCC mua vào thuộc danh sách F0 CVCBRR trên union Bảng kê ∪ HĐĐT."
        f0_dx = "Không mở chuyên đề F0. Tập trung đối chiếu chênh lệch nguồn và ĐKKD bất thường."

    n_ngung = sum(1 for n in all_ncc.values() if abnormal_status(n["tt"]))
    if n_ngung:
        dkkd_kq = f"Phát hiện {n_ngung} NCC có trạng thái ĐKKD khác 'đang hoạt động' (kể cả ngừng/đóng MST)."
        dkkd_dx = (
            "Trọng tâm 2: Đối chiếu ngày lập từng hóa đơn mua vào với ngày cơ quan thuế thông báo ngừng hoạt động/đóng MST. "
            "Loại trừ thuế GTGT đầu vào và chi phí đối với hóa đơn lập sau thời điểm ngừng."
        )
    else:
        dkkd_kq = "Không phát hiện NCC có trạng thái ĐKKD bất thường trên danh bạ đã tra cứu (các MST tra cứu được đang hoạt động hoặc chưa có trên danh bạ)."
        dkkd_dx = "Giữ đối chiếu thường: xác minh các MST chưa tra cứu được trên danh bạ."

    lech_ncc = sum(1 for n in all_ncc.values() if abs((n["hddt_ds"] or 0) - (n["bk_ds"] or 0)) > 1)
    p3_content = [
        [1, "Rủi ro tự thân của NNT (F0 / F1 / Công an)", nnt_kq, nnt_dx],
        [2, "Nhà cung cấp dính CVCBRR F0", f0_kq, f0_dx],
        [3, "Nhà cung cấp ngừng hoạt động / đóng mã số thuế", dkkd_kq, dkkd_dx],
        [
            4,
            "Chênh lệch nguồn HĐĐT và Bảng kê",
            (
                f"Union {len(all_ncc)} NCC. HĐĐT mua {n1} dòng, Bảng kê mua {n3} dòng. "
                f"Đối tác mua chỉ HĐĐT {n_hddt_only_m}, chỉ Bảng kê {n_bk_only_m}, chung hai nguồn {n_chung_m}. "
                f"{lech_ncc} NCC có chênh lệch doanh số HĐĐT vs Bảng kê."
            ),
            "Trọng tâm 3: Chọn mẫu NCC chênh lệch lớn, đối chiếu từng số HĐ, kỳ kê khai và chứng từ thanh toán.",
        ],
    ]
    for r_i, r_data in enumerate(p3_content, start=5):
        for c_i, val in enumerate(r_data, 1):
            cell = ws_p3.cell(r_i, c_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c_i == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
            elif c_i == 2:
                cell.font = FONT_BOLD
        ws_p3.row_dimensions[r_i].height = 80
    autowidth(ws_p3)
    ws_p3.column_dimensions["A"].width = 8
    ws_p3.column_dimensions["B"].width = 36
    ws_p3.column_dimensions["C"].width = 55
    ws_p3.column_dimensions["D"].width = 55

    wb2.remove(default2)
    p2g = OUT_CT_G / OUT_FILE2
    p2v = OUT_CT_V / OUT_FILE2
    wb2.save(p2v)
    print(f"Saved File 2 vault {p2v} size {p2v.stat().st_size}")
    try:
        wb2.save(p2g)
        print(f"Saved File 2 ho so {p2g} size {p2g.stat().st_size}")
    except Exception as e:
        print(f"WARN save G file2: {e}")
        wb2.save(lp(p2g))

    summary = {
        "mst": MST_NNT,
        "ten": TEN_NNT,
        "ky": KY,
        "hddt_mua": n1,
        "hddt_ban": n2,
        "hddt_mua_skip_huy": skip_huy.get("MUA", 0),
        "hddt_ban_skip_huy": skip_huy.get("BAN", 0),
        "hddt_mua_years": dict(years_m),
        "hddt_ban_years": dict(years_b),
        "hddt_mua_2021": n_hddt_m_2021,
        "bk_mua": n3,
        "bk_ban": n4,
        "doi_tac": n_dt,
        "ncc_union": len(all_ncc),
        "ncc_f0": len(f0_nccs),
        "ncc_dkkd_bat_thuong": n_ngung,
        "inv_risk": n_inv_risk,
        "partner_chung_mua": n_chung_m,
        "partner_hddt_only_mua": n_hddt_only_m,
        "partner_bk_only_mua": n_bk_only_m,
        "nnt_f0": bool(nnt_f0),
        "nnt_f1": len(nnt_f1s),
        "nnt_ca": bool(nnt_ca),
        "nnt_tt": nnt_tt,
        "file1": str(p1g),
        "file2": str(p2g),
        "f0_list": [
            {
                "mst": n["mst"],
                "ten": n["ten"],
                "cv": n["risk_cv"],
                "hddt_count": n["hddt_count"],
                "hddt_ds": n["hddt_ds"],
                "bk_count": n["bk_count"],
                "bk_ds": n["bk_ds"],
            }
            for n in sorted(f0_nccs, key=lambda x: -x["hddt_ds"])
        ],
    }
    sp = VAULT / "scratch" / "dhh_4sheet_summary.json"
    sp.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print("SUMMARY", json.dumps(summary, ensure_ascii=False, indent=2))
    print("ALL JOBS COMPLETED")
    return summary


if __name__ == "__main__":
    run()
