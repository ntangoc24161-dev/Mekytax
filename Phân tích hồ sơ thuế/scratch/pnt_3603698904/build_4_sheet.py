# -*- coding: utf-8 -*-
"""tong-hop-hddt-bang-ke-4-sheet for PHUC NGUYEN TIEN 3603698904, ky 2020-2025."""
from __future__ import annotations

import os
import re
import shutil
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603698904"
TEN = "CÔNG TY TNHH PHÚC NGUYÊN TIẾN"
KY = "2020-2025"
TODAY = date(2026, 11, 25)

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Dang kiem tra\KE HOACH\PHUC_NGUYEN_TIEN_3603698904-chua so"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
OUT = ROOT / "4. KET QUA PHAN TICH"
CAU = OUT / "cau-thanh"
SAO = ROOT / "7. Sao luu tong hop"
VAULT_CT = VAULT / "ket-qua" / "cau-thanh"
HDDT_PATH = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT.xlsx"
B25 = ROOT / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
DANHBA = ROOT / "3. DU LIEU BO SUNG & TRA CUU QLT" / "3.2. Danh ba trang thai NNT" / "Danh sach NCC.xlsx"
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FONT_W = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
FONT_TITLE = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")
FONT_RED = Font(name="Times New Roman", size=11, bold=True, color="C00000")
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
FILL_NAVY = PatternFill("solid", fgColor="1F4E78")
FILL_TOT = PatternFill("solid", fgColor="DCE6F1")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
DATE_FMT = "dd/mm/yyyy"
NUM = "#,##0"


def lp(p) -> str:
    s = str(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s


def chuan_hoa_mst(val) -> str:
    if val is None:
        return ""
    s = str(val).strip().replace(" ", "")
    if s.endswith(".0") and s[:-2].isdigit():
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
    return re.sub(r"\D", "", str(s or "").split("-")[0])


def sanitize(val):
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip()
        return "'" + s if s.startswith("=") else s
    return val


def as_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            return datetime(*xlrd.xldate_as_tuple(float(v), 0)).date()
        except Exception:
            return None
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def nnum(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("%", "")
    if s.count(".") > 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def fold(s) -> str:
    return str(s or "").strip().casefold()


def is_label_row(so, ten) -> bool:
    blob = " ".join([str(so or ""), str(ten or "")]).casefold().strip()
    if not blob:
        return True
    if "công ty" in blob or "cong ty" in blob or "tổng công ty" in blob:
        return False
    if blob in {"tổng", "tong", "tổng cộng", "tong cong"}:
        return True
    if blob.startswith("tổng cộng") or blob.startswith("tong cong"):
        return True
    if "nhóm hhdv" in blob or blob.startswith("nhóm"):
        return True
    if "hàng hóa, dịch vụ" in blob or "hàng hoá, dịch vụ" in blob:
        return True
    return False


def quy_of(d: date | None, fallback=None):
    if d:
        return (d.month - 1) // 3 + 1
    return fallback


def put(ws, r, c, val, fmt=None, bold=False, fill=None, wrap=False):
    cell = ws.cell(r, c)
    if isinstance(cell, MergedCell):
        return
    cell.value = sanitize(val)
    cell.font = FONT_B if bold else FONT
    cell.border = THIN
    if fill:
        cell.fill = fill
    if wrap:
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    if fmt:
        cell.number_format = fmt
    elif isinstance(val, datetime):
        cell.value = val.date()
        cell.number_format = DATE_FMT
    elif isinstance(val, date):
        cell.number_format = DATE_FMT
    elif isinstance(val, (int, float)):
        cell.number_format = NUM


def autowidth(ws, cap=55, min_w=10, max_row=250):
    widths = {}
    mr = min(ws.max_row or 1, max_row)
    mc = ws.max_column or 1
    for row in ws.iter_rows(max_row=mr, max_col=mc):
        for cell in row:
            if cell.value is None or isinstance(cell, MergedCell):
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            widths[cell.column] = max(widths.get(cell.column, min_w), min(len(s) + 2, cap))
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


def style_header(ws, r, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(r, c)
        cell.font = FONT_W
        cell.fill = FILL_NAVY
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        cell.border = THIN


def write_sheet_rows(ws, title, headers, rows, money_cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    put(ws, 1, 1, title, bold=True)
    ws["A1"].font = FONT_TITLE
    for i, h in enumerate(headers, 1):
        put(ws, 3, i, h, bold=True)
    style_header(ws, 3, len(headers))
    r = 4
    for rec in rows:
        rec[0] = r - 3
        for c, v in enumerate(rec, 1):
            fmt = NUM if c in money_cols else None
            put(ws, r, c, v, fmt=fmt)
        r += 1
    if rows:
        put(ws, r, 1, "TỔNG CỘNG", bold=True, fill=FILL_TOT)
        for c in money_cols:
            col = get_column_letter(c)
            cell = ws.cell(r, c, value=f"=SUM({col}4:{col}{r-1})")
            cell.font = FONT_B
            cell.fill = FILL_TOT
            cell.number_format = NUM
            cell.border = THIN
        for c in range(1, len(headers) + 1):
            ws.cell(r, c).fill = FILL_TOT
            ws.cell(r, c).border = THIN
            ws.cell(r, c).font = FONT_B
    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_title_rows = "1:3"
    autowidth(ws)
    return r


def rows_xlsx(path: Path):
    wb = load_workbook(lp(path), read_only=True, data_only=True)
    out = []
    for name in wb.sheetnames:
        ws = wb[name]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        out.append((name, rows))
    wb.close()
    return out


def rows_xls(path: Path):
    book = xlrd.open_workbook(lp(path))
    out = []
    for name in book.sheet_names():
        sh = book.sheet_by_name(name)
        rows = [[sh.cell_value(i, j) for j in range(sh.ncols)] for i in range(sh.nrows)]
        out.append((name, rows))
    return out


def sheet_chieu(sname: str, first_rows: list, fname: str) -> str | None:
    blob = " ".join(
        [sname, fname]
        + [" ".join(str(c or "") for c in (first_rows[i] if i < len(first_rows) else [])) for i in range(min(4, len(first_rows)))]
    ).casefold()
    if any(k in sname.casefold() for k in ["đầu ra", "dau ra", "bán ra", "ban ra", "bkbr"]):
        return "BAN"
    if any(k in sname.casefold() for k in ["đầu vào", "dau vao", "mua vào", "mua vao", "bkmv"]):
        return "MUA"
    if "bán ra" in blob or "ban ra" in blob or "dịch vụ bán" in blob:
        return "BAN"
    if "mua vào" in blob or "mua vao" in blob or "đầu vào" in blob:
        return "MUA"
    fn = fname.casefold()
    if "ban ra" in fn or "bán ra" in fn:
        return "BAN"
    if "mua vao" in fn or "mua vào" in fn:
        return "MUA"
    return None


def find_header(rows, chieu: str):
    for i, row in enumerate(rows[:25]):
        joined = " ".join(fold(c) for c in row)
        if "số hóa đơn" in joined or "số hoá đơn" in joined or "so hoa don" in joined:
            return i
        if "số hđ" in joined and "ngày" in joined:
            return i
    return None


def map_cols(header_rows, chieu: str):
    idx = {}
    for row in header_rows:
        for j, c in enumerate(row):
            t = fold(c).replace("\n", " ")
            if not t:
                continue
            if idx.get("so") is None and any(k in t for k in ["số hóa đơn", "số hoá đơn", "số hđ", "so hd"]):
                if "mẫu" not in t and "mau" not in t:
                    idx["so"] = j
            if idx.get("kh") is None and ("ký hiệu" in t or "kí hiệu" in t):
                if "mẫu" in t or "mau" in t:
                    idx["mau"] = j
                else:
                    idx["kh"] = j
            if idx.get("ngay") is None and ("ngày" in t or "ngay" in t) and ("lập" in t or "hóa đơn" in t or t in {"ngày", "ngay"}):
                idx["ngay"] = j
            if idx.get("mst") is None and ("mã số thuế" in t or t == "mst"):
                idx["mst"] = j
            if idx.get("ten") is None and (
                ("tên người bán" in t and chieu == "MUA")
                or ("tên người mua" in t and chieu == "BAN")
                or ("tên người" in t)
            ):
                idx["ten"] = j
            if idx.get("mh") is None and "mặt hàng" in t:
                idx["mh"] = j
            if idx.get("ds") is None and any(
                k in t for k in ["doanh số", "giá trị hhdv", "chưa có thuế", "chưa thuế"]
            ):
                idx["ds"] = j
            if idx.get("ts") is None and "thuế suất" in t:
                idx["ts"] = j
            if idx.get("thue") is None and ("thuế gtgt" in t or "tiền thuế" in t) and "suất" not in t and "mã số thuế" not in t and "chưa" not in t:
                idx["thue"] = j
    return idx


def cell(row, i):
    if i is None or i >= len(row):
        return None
    return row[i]


def extract_sheet(path: Path, sname: str, rows: list, chieu: str, nam_hint, quy_hint) -> list[dict]:
    hi = find_header(rows, chieu)
    if hi is None:
        return []
    idx = map_cols(rows[max(0, hi - 1) : hi + 3], chieu)
    if idx.get("so") is None:
        idx.setdefault("so", 1)
        idx.setdefault("ngay", 2)
        idx.setdefault("ten", 3)
        idx.setdefault("mst", 4)
        idx.setdefault("ds", 6)
        idx.setdefault("thue", 8)
    start = hi + 1
    out = []
    for row in rows[start:]:
        so = cell(row, idx.get("so"))
        ten = cell(row, idx.get("ten"))
        mst = cell(row, idx.get("mst"))
        ngay = as_date(cell(row, idx.get("ngay")))
        if is_label_row(so, ten):
            continue
        so_s = str(so or "").strip()
        if so_s.endswith(".0") and so_s[:-2].isdigit():
            so_s = so_s[:-2]
        if not so_s or so_s.lower() in {"none", "stt"}:
            continue
        if not ngay and not mst and not nnum(cell(row, idx.get("ds"))):
            continue
        ds = nnum(cell(row, idx.get("ds"))) or 0
        thue = nnum(cell(row, idx.get("thue"))) or 0
        ts = cell(row, idx.get("ts"))
        ts_n = nnum(ts)
        if ts_n is not None and ts_n > 1:
            ts_n = ts_n / 100.0
        mst_goc = str(mst or "").strip()
        nam = ngay.year if ngay else nam_hint
        out.append(
            {
                "Chieu": chieu,
                "Nam": nam,
                "Quy": quy_of(ngay, quy_hint),
                "Thang": ngay.month if ngay else None,
                "Ky_hieu_HD": str(cell(row, idx.get("kh")) or "").strip(),
                "So_HD": so_s,
                "Ngay_HD": ngay,
                "MST_goc": mst_goc,
                "MST_chuan": chuan_hoa_mst(mst_goc),
                "Ten_DoiTac": str(ten or "").strip(),
                "Mat_hang": str(cell(row, idx.get("mh")) or "").strip(),
                "GT_chua_thue": ds,
                "Thue_suat": ts_n if ts_n is not None else (str(ts).strip() if ts not in (None, "") else ""),
                "Thue_GTGT": thue,
                "Tong_tien": ds + thue,
                "File_nguon": path.name,
            }
        )
    return out


def use_bk_file(p: Path) -> bool:
    n = p.name
    if p.suffix.lower() not in {".xlsx", ".xls"}:
        return False
    if n.startswith("Bang_ke_mua_vao_ban_ra_tong_hop"):
        return False
    if n.upper().startswith("TH HD"):
        return False
    if n.startswith("2020_") or n.startswith("2021_") or n.startswith("TH_"):
        return True
    if "25" in n and ("mua" in n.casefold() or "mua" in fold(n)):
        return True
    return False


def parse_all_bang_ke() -> list[dict]:
    files = [p for p in B25.iterdir() if p.is_file() and use_bk_file(p)]
    print("BK source files", len(files))
    all_rows = []
    for p in sorted(files, key=lambda x: x.name):
        nam_hint = None
        quy_hint = None
        m = re.search(r"(20\d{2})", p.name)
        if m:
            nam_hint = int(m.group(1))
        elif "25" in p.name:
            nam_hint = 2025
        m2 = re.search(r"Q(?:UY\s*)?([1-4IV]+)", p.name, re.I)
        if m2:
            tok = m2.group(1).upper()
            quy_hint = {"1": 1, "2": 2, "3": 3, "4": 4, "I": 1, "II": 2, "III": 3, "IV": 4}.get(tok)
        try:
            sheets = rows_xls(p) if p.suffix.lower() == ".xls" else rows_xlsx(p)
        except Exception as e:
            print("  ERR read", p.name, e)
            continue
        n_file = 0
        for sname, rows in sheets:
            chieu = sheet_chieu(sname, rows, p.name)
            if not chieu:
                continue
            recs = extract_sheet(p, sname, rows, chieu, nam_hint, quy_hint)
            n_file += len(recs)
            all_rows.extend(recs)
        print(f"  {p.name} -> {n_file}")
    seen = {}
    uniq = []
    for r in all_rows:
        k = (r["Chieu"], r["So_HD"], r["MST_chuan"], str(r["Ngay_HD"]))
        if k in seen:
            continue
        seen[k] = True
        uniq.append(r)
    print(
        "BK TOTAL",
        len(uniq),
        "MUA",
        sum(1 for r in uniq if r["Chieu"] == "MUA"),
        "BAN",
        sum(1 for r in uniq if r["Chieu"] == "BAN"),
        "raw",
        len(all_rows),
    )
    return uniq


def load_hddt():
    wb = load_workbook(lp(HDDT_PATH), data_only=True)
    out = {}
    for chieu, sname, mst_at in (("MUA", "Mua vao", 5), ("BAN", "Ban ra", 7)):
        ws = wb[sname]
        rows = list(ws.iter_rows(values_only=True))
        header = list(rows[0])
        data = []
        skip = 0
        for row in rows[1:]:
            if not any(row):
                continue
            stt0 = str(row[0] or "").strip().casefold()
            if stt0.startswith("tổng") or stt0.startswith("tong"):
                skip += 1
                continue
            st = str(row[20] if len(row) > 20 else "" or "").casefold()
            kq = str(row[21] if len(row) > 21 else "" or "").casefold()
            if any(k in st for k in ["bị xóa", "bi xoa", "bị hủy", "bi huy", "bị thay thế", "hủy bỏ"]):
                skip += 1
                continue
            if "không đủ điều kiện" in kq:
                skip += 1
                continue
            data.append(list(row))
        print(f"HDDT {chieu} keep {len(data)} skip {skip}")
        out[chieu] = (header, data, mst_at)
    wb.close()
    return out


def load_danhba():
    mp = {}
    if not DANHBA.exists():
        print("Danh ba missing")
        return mp
    wb = load_workbook(lp(DANHBA), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    n = 0
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 4:
            continue
        mst = chuan_hoa_mst(row[3] if len(row) > 3 else "")
        if not mst:
            continue
        rec = {
            "tt": str(row[44] or "").strip() if len(row) > 44 else "",
            "cqt": str(row[1] or "").strip() if len(row) > 1 else "",
            "ten": str(row[4] or "").strip() if len(row) > 4 else "",
        }
        mp[mst] = rec
        mp[mst_clean(mst)] = rec
        n += 1
    wb.close()
    print("Danh ba", n)
    return mp


def db_info(mp, mst):
    m = chuan_hoa_mst(mst)
    hit = mp.get(m) or mp.get(mst_clean(m))
    if not hit:
        return "", ""
    return hit.get("tt") or "", hit.get("cqt") or ""


def load_cvcbrr_and_ca():
    f0_map, f1_map, ca_map = {}, {}, {}
    if CVCBRR.exists():
        wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
        if "F0" in wb.sheetnames:
            for r in wb["F0"].iter_rows(min_row=8, values_only=True):
                m = mst_clean(r[1] if r and len(r) > 1 else "")
                if m:
                    f0_map[m] = {
                        "cv": str(r[3] or "").strip() if len(r) > 3 else "",
                        "ngay": str(r[4])[:10] if len(r) > 4 and r[4] else "",
                        "co_quan": str(r[5] or "").strip() if len(r) > 5 else "",
                        "ten": str(r[2] or "").strip() if len(r) > 2 else "",
                    }
        if "F1" in wb.sheetnames:
            for r in wb["F1"].iter_rows(min_row=8, values_only=True):
                m_ban = mst_clean(r[1] if r and len(r) > 1 else "")
                m_mua = mst_clean(r[8] if r and len(r) > 8 else "")
                rec = {
                    "cv": str(r[3] or "").strip() if len(r) > 3 else "",
                    "ngay": str(r[4])[:10] if len(r) > 4 and r[4] else "",
                    "co_quan": str(r[5] or "").strip() if len(r) > 5 else "",
                    "m_mua": m_mua,
                }
                if m_ban:
                    f1_map[m_ban] = rec
                if m_mua:
                    f1_map.setdefault(m_mua, rec)
        wb.close()
    print("F0", len(f0_map), "F1", len(f1_map))
    if CA.exists():
        wb = load_workbook(lp(CA), read_only=True, data_only=True)
        if "NHAP_HO_SO" in wb.sheetnames:
            last_cv = last_ngay = last_cq = ""
            for r in list(wb["NHAP_HO_SO"].iter_rows(values_only=True))[4:]:
                if len(r) > 2 and r[2]:
                    last_cv = str(r[2]).strip()
                    last_ngay = str(r[3])[:10] if len(r) > 3 and r[3] else ""
                    last_cq = str(r[4] or "").strip() if len(r) > 4 else ""
                mst = mst_clean(r[9]) if len(r) > 9 and r[9] else ""
                if mst:
                    ca_map[mst] = {"cv": last_cv, "ngay": last_ngay, "co_quan": last_cq}
        wb.close()
    print("CA", len(ca_map))
    return f0_map, f1_map, ca_map


def cv_text(info) -> str:
    if not info:
        return ""
    parts = []
    if info.get("cv"):
        parts.append(f"CV {info['cv']}")
    if info.get("ngay"):
        parts.append(f"ngày {info['ngay']}")
    if info.get("co_quan"):
        parts.append(f"({info['co_quan']})")
    return " ".join(parts)


def f0_flag(mst, f0_map) -> str:
    info = f0_map.get(mst_clean(mst))
    return cv_text(info) if info else ""


def emit_hddt(ws, title, header, data, mst_at, money_idx, db, f0_map, is_mua: bool):
    extra = ["MST_chuan", "Tình trạng hoạt động", "Cơ quan thuế quản lý"]
    if is_mua:
        extra.append("Cảnh báo rủi ro")
    new_h = list(header[: mst_at + 1]) + extra[:1] + list(header[mst_at + 1 :]) + extra[1:]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(new_h))
    put(ws, 1, 1, title, bold=True)
    ws["A1"].font = FONT_TITLE
    for i, t in enumerate(new_h, 1):
        put(ws, 3, i, t, bold=True)
    style_header(ws, 3, len(new_h))
    r = 4
    stt = 1
    n_f0 = 0
    for row in data:
        row = list(row) + [None] * max(0, 22 - len(row))
        mst_c = chuan_hoa_mst(row[mst_at] if mst_at < len(row) else "")
        tt, cqt = db_info(db, mst_c)
        flag = f0_flag(mst_c, f0_map) if is_mua else ""
        if flag:
            n_f0 += 1
        new = list(row[: mst_at + 1]) + [mst_c] + list(row[mst_at + 1 :]) + [tt, cqt]
        if is_mua:
            new.append(flag)
        new[0] = stt
        stt += 1
        for c, v in enumerate(new, 1):
            orig_i = c - 1
            if orig_i > mst_at:
                orig_i -= 1
            fmt = NUM if orig_i in money_idx else (DATE_FMT if isinstance(v, (date, datetime)) else None)
            if isinstance(v, str) and re.match(r"\d{1,2}/\d{1,2}/\d{4}$", v):
                d = as_date(v)
                if d:
                    v, fmt = d, DATE_FMT
            put(ws, r, c, v, fmt=fmt)
            if flag and c == len(new):
                ws.cell(r, c).fill = FILL_RED
                ws.cell(r, c).font = FONT_RED
        r += 1
    if r > 4:
        put(ws, r, 1, "TỔNG CỘNG", bold=True, fill=FILL_TOT)
        for orig in money_idx:
            out_c = (orig + 1) if orig <= mst_at else (orig + 2)
            col = get_column_letter(out_c)
            cell = ws.cell(r, out_c, value=f"=SUM({col}4:{col}{r-1})")
            cell.font = FONT_B
            cell.fill = FILL_TOT
            cell.number_format = NUM
            cell.border = THIN
        for c in range(1, len(new_h) + 1):
            ws.cell(r, c).fill = FILL_TOT
            ws.cell(r, c).border = THIN
    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    autowidth(ws)
    return stt - 1, n_f0


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


def main():
    print("START 4-sheet", MST, KY)
    db = load_danhba()
    f0_map, f1_map, ca_map = load_cvcbrr_and_ca()
    hddt = load_hddt()
    bk_recs = parse_all_bang_ke()

    h_mua, d_mua, mst_mua = hddt["MUA"]
    h_ban, d_ban, mst_ban = hddt["BAN"]

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "HDDT_Mua_vao"
    n1, n_f0_hd = emit_hddt(
        ws1,
        f"TỔNG HỢP HĐĐT MUA VÀO - {TEN} (MST {MST}) - {KY}",
        h_mua,
        d_mua,
        mst_mua,
        {13, 14, 17},
        db,
        f0_map,
        True,
    )
    ws2 = wb.create_sheet("HDDT_Ban_ra")
    n2, _ = emit_hddt(
        ws2,
        f"TỔNG HỢP HĐĐT BÁN RA - {TEN} (MST {MST}) - {KY}",
        h_ban,
        d_ban,
        mst_ban,
        {13, 14, 17},
        db,
        f0_map,
        False,
    )

    bk_mua = [r for r in bk_recs if r["Chieu"] == "MUA"]
    bk_ban = [r for r in bk_recs if r["Chieu"] == "BAN"]
    bk_mua.sort(key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))
    bk_ban.sort(key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))
    bk_h = [
        "STT",
        "Nam",
        "Quy",
        "Thang",
        "Ky_hieu_HD",
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
        "Tình trạng hoạt động",
        "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro",
        "File_nguon",
    ]

    def pack_bk(rows, with_f0: bool):
        packed = []
        nf = 0
        for r in rows:
            tt, cqt = db_info(db, r["MST_chuan"])
            flag = f0_flag(r["MST_chuan"], f0_map) if with_f0 else ""
            if flag:
                nf += 1
            packed.append(
                [
                    None,
                    r["Nam"],
                    r["Quy"],
                    r["Thang"],
                    r["Ky_hieu_HD"],
                    r["So_HD"],
                    r["Ngay_HD"],
                    r["MST_goc"],
                    r["MST_chuan"],
                    r["Ten_DoiTac"],
                    r["Mat_hang"],
                    r["GT_chua_thue"],
                    r["Thue_suat"],
                    r["Thue_GTGT"],
                    r["Tong_tien"],
                    tt,
                    cqt,
                    flag,
                    r["File_nguon"],
                ]
            )
        return packed, nf

    packed_m, n_f0_bk = pack_bk(bk_mua, True)
    packed_b, _ = pack_bk(bk_ban, False)
    ws3 = wb.create_sheet("Bang_ke_Mua_vao")
    write_sheet_rows(ws3, f"BẢNG KÊ MUA VÀO - {TEN} (MST {MST}) - {KY}", bk_h, packed_m, {12, 14, 15})
    for rec, row in zip(packed_m, range(4, 4 + len(packed_m))):
        if rec[17]:
            ws3.cell(row, 18).fill = FILL_RED
            ws3.cell(row, 18).font = FONT_RED
    ws4 = wb.create_sheet("Bang_ke_Ban_ra")
    write_sheet_rows(ws4, f"BẢNG KÊ BÁN RA - {TEN} (MST {MST}) - {KY}", bk_h, packed_b, {12, 14, 15})

    partners = {}

    def add_p(mst, ten, src):
        m = chuan_hoa_mst(mst)
        if not m:
            return
        rec = partners.setdefault(m, {"mst": m, "ten": ten or "", "src": set()})
        if ten and len(str(ten)) > len(rec["ten"]):
            rec["ten"] = str(ten)
        rec["src"].add(src)

    for row in d_mua:
        add_p(row[5], row[6], "HDDT_Mua")
    for row in d_ban:
        add_p(row[7], row[8], "HDDT_Ban")
    for r in bk_mua:
        add_p(r["MST_chuan"], r["Ten_DoiTac"], "Bang_ke_Mua")
    for r in bk_ban:
        add_p(r["MST_chuan"], r["Ten_DoiTac"], "Bang_ke_Ban")

    ws5 = wb.create_sheet("Tong_hop_doi_tac")
    h5 = [
        "STT",
        "MST_chuan",
        "Ten_doi_tac",
        "Ghi_chu_nguon",
        "Tình trạng hoạt động",
        "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro",
    ]
    packed_p = []
    src_count = Counter()
    f0_partners = []
    for rec in sorted(partners.values(), key=lambda x: x["mst"]):
        lab = src_label(rec["src"])
        src_count[lab] += 1
        tt, cqt = db_info(db, rec["mst"])
        flag = f0_flag(rec["mst"], f0_map) if ("Mua" in lab or "HDDT_Mua" in rec["src"] or "Bang_ke_Mua" in rec["src"]) else ""
        if flag:
            f0_partners.append((rec["mst"], rec["ten"], flag, lab))
        packed_p.append([None, rec["mst"], rec["ten"], lab, tt, cqt, flag])
    write_sheet_rows(ws5, f"TỔNG HỢP ĐỐI TÁC - {TEN} (MST {MST})", h5, packed_p, set())
    for rec, row in zip(packed_p, range(4, 4 + len(packed_p))):
        if rec[6]:
            ws5.cell(row, 7).fill = FILL_RED
            ws5.cell(row, 7).font = FONT_RED

    nnt_f0 = MST in f0_map or mst_clean(MST) in f0_map
    nnt_f1 = MST in f1_map or mst_clean(MST) in f1_map
    nnt_ca = MST in ca_map or mst_clean(MST) in ca_map
    tt_nnt, cqt_nnt = db_info(db, MST)

    ws6 = wb.create_sheet("Tong_hop_trang_thai_rui_ro")
    put(ws6, 1, 1, f"TỔNG HỢP TRẠNG THÁI & RỦI RO - {TEN} (MST {MST}) - {KY}", bold=True)
    ws6["A1"].font = FONT_TITLE
    ws6.merge_cells("A1:F1")
    put(ws6, 3, 1, "A. Người nộp thuế đang kiểm tra", bold=True, fill=FILL_NAVY)
    ws6.cell(3, 1).font = FONT_W
    for i, h in enumerate(["Hạng mục", "Kết quả", "Chi tiết"], 1):
        put(ws6, 4, i, h, bold=True)
    style_header(ws6, 4, 3)
    nnt_rows = [
        ("MST / Tên", MST, TEN),
        ("Tình trạng ĐKKD (danh bạ 3.2)", tt_nnt or "Không có trên danh bạ file này", cqt_nnt),
        ("F0 CVCBRR", "Có" if nnt_f0 else "Không", cv_text(f0_map.get(mst_clean(MST)))),
        ("F1 CVCBRR", "Có" if nnt_f1 else "Không", cv_text(f1_map.get(mst_clean(MST)))),
        ("Sổ Công an", "Có" if nnt_ca else "Không", cv_text(ca_map.get(mst_clean(MST)))),
        ("Nguồn sổ", "CVCBRR + CA ngày chạy 25/11/2026", ""),
    ]
    for i, row in enumerate(nnt_rows):
        for c, v in enumerate(row, 1):
            put(ws6, 5 + i, c, v, wrap=c == 3)
            if row[1] == "Có":
                ws6.cell(5 + i, 2).fill = FILL_RED
                ws6.cell(5 + i, 2).font = FONT_RED

    put(ws6, 12, 1, "B. Thống kê ĐKKD đối tác (map danh bạ 3.2)", bold=True, fill=FILL_NAVY)
    ws6.cell(12, 1).font = FONT_W
    tt_count = Counter()
    mapped = 0
    for rec in partners.values():
        tt, _ = db_info(db, rec["mst"])
        if tt:
            mapped += 1
            tt_count[tt] += 1
        else:
            tt_count["Chưa map danh bạ"] += 1
    put(ws6, 13, 1, "Trạng thái", bold=True)
    put(ws6, 13, 2, "Số MST", bold=True)
    style_header(ws6, 13, 2)
    rr = 14
    for k, v in tt_count.most_common():
        put(ws6, rr, 1, k)
        put(ws6, rr, 2, v, fmt=NUM)
        rr += 1
    put(ws6, rr, 1, f"Đã map danh bạ {mapped}/{len(partners)} MST", bold=True)

    rr += 2
    put(ws6, rr, 1, "C. NCC mua vào dính F0 (chỉ F0, không gắn F1/CA cho NCC)", bold=True, fill=FILL_NAVY)
    ws6.cell(rr, 1).font = FONT_W
    rr += 1
    for i, h in enumerate(["STT", "MST", "Tên", "Cảnh báo rủi ro", "Nguồn"], 1):
        put(ws6, rr, i, h, bold=True)
    style_header(ws6, rr, 5)
    rr += 1
    if f0_partners:
        for i, (m, ten, flag, lab) in enumerate(f0_partners, 1):
            put(ws6, rr, 1, i, fmt=NUM)
            put(ws6, rr, 2, m)
            put(ws6, rr, 3, ten)
            put(ws6, rr, 4, flag)
            put(ws6, rr, 5, lab)
            ws6.cell(rr, 4).fill = FILL_RED
            rr += 1
    else:
        put(ws6, rr, 1, "Không có NCC mua vào trùng F0 trên kỳ 2020-2025.")
        rr += 1

    rr += 2
    put(ws6, rr, 1, "D. Số dòng nguồn", bold=True, fill=FILL_NAVY)
    ws6.cell(rr, 1).font = FONT_W
    rr += 1
    stats = [
        ("HĐĐT mua (tính, đã lọc hủy)", n1),
        ("HĐĐT bán (tính, đã lọc hủy)", n2),
        ("Bảng kê mua", len(bk_mua)),
        ("Bảng kê bán", len(bk_ban)),
        ("Đối tác unique", len(partners)),
        ("HĐĐT mua dính F0 (dòng)", n_f0_hd),
        ("Bảng kê mua dính F0 (dòng)", n_f0_bk),
        ("NCC F0 unique", len(f0_partners)),
    ]
    put(ws6, rr, 1, "Chỉ tiêu", bold=True)
    put(ws6, rr, 2, "Giá trị", bold=True)
    style_header(ws6, rr, 2)
    rr += 1
    for a, b in stats:
        put(ws6, rr, 1, a)
        put(ws6, rr, 2, b, fmt=NUM)
        rr += 1
    rr += 1
    put(
        ws6,
        rr,
        1,
        "Ghi chú: bảng kê lấy file năm 2020-2024 + quý 2025 mua vào; bỏ file tổng hợp 05/10 và các kỳ quý 2020-2024 để khỏi đếm trùng. Chưa thấy bảng kê bán ra 2025 trong 2.5. Danh bạ 3.2 là file Danh sach NCC.xlsx (325 MST), không phải danh bạ toàn quốc.",
        wrap=True,
        fill=FILL_YEL,
    )
    ws6.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=5)
    autowidth(ws6)
    ws6.freeze_panes = "A5"

    fname = f"Tong_hop_HDDT_BangKe_{MST}_{KY}.xlsx"
    dests = [CAU / fname, SAO / fname, VAULT_CT / fname]
    for d in dests:
        d.parent.mkdir(parents=True, exist_ok=True)
        wb.save(lp(d))
        print("saved", d, d.stat().st_size)
    print("COUNTS HDDT_M", n1, "HDDT_B", n2, "BK_M", len(bk_mua), "BK_B", len(bk_ban), "DT", len(partners))
    print("SRC", dict(src_count))
    print("F0 partners", len(f0_partners), "NNT F0/F1/CA", nnt_f0, nnt_f1, nnt_ca)
    print("DONE")


if __name__ == "__main__":
    main()
