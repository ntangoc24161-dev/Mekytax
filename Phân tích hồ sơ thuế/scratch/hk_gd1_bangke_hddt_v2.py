# -*- coding: utf-8 -*-
"""Huy Khang: gộp bảng kê + HĐĐT 4 sheet, chạy GĐ1 v2 sau khi NNT nộp sổ."""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603755214"
TEN = "CÔNG TY TNHH THIẾT BỊ CÔNG NGHIỆP HUY KHANG"
KY = "2020-2026"
KY_BK = "2020-2025"
VER = "v2"
TODAY = date(2026, 11, 25)

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Ke hoach nam\Bo sung\HUY_KHANG_3603755214"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
OUT = ROOT / "4. KET QUA PHAN TICH"
CAU = OUT / "cau-thanh"
B2 = ROOT / "2. HO SO NNT CUNG CAP"
B25 = B2 / "2.5. Bang ke mua vao ban ra"
HDDT_DIR = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu"
SRC_DUMP = ROOT / "DU LIEU HUY KHANG GOI THUE"
VAULT_CT = VAULT / "ket-qua" / "cau-thanh"

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FONT_W = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
FONT_TITLE = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
FILL_NAVY = PatternFill("solid", fgColor="1F4E78")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
FILL_GRN = PatternFill("solid", fgColor="E2EFDA")
FILL_TOT = PatternFill("solid", fgColor="DCE6F1")
DATE_FMT = "dd/mm/yyyy"
NUM = "#,##0"
NUM2 = "#,##0.00"


def lp(p) -> str:
    s = str(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def chuan_hoa_mst(val) -> str:
    if val is None:
        return ""
    s = str(val).strip().replace(" ", "").replace(".0", "")
    if s.endswith(".0"):
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


def sanitize(val):
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip()
        if s.startswith("="):
            return "'" + s
        return s
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
    s = s.replace(".", "").replace(",", ".") if s.count(".") > 1 else s.replace(",", ".")
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
    if "hàng hoá, dịch vụ" in blob:
        return True
    return False


def quy_of(d: date | None, fallback=None):
    if d:
        return (d.month - 1) // 3 + 1
    return fallback


def put(ws, r, c, val, fmt=None, bold=False, fill=None):
    cell = ws.cell(r, c)
    if isinstance(cell, MergedCell):
        return
    cell.value = sanitize(val)
    cell.font = FONT_B if bold else FONT
    cell.border = THIN
    if fill:
        cell.fill = fill
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


def clear_example(ws, start=3):
    max_r = ws.max_row or start
    drop = []
    for r in range(start, max_r + 1):
        joined = " ".join(str(ws.cell(r, c).value or "") for c in range(1, (ws.max_column or 1) + 1))
        if "VI_DU" in joined or "CÔNG TY A" in joined or "CÔNG TY X" in joined:
            drop.append(r)
    for r in reversed(drop):
        ws.delete_rows(r)


# ---------------------------------------------------------------------------
# Cleanup leftover dump (already copied)
# ---------------------------------------------------------------------------
def cleanup_dump():
    if not SRC_DUMP.exists():
        print("DUMP already gone")
        return
    left = [p for p in SRC_DUMP.rglob("*") if p.is_file()]
    print("DUMP leftover files", len(left))
    for p in left:
        try:
            os.remove(lp(p))
            print("  del", p.name)
        except Exception as e:
            print("  keep", p, e)
    # remove empty dirs
    try:
        shutil.rmtree(lp(SRC_DUMP), ignore_errors=True)
    except Exception:
        pass
    print("DUMP exists", SRC_DUMP.exists())


# ---------------------------------------------------------------------------
# Parse bang ke
# ---------------------------------------------------------------------------
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
        rows = []
        for i in range(sh.nrows):
            rows.append([sh.cell_value(i, j) for j in range(sh.ncols)])
        out.append((name, rows))
    return out


def sheet_chieu(sname: str, first_rows: list, fname: str) -> str | None:
    blob = " ".join([sname, fname] + [" ".join(str(c or "") for c in (first_rows[i] if i < len(first_rows) else [])) for i in range(min(4, len(first_rows)))]).casefold()
    if "tờ khai thuế" in blob and "bảng kê" not in sname.casefold():
        if "giảm thuế" in blob:
            return None
        if "tờ khai" in sname.casefold():
            return None
    if any(k in sname.casefold() for k in ["đầu ra", "dau ra", "bán ra", "ban ra", "bkbr"]):
        return "BAN"
    if any(k in sname.casefold() for k in ["đầu vào", "dau vao", "mua vào", "mua vao", "bkmv", "của h"]):
        return "MUA"
    if "bán ra" in blob or "ban ra" in blob or "dịch vụ bán" in blob:
        return "BAN"
    if "mua vào" in blob or "mua vao" in blob or "đầu vào" in blob:
        return "MUA"
    fn = fname.casefold()
    if "ban ra" in fn or "dau ra" in fn or "bán ra" in fn:
        return "BAN"
    if "mua vao" in fn or "dau vao" in fn or "mua vào" in fn:
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
    """Scan 1-3 header rows."""
    idx = {}
    for row in header_rows:
        for j, c in enumerate(row):
            t = fold(c).replace("\n", " ")
            if not t:
                continue
            if idx.get("so") is None and any(k in t for k in ["số hóa đơn", "số hoá đơn", "số hđ", "so hd"]):
                if "mẫu" not in t and "mau" not in t:
                    idx["so"] = j
            if idx.get("kh") is None and ("ký hiệu" in t or "kí hiệu" in t or "số hiệu" in t or "mã hóa đơn" in t):
                if "mẫu" in t or "mau" in t:
                    idx["mau"] = j
                else:
                    idx["kh"] = j
            if idx.get("ngay") is None and ("ngày" in t or "ngay" in t) and ("lập" in t or "phát hành" in t or "hóa đơn" in t or t in {"ngày", "ngay"}):
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
                k in t
                for k in [
                    "doanh số",
                    "doanh thu chưa",
                    "giá trị hhdv",
                    "chưa có thuế",
                    "chưa thuế",
                    "revenue excluding",
                ]
            ):
                idx["ds"] = j
            if idx.get("ts") is None and "thuế suất" in t:
                idx["ts"] = j
            if idx.get("thue") is None and (
                t in {"thuế gtgt", "vat amount"}
                or "thuế gtgt" in t
                or "tiền thuế" in t
                or "thuế gtgt đủ" in t
            ) and "suất" not in t and "mã số thuế" not in t and "chưa" not in t and "doanh số" not in t and "doanh thu" not in t:
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
    # fallbacks by format
    if idx.get("so") is None:
        # HTKK classic: so at col E (4)
        idx.setdefault("kh", 3)
        idx.setdefault("so", 4)
        idx.setdefault("ngay", 5)
        idx.setdefault("ten", 6)
        idx.setdefault("mst", 7)
        idx.setdefault("ds", 8)
        idx.setdefault("thue", 9)
    start = hi + 1
    # skip subheader rows like [1] [2]
    while start < min(hi + 4, len(rows)):
        joined = " ".join(fold(c) for c in rows[start])
        if joined.startswith("[1]") or joined.startswith("(1)") or "ký hiệu mẫu" in joined:
            start += 1
            continue
        break
    out = []
    for ridx, row in enumerate(rows[start:], start=start + 1):
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
        quy = quy_of(ngay, quy_hint)
        out.append(
            {
                "Chieu": chieu,
                "Nam": nam,
                "Quy": quy,
                "Thang": ngay.month if ngay else None,
                "Ky_hieu_HD": str(cell(row, idx.get("kh")) or "").strip(),
                "So_HD": so_s,
                "Ngay_HD": ngay,
                "MST_goc": mst_goc,
                "MST_chuan": chuan_hoa_mst(mst_goc),
                "Ten_DoiTac": str(ten or "").strip(),
                "Mat_hang": str(cell(row, idx.get("mh")) or "").strip(),
                "DVT": "",
                "So_luong": None,
                "GT_chua_thue": ds,
                "Thue_suat": ts_n if ts_n is not None else (str(ts).strip() if ts not in (None, "") else ""),
                "Thue_GTGT": thue,
                "Tong_tien": ds + thue,
                "File_nguon": path.name,
                "Sheet": sname,
            }
        )
    return out


def parse_all_bang_ke() -> list[dict]:
    files = [p for p in B25.rglob("*") if p.is_file() and p.suffix.lower() in {".xlsx", ".xls"} and not p.name.startswith("Bang_ke_mua_vao_ban_ra_tong_hop") and not p.name.startswith("Tong_hop")]
    print("BK source files", len(files))
    all_rows = []
    for p in sorted(files):
        nam_hint = None
        quy_hint = None
        m = re.search(r"(20\d{2})", str(p.relative_to(B25)))
        if m:
            nam_hint = int(m.group(1))
        m2 = re.search(r"Q([1-4])", p.name, re.I)
        if m2:
            quy_hint = int(m2.group(1))
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
        print(f"  {p.relative_to(B25)} -> {n_file}")
    print("BK TOTAL", len(all_rows), "MUA", sum(1 for r in all_rows if r["Chieu"] == "MUA"), "BAN", sum(1 for r in all_rows if r["Chieu"] == "BAN"))
    return all_rows


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
    # total formula
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


def write_bang_ke_file(recs: list[dict]) -> Path:
    mua = [r for r in recs if r["Chieu"] == "MUA"]
    ban = [r for r in recs if r["Chieu"] == "BAN"]
    mua.sort(key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))
    ban.sort(key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))
    headers = [
        "STT", "Nam", "Quy", "Thang", "Ky_hieu_HD", "So_HD", "Ngay_HD",
        "MST_goc", "MST_chuan", "Ten_DoiTac", "Mat_hang",
        "GT_chua_thue", "Thue_suat", "Thue_GTGT", "Tong_tien", "File_nguon",
    ]
    money = {12, 14, 15}

    def pack(rows):
        out = []
        for r in rows:
            out.append([
                None, r["Nam"], r["Quy"], r["Thang"], r["Ky_hieu_HD"], r["So_HD"], r["Ngay_HD"],
                r["MST_goc"], r["MST_chuan"], r["Ten_DoiTac"], r["Mat_hang"],
                r["GT_chua_thue"], r["Thue_suat"], r["Thue_GTGT"], r["Tong_tien"], r["File_nguon"],
            ])
        return out

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "mua vao"
    write_sheet_rows(
        ws1,
        f"BẢNG KÊ MUA VÀO - {TEN} (MST {MST}) - {KY_BK}",
        headers,
        pack(mua),
        money,
    )
    ws2 = wb.create_sheet("ban ra")
    write_sheet_rows(
        ws2,
        f"BẢNG KÊ BÁN RA - {TEN} (MST {MST}) - {KY_BK}",
        headers,
        pack(ban),
        money,
    )
    dests = [
        B25 / f"Bang_ke_mua_vao_ban_ra_tong_hop_{KY_BK}.xlsx",
        CAU / f"Bang_ke_mua_vao_ban_ra_tong_hop_{KY_BK}.xlsx",
        VAULT_CT / f"Bang_ke_mua_vao_ban_ra_tong_hop_{MST}_{KY_BK}.xlsx",
    ]
    for d in dests:
        d.parent.mkdir(parents=True, exist_ok=True)
        wb.save(lp(d))
        print("saved", d, d.stat().st_size)
    return dests[0]


# ---------------------------------------------------------------------------
# HDDT
# ---------------------------------------------------------------------------
def load_hddt(path: Path, chieu: str) -> tuple[list, list[tuple]]:
    wb = load_workbook(lp(path), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header_i = None
    for i, row in enumerate(rows[:12]):
        joined = " ".join(fold(c) for c in row)
        if "số hóa đơn" in joined and "mst" in joined:
            header_i = i
            break
    if header_i is None:
        raise RuntimeError(f"no header in {path}")
    header = list(rows[header_i])
    data = []
    skip = 0
    for row in rows[header_i + 1 :]:
        if not any(row):
            continue
        stt0 = str(row[0] if row else "").strip().casefold()
        if stt0.startswith("tổng") or stt0.startswith("tong"):
            skip += 1
            continue
        st = str(row[20] if len(row) > 20 else "" or "").casefold()
        if any(k in st for k in ["bị xóa", "bi xoa", "bị hủy", "bi huy", "bị thay thế", "bi thay the", "hủy bỏ"]):
            skip += 1
            continue
        data.append(row)
    print(f"HDDT {chieu} keep {len(data)} skip {skip} header@{header_i}")
    return header, data


def write_4_sheet(bk_recs: list[dict]):
    hddt_files = list(HDDT_DIR.glob("*.xlsx"))
    p_mua = next(p for p in hddt_files if "mua" in p.name.casefold())
    p_ban = next(p for p in hddt_files if "ban" in p.name.casefold() or "bán" in p.name.casefold())
    h_mua, d_mua = load_hddt(p_mua, "MUA")
    h_ban, d_ban = load_hddt(p_ban, "BAN")

    def hddt_out_header(h, mst_at):
        return list(h[: mst_at + 1]) + ["MST_chuan"] + list(h[mst_at + 1 :]) + ["Tình trạng hoạt động", "Cơ quan thuế quản lý"]

    def emit_hddt(ws, title, header, data, mst_at, money_idx):
        new_h = hddt_out_header(header, mst_at)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(new_h))
        put(ws, 1, 1, title, bold=True)
        ws["A1"].font = FONT_TITLE
        for i, t in enumerate(new_h, 1):
            put(ws, 3, i, t, bold=True)
        style_header(ws, 3, len(new_h))
        r = 4
        stt = 1
        for row in data:
            row = list(row)
            mst_c = chuan_hoa_mst(row[mst_at] if mst_at < len(row) else "")
            new = list(row[: mst_at + 1]) + [mst_c] + list(row[mst_at + 1 :]) + ["", ""]
            new[0] = stt
            stt += 1
            for c, v in enumerate(new, 1):
                fmt = NUM if (c - 1) in money_idx or c in money_idx else None
                # money after insert: original money cols shift +1 after mst_at
                orig_i = c - 1
                if orig_i > mst_at:
                    orig_i -= 1
                fmt = NUM if orig_i in money_idx else (DATE_FMT if isinstance(v, (date, datetime)) else None)
                put(ws, r, c, v, fmt=fmt)
            r += 1
        if r > 4:
            put(ws, r, 1, "TỔNG CỘNG", bold=True, fill=FILL_TOT)
            for orig in money_idx:
                # column in output: orig if orig <= mst_at else orig+1, 1-based
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
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        autowidth(ws)
        return stt - 1

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "HDDT_Mua_vao"
    n1 = emit_hddt(
        ws1,
        f"TỔNG HỢP HĐĐT MUA VÀO - {TEN} (MST {MST}) - {KY}",
        h_mua,
        d_mua,
        mst_at=5,
        money_idx={15, 17, 18},
    )
    ws2 = wb.create_sheet("HDDT_Ban_ra")
    n2 = emit_hddt(
        ws2,
        f"TỔNG HỢP HĐĐT BÁN RA - {TEN} (MST {MST}) - {KY}",
        h_ban,
        d_ban,
        mst_at=8,
        money_idx={15, 17, 18},
    )

    bk_mua = [r for r in bk_recs if r["Chieu"] == "MUA"]
    bk_ban = [r for r in bk_recs if r["Chieu"] == "BAN"]
    bk_mua.sort(key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))
    bk_ban.sort(key=lambda x: (x["Ngay_HD"] or date.min, x["So_HD"]))
    bk_h = [
        "STT", "Nam", "Quy", "Thang", "Ky_hieu_HD", "So_HD", "Ngay_HD",
        "MST_goc", "MST_chuan", "Ten_DoiTac", "Mat_hang",
        "GT_chua_thue", "Thue_suat", "Thue_GTGT", "Tong_tien",
        "Tình trạng hoạt động", "Cơ quan thuế quản lý", "File_nguon",
    ]

    def emit_bk(ws, title, rows):
        packed = []
        for r in rows:
            packed.append([
                None, r["Nam"], r["Quy"], r["Thang"], r["Ky_hieu_HD"], r["So_HD"], r["Ngay_HD"],
                r["MST_goc"], r["MST_chuan"], r["Ten_DoiTac"], r["Mat_hang"],
                r["GT_chua_thue"], r["Thue_suat"], r["Thue_GTGT"], r["Tong_tien"],
                "", "", r["File_nguon"],
            ])
        write_sheet_rows(ws, title, bk_h, packed, {12, 14, 15})
        return len(rows)

    ws3 = wb.create_sheet("Bang_ke_Mua_vao")
    n3 = emit_bk(ws3, f"BẢNG KÊ MUA VÀO - {TEN} (MST {MST}) - {KY_BK}", bk_mua)
    ws4 = wb.create_sheet("Bang_ke_Ban_ra")
    n4 = emit_bk(ws4, f"BẢNG KÊ BÁN RA - {TEN} (MST {MST}) - {KY_BK}", bk_ban)

    # Tong_hop_doi_tac
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
        add_p(row[8], row[9], "HDDT_Ban")
    for r in bk_mua:
        add_p(r["MST_chuan"], r["Ten_DoiTac"], "Bang_ke_Mua")
    for r in bk_ban:
        add_p(r["MST_chuan"], r["Ten_DoiTac"], "Bang_ke_Ban")

    def src_label(s):
        has_h_m = "HDDT_Mua" in s
        has_h_b = "HDDT_Ban" in s
        has_k_m = "Bang_ke_Mua" in s
        has_k_b = "Bang_ke_Ban" in s
        labels = []
        if has_h_m and has_k_m:
            labels.append("Chung_Mua")
        elif has_h_m:
            labels.append("HDDT_Mua")
        elif has_k_m:
            labels.append("Bang_ke_Mua")
        if has_h_b and has_k_b:
            labels.append("Chung_Ban")
        elif has_h_b:
            labels.append("HDDT_Ban")
        elif has_k_b:
            labels.append("Bang_ke_Ban")
        return "; ".join(labels)

    ws5 = wb.create_sheet("Tong_hop_doi_tac")
    h5 = ["STT", "MST_chuan", "Ten_doi_tac", "Ghi_chu_nguon"]
    packed = []
    for rec in sorted(partners.values(), key=lambda x: x["mst"]):
        packed.append([None, rec["mst"], rec["ten"], src_label(rec["src"])])
    write_sheet_rows(ws5, f"TỔNG HỢP ĐỐI TÁC - {TEN} (MST {MST})", h5, packed, set())

    dests = [
        CAU / f"Tong_hop_HDDT_BangKe_{MST}_{KY}.xlsx",
        VAULT_CT / f"Tong_hop_HDDT_BangKe_{MST}_{KY}.xlsx",
    ]
    for d in dests:
        d.parent.mkdir(parents=True, exist_ok=True)
        wb.save(lp(d))
        print("saved", d, d.stat().st_size)
    print("4-sheet counts HDDT_M", n1, "HDDT_B", n2, "BK_M", n3, "BK_B", n4, "DT", len(partners))
    return dests[0]


# ---------------------------------------------------------------------------
# GĐ1 v2
# ---------------------------------------------------------------------------
def classify_nnt(rel: str, name: str) -> str:
    n = name.casefold()
    d = rel.replace("/", "\\").casefold()
    if "2.5" in d or "bang ke" in n or "bangke" in n:
        return "IN02_BangKe"
    if "2.2" in d or any(k in n for k in ["ton_kho", "ton kho", "156", "154", "152"]):
        if re.search(r"\b(111|112|131|133|331|511|632)\b", n):
            pass
        else:
            return "IN09_Kho_GiaThanh"
    if re.search(r"(511|632)", n) and "2.1" in d:
        return "IN13_So_511_632"
    if re.search(r"\b111\b", n) or "tien mat" in n:
        return "IN10_Chi_tien_mat"
    if "sao ke" in n or "unc" in n:
        return "IN11_Sao_ke_NH"
    if "2.4" in d:
        return "IN08_So_KeToan"
    return "IN08_So_KeToan"


def tk_from_name(name: str) -> str:
    m = re.search(r"(?:tk|tai khoan|tài khoản)?\s*(\d{3})\b", name, re.I)
    return m.group(1) if m else ""


def fill_gd1(bk_recs: list[dict]):
    nnt_files = [p for p in B2.rglob("*") if p.is_file()]
    print("NNT files", len(nnt_files))
    kq12_src = MAU / "KQ1.2_Kiem_ke_ho_so_NNT.xlsx"
    kq13_src = OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_v1.xlsx"
    kq12 = OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_{VER}.xlsx"
    kq13 = OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx"
    shutil.copy2(lp(kq12_src), lp(kq12))
    shutil.copy2(lp(kq13_src), lp(kq13))

    # ----- KQ1.2 -----
    wb = load_workbook(lp(kq12))
    ws = wb["IN00_Manifest"]
    clear_example(ws, 3)
    # wipe leftover
    if ws.max_row and ws.max_row >= 3:
        ws.delete_rows(3, ws.max_row - 2)
    r = 3
    by_in = defaultdict(list)
    for i, p in enumerate(sorted(nnt_files, key=lambda x: str(x)), 1):
        rel = str(p.relative_to(ROOT))
        ma = classify_nnt(rel, p.name)
        by_in[ma].append(p)
        try:
            h = sha256(p)
        except Exception:
            h = ""
        try:
            mtime = datetime.fromtimestamp(os.stat(lp(p)).st_mtime).date()
        except Exception:
            mtime = TODAY
        put(ws, r, 1, f"N{i:03d}")
        put(ws, r, 2, ma)
        put(ws, r, 3, "DN")
        put(ws, r, 4, p.suffix.lower().lstrip(".") or "file")
        put(ws, r, 5, mtime)
        put(ws, r, 6, 2020)
        put(ws, r, 7, 2025)
        put(ws, r, 8, h)
        put(ws, r, 9, "Không")
        put(ws, r, 10, "Có")
        put(ws, r, 11, None)
        put(ws, r, 12, 0)
        put(ws, r, 13, "Giữ nguyên")
        put(ws, r, 14, rel)
        put(ws, r, 15, p.name)
        r += 1
    autowidth(ws, max_row=40)

    ws = wb["IN02_BangKe"]
    clear_example(ws, 3)
    if ws.max_row and ws.max_row >= 3:
        ws.delete_rows(3, ws.max_row - 2)
    headers = [ws.cell(2, c).value for c in range(1, 25)]
    r = 3
    for rec in bk_recs:
        vals = [
            rec["Chieu"], rec["Nam"], rec["Quy"], rec["Thang"], rec["Ky_hieu_HD"], rec["So_HD"], rec["Ngay_HD"],
            rec["MST_goc"], rec["MST_chuan"], rec["Ten_DoiTac"], rec["Mat_hang"], rec["DVT"], rec["So_luong"],
            rec["GT_chua_thue"], rec["Thue_suat"], rec["Thue_GTGT"], rec["Tong_tien"],
            "", "", "", "", "", rec["File_nguon"], "",
        ]
        for c, v in enumerate(vals, 1):
            fmt = NUM if c in {14, 16, 17} else None
            put(ws, r, c, v, fmt=fmt)
        r += 1
    print("IN02 rows", r - 3)

    def fill_file_sheet(sheet, files, extra):
        ws = wb[sheet]
        clear_example(ws, 3)
        if ws.max_row and ws.max_row >= 3:
            ws.delete_rows(3, ws.max_row - 2)
        r = 3
        for p in files:
            rec = extra(p)
            for c, v in enumerate(rec, 1):
                put(ws, r, c, v)
            r += 1
        return r - 3

    n08 = fill_file_sheet(
        "IN08_So_KeToan",
        [p for p in by_in["IN08_So_KeToan"] if "2.1" in str(p)],
        lambda p: [
            int(m.group(1)) if (m := re.search(r"(20\d{2})", str(p))) else None,
            None,
            tk_from_name(p.name) or "",
            None, None, None, None, None,
            "Sổ NNT",
            p.name,
        ],
    )
    n09 = fill_file_sheet(
        "IN09_Kho_GiaThanh",
        by_in["IN09_Kho_GiaThanh"],
        lambda p: [
            int(m.group(1)) if (m := re.search(r"(20\d{2})", str(p))) else None,
            "HTK",
            None, None, None, None, None, None, None, None, None, None, None, None, None,
            p.name,
        ],
    )
    n10 = fill_file_sheet(
        "IN10_Chi_tien_mat",
        by_in["IN10_Chi_tien_mat"],
        lambda p: [None, int(m.group(1)) if (m := re.search(r"(20\d{2})", str(p))) else None, "", "111", "", "", "Sổ chi tiết TK 111", None, "", "Chưa rõ", "GĐ1 kiểm kê file", p.name],
    )
    ws = wb["IN11_Sao_ke_NH"]
    clear_example(ws, 3)
    put(ws, 3, 14, "Chưa nhận sao kê ngân hàng / UNC (chỉ có sổ 112 kế toán)")
    ws = wb["IN12_GDLK"]
    clear_example(ws, 3)
    put(ws, 3, 9, "Chưa thấy phụ lục GDLK")
    n13 = fill_file_sheet(
        "IN13_So_511_632",
        by_in["IN13_So_511_632"],
        lambda p: [
            int(m.group(1)) if (m := re.search(r"(20\d{2})", str(p))) else None,
            None, None, None,
            tk_from_name(p.name) or ("511" if "511" in p.name else "632"),
            "", "", "Sổ chi tiết", "", None, None, None, "", "", "", None, None, None, None,
            p.name,
        ],
    )
    for name in wb.sheetnames:
        autowidth(wb[name], max_row=30)
    wb.save(lp(kq12))
    print("saved KQ1.2", n08, n09, n10, n13)

    # ----- KQ1.3 -----
    mua_n = sum(1 for x in bk_recs if x["Chieu"] == "MUA")
    ban_n = sum(1 for x in bk_recs if x["Chieu"] == "BAN")
    years_bk = sorted({x["Nam"] for x in bk_recs if x["Nam"]})
    wb = load_workbook(lp(kq13))
    ws = wb["KQ1_Kiem_ke"]
    put(ws, 2, 2, MST, bold=True)
    put(ws, 2, 8, TEN, bold=True)
    put(ws, 3, 2, KY)
    put(ws, 3, 8, "D", bold=True)
    put(ws, 4, 2, f"IN Đủ: IN01, IN02, IN03, IN04, IN08, IN09, IN10, IN13, IN14 | Mỏng: IN05, IN07 | Thiếu: IN11, IN12. Bảng kê {mua_n} mua / {ban_n} bán, năm {years_bk}")
    put(ws, 4, 8, TODAY)
    put(ws, 5, 2, VER)
    put(ws, 5, 8, "1.4 vòng 2")
    put(ws, 6, 2, TODAY)
    put(ws, 6, 8, "Có - NNT nộp 2020-2025")
    put(ws, 7, 2, "GĐ3 một phần: 511-632, kho, công nợ, 111. Chưa chạy đủ vì thiếu sao kê/UNC (IN11)")
    put(ws, 7, 8, "Bảng kê mỏng 2020 (Q4) và 2021 (Q1); 2022-2025 đủ quý")

    def eval_row(row, *vals):
        for i, v in enumerate(vals, 6):
            put(ws, row, i, v)
        muc = vals[6]
        if muc == "Đủ":
            ws.cell(row, 12).fill = FILL_GRN
        elif muc == "Mỏng":
            ws.cell(row, 12).fill = FILL_YEL
        else:
            ws.cell(row, 12).fill = FILL_RED

    n08_so = len([p for p in by_in["IN08_So_KeToan"] if "2.1" in str(p)])
    # keep IN01/03/04/05/07/14 from v1; overwrite IN02, IN08-IN13
    eval_row(11, "Excel NNT", "Q4/2020; Q1/2021; 2022-2025", "2020 Q1-Q3; 2021 Q2-Q4; 2026", mua_n + ban_n, 0, "Không", "Mỏng", "Đã gộp bảng kê, lọc Tổng cộng", "Đã chuẩn hóa MST", "IN02 đủ mua+bán", 0, "2020 chỉ Q4, 2021 chỉ Q1")
    eval_row(16, "Excel NNT", "2020-2025", "Sổ 2020 gộp 1 file .xls", n08_so, 0, "Không", "Đủ", "NKC, sổ cái, 131/331/111/112/133", "Chưa chuẩn hóa dòng", "GĐ1 kiểm kê file, chưa bóc PS", 0, "File 2020 .xls 59MB chưa bóc dòng")
    eval_row(17, "Excel NNT", "2020-2025", "Định mức/giá thành 154 chưa tách", len(by_in["IN09_Kho_GiaThanh"]), 0, "Không", "Đủ", "XNT/tồn kho theo tháng/năm", "Chưa chuẩn hóa dòng", "", 0, "")
    eval_row(18, "Excel NNT", "2021-2025", "2020 nằm trong sổ gộp", len(by_in["IN10_Chi_tien_mat"]), 0, "Không", "Đủ", "Sổ chi tiết TK 111", "Chưa chuẩn hóa dòng", "", 0, "")
    eval_row(19, "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Chỉ có sổ 112 kế toán, chưa có sao kê/UNC", "Chưa chuẩn hóa", "", None, "Chặn TT05/TT06/GT16")
    eval_row(20, "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Chưa thấy PL GDLK", "Chưa chuẩn hóa", "", None, "")
    eval_row(21, "Excel NNT", "2021-2025", "2020 nằm trong sổ gộp", len(by_in["IN13_So_511_632"]), 0, "Không", "Đủ", "Sổ 511 và 632 theo năm", "Chưa chuẩn hóa dòng", "Mở DT01-DT03", 0, "")

    ws = wb["KQ1_Kha_nang_phan_tich"]
    put(ws, 2, 2, MST)
    put(ws, 2, 6, KY)
    put(ws, 4, 2, VER)
    put(ws, 4, 6, "1.4")
    put(ws, 5, 2, TODAY)
    put(ws, 5, 6, "Mức D - có sổ + kho, thiếu sao kê/UNC")
    have = {"IN01", "IN02", "IN03", "IN04", "IN05", "IN08", "IN09", "IN10", "IN13", "IN14"}
    have_part = {"IN07"}
    n_ngay = n_mot = n_chan = 0
    for r in range(10, 90):
        ma = ws.cell(r, 1).value
        if not ma:
            continue
        need = str(ws.cell(r, 4).value or "")
        ins = set(re.findall(r"IN\d{2}", need.replace("–", "-").replace("—", "-")))
        if "IN01-IN08" in need.replace("–", "-") or "IN01–IN08" in need:
            ins |= {f"IN{i:02d}" for i in range(1, 9)}
        missing = [i for i in sorted(ins) if i not in have and i not in have_part]
        if ma == "HD07":
            kha, thieu, goi, n_chan = "Bị chặn", "Tham chiếu ngành (K13)", "TREO theo K13", n_chan + 1
        elif ma == "TN20":
            kha, thieu, goi, n_chan = "Bị chặn", "Chưa triển khai / thiếu khấu hao chi tiết", "CHƯA_TRIỂN_KHAI", n_chan + 1
        elif not ins or all(i in have for i in ins):
            kha, thieu, goi, n_ngay = "Chạy ngay", "", "Đủ nguồn sau NNT nộp sổ", n_ngay + 1
        elif any(i in have or i in have_part for i in ins) and missing:
            kha, thieu, goi, n_mot = "Chạy một phần", ", ".join(missing), "Chạy phần có nguồn", n_mot + 1
        else:
            kha, thieu, goi, n_chan = "Bị chặn", ", ".join(missing) if missing else "Thiếu nguồn", "Xin NNT / trích CQT", n_chan + 1
        put(ws, r, 8, kha)
        put(ws, r, 9, thieu)
        put(ws, r, 10, thieu)
        put(ws, r, 11, goi)
        ws.cell(r, 8).fill = FILL_GRN if kha == "Chạy ngay" else (FILL_YEL if kha == "Chạy một phần" else FILL_RED)
    put(ws, 3, 2, n_ngay)
    put(ws, 3, 5, n_mot)
    put(ws, 3, 6, n_chan)
    print("rules ngay/mot/chan", n_ngay, n_mot, n_chan)

    ws = wb["KQ1_De_nghi_bo_sung"]
    put(ws, 2, 2, MST)
    put(ws, 2, 6, KY)
    put(ws, 3, 2, VER)
    put(ws, 3, 6, TODAY)
    # clear old items
    for rr in range(10, 30):
        for c in range(1, 12):
            ws.cell(rr, c).value = None
    items = [
        (1, "Sao kê ngân hàng Excel + UNC các món ≥ 20tr/5tr", "IN11", "2.3", "NNT cung cấp", 5, "GT16, TT02, TT03, TT05, TT06", "Còn thiếu để lên mức E"),
        (2, "Bảng kê GTGT Q1-Q3/2020 và Q2-Q4/2021", "IN02", "2.5", "NNT cung cấp", 2, "GT02, GT04", "Hiện chỉ có Q4/2020 và Q1/2021"),
        (3, "Danh bạ trạng thái ĐKKD toàn bộ NCC", "IN07", "3.2", "CQT trích", 3, "HD02, HD05", ""),
        (4, "Phụ lục GDLK (nếu có giao dịch liên kết)", "IN12", "2.1", "NNT cung cấp", 2, "TC16, TC18", ""),
        (5, "Bảng kê / HĐĐT năm 2026 nếu kỳ kiểm tra phủ 2026", "IN02", "2.5", "NNT / CQT", 1, "GT01-GT04", "XML GTGT 2026 đã có, bảng kê chưa có"),
    ]
    for i, it in enumerate(items):
        for c, v in enumerate(it, 1):
            put(ws, 10 + i, c, v)
    autowidth(wb["KQ1_Kiem_ke"], max_row=30)
    autowidth(wb["KQ1_Kha_nang_phan_tich"], max_row=90)
    autowidth(wb["KQ1_De_nghi_bo_sung"], max_row=20)
    wb.save(lp(kq13))
    print("saved KQ1.3")
    return {
        "nnt_files": len(nnt_files),
        "in02": mua_n + ban_n,
        "mua": mua_n,
        "ban": ban_n,
        "n08": n08,
        "n09": n09,
        "n10": n10,
        "n13": n13,
        "ngay": n_ngay,
        "mot": n_mot,
        "chan": n_chan,
        "kq12": str(kq12),
        "kq13": str(kq13),
        "years_bk": years_bk,
        "by_in": {k: len(v) for k, v in by_in.items()},
    }


def main():
    CAU.mkdir(parents=True, exist_ok=True)
    VAULT_CT.mkdir(parents=True, exist_ok=True)
    cleanup_dump()
    recs = parse_all_bang_ke()
    write_bang_ke_file(recs)
    write_4_sheet(recs)
    summary = fill_gd1(recs)
    print("SUMMARY", summary)


if __name__ == "__main__":
    main()
