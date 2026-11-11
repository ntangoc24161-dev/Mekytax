# -*- coding: utf-8 -*-
"""Phan tich ho so thue v1.4 - CTY TNHH DAI HUY HOANG MST 3602179712, ky 2021-2025."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

MST = "3602179712"
TEN = "CÔNG TY TNHH ĐẠI HUY HOÀNG"
KY = "2021-2025"
YEARS = [2021, 2022, 2023, 2024, 2025]
VER = "v1"
TODAY = date(2026, 11, 9)
LOAI_HINH = "San xuat (go tram) + cho thue xuong / ban dien"

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\DAI_HUY_HOANG_3602179712"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
ENGINE = VAULT / "skills" / "phan-tich-ho-so-thue" / "engine"
OUT = ROOT / "4. KET QUA PHAN TICH"
CVC = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
DANHBA = ROOT / "3. DU LIEU BO SUNG & TRA CUU QLT" / "3.2. Danh ba trang thai NNT"

sys.path.insert(0, str(ENGINE))
from ingest import xml_tkhai as xt  # noqa: E402
from ingest.xml_tkhai import file_hash, parse_xml, sniff_xml  # noqa: E402

# HTKK cũ / TT133: mã tờ khai khác bộ 842/892/953/402
xt.MA_LOAI.update(
    {
        "01": "gtgt",
        "03": "tndn",
        "684": "bctc",
        "50": "bctc",
        "395": "tncn",
    }
)

FONT = Font(name="Times New Roman", size=10)
FONT_B = Font(name="Times New Roman", size=10, bold=True)
FONT_H = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
NAVY = PatternFill("solid", fgColor="1F4E78")
PALE = PatternFill("solid", fgColor="D9E1F2")
RED = PatternFill("solid", fgColor="FCE4D6")
YEL = PatternFill("solid", fgColor="FFF2CC")
GRN = PatternFill("solid", fgColor="E2EFDA")
NOTE = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
RED_F = Font(name="Times New Roman", size=10, color="C00000")
ORG_F = Font(name="Times New Roman", size=10, color="B25900")
GRN_F = Font(name="Times New Roman", size=10, color="385723")
DATE_FMT = "dd/mm/yyyy"
NUM = "#,##0"
NUM2 = "#,##0.00"
PCT = "0.00%"
WRAP = Alignment(vertical="center", wrap_text=True)
CENTER = Alignment(vertical="center", horizontal="center", wrap_text=True)

SKIP_DIR = {"0. bien ban", "0. bien ban".upper(), "7. sao luu tong hop", "ket-qua", "4. ket qua phan tich"}
HUY_RE = re.compile(r"huy|xoá|xóa|thay the|thay thế|da bi|đã bị", re.I)
TONG_RE = re.compile(r"^\s*(tổng|tong|cộng|cong)\b", re.I)
MST_RE = re.compile(r"(\d{10})(?:[-\s]?(\d{3}))?")


def lp(p: Path | str) -> str:
    s = str(p)
    if not s.startswith("\\\\?\\"):
        s = "\\\\?\\" + s
    return s


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def nnum(v, default=None):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\xa0", "")
    s = s.replace(".", "").replace(",", ".") if s.count(",") == 1 and s.count(".") > 1 else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return default


def as_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
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


def mst_chuan(s) -> str:
    t = re.sub(r"[^0-9]", "", str(s or ""))
    if len(t) >= 13:
        return t[:10] + "-" + t[10:13]
    if len(t) >= 10:
        return t[:10]
    return t


def mst_digits(s) -> str:
    return re.sub(r"[^0-9]", "", str(s or ""))[:13]


def so_hd_norm(s) -> str:
    t = str(s or "").strip()
    t = re.sub(r"^0+", "", t)
    return t or "0"


def hd_from_text(s) -> str:
    t = str(s or "")
    m = re.search(r"(?:HĐ|HD|HDBR|PXK)[^\d]{0,6}(\d{1,8})", t, re.I)
    if m:
        return so_hd_norm(m.group(1))
    m = re.search(r"(\d{1,8})", t)
    return so_hd_norm(m.group(1)) if m else ""


def year_of(d) -> int | None:
    if d is None:
        return None
    if isinstance(d, date):
        return d.year
    return None


def quy_of(d: date | None) -> int | None:
    if not d:
        return None
    return (d.month - 1) // 3 + 1


def in_ky(d: date | None) -> bool:
    y = year_of(d)
    return y in YEARS if y else False


def is_huy(status) -> bool:
    return bool(HUY_RE.search(str(status or "")))


def safe_text(s) -> str:
    return str(s or "").replace("\x00", "").strip()


def apply_font(ws: Worksheet):
    ws.sheet_view.showGridLines = True
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.font is None or cell.font.name != "Times New Roman":
                b = bool(cell.font and cell.font.bold)
                sz = cell.font.size if cell.font and cell.font.size else 10
                color = cell.font.color if cell.font else None
                cell.font = Font(name="Times New Roman", size=sz, bold=b, color=color)


def autofit(ws: Worksheet, cap=55):
    widths = {}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 400)):
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            v = str(cell.value)
            widths[cell.column] = min(cap, max(widths.get(cell.column, 10), min(len(v) + 2, cap)))
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(10, w)


def put(ws, r, c, value, bold=False, num=False, pct=False, dt=False, fill=None, font=None, wrap=True, center=False):
    cell = ws.cell(r, c)
    if isinstance(cell, MergedCell):
        for mr in list(ws.merged_cells.ranges):
            if cell.coordinate in mr:
                cell = ws.cell(mr.min_row, mr.min_col)
                break
        else:
            return None
    cell.value = value
    cell.font = font or (FONT_B if bold else FONT)
    cell.alignment = CENTER if center else WRAP
    cell.border = THIN
    if fill:
        cell.fill = fill
    if dt:
        cell.number_format = DATE_FMT
        if isinstance(value, datetime):
            cell.value = value.date()
    elif pct and isinstance(value, (int, float)):
        cell.number_format = PCT
    elif num and isinstance(value, (int, float)):
        cell.number_format = NUM if float(value).is_integer() or abs(value) >= 1 else NUM2
    return cell


def clear_example(ws, header_row=2):
    """Delete example row immediately under header if it looks like VI_DU / 0000000001."""
    r = header_row + 1
    vals = [ws.cell(r, c).value for c in range(1, 8)]
    joined = " ".join(str(v) for v in vals if v)
    if "VI_DU" in joined or "0000000001" in joined or "ví dụ" in joined.lower() or "CÔNG TY A" in joined:
        ws.delete_rows(r, 1)


def copy_wb(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(lp(src), lp(dst))


def header_map(ws, row=2):
    m = {}
    for c in range(1, (ws.max_column or 1) + 1):
        v = ws.cell(row, c).value
        if v:
            m[str(v).strip()] = c
    return m


def write_rows(ws, header_row, rows, start=None):
    hm = header_map(ws, header_row)
    r0 = start or header_row + 1
    # wipe leftover template formulas below
    max_r = ws.max_row or r0
    if max_r > r0:
        ws.delete_rows(r0, max_r - r0 + 1)
    for i, rec in enumerate(rows):
        rr = r0 + i
        for k, v in rec.items():
            if k not in hm:
                continue
            c = hm[k]
            if isinstance(v, date) and not isinstance(v, datetime):
                put(ws, rr, c, v, dt=True, center=True)
            elif isinstance(v, datetime):
                put(ws, rr, c, v.date(), dt=True, center=True)
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                put(ws, rr, c, v, num=True)
            else:
                put(ws, rr, c, v)
    return r0 + len(rows) - 1 if rows else r0 - 1


# ---------------------------------------------------------------------------
# SCAN
# ---------------------------------------------------------------------------
def classify_file(rel: str, name: str) -> str:
    u = (rel + "/" + name).lower()
    n = name.lower()
    if "1.2" in u and "hddt" in n:
        return "IN01"
    if "2.5" in u or "bang ke" in u:
        return "IN02"
    if n.endswith(".xml") or n.endswith(".xml.xml"):
        if "gtgt" in n or "01/gtgt" in n:
            return "IN03"
        if "bctc" in n or "b01" in n or "133_b01" in n:
            return "IN04"
        if "tndn" in n:
            return "IN04"
        if "tncn" in n or "qtt-tncn" in n or "05_" in n:
            return "IN05"
        return "IN03"
    if "1.3" in u or n.startswith("dai huy hoang-ptdn"):
        return "IN14"
    if "1.4" in u or "nld" in n:
        return "IN05"
    if "3.1" in u or "cvcbrr" in n:
        return "IN07"
    if "3.2" in u or "danh_sach_doi_tac" in n or "danh ba" in u:
        return "IN07"
    if "gdlk" in n or "lien ket" in n:
        return "IN12"
    if re.search(r"(^|[^\d])511", n) or "511" in n:
        return "IN13"
    if re.search(r"(^|[^\d])632", n) or "632" in n:
        return "IN13"
    if "111" in n and "112" not in n:
        return "IN10"
    if "nxt" in n or "152" in n or "154" in n or "155" in n or "kho" in u and "2.2" in u:
        return "IN09"
    if "2.2" in u:
        return "IN09"
    if "2.3" in u or "sao ke" in n or "unc" in n:
        return "IN11"
    if "nkc" in n or "2.1" in u:
        return "IN08"
    if "2.4" in u:
        return "IN08"
    return "KHAC"


def scan_all():
    files = []
    for p in ROOT.rglob("*"):
        if not p.is_file():
            continue
        rel = str(p.relative_to(ROOT))
        low_parts = {x.lower() for x in p.parts}
        if any("ket qua phan tich" in x or "sao luu" in x or "bien ban" in x for x in low_parts):
            if "1. du lieu" not in rel.lower() and "2. ho so" not in rel.lower() and "3. du lieu" not in rel.lower():
                continue
        try:
            st = p.stat()
        except OSError:
            continue
        ma = classify_file(rel.replace("\\", "/"), p.name)
        files.append(
            {
                "path": p,
                "rel": rel,
                "name": p.name,
                "ma_in": ma,
                "ext": p.suffix.lower(),
                "size": st.st_size,
                "mtime": datetime.fromtimestamp(st.st_mtime),
                "sha": sha256(p) if st.st_size < 80_000_000 else "",
            }
        )
    return files


# ---------------------------------------------------------------------------
# XML
# ---------------------------------------------------------------------------
def parse_all_xml():
    xml_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"
    tks = []
    skipped = []
    seen = set()
    for p in xml_dir.rglob("*"):
        if not p.is_file():
            continue
        if ".xml" not in "".join(p.suffixes).lower():
            continue
        if not sniff_xml(p):
            skipped.append(p.name)
            continue
        h = file_hash(p)
        if h in seen:
            continue
        seen.add(h)
        tk = parse_xml(p)
        if tk is None:
            skipped.append("fail:" + p.name)
            continue
        tks.append(tk)
    # latest per (loai, ky)
    best = {}
    for tk in tks:
        key = (tk.loai, tk.ky)
        old = best.get(key)
        score = (1 if tk.loai_tkhai == "B" else 0, tk.so_lan)
        if old is None or score >= (1 if old.loai_tkhai == "B" else 0, old.so_lan):
            best[key] = tk
    return list(best.values()), skipped


def parse_tncn_nld(path: Path, nam: int) -> list[dict]:
    rows = []
    try:
        tree = ET.parse(lp(path))
        root = tree.getRoot() if hasattr(tree, "getRoot") else tree.getroot()
    except Exception:
        return rows

    def ln(tag):
        return tag.split("}", 1)[-1] if tag else ""

    def txt(el):
        return (el.text or "").strip() if el is not None else ""

    # walk potential employee nodes
    for el in root.iter():
        kids = {ln(ch.tag): ch for ch in list(el)}
        keys = set(kids)
        if not ({"hoTen", "tenNNT", "tenNld", "hoVaTen"} & keys):
            continue
        ho = txt(kids.get("hoTen") or kids.get("tenNNT") or kids.get("tenNld") or kids.get("hoVaTen"))
        if not ho:
            continue
        cccd = txt(kids.get("soCMND") or kids.get("cccd") or kids.get("soDDCN") or kids.get("maSoThueTNCN"))
        mst_cn = txt(kids.get("mst") or kids.get("mstTNCN") or kids.get("maSoThue"))
        tnct = nnum(txt(kids.get("tongTNCT") or kids.get("ct21") or kids.get("thuNhapChiuThue")), 0) or 0
        thue = nnum(txt(kids.get("thueTNCN") or kids.get("thueDaKhauTru") or kids.get("ct36")), 0) or 0
        rows.append(
            {
                "Loai_dong": "NLD",
                "Nam": nam,
                "Ho_ten": ho,
                "CCCD": cccd,
                "MST_CN": mst_cn,
                "Nguon_BK": "05-1",
                "Loai_thu_nhap": "",
                "Tong_TNCT": tnct,
                "Thue_da_KT": thue,
                "Co_trich_10": "",
                "File_nguon": path.name,
            }
        )
    return rows


# ---------------------------------------------------------------------------
# HDDT / Bang ke
# ---------------------------------------------------------------------------
def parse_hddt(path: Path) -> list[dict]:
    wb = load_workbook(lp(path), read_only=True, data_only=True)
    out = []
    for sheet, chieu in (("Mua vao", "MUA"), ("Ban ra", "BAN")):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        rows = ws.iter_rows(values_only=True)
        header = next(rows)
        idx = {str(h).strip(): i for i, h in enumerate(header) if h}

        def col(*cands):
            for c in cands:
                for k, i in idx.items():
                    if c.lower() in k.lower():
                        return i
            return None

        i_kh = col("Ký hiệu hóa đơn")
        i_so = col("Số hóa đơn")
        i_ng = col("Ngày lập")
        i_mst_ban = col("MST người bán", "MST người xuất")
        i_ten_ban = col("Tên người bán", "Tên người xuất")
        i_mst_mua = col("MST người mua", "MST người nhận")
        i_ten_mua = col("Tên người mua", "Tên người nhận")
        i_ds = col("Tổng tiền chưa thuế")
        i_thue = col("Tổng tiền thuế")
        i_tt = col("Trạng thái hóa đơn")
        i_mh = col("Tên hàng", "Hàng hóa")
        for rec in rows:
            if not rec or not any(rec):
                continue
            ngay = as_date(rec[i_ng] if i_ng is not None else None)
            if not in_ky(ngay):
                continue
            stt = rec[i_tt] if i_tt is not None else ""
            tinh = "Không" if is_huy(stt) else "Có"
            mst_ban = mst_chuan(rec[i_mst_ban] if i_mst_ban is not None else "")
            mst_mua = mst_chuan(rec[i_mst_mua] if i_mst_mua is not None else "")
            dt_mst = mst_ban if chieu == "MUA" else mst_mua
            out.append(
                {
                    "Chieu": chieu,
                    "Nam": ngay.year if ngay else None,
                    "Quy": quy_of(ngay),
                    "Thang": ngay.month if ngay else None,
                    "Ky_hieu_HD": rec[i_kh] if i_kh is not None else "",
                    "So_HD": str(rec[i_so] or "").strip(),
                    "Ngay_HD": ngay,
                    "MST_NB": mst_ban,
                    "Ten_NB": rec[i_ten_ban] if i_ten_ban is not None else "",
                    "MST_NM": mst_mua,
                    "Ten_NM": rec[i_ten_mua] if i_ten_mua is not None else "",
                    "MST_DoiTac_chuan": dt_mst,
                    "Mat_hang": rec[i_mh] if i_mh is not None else "",
                    "GT_chua_thue": nnum(rec[i_ds], 0) or 0,
                    "Thue_GTGT": nnum(rec[i_thue], 0) or 0,
                    "Tinh_vao_tong": tinh,
                    "Trang_thai_CQT": stt,
                    "File_nguon": path.name,
                }
            )
    wb.close()
    return out


def parse_bang_ke_folder(folder: Path) -> list[dict]:
    out = []
    for p in folder.rglob("*.xls"):
        try:
            book = xlrd.open_workbook(lp(p))
        except Exception:
            continue
        # year/quy from filename
        m = re.search(r"(20\d{2})\s*Q([1-4])", p.name, re.I)
        nam = int(m.group(1)) if m else None
        quy = int(m.group(2)) if m else None
        for sname in book.sheet_names():
            chieu = "MUA" if sname.upper() in {"VAO", "MUA"} or "vao" in sname.lower() else None
            if sname.upper() in {"RA", "BAN"} or "ban" in sname.lower() or sname.upper() == "RA":
                chieu = "BAN"
            if not chieu:
                continue
            sh = book.sheet_by_name(sname)
            # header row
            hdr = None
            for i in range(min(6, sh.nrows)):
                vals = [str(sh.cell_value(i, j)).strip().upper() for j in range(min(12, sh.ncols))]
                if "SO_HD" in vals or "SỐ HĐ" in " ".join(vals) or "SO HD" in " ".join(vals):
                    hdr = {vals[j]: j for j in range(len(vals))}
                    start = i + 1
                    break
            if hdr is None:
                continue

            def g(row, *keys):
                for k in keys:
                    if k in hdr:
                        return sh.cell_value(row, hdr[k])
                return None

            for r in range(start, sh.nrows):
                so = g(r, "SO_HD", "SO HD")
                ten = str(g(r, "TENKH") or "")
                if not so and not ten:
                    continue
                if TONG_RE.search(ten) or TONG_RE.search(str(so or "")):
                    continue
                ngay = as_date(g(r, "NGAY_HD"))
                if nam and (not ngay or ngay.year != nam) and ngay is None:
                    # xlrd date
                    raw = g(r, "NGAY_HD")
                    if isinstance(raw, float):
                        try:
                            ngay = datetime(*xlrd.xldate_as_tuple(raw, book.datemode)).date()
                        except Exception:
                            ngay = None
                if ngay and not in_ky(ngay):
                    continue
                y = ngay.year if ngay else nam
                if y not in YEARS:
                    continue
                mst = mst_chuan(g(r, "MS_DN", "MST"))
                ds = nnum(g(r, "TTVND"), 0) or 0
                thue = nnum(g(r, "THUEVND"), 0) or 0
                ts = g(r, "TS_GTGT")
                out.append(
                    {
                        "Chieu": chieu,
                        "Nam": y,
                        "Quy": quy_of(ngay) or quy,
                        "Thang": ngay.month if ngay else None,
                        "Ky_hieu_HD": "",
                        "So_HD": str(so or "").strip(),
                        "Ngay_HD": ngay,
                        "MST_goc": str(g(r, "MS_DN") or ""),
                        "MST_chuan": mst,
                        "Ten_DoiTac": ten,
                        "Mat_hang": g(r, "MATHANG") or "",
                        "GT_chua_thue": ds,
                        "Thue_suat": nnum(ts, None),
                        "Thue_GTGT": thue,
                        "Trang_thai_DC": "",
                        "File_nguon": p.name,
                    }
                )
    return out


# ---------------------------------------------------------------------------
# 511 / 632 / 111 / kho
# ---------------------------------------------------------------------------
def _pick_ledger_sheet(wb):
    if "SHVND" in wb.sheetnames:
        return wb["SHVND"]
    for n in wb.sheetnames:
        if n.upper() not in {"SHN", "SHUSD"}:
            return wb[n]
    return wb[wb.sheetnames[0]]


def parse_ledger(path: Path, tk: str) -> list[dict]:
    """Đọc sổ chi tiết MISA S03b: SHVND (Nợ/Có VND) hoặc IN (Nợ/Có sát TK ĐU)."""
    try:
        wb = load_workbook(lp(path), read_only=True, data_only=True)
    except Exception:
        return []
    ws = _pick_ledger_sheet(wb)
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdr_i = None
    for i, row in enumerate(rows[:40]):
        joined = " ".join(str(c) for c in (row or [])[:12] if c)
        if "Ngày" in joined and "Chứng từ" in joined:
            hdr_i = i
            break
    if hdr_i is None:
        return []
    sub = list(rows[hdr_i + 1]) if hdr_i + 1 < len(rows) else []
    nos = [j for j, h in enumerate(sub) if str(h or "").strip().lower() in {"nợ", "no"}]
    cos = [j for j, h in enumerate(sub) if str(h or "").strip().lower() in {"có", "co"}]
    dus = [
        j
        for j, h in enumerate(sub)
        if str(h or "").strip().upper() in {"DU", "ĐU", "TK ĐU", "TK DU"}
    ]
    if nos and cos:
        i_no, i_co = nos[-1], cos[-1]
    else:
        i_no, i_co = (7, 8) if max(len(sub), 8) >= 9 else (5, 6)
    i_tk = dus[-1] if dus else (i_no - 1 if i_no else 4)
    out = []
    for row in rows[hdr_i + 2 :]:
        if not row:
            continue
        dg = str(row[3] or "") if len(row) > 3 else ""
        if TONG_RE.search(dg) or "dư" in dg.lower() or "DU DAU" in dg.upper() or "DƯ" in dg.upper():
            continue
        ngay = as_date(row[0] if len(row) > 0 else None) or as_date(row[2] if len(row) > 2 else None)
        if not ngay or not in_ky(ngay):
            continue
        so_ct = str(row[1] or "")
        ngay_ct = as_date(row[2] if len(row) > 2 else None) or ngay
        tk_du = str(row[i_tk] or "") if i_tk is not None and len(row) > i_tk else ""
        ps_no = nnum(row[i_no] if len(row) > i_no else 0, 0) or 0
        ps_co = nnum(row[i_co] if len(row) > i_co else 0, 0) or 0
        if ps_no == 0 and ps_co == 0:
            continue
        so_hd = hd_from_text(so_ct) or hd_from_text(dg)
        rec = {
            "Nam": ngay_ct.year,
            "Thang": ngay_ct.month,
            "Ngay_CT": ngay_ct,
            "So_CT": so_ct,
            "TK": tk,
            "TK_doi_ung": tk_du,
            "Ma_hang": "",
            "Dien_giai": dg,
            "MST_doi_tuong": "",
            "So_luong": None,
            "PS_Co": ps_co,
            "PS_No": ps_no,
            "Ky_hieu_HD_LQ": "",
            "So_HD_LQ": so_hd,
            "File_nguon": path.name,
        }
        if tk == "111":
            if ps_co <= 0:
                continue
            out.append(
                {
                    "Ngay_CT": ngay_ct,
                    "Nam": ngay_ct.year,
                    "So_phieu_chi": so_ct,
                    "TK_doi_ung": tk_du,
                    "MST_doi_tuong": "",
                    "Ten_doi_tuong": "",
                    "Dien_giai": dg,
                    "So_tien": ps_co,
                    "Nhom_nghiep_vu": "",
                    "Co_HD_hop_le": "Chưa rõ",
                    "Danh_gia": "",
                    "File_nguon": path.name,
                }
            )
        else:
            out.append(rec)
    return out


def parse_so_511(path: Path, nam_hint=None) -> list[dict]:
    return parse_ledger(path, "511")


def parse_so_632(path: Path) -> list[dict]:
    return parse_ledger(path, "632")


def parse_so_111(path: Path) -> list[dict]:
    return parse_ledger(path, "111")


def match_511_632(r511: list[dict], r632: list[dict]):
    by_key = defaultdict(list)
    for r in r632:
        k1 = (so_hd_norm(r.get("So_HD_LQ")), r.get("Ngay_CT"))
        by_key[k1].append(r)
    used = set()
    matched = []
    unmatched_511 = []
    for a in r511:
        if (a.get("PS_Co") or 0) <= 0:
            continue
        k1 = (so_hd_norm(a.get("So_HD_LQ")), a.get("Ngay_CT"))
        hit = None
        for b in by_key.get(k1, []):
            if id(b) not in used:
                hit = b
                break
        if hit is None:
            # key 2: HD only
            hd = so_hd_norm(a.get("So_HD_LQ"))
            for b in r632:
                if id(b) in used:
                    continue
                if so_hd_norm(b.get("So_HD_LQ")) == hd and hd not in {"", "0"}:
                    hit = b
                    break
        if hit is None:
            unmatched_511.append(a)
            lai = a.get("PS_Co") or 0
            matched.append({**a, "PS_632": None, "Lai_gop": lai, "Trang_thai": "KHÔNG_CÓ_632", "P_le": lai})
        else:
            used.add(id(hit))
            gv = hit.get("PS_No") or 0
            dt = a.get("PS_Co") or 0
            lai = dt - gv
            st = "Khớp HĐ + ngày CT + mặt hàng" if k1 in by_key else "Khớp HĐ + mặt hàng (khác ngày CT)"
            matched.append({**a, "PS_632": gv, "Lai_gop": lai, "Trang_thai": st, "P_le": lai})
    unmatched_632 = [b for b in r632 if id(b) not in used and (b.get("PS_No") or 0) > 0]
    return matched, unmatched_511, unmatched_632


# ---------------------------------------------------------------------------
# CVCBRR / CA / danh ba
# ---------------------------------------------------------------------------
def load_f0_f1():
    f0, f1 = {}, {}
    wb = load_workbook(lp(CVC), read_only=True, data_only=True)
    ws = wb["F0"]
    for i, row in enumerate(ws.iter_rows(min_row=8, values_only=True), 8):
        mst = mst_chuan(row[1] if row else "")
        if not mst:
            continue
        f0[mst_digits(mst)] = {
            "mst": mst,
            "ten": row[2],
            "so_cv": row[3],
            "ngay_cv": as_date(row[4]),
            "dv": row[5],
            "nd": row[6],
            "nhom": row[8],
        }
    ws = wb["F1"]
    for i, row in enumerate(ws.iter_rows(min_row=8, values_only=True), 8):
        mst_mua = mst_chuan(row[8] if row and len(row) > 8 else "")
        if not mst_mua:
            continue
        f1.setdefault(mst_digits(mst_mua), []).append(
            {
                "mst_ban": mst_chuan(row[1]),
                "ten_ban": row[2],
                "so_cv": row[3],
                "ten_mua": row[7],
                "mst_mua": mst_mua,
            }
        )
    wb.close()
    return f0, f1


def load_ca_set():
    s = set()
    if not CA.exists():
        return s
    try:
        wb = load_workbook(lp(CA), read_only=True, data_only=True)
    except Exception:
        return s
    for name in wb.sheetnames:
        if name.upper() in {"DASHBOARD", "DANH_MUC"}:
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        idx = None
        start = 0
        for i, row in enumerate(rows[:20]):
            for j, h in enumerate(row or []):
                hs = str(h or "").strip().lower()
                if hs in {"mst", "mã số thuế", "ma so thue"} or hs == "mst":
                    idx = j
                    start = i + 1
                    break
            if idx is not None:
                break
        if idx is None:
            continue
        for rec in rows[start:]:
            if rec and idx < len(rec) and rec[idx]:
                d = mst_digits(rec[idx])
                if len(d) >= 10:
                    s.add(d[:10])
    wb.close()
    return s


def load_dkkd():
    m = {}
    if not DANHBA.exists():
        return m
    for p in DANHBA.glob("*.xlsx"):
        try:
            wb = load_workbook(lp(p), read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        hdr = None
        start = 0
        for i, row in enumerate(rows[:15]):
            joined = " ".join(str(c).lower() for c in row if c)
            if "mst" in joined:
                hdr = [str(c or "").strip().lower() for c in row]
                start = i + 1
                break
        if not hdr:
            continue

        def col(*cands):
            for c in cands:
                for j, h in enumerate(hdr):
                    if c in h:
                        return j
            return None

        i_mst = col("mst", "mã số thuế", "ma so thue")
        i_ten = col("tên", "ten")
        i_tt = col("trạng", "trang thai", "tình trạng", "tinh trang", "đkkd")
        if i_mst is None:
            continue
        for row in rows[start:]:
            if not row or not row[i_mst]:
                continue
            mst = mst_digits(row[i_mst])
            m[mst] = {
                "ten": row[i_ten] if i_ten is not None else "",
                "tt": row[i_tt] if i_tt is not None else "",
            }
    return m


# ---------------------------------------------------------------------------
# ANALYTICS
# ---------------------------------------------------------------------------
def sum_where(rows, pred, field):
    return sum((nnum(r.get(field), 0) or 0) for r in rows if pred(r))


def gtgt_by_ky(tks):
    d = {}
    for tk in tks:
        if tk.loai != "gtgt":
            continue
        if tk.nam not in YEARS:
            continue
        d[tk.ky] = tk
    return d


def bctc_map(tks):
    d = {}
    for tk in tks:
        if tk.loai != "bctc":
            continue
        if tk.nam not in YEARS:
            continue
        d[tk.nam] = tk
    return d


def tndn_map(tks):
    d = {}
    for tk in tks:
        if tk.loai != "tndn":
            continue
        if tk.nam not in YEARS:
            continue
        d[tk.nam] = tk
    return d


def tncn_map(tks):
    d = {}
    for tk in tks:
        if tk.loai != "tncn":
            continue
        if tk.nam not in YEARS:
            continue
        # prefer annual QTT
        if tk.ky and "Q" in tk.ky.upper() and len(tk.ky) <= 8:
            d.setdefault(("Q", tk.nam, tk.ky), tk)
        else:
            d[("Y", tk.nam)] = tk
    return d


# ---------------------------------------------------------------------------
# FILL KQ FILES
# ---------------------------------------------------------------------------
def fill_manifest(ws, files, ma_filter, loai_nguon):
    clear_example(ws, 2)
    rows = []
    n = 0
    for f in files:
        if f["ma_in"] not in ma_filter:
            continue
        n += 1
        ky_tu = ky_den = ""
        m = re.search(r"(20\d{2})", f["name"])
        if m:
            ky_tu = ky_den = m.group(1)
        rows.append(
            {
                "Ma_nguon": f"N{n:03d}",
                "Ma_IN": f["ma_in"],
                "Loai_nguon": loai_nguon,
                "Dinh_dang": "Cau_truc" if f["ext"] in {".xml", ".xlsx", ".xls"} else f["ext"],
                "Ngay_nhan": f["mtime"].date(),
                "Ky_tu": ky_tu,
                "Ky_den": ky_den,
                "Ma_bam": f["sha"][:16] if f["sha"] else "",
                "Ban_bo_sung": "Không",
                "La_ban_moi_nhat": "Có",
                "So_dong_doc": "",
                "So_dong_loi": 0,
                "Xu_ly_nguon": "Giữ nguyên",
                "Ghi_chu": "",
                "File_nguon": f["rel"],
            }
        )
    write_rows(ws, 2, rows)
    apply_font(ws)
    autofit(ws)
    return len(rows)


def muc_in(has, n_rows, pdf=False):
    if pdf:
        return "PDF chưa trích"
    if not has:
        return "Chưa nhận"
    if n_rows == 0:
        return "Mỏng"
    if n_rows < 12:
        return "Mỏng"
    return "Đủ"


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    print("SCAN files...")
    files = scan_all()
    print(" files kept", len(files))

    print("XML...")
    tks, xml_skip = parse_all_xml()
    for tk in tks:
        if tk.loai != "tndn":
            continue
        try:
            root = ET.parse(tk.file).getroot()
        except Exception:
            continue
        deep = xt._deep_ct(root)
        for k, v in deep.items():
            if k not in tk.chi_tieu or not tk.chi_tieu[k]:
                tk.chi_tieu[k] = v
    print("  to_khai", len(tks), "skip", len(xml_skip))
    gtgt = [t for t in tks if t.loai == "gtgt" and t.nam in YEARS]
    bctc = [t for t in tks if t.loai == "bctc" and t.nam in YEARS]
    tndn = [t for t in tks if t.loai == "tndn" and t.nam in YEARS]
    tncn = [t for t in tks if t.loai == "tncn" and t.nam in YEARS]
    print("  gtgt", len(gtgt), "bctc", len(bctc), "tndn", len(tndn), "tncn", len(tncn))

    hddt_path = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT.xlsx"
    print("HDDT...")
    in01 = parse_hddt(hddt_path) if hddt_path.exists() else []
    print("  hddt", len(in01))

    bk_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
    print("Bang ke...")
    in02 = parse_bang_ke_folder(bk_dir)
    print("  bk", len(in02))

    print("CVCBRR...")
    f0, f1 = load_f0_f1()
    ca_set = load_ca_set()
    dkkd = load_dkkd()
    print("  F0", len(f0), "F1 keys", len(f1), "CA", len(ca_set), "DKKD", len(dkkd))
    nnt_f1 = f1.get(MST, [])
    nnt_on_ca = MST in ca_set or MST[:10] in ca_set

    print("511/632/111...")
    so_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.1. So ke toan tong hop va chi tiet"
    year_dirs = [p for p in so_dir.iterdir() if p.is_dir() and "so kt" in p.name.lower()]
    scan_dirs = year_dirs if year_dirs else [so_dir]
    r511, r632, r111 = [], [], []
    seen_so = set()
    for d in scan_dirs:
        for p in d.glob("*.xlsx"):
            h = sha256(p)
            if h in seen_so:
                continue
            seen_so.add(h)
            n = p.name.lower().replace(" ", "")
            if "632" in n:
                r632.extend(parse_so_632(p))
            elif "111" in n and "112" not in n:
                r111.extend(parse_so_111(p))
            elif "511" in n or re.search(r"515", n):
                r511.extend(parse_so_511(p))
    print("  511", len(r511), "632", len(r632), "111", len(r111), "dirs", [d.name for d in scan_dirs])
    matched, un511, un632 = match_511_632(r511, r632)

    # kho: look for NXT files - store file-level only if parse is thin
    kho_files = list((ROOT / "2. HO SO NNT CUNG CAP" / "2.2. Kho - Chi phi - Gia thanh").rglob("*.xlsx"))

    # IN03 rows
    in03_rows = []
    for tk in sorted(gtgt, key=lambda x: (x.nam, x.ky)):
        ct = tk.chi_tieu
        in03_rows.append(
            {
                "Nam": tk.nam,
                "Ky": tk.ky,
                "Loai_TK": "Bổ sung" if tk.loai_tkhai == "B" else "Chính thức",
                "Lan_BS": tk.so_lan,
                "La_ban_hieu_luc": "Có",
                "ct22": ct.get("22"),
                "ct23": ct.get("23"),
                "ct24": ct.get("24"),
                "ct25": ct.get("25"),
                "ct26": ct.get("26"),
                "ct32": ct.get("32"),
                "ct33": ct.get("33"),
                "ct34": ct.get("34"),
                "ct35": ct.get("35"),
                "ct36": ct.get("36"),
                "ct37": ct.get("37"),
                "ct40": ct.get("40"),
                "ct43": ct.get("43"),
                "File_nguon": Path(tk.file).name,
            }
        )

    in04_rows = []
    labels = {
        "01": "Doanh thu ban hang va cung cap dich vu",
        "11": "Gia von hang ban (TT133 ct11)",
        "02": "Cac khoan giam tru DT (TT133 ct02)",
        "10": "Doanh thu thuan (TT133 ct10)",
        "20": "Loi nhuan gop (TT133 ct20)",
        "21": "Doanh thu HDTD",
        "22": "Chi phi tai chinh",
        "23": "Chi phi lai vay",
        "24": "Chi phi QLDN",
        "50": "Loi nhuan truoc thue",
        "51": "Chi phi thue TNDN",
        "110": "Tien",
        "131": "Phai thu KH",
        "140": "Hang ton kho",
        "211": "TSCD",
        "311": "Phai tra nguoi ban",
        "411": "Von CSH",
    }
    for tk in bctc:
        che_do = "TT133"
        # detect
        src = Path(tk.file).name.upper()
        if "B01B_DNN" in src or "DNN" in src:
            che_do = "TT133"
        elif "B01-DN" in src or "B01_DN" in src:
            che_do = "TT200"
        for ma, ten in labels.items():
            val = tk.bctc_kqkd.get(ma)
            if val is None:
                val = tk.bctc_cdkt_cuoi.get(ma)
            if val is None:
                continue
            bao = "KQKD" if ma in {"01", "02", "10", "11", "20", "21", "22", "23", "24", "50", "51"} else "CDKT"
            in04_rows.append(
                {
                    "Nam": tk.nam,
                    "Bao_cao": bao,
                    "Ma_so": ma,
                    "Chi_tieu": ten,
                    "So_nam_nay": val,
                    "So_nam_truoc": tk.bctc_cdkt_dau.get(ma) if bao == "CDKT" else None,
                    "Che_do_KT": che_do,
                    "File_nguon": Path(tk.file).name,
                }
            )
        for tk2 in tndn:
            if tk2.nam != tk.nam:
                continue
            in04_rows.append(
                {
                    "Nam": tk.nam,
                    "Bao_cao": "03/TNDN",
                    "Ma_so": "A1",
                    "Chi_tieu": "LN ke toan truoc thue (03/TNDN A1)",
                    "So_nam_nay": tk2.chi_tieu.get("A1") or tk2.chi_tieu.get("ctA1"),
                    "So_nam_truoc": None,
                    "Che_do_KT": che_do,
                    "File_nguon": Path(tk2.file).name,
                }
            )
            in04_rows.append(
                {
                    "Nam": tk.nam,
                    "Bao_cao": "03/TNDN",
                    "Ma_so": "C9",
                    "Chi_tieu": "Thue TNDN phai nop (C9)",
                    "So_nam_nay": tk2.chi_tieu.get("C9") or tk2.chi_tieu.get("ctC9"),
                    "So_nam_truoc": None,
                    "Che_do_KT": che_do,
                    "File_nguon": Path(tk2.file).name,
                }
            )
    bctc_nams = {r["Nam"] for r in in04_rows if r["Bao_cao"] != "03/TNDN"}
    for tk2 in tndn:
        if tk2.nam not in YEARS or tk2.nam in bctc_nams:
            continue
        in04_rows.append(
            {
                "Nam": tk2.nam,
                "Bao_cao": "03/TNDN",
                "Ma_so": "A1",
                "Chi_tieu": "LN ke toan truoc thue (03/TNDN A1)",
                "So_nam_nay": tk2.chi_tieu.get("A1"),
                "So_nam_truoc": None,
                "Che_do_KT": "",
                "File_nguon": Path(tk2.file).name,
            }
        )

    in05_rows = []
    for tk in tncn:
        if tk.nam not in YEARS:
            continue
        loai = "QTT" if "Q" not in (tk.ky or "").upper() or len(tk.ky) > 6 else "TAM"
        # annual files often ky = 2024
        if re.fullmatch(r"20\d{2}", tk.ky or ""):
            loai = "QTT"
        in05_rows.append(
            {
                "Loai_dong": "QTT" if loai == "QTT" else "NLD",
                "Nam": tk.nam,
                "Ho_ten": "",
                "CCCD": "",
                "MST_CN": "",
                "Nguon_BK": "05/QTT-TNCN" if loai == "QTT" else "05-1",
                "Loai_thu_nhap": tk.ky,
                "Tong_TNCT": tk.chi_tieu.get("21") or tk.chi_tieu.get("26") or tk.chi_tieu.get("20"),
                "Thue_da_KT": tk.chi_tieu.get("36") or tk.chi_tieu.get("40") or tk.chi_tieu.get("35"),
                "Co_trich_10": "",
                "Mau_TK": tk.ma_tkhai,
                "File_nguon": Path(tk.file).name,
            }
        )
        in05_rows.extend(parse_tncn_nld(Path(tk.file), tk.nam))

    # TMS PTDN as IN14 narrative rows
    in14_rows = []
    tms_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.3. Nghia vu nop thue & Thong bao"
    for p in tms_dir.glob("*.xlsx"):
        in14_rows.append(
            {
                "MST": MST,
                "Loai_ban_ghi": "BCTC_TMS" if "PTDN01" in p.name else "KQKD_TMS" if "PTDN02" in p.name else "TMS",
                "Sac_thue": "",
                "Ky_thue": KY,
                "Nam": "",
                "Noi_dung": p.name,
                "Nguon_TMS": p.name,
                "File_nguon": p.name,
            }
        )

    # NCC from IN02 + IN01
    ncc_agg = defaultdict(lambda: {"ten": "", "gt": 0.0, "thue": 0.0, "so_hd": 0, "src": set(), "rows": []})
    for r in in02:
        if r["Chieu"] != "MUA":
            continue
        k = mst_digits(r["MST_chuan"])
        if len(k) < 10:
            continue
        a = ncc_agg[k]
        a["ten"] = a["ten"] or r["Ten_DoiTac"]
        a["gt"] += r["GT_chua_thue"] or 0
        a["thue"] += r["Thue_GTGT"] or 0
        a["so_hd"] += 1
        a["src"].add("Bảng kê")
        a["rows"].append(r)
    for r in in01:
        if r["Chieu"] != "MUA" or r["Tinh_vao_tong"] != "Có":
            continue
        k = mst_digits(r["MST_DoiTac_chuan"])
        if len(k) < 10:
            continue
        a = ncc_agg[k]
        a["ten"] = a["ten"] or r["Ten_NB"]
        a["src"].add("HĐĐT")
        if "Bảng kê" not in a["src"]:
            a["gt"] += r["GT_chua_thue"] or 0
            a["thue"] += r["Thue_GTGT"] or 0
            a["so_hd"] += 1
        a["rows"].append(r)

    def ncc_flag(mst_d):
        nhom = ""
        nguon = ""
        cv = ""
        if mst_d in f0 or mst_d[:10] in f0:
            info = f0.get(mst_d) or f0.get(mst_d[:10])
            nhom = "F0"
            nguon = "CANH_BAO"
            cv = info.get("so_cv") if info else ""
            if info and info.get("nhom") and "Công an" in str(info.get("nhom")):
                nguon = "CA"
        if mst_d[:10] in ca_set:
            if not nhom:
                nhom = "F0"
            nguon = nguon or "CA"
        tt = (dkkd.get(mst_d) or dkkd.get(mst_d[:10]) or {}).get("tt") or ""
        return nhom, nguon, cv, tt

    in07_rows = []
    # NNT itself
    nnt_nhom = "F1" if nnt_f1 else ""
    if nnt_on_ca:
        nnt_nhom = nnt_nhom or "F0"
    in07_rows.append(
        {
            "MST": MST,
            "Ten": TEN,
            "Loai_canh_bao": nnt_nhom or "Không",
            "Nhom_F": nnt_nhom,
            "Nguon_F0": "CANH_BAO" if nnt_f1 else ("CA" if nnt_on_ca else ""),
            "So_CV": nnt_f1[0]["so_cv"] if nnt_f1 else "",
            "Tinh_trang_DKKD": (dkkd.get(MST) or {}).get("tt") or "",
            "Nguon_tra_cuu": "SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx",
            "Ngay_CV": None,
            "Co_quan_BH": "",
            "Noi_dung": "",
            "Nganh_nghe": "",
            "Dia_chi": "",
            "Nguoi_dai_dien": "",
            "Ngay_hieu_luc_TT": None,
            "Ngay_cap_MST": None,
        }
    )
    for k, a in sorted(ncc_agg.items(), key=lambda x: -x[1]["gt"]):
        nhom, nguon, cv, tt = ncc_flag(k)
        info = f0.get(k) or f0.get(k[:10]) or {}
        in07_rows.append(
            {
                "MST": mst_chuan(k),
                "Ten": a["ten"],
                "Loai_canh_bao": nhom or ("ĐKKD" if tt and any(x in str(tt).lower() for x in ["ngừng", "ngung", "đóng", "dong", "bỏ"]) else "Không"),
                "Nhom_F": nhom,
                "Nguon_F0": nguon,
                "So_CV": cv,
                "Tinh_trang_DKKD": tt,
                "Nguon_tra_cuu": "Sổ CVCBRR + danh bạ",
                "Ngay_CV": info.get("ngay_cv"),
                "Co_quan_BH": info.get("dv") or "",
                "Noi_dung": info.get("nd") or "",
                "Nganh_nghe": "",
                "Dia_chi": "",
                "Nguoi_dai_dien": "",
                "Ngay_hieu_luc_TT": None,
                "Ngay_cap_MST": None,
            }
        )

    # totals
    def tot_hd(chieu, nam=None):
        return sum(
            r["GT_chua_thue"]
            for r in in01
            if r["Chieu"] == chieu and r["Tinh_vao_tong"] == "Có" and (nam is None or r["Nam"] == nam)
        )

    def tot_bk(chieu, nam=None):
        return sum(r["GT_chua_thue"] for r in in02 if r["Chieu"] == chieu and (nam is None or r["Nam"] == nam))

    def tot_tk(nam, field):
        s = 0
        n = 0
        for r in in03_rows:
            if r["Nam"] == nam and r.get(field) is not None:
                s += r[field] or 0
                n += 1
        return s if n else None

    def bctc_val(nam, ma):
        for r in in04_rows:
            if r["Nam"] == nam and str(r["Ma_so"]) == str(ma):
                return r["So_nam_nay"]
        return None

    def bctc_gv(nam):
        v = bctc_val(nam, "11")
        if v is not None:
            return v
        return bctc_val(nam, "02")

    # 511/632 yearly
    y511 = {y: sum(r["PS_Co"] or 0 for r in r511 if r["Nam"] == y) for y in YEARS}
    y632 = {y: sum(r["PS_No"] or 0 for r in r632 if r["Nam"] == y) for y in YEARS}
    y_match = {y: [m for m in matched if m.get("Nam") == y] for y in YEARS}

    # cash threshold
    cash_hits = []
    for r in in02:
        if r["Chieu"] != "MUA":
            continue
        d = r["Ngay_HD"]
        gt = (r["GT_chua_thue"] or 0) + (r["Thue_GTGT"] or 0)
        if not d:
            continue
        th = 20_000_000 if d <= date(2025, 6, 30) else 5_000_000
        if gt >= th:
            cash_hits.append({**r, "Nguong": th, "Tong_TT": gt})

    # top5
    tot_mua = sum(a["gt"] for a in ncc_agg.values()) or 1
    top5 = sorted(ncc_agg.items(), key=lambda x: -x[1]["gt"])[:5]
    top5_pct = sum(a["gt"] for _, a in top5) / tot_mua if tot_mua else 0
    f0_ncc = [(k, a) for k, a in ncc_agg.items() if ncc_flag(k)[0] == "F0"]
    f0_gt = sum(a["gt"] for _, a in f0_ncc)
    f0_thue = sum(a["thue"] for _, a in f0_ncc)

    # level
    has = {
        "IN01": len(in01) > 0,
        "IN02": len(in02) > 0,
        "IN03": len(in03_rows) > 0,
        "IN04": len(bctc) > 0,
        "IN05": len(tncn) > 0,
        "IN07": True,
        "IN08": any(f["ma_in"] == "IN08" for f in files),
        "IN09": len(kho_files) > 0,
        "IN10": len(r111) > 0,
        "IN11": False,
        "IN12": False,
        "IN13": len(r511) > 0 and len(r632) > 0,
        "IN14": len(in14_rows) > 0,
    }
    if has["IN01"] and has["IN03"] and has["IN04"]:
        muc = "A"
    if has["IN02"]:
        muc = "B"
    if has["IN08"] and has["IN13"]:
        muc = "C"
    if has["IN09"]:
        muc = "D"
    if has["IN10"] and has["IN11"]:
        muc = "E"
    # we have 111 but no bank -> D
    muc = "D"

    names = {
        "kq11": OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_{VER}.xlsx",
        "kq12": OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_{VER}.xlsx",
        "kq13": OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx",
        "kq21": OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_{VER}.xlsx",
        "kq22": OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_{VER}.xlsx",
        "kq3": OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_{VER}.xlsx",
        "dau": OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}.xlsx",
        "giai": OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}.xlsx",
    }

    print("Copy templates...")
    copy_wb(MAU / "KQ1.1_Kiem_ke_CQT.xlsx", names["kq11"])
    copy_wb(MAU / "KQ1.2_Kiem_ke_ho_so_NNT.xlsx", names["kq12"])
    copy_wb(MAU / "KQ1.3_Danh_gia_ho_so.xlsx", names["kq13"])
    copy_wb(MAU / "KQ2.1_Phan_tich_CQT.xlsx", names["kq21"])
    copy_wb(MAU / "KQ2.2_Phan_tich_ho_so_cty.xlsx", names["kq22"])
    copy_wb(MAU / "KQ3_Ket_luan_danh_gia.xlsx", names["kq3"])

    # ---- KQ1.1 ----
    print("Fill KQ1.1")
    wb = load_workbook(lp(names["kq11"]))
    fill_manifest(wb["IN00_Manifest"], files, {"IN01", "IN03", "IN04", "IN05", "IN14"}, "CQT")
    ws = wb["IN01_HDDT"]
    clear_example(ws, 2)
    write_rows(ws, 2, in01)
    apply_font(ws)
    autofit(ws)
    ws = wb["IN03_ToKhai_GTGT"]
    clear_example(ws, 2)
    write_rows(ws, 2, in03_rows)
    apply_font(ws)
    autofit(ws)
    ws = wb["IN04_BCTC_TNDN"]
    clear_example(ws, 2)
    write_rows(ws, 2, in04_rows)
    apply_font(ws)
    autofit(ws)
    ws = wb["IN05_NLD_TNCN"]
    # keep formulas? overwrite data rows
    write_rows(ws, 2, in05_rows)
    apply_font(ws)
    autofit(ws)
    ws = wb["IN14_TMS_Nghia_vu"]
    clear_example(ws, 2)
    write_rows(ws, 2, in14_rows)
    apply_font(ws)
    autofit(ws)
    wb.save(lp(names["kq11"]))
    wb.close()

    # ---- KQ1.2 ----
    print("Fill KQ1.2")
    wb = load_workbook(lp(names["kq12"]))
    fill_manifest(wb["IN00_Manifest"], files, {"IN02", "IN08", "IN09", "IN10", "IN11", "IN12", "IN13"}, "DN")
    ws = wb["IN02_BangKe"]
    clear_example(ws, 2)
    write_rows(ws, 2, in02)
    apply_font(ws)
    autofit(ws)
    # IN08: yearly NKC file-level
    in08_rows = []
    for p in so_dir.glob("NKC *.xlsx"):
        m = re.search(r"(20\d{2})", p.name)
        nam = int(m.group(1)) if m else None
        if nam not in YEARS:
            continue
        in08_rows.append(
            {
                "Nam": nam,
                "Thang": "",
                "TK": "NKC",
                "PS_No": None,
                "PS_Co": None,
                "Nguon_so": "NKC",
                "File_nguon": p.name,
            }
        )
    ws = wb["IN08_So_KeToan"]
    clear_example(ws, 2)
    write_rows(ws, 2, in08_rows)
    apply_font(ws)
    ws = wb["IN09_Kho_GiaThanh"]
    clear_example(ws, 2)
    kho_rows = []
    for p in kho_files:
        m = re.search(r"(20\d{2})", str(p))
        nam = int(m.group(1)) if m else None
        kho_rows.append({"Nam": nam, "Nhom_hang": p.stem[:40], "File_nguon": p.name})
    write_rows(ws, 2, kho_rows)
    apply_font(ws)
    ws = wb["IN10_Chi_tien_mat"]
    clear_example(ws, 2)
    write_rows(ws, 2, r111)
    apply_font(ws)
    autofit(ws)
    ws = wb["IN11_Sao_ke_NH"]
    clear_example(ws, 2)
    put(ws, 3, 1, None)
    put(ws, 3, 14, "Chưa nhận - thư mục 2.3 trống")
    ws = wb["IN12_GDLK"]
    clear_example(ws, 2)
    put(ws, 3, 9, "Chưa nhận phụ lục GDLK")
    ws = wb["IN13_So_511_632"]
    write_rows(ws, 2, r511 + r632)
    apply_font(ws)
    autofit(ws)
    wb.save(lp(names["kq12"]))
    wb.close()

    # ---- KQ1.3 ----
    print("Fill KQ1.3")
    wb = load_workbook(lp(names["kq13"]))
    ws = wb["KQ1_Kiem_ke"]
    put(ws, 2, 2, MST, bold=True)
    put(ws, 2, 8, TEN, bold=True)
    put(ws, 3, 2, KY, bold=True)
    put(ws, 3, 8, muc, bold=True)
    put(ws, 5, 2, VER)
    put(ws, 5, 8, "1.4")
    put(ws, 6, 2, TODAY, dt=True)
    put(ws, 6, 8, "Một phần")
    put(ws, 7, 2, "KQ2.1, KQ2.2, KQ3")
    put(ws, 7, 8, "Kỳ suy ra 2021-2025 từ sổ KT + bảng kê + XML (K12). Thiếu sao kê NH (IN11) và PL GDLK (IN12).")

    inv = [
        ("IN01", has["IN01"], len(in01), "CQT", "Excel", "2022-2025", "HDDT.xlsx"),
        ("IN02", has["IN02"], len(in02), "DN", "Excel .xls", "2021-2025", "20 kỳ Q1-Q4"),
        ("IN03", has["IN03"], len(in03_rows), "CQT", "XML", "2021-2025", "Tờ khai GTGT"),
        ("IN04", has["IN04"], len(in04_rows), "CQT", "XML", "2021-2025", "BCTC+TNDN"),
        ("IN05", has["IN05"], len(in05_rows), "CQT", "XML", "2021-2025", "TNCN"),
        ("IN07", True, len(in07_rows), "CQT", "Sổ CVCBRR", "2026-11-09", "F0/F1 + CA + danh bạ"),
        ("IN08", has["IN08"], len(in08_rows), "DN", "Excel", "2021-2025", "NKC + sổ cái"),
        ("IN09", has["IN09"], len(kho_rows), "DN", "Excel", "2021-2025", "NXT/152/154/155"),
        ("IN10", has["IN10"], len(r111), "DN", "Excel", "2021-2025", "TK 1111"),
        ("IN11", False, 0, "DN", "", "", "2.3 trống"),
        ("IN12", False, 0, "DN", "", "", "Chưa nhận"),
        ("IN13", has["IN13"], len(r511) + len(r632), "DN", "Excel", "2021-2025", "511+632"),
        ("IN14", has["IN14"], len(in14_rows), "CQT", "Excel TMS", "2016-2025", "PTDN01/02"),
    ]
    for i, (ma, ok, n, ng, dd, ky, ghi) in enumerate(inv):
        r = 10 + i
        muc_d = muc_in(ok, n)
        put(ws, r, 7, ky)
        put(ws, r, 9, n, num=True)
        put(ws, r, 12, muc_d, bold=True)
        put(ws, r, 13, ghi)
        put(ws, r, 6, dd)
        put(ws, r, 5, ng)

    du = sum(1 for x in inv if muc_in(x[1], x[2]) == "Đủ")
    mong = sum(1 for x in inv if muc_in(x[1], x[2]) == "Mỏng")
    thieu = sum(1 for x in inv if muc_in(x[1], x[2]) in {"Thiếu", "Chưa nhận"})
    put(ws, 4, 2, du, num=True)
    put(ws, 4, 3, mong, num=True)
    put(ws, 4, 4, thieu, num=True)

    ws = wb["KQ1_Kha_nang_phan_tich"]
    put(ws, 2, 2, MST)
    put(ws, 2, 6, KY)
    put(ws, 4, 2, VER)
    put(ws, 4, 6, "1.4")
    put(ws, 5, 2, TODAY, dt=True)

    def kha_nang(ma, need):
        need_ins = re.findall(r"IN\d{2}", need or "")
        missing = [i for i in need_ins if not has.get(i, False) and i not in {"IN07"}]
        # IN07 we built
        if "IN11" in need_ins and not has["IN11"]:
            missing.append("IN11")
        if "IN12" in need_ins and not has["IN12"]:
            missing.append("IN12")
        missing = sorted(set(missing))
        if not missing:
            return "Chạy được ngay", "", ""
        if any(has.get(i, False) or i == "IN07" for i in need_ins) and missing:
            return "Chạy một phần", ",".join(missing), ",".join(missing)
        return "Bị chặn", ",".join(missing), ",".join(missing)

    n_now = n_part = n_block = 0
    for r in range(10, 90):
        ma = ws.cell(r, 1).value
        if not ma:
            continue
        need = str(ws.cell(r, 4).value or "")
        kn, thieu_f, thieu_in = kha_nang(str(ma), need)
        put(ws, r, 8, kn)
        put(ws, r, 9, thieu_f)
        put(ws, r, 10, thieu_in)
        if kn == "Chạy được ngay":
            n_now += 1
        elif kn == "Chạy một phần":
            n_part += 1
        else:
            n_block += 1
    put(ws, 3, 2, n_now, num=True)
    put(ws, 3, 6, n_part, num=True)
    put(ws, 3, 7, n_block, num=True)

    ws = wb["KQ1_De_nghi_bo_sung"]
    put(ws, 2, 2, MST)
    put(ws, 2, 6, KY)
    put(ws, 3, 2, VER)
    put(ws, 3, 6, TODAY, dt=True)
    de_nghi = [
        (1, "Sao kê ngân hàng Excel + UNC các món >= 20tr (den 30/06/2025) va >= 5tr (tu 01/07/2025)", "IN11", "2.3", "File Excel dong tien, khong chi PDF scan", 5, "GT16, TT02, TT03, TT05, TT06"),
        (2, "Phu luc giao dich lien ket (neu co ben lien ket)", "IN12", "2.x", "PL 01 NĐ 132", 2, "LK01, LK02"),
        (3, "Hop dong kinh te / bien ban nghiem thu cho NCC trong diem va F0", "IN09", "2.2/3.4", "Gan mat xich 2 chuoi bang chung", 0, "K11"),
    ]
    for i, row in enumerate(de_nghi):
        for c, v in enumerate(row, 1):
            put(ws, 10 + i, c, v, num=isinstance(v, (int, float)))
    apply_font(wb["KQ1_Kiem_ke"])
    apply_font(wb["KQ1_Kha_nang_phan_tich"])
    wb.save(lp(names["kq13"]))
    wb.close()

    # ---- KQ2.1 ----
    print("Fill KQ2.1")
    wb = load_workbook(lp(names["kq21"]))
    ws = wb["KQ2_Doi_chieu_GTGT"]
    put(ws, 2, 1, f"Người nộp thuế: {TEN} - MST: {MST} - Kỳ: {KY}", bold=True)
    # year blocks: 2016 r9, 14 rows/year
    def year_start(y):
        return 9 + (y - 2016) * 14

    period_rows = ["Q1", "Q2", "Q3", "Q4", "T05", "T06", "T07", "T08", "T09", "T10", "T11", "T12"]
    for y in YEARS:
        base = year_start(y)
        for tk in gtgt:
            if tk.nam != y:
                continue
            ky = (tk.ky or "").upper().replace(" ", "")
            # Q1/2021 or 01/2021 or 2021
            idx = None
            mq = re.search(r"Q([1-4])", ky)
            if mq:
                idx = int(mq.group(1)) - 1
            else:
                mt = re.search(r"(?:^|/)(0?[1-9]|1[0-2])(?:/|$)", ky)
                if mt:
                    mth = int(mt.group(1))
                    if mth <= 4:
                        idx = mth - 1  # unlikely monthly
                    else:
                        # T05 is index 4
                        idx = mth - 1
            if idx is None:
                continue
            rr = base + 1 + idx  # Q1 = base+1
            ct = tk.chi_tieu
            dt = ct.get("34") if ct.get("34") is not None else ct.get("32")
            th = ct.get("35") if ct.get("35") is not None else ct.get("33")
            put(ws, rr, 5, dt, num=True)  # 8-10% col E
            put(ws, rr, 9, th, num=True)  # thue 8-10% col I
            put(ws, rr, 11, ct.get("23"), num=True)
            put(ws, rr, 12, ct.get("24"), num=True)
            put(ws, rr, 17, ct.get("40"), num=True)
            put(ws, rr, 18, ct.get("43"), num=True)

    # Phần II comment at end
    last = ws.max_row + 2
    put(ws, last, 1, "Nhận xét / nguyên nhân trọng yếu", bold=True, fill=NOTE)
    notes = []
    for y in YEARS:
        tk34 = tot_tk(y, "ct34")
        if tk34 is None:
            tk34 = tot_tk(y, "ct32")
        hd = tot_hd("BAN", y)
        bk = tot_bk("BAN", y)
        if tk34 is None:
            notes.append(f"{y}: tờ khai GTGT thiếu chỉ tiêu [34]/[32] - không điền 0.")
            continue
        d_hd = (tk34 or 0) - hd
        d_bk = (tk34 or 0) - bk
        notes.append(
            f"{y}: TK[34]={tk34:,.0f} | HĐĐT bán={hd:,.0f} lệch {d_hd:,.0f} | BK bán={bk:,.0f} lệch {d_bk:,.0f}."
        )
    for i, line in enumerate(notes):
        put(ws, last + 1 + i, 1, line, wrap=True, fill=NOTE)
        ws.merge_cells(start_row=last + 1 + i, start_column=1, end_row=last + 1 + i, end_column=10)

    # IN07
    ws = wb["IN07_NCC_CanhBao"]
    clear_example(ws, 2)
    write_rows(ws, 2, in07_rows)
    apply_font(ws)
    autofit(ws)

    # NCC mua vao - 100% invoices of valid mua
    ws = wb["KQ2_NCC_mua_vao"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)
    # find header row 7
    hdr_row = 7
    # wipe from 8
    if ws.max_row > 8:
        ws.delete_rows(8, ws.max_row - 7)
    r = 8
    inv_rows = []
    for rec in in02:
        if rec["Chieu"] != "MUA":
            continue
        inv_rows.append(rec)
    # also HDDT-only
    bk_keys = {(so_hd_norm(x["So_HD"]), mst_digits(x["MST_chuan"])) for x in in02 if x["Chieu"] == "MUA"}
    for rec in in01:
        if rec["Chieu"] != "MUA" or rec["Tinh_vao_tong"] != "Có":
            continue
        k = (so_hd_norm(rec["So_HD"]), mst_digits(rec["MST_DoiTac_chuan"]))
        if k not in bk_keys:
            inv_rows.append(
                {
                    "MST_chuan": rec["MST_DoiTac_chuan"],
                    "Ten_DoiTac": rec["Ten_NB"],
                    "So_HD": rec["So_HD"],
                    "Ngay_HD": rec["Ngay_HD"],
                    "Mat_hang": rec.get("Mat_hang"),
                    "GT_chua_thue": rec["GT_chua_thue"],
                    "Thue_GTGT": rec["Thue_GTGT"],
                    "_src": "HĐĐT",
                }
            )
        else:
            rec["_src"] = "Chung"
    for rec in in02:
        if rec["Chieu"] != "MUA":
            continue
        rec["_src"] = rec.get("_src") or "Bảng kê"
        k = (so_hd_norm(rec["So_HD"]), mst_digits(rec["MST_chuan"]))
        if any(
            so_hd_norm(h["So_HD"]) == k[0] and mst_digits(h["MST_DoiTac_chuan"]) == k[1]
            for h in in01
            if h["Chieu"] == "MUA"
        ):
            rec["_src"] = "Chung"

    for rec in inv_rows:
        mst = rec.get("MST_chuan") or rec.get("MST_DoiTac_chuan")
        nhom, nguon, cv, tt = ncc_flag(mst_digits(mst))
        src = rec.get("_src") or "Bảng kê"
        put(ws, r, 1, r - 7, num=True)
        put(ws, r, 2, mst_chuan(mst))
        put(ws, r, 3, rec.get("Ten_DoiTac") or rec.get("Ten_NB"))
        put(ws, r, 4, rec.get("So_HD"))
        put(ws, r, 5, rec.get("Ngay_HD"), dt=True, center=True)
        put(ws, r, 6, rec.get("Mat_hang"))
        put(ws, r, 7, rec.get("GT_chua_thue") or 0, num=True)
        put(ws, r, 8, rec.get("Thue_GTGT") or 0, num=True)
        put(ws, r, 9, src)
        put(ws, r, 10, f"Có - {cv}" if cv else ("Có" if nhom else "Không"))
        put(ws, r, 11, nhom or "Không")
        put(ws, r, 12, tt)
        if nhom == "F0":
            for c in range(1, 13):
                ws.cell(r, c).fill = RED
                ws.cell(r, c).font = RED_F
        r += 1
    # comment HD04
    put(ws, r + 1, 1, "Nhận xét module (3.4)", bold=True, fill=NOTE)
    put(
        ws,
        r + 2,
        1,
        f"HD04 Top-5 NCC / tong mua = {top5_pct:.2%} (nguong 50%). So NCC={len(ncc_agg)}. F0={len(f0_ncc)} doanh so {f0_gt:,.0f} thue {f0_thue:,.0f}. NNT tren F1: {'Có' if nnt_f1 else 'Không'}; tren CA: {'Có' if nnt_on_ca else 'Không'}.",
        wrap=True,
        fill=NOTE,
    )
    ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=12)

    # TNCN
    ws = wb["KQ2_TNCN_LD2"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)
    rr = 8
    for y in YEARS:
        qtt = [t for t in tncn if t.nam == y and re.fullmatch(r"20\d{2}", t.ky or "")]
        if not qtt:
            qtt = [t for t in tncn if t.nam == y]
        nld_y = [x for x in in05_rows if x.get("Nam") == y and x.get("Loai_dong") == "NLD"]
        cccd = [x.get("CCCD") for x in nld_y if x.get("CCCD")]
        dup = len(cccd) - len(set(cccd)) if cccd else 0
        put(ws, rr, 1, y, num=True)
        put(ws, rr, 2, len(nld_y) or (qtt[0].chi_tieu.get("21") if qtt else None), num=True)
        put(ws, rr, 4, qtt[0].chi_tieu.get("26") or qtt[0].chi_tieu.get("21") if qtt else None, num=True)
        put(ws, rr, 5, qtt[0].chi_tieu.get("36") if qtt else None, num=True)
        put(ws, rr, 7, dup, num=True)
        rr += 1
    put(ws, rr + 1, 1, "Nhận xét: QTT TNCN đọc từ XML; danh sách NLD chi tiết phụ thuộc cấu trúc 05-1 trong XML. Trùng CCCD đếm trên dòng NLD đã bóc được.", wrap=True, fill=NOTE)

    # EBITDA sheet add
    if "KQ2_EBITDA_Phan_tich" not in wb.sheetnames:
        ews = wb.create_sheet("KQ2_EBITDA_Phan_tich", 1)
    else:
        ews = wb["KQ2_EBITDA_Phan_tich"]
    ews.sheet_view.showGridLines = True
    put(ews, 1, 1, "KQ2_EBITDA_Phan_tich - NĐ 132/2020 (30% EBITDA)", bold=True, fill=NAVY, font=FONT_H)
    ews.merge_cells("A1:G1")
    headers = ["Chỉ tiêu"] + [str(y) for y in YEARS] + ["Ghi chú"]
    for i, h in enumerate(headers, 1):
        put(ews, 3, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    ebitda_rows = [
        "Loi nhuan truoc thue [50]",
        "Chi phi lai vay (635 / [23])",
        "Khau hao TSCD",
        "EBITDA",
        "Tran 30% EBITDA",
        "Vuot nguong lai vay",
        "Co GDLK?",
    ]
    for i, name in enumerate(ebitda_rows):
        put(ews, 4 + i, 1, name)
        for j, y in enumerate(YEARS):
            if i == 0:
                put(ews, 4 + i, 2 + j, bctc_val(y, "50"), num=True)
            elif i == 1:
                put(ews, 4 + i, 2 + j, bctc_val(y, "23"), num=True)
            elif i == 6:
                put(ews, 4 + i, 2 + j, "Chưa đủ nguồn IN12")
    put(ews, 12, 1, "Nhan xet: chi phi lai vay lay tu BCTC neu co ma [23]; thieu so khau hao chi tiet -> EBITDA chua chot. IN12 chua nhan.", wrap=True, fill=NOTE)

    # Tong hop co - leave formulas if present
    ws = wb["KQ2_Tong_hop_co"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)

    apply_font(wb["KQ2_Doi_chieu_GTGT"])
    wb.save(lp(names["kq21"]))
    wb.close()

    # ---- KQ2.2 ----
    print("Fill KQ2.2")
    wb = load_workbook(lp(names["kq22"]))
    ws = wb["KQ2_511_632"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)
    # rows 6-... indicators, cols B total, C-F years
    labels_511 = [
        (6, "Doanh thu TK 511 (PS Có)", [y511[y] for y in YEARS]),
        (7, "Giá vốn TK 632 (PS Nợ)", [y632[y] for y in YEARS]),
    ]
    tot511 = sum(y511.values())
    tot632 = sum(y632.values())
    put(ws, 6, 2, tot511, num=True)
    put(ws, 7, 2, tot632, num=True)
    put(ws, 8, 2, tot511 - tot632, num=True)
    put(ws, 9, 2, (tot511 - tot632) / tot511 if tot511 else None, pct=True)
    n_khop = sum(1 for m in matched if m["Trang_thai"].startswith("Khớp HĐ + ngày"))
    n_p = sum(1 for m in matched if (m.get("Lai_gop") or 0) <= 0)
    put(ws, 10, 2, n_khop, num=True)
    put(ws, 11, 2, n_p, num=True)
    for j, y in enumerate(YEARS):
        put(ws, 5, 3 + j, y, bold=True, center=True)
        put(ws, 6, 3 + j, y511[y], num=True)
        put(ws, 7, 3 + j, y632[y], num=True)
        lai = y511[y] - y632[y]
        put(ws, 8, 3 + j, lai, num=True)
        put(ws, 9, 3 + j, lai / y511[y] if y511[y] else None, pct=True)
        mm = y_match[y]
        put(ws, 10, 3 + j, sum(1 for m in mm if str(m["Trang_thai"]).startswith("Khớp HĐ + ngày")), num=True)
        put(ws, 11, 3 + j, sum(1 for m in mm if (m.get("Lai_gop") or 0) <= 0), num=True)
        tl = lai / y511[y] if y511[y] else None
        dg = ""
        if tl is not None and tl < 0:
            dg = "Lãi gộp âm - rủi ro giá vốn/DT"
        elif n_p:
            dg = "Có dòng P<=0"
        else:
            dg = "Theo dõi"
        put(ws, 6, 7, "Đối chiếu khóa HĐ+ngày từ sổ SHVND/IN; DN SX go tram + thue xuong.")
    put(ws, 16, 1, f"Khớp khóa 1: {n_khop}/{len(r511)} dòng 511. Bán dưới giá/hòa vốn: {n_p} dòng. 632 không khớp 511: {len(un632)}.", wrap=True, fill=NOTE)

    if "KQ2_511_632_Chi_tiet" not in wb.sheetnames:
        cws = wb.create_sheet("KQ2_511_632_Chi_tiet")
    else:
        cws = wb["KQ2_511_632_Chi_tiet"]
    heads = ["Nam", "Ngay_CT", "So_CT", "So_HD", "Dien_giai", "DT_511", "GV_632", "Lai_gop", "Trang_thai"]
    for i, h in enumerate(heads, 1):
        put(cws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    rr = 2
    for m in matched:
        put(cws, rr, 1, m.get("Nam"), num=True)
        put(cws, rr, 2, m.get("Ngay_CT"), dt=True, center=True)
        put(cws, rr, 3, m.get("So_CT"))
        put(cws, rr, 4, m.get("So_HD_LQ"))
        put(cws, rr, 5, m.get("Dien_giai"))
        put(cws, rr, 6, m.get("PS_Co") or 0, num=True)
        put(cws, rr, 7, m.get("PS_632"), num=True)
        put(cws, rr, 8, m.get("Lai_gop"), num=True)
        put(cws, rr, 9, m.get("Trang_thai"))
        if (m.get("Lai_gop") or 0) <= 0:
            for c in range(1, 10):
                cws.cell(rr, c).fill = RED
                cws.cell(rr, c).font = RED_F
        rr += 1
    apply_font(cws)
    autofit(cws)

    ws = wb["KQ2_Kho_Gia_thanh"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)
    put(ws, 7, 12, f"Có {len(kho_files)} file NXT/152/154/155. Chưa cộng tồn CK vì cấu trúc sổ kho nhiều file; KH04 để KQ3 đối chiếu ct140 TT133.", wrap=True)
    put(ws, 20, 1, "Nhận xét TC3: BCTC TT133 dùng ct140 (không dùng ct141=0). Thiếu bảng tổng hợp XNT 1 dòng/năm nên chưa chốt khớp tồn CK vs BCTC.", wrap=True, fill=NOTE)

    ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)
    for j, y in enumerate(YEARS):
        put(ws, 7 + j, 1, y, num=True)
        put(ws, 7 + j, 8, y632[y], num=True)
        dt = bctc_val(y, "01")
        gv = bctc_gv(y)
        put(ws, 7 + j, 11, f"BCTC [01]={dt} [02]={gv}" if dt is not None else "Thieu BCTC")
    put(ws, 20, 1, "DN vừa SX vừa cho thuê/bán điện. Tỷ lệ GV/DT cần tách hoạt động gỗ vs thuê xưởng khi kiểm tra thực địa.", wrap=True, fill=NOTE)

    ws = wb["KQ2_Cong_no_Tien_mat"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)
    put(ws, 8, 11, "IN11 chưa nhận - không kết luận thiếu UNC.", wrap=True)
    # cash hits summary
    put(ws, 40, 1, f"TT01: {len(cash_hits)} HĐ mua >= ngưỡng tiền mặt (20tr đến 30/06/2025, 5tr từ 01/07/2025). TT02/GT16 chờ sao kê.", wrap=True, fill=NOTE)
    put(ws, 41, 1, f"IN10 chi TK 111: {len(r111)} dòng PS Có. IN11: Chưa nhận.", wrap=True, fill=NOTE)

    ws = wb["KQ2_Tong_hop_co"]
    put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên {VER}", bold=True)
    wb.save(lp(names["kq22"]))
    wb.close()

    # ---- KQ3 ----
    print("Fill KQ3")
    wb = load_workbook(lp(names["kq3"]))
    ws = wb["KQ3"]
    put(ws, 2, 2, MST, bold=True)
    put(ws, 2, 7, TEN, bold=True)
    put(ws, 3, 2, KY)
    put(ws, 3, 7, muc)
    put(ws, 4, 2, VER)
    put(ws, 4, 7, "1.4")
    put(ws, 5, 2, TODAY, dt=True)
    put(ws, 5, 7, f"KQ1.3 {VER}; KQ2.1 {VER}; KQ2.2 {VER}")

    # load rule catalog from KQ1.3 kha nang
    cat = []
    wb13 = load_workbook(lp(names["kq13"]), data_only=True)
    wsk = wb13["KQ1_Kha_nang_phan_tich"]
    for r in range(10, 90):
        ma = wsk.cell(r, 1).value
        if not ma:
            continue
        cat.append(
            {
                "ma": str(ma),
                "nhom": wsk.cell(r, 2).value,
                "ten": wsk.cell(r, 3).value,
                "need": str(wsk.cell(r, 4).value or ""),
                "tep": wsk.cell(r, 6).value,
                "sheet": wsk.cell(r, 7).value,
                "kn": wsk.cell(r, 8).value,
                "thieu": wsk.cell(r, 10).value,
            }
        )
    wb13.close()

    findings = []  # for extra files

    def add_rule(rdef, **kw):
        return {**rdef, **kw}

    rows_kq3 = []
    for i, rd in enumerate(cat):
        ma = rd["ma"]
        need = rd["need"]
        missing = [x for x in re.findall(r"IN\d{2}", need) if x not in {"IN07"} and not has.get(x, False)]
        status = "ĐÃ_CHẠY"
        ket = "Không"
        nhan = "Chắc"
        muc_rr = "Thấp"
        diem = 0
        hs = 1
        va = vb = do = None
        ca_a = cb_b = ""
        v_tien = w_tien = None
        sac = ""
        thieu_txt = ""
        nguon = ""
        if missing:
            status = "KHÔNG_ĐỦ_DỮ_LIỆU"
            ket = "Chưa kết luận"
            nhan = ""
            muc_rr = ""
            thieu_txt = ",".join(sorted(set(missing)))
        # specific
        if ma == "HD07":
            status = "TREO"
            ket = "Chưa kết luận"
            thieu_txt = "thieu so lieu nganh (K13)"
        if ma == "TN20":
            status = "CHƯA_TRIỂN_KHAI"
            ket = "Chưa kết luận"

        # compute where possible
        if ma == "GT01" and status == "ĐÃ_CHẠY":
            diffs = []
            for y in YEARS:
                a = tot_tk(y, "ct34")
                if a is None:
                    a = tot_tk(y, "ct32")
                b = tot_hd("BAN", y)
                if a is None:
                    continue
                diffs.append((y, a, b, a - b))
            if diffs:
                y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
                ca_a, va, cb_b, vb, do = f"TK[34] {y}", a, "HDDT ban", b, d
                nguon = "KQ2.1 KQ2_Doi_chieu_GTGT"
                if abs(d) > 1:
                    ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
                    v_tien = abs(d)
                    w_tien = abs(d)
                    findings.append(("GT01", y, d, a, b))
        if ma == "GT02" and status == "ĐÃ_CHẠY":
            diffs = []
            for y in YEARS:
                a = tot_tk(y, "ct34")
                if a is None:
                    a = tot_tk(y, "ct32")
                b = tot_bk("BAN", y)
                if a is None:
                    continue
                diffs.append((y, a, b, a - b))
            if diffs:
                y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
                ca_a, va, cb_b, vb, do = f"TK[34] {y}", a, "BK ban", b, d
                nguon = "KQ2.1"
                if abs(d) > 1:
                    ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 8, "GTGT"
                    v_tien = abs(d)
                    w_tien = 0  # T2 avoid double count with GT01
                    findings.append(("GT02", y, d, a, b))
        if ma == "GT04" and status == "ĐÃ_CHẠY":
            diffs = []
            for y in YEARS:
                a = tot_tk(y, "ct23")
                b = tot_hd("MUA", y)
                if a is None:
                    continue
                diffs.append((y, a, b, a - b))
            if diffs:
                y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
                ca_a, va, cb_b, vb, do = f"TK[23] {y}", a, "HDDT mua", b, d
                if abs(d) > 1:
                    ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 8, "GTGT"
                    v_tien = abs(d)
                    w_tien = abs(d)
                    findings.append(("GT04", y, d, a, b))
        if ma == "HD01" and status == "ĐÃ_CHẠY":
            ca_a, va, cb_b, vb, do = "NCC F0", len(f0_ncc), "Doanh so F0", f0_gt, f0_gt
            nguon = "KQ2.1 IN07/NCC"
            if f0_ncc:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
                v_tien = f0_thue
                w_tien = f0_thue
            if nnt_f1:
                ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Cao", max(diem, 9)
        if ma == "HD04" and status == "ĐÃ_CHẠY":
            ca_a, va, cb_b, vb, do = "Top5/tong mua", top5_pct, "Nguong P14", 0.5, top5_pct - 0.5
            if top5_pct > 0.5:
                ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Trung bình", 6
        if ma == "HD08" and status == "ĐÃ_CHẠY":
            ca_a, va = "NNT tren F1", 1 if nnt_f1 else 0
            if nnt_f1:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
        if ma == "TT01" and status == "ĐÃ_CHẠY":
            ca_a, va = "HD mua >= nguong", len(cash_hits)
            do = len(cash_hits)
            nhan = "Sơ bộ"
            if cash_hits:
                ket, muc_rr, diem = "Có phát hiện", "Trung bình", 5
        if ma == "TN01" and status == "ĐÃ_CHẠY":
            diffs = []
            for y in YEARS:
                a = tot_tk(y, "ct34")
                if a is None:
                    a = tot_tk(y, "ct32")
                b = bctc_val(y, "01")
                if a is None or b is None:
                    continue
                diffs.append((y, a, b, a - b))
            if diffs:
                y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
                ca_a, va, cb_b, vb, do = f"GTGT[34] {y}", a, "BCTC[01]", b, d
                nhan = "Chắc"  # now have 511
                if abs(d) > 1:
                    ket, muc_rr, diem, sac = "Có phát hiện", "Cao", 7, "TNDN"
                    v_tien = abs(d)
                    w_tien = 0
        if ma == "TN03" and status == "ĐÃ_CHẠY":
            for y in YEARS:
                dt = bctc_val(y, "01")
                gv = bctc_gv(y)
                if dt and gv is not None and dt - gv < 0:
                    ket, nhan, muc_rr, diem = "Có phát hiện", "Sơ bộ - chưa so ngành", "Cao", 7
                    ca_a, va, cb_b, vb, do = f"DT {y}", dt, "GV", gv, dt - gv
                    break
            else:
                nhan = "Sơ bộ - chưa so ngành"
        if ma == "DT01" and not missing:
            status = "ĐÃ_CHẠY"
            ket = "Có phát hiện" if n_p else "Không"
            ca_a, va = "Dong P<=0", n_p
            nhan, muc_rr, diem = "Chắc", ("Cao" if n_p else "Thấp"), (8 if n_p else 0)
            sac = "TNDN" if n_p else ""
        if ma == "DT02" and not missing:
            status = "ĐÃ_CHẠY"
            n_miss = sum(1 for m in matched if m["Trang_thai"] == "KHÔNG_CÓ_632")
            ket = "Có phát hiện" if n_miss else "Không"
            ca_a, va = "511 khong 632", n_miss
            nhan = "Chắc"
            if n_miss:
                muc_rr, diem = "Trung bình", 6
        if ma == "DT03" and not missing:
            status = "ĐÃ_CHẠY"
            ket = "Có phát hiện" if un632 else "Không"
            ca_a, va = "632 khong 511", len(un632)
            nhan = "Chắc"
            if un632:
                muc_rr, diem = "Trung bình", 6
        if ma == "KH04":
            if has["IN09"] and has["IN04"]:
                status = "ĐÃ_CHẠY"
                ket = "Chưa kết luận"
                thieu_txt = "Chua cong duoc ton CK 152+155+156 tu nhieu file NXT"
                nhan = "Sơ bộ"
            else:
                status = "KHÔNG_ĐỦ_DỮ_LIỆU"
                ket = "Chưa kết luận"
        if ma == "LS01" or ma == "LS02":
            status = "ĐÃ_CHẠY" if has["IN14"] else "KHÔNG_ĐỦ_DỮ_LIỆU"
            ket = "Chưa kết luận"
            thieu_txt = "PTDN01/02 la BCTC TMS, chua co bang no thue/cham nop dong"
            nhan = ""
        if ma == "DQ02" and status == "ĐÃ_CHẠY":
            ket = "Không"
            nhan = "Chắc"
            ca_a, va = "Dong Tong cong da loc", 0

        hs = {"Cao": 3, "Trung bình": 2, "Thấp": 1}.get(muc_rr, 0)
        if ket != "Có phát hiện":
            w_tien = None if status != "ĐÃ_CHẠY" else w_tien
        ma_ph = f"{ma}|{MST}|{KY}"
        rows_kq3.append(
            {
                "Mã_PH": ma_ph,
                "Tệp phân tích": rd["tep"] or "Trực tiếp",
                "Nhóm": rd["nhom"],
                "Mã quy tắc": ma,
                "Tên quy tắc": rd["ten"],
                "Đối tượng": TEN,
                "Kỳ": KY,
                "Chỉ tiêu A": ca_a,
                "Giá trị A": va,
                "Chỉ tiêu B": cb_b,
                "Giá trị B": vb,
                "Giá trị đo": do,
                "Ngưỡng": rd.get("need"),
                "Trạng thái chạy": status,
                "Kết quả": ket,
                "Nhãn": nhan,
                "Mức rủi ro": muc_rr,
                "Hệ số trọng yếu": hs,
                "Điểm": diem,
                "Sắc thuế": sac,
                "Số tiền liên quan (đồng)": v_tien,
                "Số tiền tính vào tổng (đồng)": w_tien if ket == "Có phát hiện" else 0 if status == "ĐÃ_CHẠY" else None,
                "Nguồn": nguon,
                "Còn thiếu để kết luận": thieu_txt,
                "Nội dung kiểm tra tại trụ sở": "Xem chung tu goc" if ket == "Có phát hiện" else "",
                "Chứng từ cần xem": "",
                "Ghi chú": "",
            }
        )

    # write from row 10
    hm = header_map(ws, 9)
    if ws.max_row > 10:
        ws.delete_rows(10, ws.max_row - 9)
    for i, rec in enumerate(rows_kq3):
        rr = 10 + i
        for k, v in rec.items():
            if k not in hm:
                continue
            c = hm[k]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                put(ws, rr, c, v, num=True)
            else:
                put(ws, rr, c, v)
        if rec.get("Mức rủi ro") == "Cao":
            ws.cell(rr, hm["Mức rủi ro"]).fill = RED
            ws.cell(rr, hm["Mức rủi ro"]).font = RED_F
        elif rec.get("Mức rủi ro") == "Trung bình":
            ws.cell(rr, hm["Mức rủi ro"]).fill = YEL
            ws.cell(rr, hm["Mức rủi ro"]).font = ORG_F
        if rec.get("Trạng thái chạy") == "KHÔNG_ĐỦ_DỮ_LIỆU":
            ws.cell(rr, hm["Trạng thái chạy"]).fill = YEL

    w_gtgt = sum((r.get("Số tiền tính vào tổng (đồng)") or 0) for r in rows_kq3 if r.get("Sắc thuế") == "GTGT")
    w_tndn = sum((r.get("Số tiền tính vào tổng (đồng)") or 0) for r in rows_kq3 if r.get("Sắc thuế") == "TNDN")
    put(ws, 7, 2, w_gtgt, num=True)
    n_cao = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Cao" and r.get("Kết quả") == "Có phát hiện")
    n_tb = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Trung bình" and r.get("Kết quả") == "Có phát hiện")
    put(ws, 6, 8, n_cao, num=True)
    put(ws, 6, 9, n_tb, num=True)

    # 28 TC
    ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
    # delete example row 4
    if str(ws.cell(4, 1).value or "").startswith("TC"):
        ws.delete_rows(4, 1)
    tc = [
        ("TC01", "Chênh lệch DT GTGT [34] vs DT BCTC [01]", "Doanh thu", "lệch ≠ 0"),
        ("TC02", "Người mua trả tiền trước > 5% DT", "Cong no", "5%"),
        ("TC03", "HTK + XDCB > 30% DT", "Hang ton kho", "30% - B01-DNN dùng ct140"),
        ("TC04", "DT chưa thực hiện > 0", "Doanh thu", "> 0"),
        ("TC05", "DT thuần - giá vốn < 0", "Gia von", "< 0"),
        ("TC06", "Chi phí phải trả trích trước lớn", "Chi phi", "[315]"),
        ("TC07", "Lãi vay chưa vốn hóa / CP tra truoc", "Chi phi", "[242]"),
        ("TC08", "Ưu đãi / miễn giảm TNDN", "TNDN", "C2/C11"),
        ("TC09", "VAT khấu trừ vs BCTC [152]", "GTGT", "[22]/[43] vs [152]"),
        ("TC10", "VAT khấu trừ vs 10%×(HTK+TSCĐ)", "GTGT", "×10%"),
        ("TC11", "Có DT KCT [26] nhưng [25]=[24]", "GTGT", ""),
        ("TC12", "Điều chỉnh giảm LN trước thuế", "TNDN", "B7/B8"),
        ("TC13", "Trích lập dự phòng", "BCTC", ""),
        ("TC14", "CN von [22]>0 va [25]<0", "BCTC", ""),
        ("TC15", "Bù trừ lãi BĐS / lỗ SXKD", "TNDN", ""),
        ("TC16", "Chuyển giá [251][252]>0 và [421]<0", "GDLK", ""),
        ("TC17", "CP lãi vay vs 30% EBITDA", "EBITDA", "NĐ 132"),
        ("TC18", "B11>0 không có PL GDLK", "GDLK", "IN12"),
        ("TC19", "Thuế TNDN hoãn lại", "TNDN", ""),
        ("TC20", "Phải trả người bán lớn", "Cong no", "[311]+[331]"),
        ("TC21", "Lãi vay khi thiếu VĐL", "Tai chinh", "bo"),
        ("TC22", "QTT TNCN vs TK tháng/quý", "TNCN", ""),
        ("TC23", "Dòng tiền chi NLĐ vs TNCT", "TNCN", "20%"),
        ("TC24", "KHBS vs hạch toán TMS", "TMS", "khong chay"),
        ("TC25", "Bổ sung GTGT chưa điều chỉnh", "GTGT", ""),
        ("TC26", "Tỷ trọng HĐ mua NNT TT 03/06", "Hoa don", "can DKKD"),
        ("TC27", "DT TK GTGT vs HĐĐT bán ra", "Hoa don", "lệch ≠ 0"),
        ("TC28", "Phân tích ngang/dọc BCTC - GV tăng nhanh hơn DT", "Gia von", ""),
    ]
    for i, (id_, ten, phan, nguong) in enumerate(tc):
        kq = "Chưa đủ dữ liệu để chốt"
        muc_tc = "Thấp"
        kn = "Bổ sung chứng từ nếu thiếu nguồn"
        if id_ == "TC01":
            kq = "Có chênh GTGT[34] vs BCTC[01] ở một số năm (xem GT/TN01)."
            muc_tc = "Cao"
            kn = "Đối chiếu sổ 511, hóa đơn bán ra từng năm, cut-off."
        elif id_ == "TC03":
            kq = "Che do TT133 - dung ct140. Chua cong duoc ton CK tu nhieu file NXT."
            muc_tc = "Trung bình"
            kn = "Lap bang XNT 152+155 (+156 neu co) roi khop BCTC ct140."
        elif id_ == "TC05":
            kq = f"Lãi gộp sổ 511-632 kỳ: {(tot511-tot632)/tot511 if tot511 else None}."
            muc_tc = "Trung bình" if tot511 and (tot511 - tot632) / tot511 < 0.05 else "Thấp"
            kn = "Rà từng dòng P<=0 trên KQ2_511_632_Chi_tiet."
        elif id_ == "TC17":
            kq = "EBITDA: chi phi lai vay / khau hao chua du de chot 30%."
            muc_tc = "Thấp"
            kn = "Lay so 635 va khau hao tu so cai."
        elif id_ == "TC18":
            kq = "IN12 chưa nhận - KHÔNG_ĐỦ_DỮ_LIỆU."
            muc_tc = "Thấp"
            kn = "Xin PL GDLK neu co ben lien ket."
        elif id_ == "TC22":
            kq = "Da doc XML QTT TNCN 2021-2025; tra cuu da nguon o 3.3 mong."
            muc_tc = "Thấp"
            kn = "Doi chieu bang luong 2.4 vs QTT."
        elif id_ == "TC24":
            kq = "Khong chay (cam TMS dong)."
            muc_tc = "Thấp"
            kn = ""
        elif id_ == "TC26":
            kq = f"Doi chieu DKKD: {sum(1 for k in ncc_agg if ncc_flag(k)[3])} NCC co trang thai."
            muc_tc = "Trung bình"
            kn = "Chon mau NCC ngung/dong MST."
        elif id_ == "TC27":
            kq = "Lech DT to khai vs HDDT ban ra - xem GT01."
            muc_tc = "Cao"
            kn = "Chon mau HD dieu chinh/huy/sai ky."
        elif id_ == "TC28":
            kq = "GV 632 vs DT 511 theo nam da co; can tach hoat dong thue xuong."
            muc_tc = "Trung bình"
            kn = "Tach doanh thu go / thue xuong / dien."
        put(ws, 4 + i, 1, id_, bold=True)
        put(ws, 4 + i, 2, ten)
        put(ws, 4 + i, 3, phan)
        put(ws, 4 + i, 4, nguong)
        put(ws, 4 + i, 5, kq, wrap=True)
        put(ws, 4 + i, 6, muc_tc, bold=True)
        put(ws, 4 + i, 7, kn, wrap=True)
        if muc_tc == "Cao":
            ws.cell(4 + i, 6).fill = RED
            ws.cell(4 + i, 6).font = RED_F
        elif muc_tc == "Trung bình":
            ws.cell(4 + i, 6).fill = YEL
            ws.cell(4 + i, 6).font = ORG_F
    apply_font(ws)
    autofit(ws)
    apply_font(wb["KQ3"])
    wb.save(lp(names["kq3"]))
    wb.close()

    # ---- extra GĐ3 files ----
    print("Fill extra GĐ3")
    wb = Workbook()
    ws = wb.active
    ws.title = "Dau_hieu"
    heads = ["STT", "Ma", "Dau hieu", "Muc", "So lieu", "Sheet KQ", "Huong kiem tra", "Tai lieu can", "Trang thai"]
    for i, h in enumerate(heads, 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    dau = [
        (1, "GT01", "Lech DT to khai GTGT vs HDDT ban ra", "Cao", "Xem GT01 theo nam", "KQ2.1/KQ3", "1. Chon HD dieu chinh 2. Sai ky 3. HD huy 4. Doi chieu BK", "HDDT portal, BK ban ra, so 511", "Chua kiem tra tai tru so"),
        (2, "GT04", "Lech GT mua vao to khai vs HDDT", "Cao", "Xem GT04", "KQ2.1/KQ3", "1. HD mua khong ke 2. Ke khong co HD 3. Huy/thay the", "HDDT mua, BK mua, so 133", "Chua kiem tra tai tru so"),
        (3, "HD01", f"NCC F0/CA: {len(f0_ncc)} MST, DS {f0_gt:,.0f}", "Cao", f"Thue {f0_thue:,.0f}", "KQ2.1 NCC", "1. Hop dong 2. Thanh toan NH 3. Nhap kho 4. Su dung", "IN07, UNC, PNK", "Thieu IN11"),
        (4, "HD04", f"Top-5 NCC {top5_pct:.2%}", "Trung bình" if top5_pct > 0.5 else "Thấp", f"{len(ncc_agg)} NCC", "KQ2.1", "Chon mau hop dong Top-5", "Hop dong, BK", "Chua kiem tra"),
        (5, "DT01", f"Ban duoi gia/hoa von {n_p} dong", "Cao" if n_p else "Thấp", f"511 {len(r511)} / 632 {len(r632)}", "KQ2.2 Chi tiet", "Doi chieu PXK vs HDBR", "So 511/632, the kho", "Da chay so"),
        (6, "TT01", f"{len(cash_hits)} HD mua vuot nguong tien mat", "Trung bình", "P10/P11", "KQ2.2 Cong no", "Doi UNC; khong ket luan thieu chung tu khi chua co sao ke", "Sao ke, UNC", "Thieu IN11"),
        (7, "IN11", "Thieu sao ke ngan hang", "Trung bình", "2.3 trong", "KQ1.3", "De nghi NNT bo sung", "Excel sao ke", "Chua nhan"),
    ]
    for i, row in enumerate(dau):
        for c, v in enumerate(row, 1):
            put(ws, 2 + i, c, v)
            if row[3] == "Cao":
                ws.cell(2 + i, 4).fill = RED
                ws.cell(2 + i, 4).font = RED_F
            elif row[3] == "Trung bình":
                ws.cell(2 + i, 4).fill = YEL
                ws.cell(2 + i, 4).font = ORG_F
    apply_font(ws)
    autofit(ws)
    wb.save(lp(names["dau"]))
    wb.close()

    wb = Workbook()
    sheets = [
        "00_Tong_hop_chi_tiet",
        "1_NCC_F0_CanhBao",
        "2_NCC_DKKD_bat_thuong",
        "3_Ty_le_gia_von",
        "4_HD_ban_duoi_gia",
        "5_Khong_vay_no_EBITDA",
        "6_Top5_NCC_95%",
        "7_Lech_HDDT_BangKe",
        "8_Cong_no_TK131_TK331",
        "9_TNCN_NLD_nhieu_nguon",
    ]
    wb.active.title = sheets[0]
    for s in sheets[1:]:
        wb.create_sheet(s)

    ws = wb["00_Tong_hop_chi_tiet"]
    put(ws, 1, 1, f"{TEN} - {MST} - {KY} - Tong hop chi tiet dau hieu", bold=True, fill=NAVY, font=FONT_H)
    ws.merge_cells("A1:H1")
    for i, h in enumerate(["Phan", "Chi tieu", "Gia tri", "Nguon"], 1):
        put(ws, 3, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    summary = [
        ("I", "So HĐĐT mua (tinh vao tong)", sum(1 for x in in01 if x["Chieu"] == "MUA" and x["Tinh_vao_tong"] == "Có"), "IN01"),
        ("I", "So HĐĐT ban (tinh vao tong)", sum(1 for x in in01 if x["Chieu"] == "BAN" and x["Tinh_vao_tong"] == "Có"), "IN01"),
        ("I", "So dong bang ke mua", sum(1 for x in in02 if x["Chieu"] == "MUA"), "IN02"),
        ("I", "So dong bang ke ban", sum(1 for x in in02 if x["Chieu"] == "BAN"), "IN02"),
        ("I", "So NCC mua", len(ncc_agg), "IN02+IN01"),
        ("II", "NCC F0", len(f0_ncc), "IN07"),
        ("II", "Doanh so F0", f0_gt, "IN07"),
        ("II", "Top5 %", top5_pct, "HD04"),
        ("III", "Tong DT 511", tot511, "IN13"),
        ("III", "Tong GV 632", tot632, "IN13"),
        ("III", "Lai gop so", tot511 - tot632, "IN13"),
        ("III", "Dong P<=0", n_p, "DT01"),
        ("IV", "HD mua vuot nguong TM", len(cash_hits), "TT01"),
        ("IV", "Sao ke NH", "Chua nhan", "IN11"),
        ("V", "NNT tren F1", "Có" if nnt_f1 else "Không", "HD08"),
    ]
    for i, row in enumerate(summary):
        for c, v in enumerate(row, 1):
            put(ws, 4 + i, c, v, num=isinstance(v, (int, float)))

    ws = wb["1_NCC_F0_CanhBao"]
    for i, h in enumerate(["MST", "Ten", "DS", "Thue", "Nhom", "Nguon", "CV"], 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    rr = 2
    for k, a in f0_ncc:
        nhom, nguon, cv, tt = ncc_flag(k)
        put(ws, rr, 1, mst_chuan(k))
        put(ws, rr, 2, a["ten"])
        put(ws, rr, 3, a["gt"], num=True)
        put(ws, rr, 4, a["thue"], num=True)
        put(ws, rr, 5, nhom)
        put(ws, rr, 6, nguon)
        put(ws, rr, 7, cv)
        for c in range(1, 8):
            ws.cell(rr, c).fill = RED
            ws.cell(rr, c).font = RED_F
        rr += 1
    if rr == 2:
        put(ws, 2, 1, "Khong co NCC F0 trong ky")

    ws = wb["2_NCC_DKKD_bat_thuong"]
    for i, h in enumerate(["MST", "Ten", "DS", "Trang thai"], 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    rr = 2
    for k, a in ncc_agg.items():
        tt = ncc_flag(k)[3]
        if tt and any(x in str(tt).lower() for x in ["ngừng", "ngung", "đóng", "dong", "bỏ", "bo dia"]):
            put(ws, rr, 1, mst_chuan(k))
            put(ws, rr, 2, a["ten"])
            put(ws, rr, 3, a["gt"], num=True)
            put(ws, rr, 4, tt)
            for c in range(1, 5):
                ws.cell(rr, c).fill = YEL
                ws.cell(rr, c).font = ORG_F
            rr += 1
    if rr == 2:
        put(ws, 2, 1, "Chua gan duoc DKKD bat thuong (danh ba mong hoac khong khop MST).")

    ws = wb["3_Ty_le_gia_von"]
    for i, h in enumerate(["Nam", "DT_511", "GV_632", "Lai gop", "Ty le", "BCTC[01]", "BCTC[02]"], 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    for j, y in enumerate(YEARS):
        put(ws, 2 + j, 1, y)
        put(ws, 2 + j, 2, y511[y], num=True)
        put(ws, 2 + j, 3, y632[y], num=True)
        put(ws, 2 + j, 4, y511[y] - y632[y], num=True)
        put(ws, 2 + j, 5, (y511[y] - y632[y]) / y511[y] if y511[y] else None, pct=True)
        put(ws, 2 + j, 6, bctc_val(y, "01"), num=True)
        put(ws, 2 + j, 7, bctc_gv(y), num=True)

    ws = wb["4_HD_ban_duoi_gia"]
    for i, h in enumerate(["Nam", "Ngay", "So HD", "DT", "GV", "Lai", "Dien giai"], 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    rr = 2
    for m in matched:
        if (m.get("Lai_gop") or 0) > 0:
            continue
        put(ws, rr, 1, m.get("Nam"))
        put(ws, rr, 2, m.get("Ngay_CT"), dt=True, center=True)
        put(ws, rr, 3, m.get("So_HD_LQ"))
        put(ws, rr, 4, m.get("PS_Co"), num=True)
        put(ws, rr, 5, m.get("PS_632"), num=True)
        put(ws, rr, 6, m.get("Lai_gop"), num=True)
        put(ws, rr, 7, m.get("Dien_giai"))
        for c in range(1, 8):
            ws.cell(rr, c).fill = RED
            ws.cell(rr, c).font = RED_F
        rr += 1
    if rr == 2:
        put(ws, 2, 1, "Khong co dong P<=0")

    ws = wb["5_Khong_vay_no_EBITDA"]
    put(ws, 1, 1, "EBITDA 4-5 nam - so BCTC [50]/[23]; khau hao chua tach.", bold=True, fill=NAVY, font=FONT_H)
    for j, y in enumerate(YEARS):
        put(ws, 3, 2 + j, y, bold=True, center=True)
        put(ws, 4, 1, "LN truoc thue [50]")
        put(ws, 4, 2 + j, bctc_val(y, "50"), num=True)
        put(ws, 5, 1, "Lai vay [23]")
        put(ws, 5, 2 + j, bctc_val(y, "23"), num=True)
    put(ws, 7, 1, "Neu lai vay = 0 tren BCTC: to xanh tuan thu NĐ 132 o muc so lieu ke khai; van can so 635.", wrap=True, fill=GRN)

    ws = wb["6_Top5_NCC_95%"]
    for i, h in enumerate(["Hang", "MST", "Ten", "DS", "Ty trong"], 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    for i, (k, a) in enumerate(top5, 1):
        put(ws, 1 + i, 1, i)
        put(ws, 1 + i, 2, mst_chuan(k))
        put(ws, 1 + i, 3, a["ten"])
        put(ws, 1 + i, 4, a["gt"], num=True)
        put(ws, 1 + i, 5, a["gt"] / tot_mua if tot_mua else None, pct=True)
    put(ws, 8, 1, f"Tong Top-5 = {top5_pct:.2%}")

    ws = wb["7_Lech_HDDT_BangKe"]
    for i, h in enumerate(["Nam", "TK[34]", "HDDT ban", "BK ban", "Lech TK-HD", "Lech TK-BK", "TK[23]", "HDDT mua", "BK mua"], 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    for j, y in enumerate(YEARS):
        a = tot_tk(y, "ct34")
        if a is None:
            a = tot_tk(y, "ct32")
        hd_b = tot_hd("BAN", y)
        bk_b = tot_bk("BAN", y)
        hd_m = tot_hd("MUA", y)
        bk_m = tot_bk("MUA", y)
        a23 = tot_tk(y, "ct23")
        put(ws, 2 + j, 1, y)
        put(ws, 2 + j, 2, a, num=True)
        put(ws, 2 + j, 3, hd_b, num=True)
        put(ws, 2 + j, 4, bk_b, num=True)
        put(ws, 2 + j, 5, None if a is None else a - hd_b, num=True)
        put(ws, 2 + j, 6, None if a is None else a - bk_b, num=True)
        put(ws, 2 + j, 7, a23, num=True)
        put(ws, 2 + j, 8, hd_m, num=True)
        put(ws, 2 + j, 9, bk_m, num=True)

    ws = wb["8_Cong_no_TK131_TK331"]
    put(ws, 1, 1, "IN11 chua nhan - khong ket luan thieu UNC (K11).", wrap=True, fill=YEL)
    put(ws, 2, 1, f"IN10 so dong chi 111: {len(r111)}")
    put(ws, 3, 1, f"TT01 HD mua vuot nguong: {len(cash_hits)}")

    ws = wb["9_TNCN_NLD_nhieu_nguon"]
    for i, h in enumerate(["Loai", "Nam", "Ho ten", "CCCD", "TNCT", "Thue", "Nguon", "File"], 1):
        put(ws, 1, i, h, bold=True, fill=NAVY, font=FONT_H, center=True)
    for i, rec in enumerate(in05_rows[:500]):
        put(ws, 2 + i, 1, rec.get("Loai_dong"))
        put(ws, 2 + i, 2, rec.get("Nam"))
        put(ws, 2 + i, 3, rec.get("Ho_ten"))
        put(ws, 2 + i, 4, rec.get("CCCD"))
        put(ws, 2 + i, 5, rec.get("Tong_TNCT"), num=True)
        put(ws, 2 + i, 6, rec.get("Thue_da_KT"), num=True)
        put(ws, 2 + i, 7, rec.get("Nguon_BK"))
        put(ws, 2 + i, 8, rec.get("File_nguon"))

    for s in wb.sheetnames:
        apply_font(wb[s])
        autofit(wb[s])
        wb[s].sheet_view.showGridLines = True
    wb.save(lp(names["giai"]))
    wb.close()

    summary = {
        "mst": MST,
        "ten": TEN,
        "ky": KY,
        "muc": muc,
        "ver": VER,
        "files": {k: str(v) for k, v in names.items()},
        "counts": {
            "hddt": len(in01),
            "bangke": len(in02),
            "gtgt": len(in03_rows),
            "bctc_rows": len(in04_rows),
            "tncn": len(in05_rows),
            "ncc": len(ncc_agg),
            "f0": len(f0_ncc),
            "r511": len(r511),
            "r632": len(r632),
            "p_le": n_p,
            "cash": len(cash_hits),
            "top5_pct": top5_pct,
            "nnt_f1": bool(nnt_f1),
            "tot511": tot511,
            "tot632": tot632,
            "w_gtgt": w_gtgt,
            "w_tndn": w_tndn,
            "n_cao": n_cao,
            "n_tb": n_tb,
            "kha_nang": {"now": n_now, "part": n_part, "block": n_block},
        },
        "by_year": {
            str(y): {
                "tk34": tot_tk(y, "ct34") if tot_tk(y, "ct34") is not None else tot_tk(y, "ct32"),
                "hddt_ban": tot_hd("BAN", y),
                "bk_ban": tot_bk("BAN", y),
                "tk23": tot_tk(y, "ct23"),
                "hddt_mua": tot_hd("MUA", y),
                "bk_mua": tot_bk("MUA", y),
                "bctc01": bctc_val(y, "01"),
                "bctc02": bctc_gv(y),
                "ps511": y511[y],
                "ps632": y632[y],
            }
            for y in YEARS
        },
    }
    js = OUT / "cau-thanh" / f"phan-tich-v14-{MST}-{KY}.json"
    js.parent.mkdir(parents=True, exist_ok=True)
    js.write_text(json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print("DONE", json.dumps(summary["counts"], ensure_ascii=False))
    return summary


if __name__ == "__main__":
    main()
