# -*- coding: utf-8 -*-
"""Giai doan 1 - Kien Thanh Tan MST 3603216935. Xuat KQ1.1, KQ1.2, KQ1.3 theo quy trinh 1.4."""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603216935"
TEN = "CÔNG TY TNHH MỘT THÀNH VIÊN KIẾN THÀNH TÂN"
KY = "2015-2025"
YEARS = list(range(2015, 2026))
TODAY = date(2026, 11, 17)
VER = "v1"
MAU_VER = "1.4"

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\1. TD\14. KIEN THANH TAN_3603216935"
)
OUT = ROOT / "4. KET QUA PHAN TICH"
VAULT_OUT = VAULT / "ket-qua" / "KIEN_THANH_TAN_3603216935"
ENGINE = VAULT / "skills" / "phan-tich-ho-so-thue" / "engine"
sys.path.insert(0, str(ENGINE))
from ingest import xml_tkhai as xt  # noqa: E402

xt.MA_LOAI["683"] = "bctc"
xt.MA_LOAI["402"] = "bctc"
xt.MA_LOAI["842"] = "gtgt"
xt.MA_LOAI["892"] = "tndn"
xt.MA_LOAI["953"] = "tncn"

CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F4E78")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
FILL_GRN = PatternFill("solid", fgColor="E2EFDA")
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
NUM = "#,##0"
DATE_FMT = "dd/mm/yyyy"
SKIP_HASH_EXT = {".bak"}


def lp(p) -> str:
    s = os.path.abspath(str(p))
    if s.startswith("\\\\?\\"):
        return s
    if s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + s[2:]
    return "\\\\?\\" + s


def sha256(path) -> str:
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def nnum(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("%", "")
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def ndate(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and v > 20000:
        try:
            return datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(v)).date()
        except Exception:
            return None
    s = str(v).strip().replace("00:00:00", "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    return None


def iso(d):
    return d.isoformat() if d else None


def norm_mst(v) -> str:
    if v is None:
        return ""
    s = str(v).strip().replace(" ", "").replace(".", "").upper()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    if "-" in s:
        a, b = s.split("-", 1)
        a = re.sub(r"\D", "", a)
        b = re.sub(r"\D", "", b)
        return f"{a}-{b}" if a and b else a
    return re.sub(r"\D", "", s)


def mst_base(s: str) -> str:
    s = norm_mst(s)
    return s.split("-")[0] if s else ""


def so_hd_key(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    s = re.sub(r"\s+", "", s)
    return s.lstrip("0") or "0"


def tnr(cell, bold=None):
    f = cell.font
    cell.font = Font(
        name="Times New Roman",
        size=f.size if f and f.size else 11,
        bold=f.bold if bold is None and f else (bool(bold) if bold is not None else False),
        italic=False,
        color=f.color if f else None,
    )


def apply_tnr(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            tnr(cell)


def autowidth(ws, cap=60, min_w=10, max_row=4000):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, max_row), max_col=ws.max_column or 1):
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            widths[cell.column] = max(widths.get(cell.column, min_w), min(len(s), cap) + 2)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


def clear_from(ws, start_row):
    if ws.max_row and ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)


def put(ws, r, c, v, bold=False, num=False, dt=False, fill=None, wrap=False):
    cell = ws.cell(r, c, v)
    cell.font = FONT_B if bold else FONT
    cell.border = THIN
    cell.alignment = Alignment(vertical="center", wrap_text=wrap)
    if fill:
        cell.fill = fill
    if num and isinstance(v, (int, float)) and not isinstance(v, bool):
        cell.number_format = NUM
    if dt and v not in (None, ""):
        cell.number_format = DATE_FMT
        cell.alignment = Alignment(vertical="center", horizontal="center")
    return cell


def write_row(ws, r, values, nums=(), dates=(), fill=None, wraps=()):
    for c, v in enumerate(values, 1):
        cell = ws.cell(r, c, v)
        tnr(cell)
        cell.border = THIN
        cell.alignment = Alignment(vertical="center", wrap_text=c in wraps)
        if fill:
            cell.fill = fill
        if c in nums and isinstance(v, (int, float)) and not isinstance(v, bool):
            cell.number_format = NUM
        if c in dates and v not in (None, ""):
            cell.number_format = DATE_FMT


def sub_placeholders(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str):
                v = cell.value
                v = v.replace("[MST]", MST).replace("[KY]", KY)
                v = v.replace("[NGAY]", TODAY.strftime("%d/%m/%Y"))
                v = v.replace("[TÊN DOANH NGHIỆP]", TEN).replace("[DN]", TEN)
                v = v.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")
                if v != cell.value:
                    cell.value = v
            tnr(cell)


def classify_file(rel: str, name: str) -> tuple[str, str]:
    u = rel.replace("/", "\\").upper()
    nu = name.upper()
    if "1. DU LIEU THO NGANH THUE" in u:
        src = "CQT"
    elif "2. HO SO NNT" in u:
        src = "DN"
    elif "3. DU LIEU BO SUNG" in u:
        src = "TRA_CUU"
    elif "0. BIEN BAN" in u:
        src = "BB"
    elif "4. KET QUA PHAN TICH" in u:
        src = "KQ"
    else:
        src = "KHAC"
    inn = "—"
    if "1.1" in u or "TO KHAI" in u:
        if nu.startswith("GTGT"):
            inn = "IN03"
        elif nu.startswith("TNDN") or nu.startswith("BCTC"):
            inn = "IN04"
        elif nu.startswith("TNCN"):
            inn = "IN05"
        elif nu.startswith("BC26") or nu.startswith("TB01") or "TB03" in nu:
            inn = "—"
        else:
            inn = "IN03"
    elif "1.2" in u or "HOA DON DIEN TU" in u:
        inn = "IN01"
    elif "1.3" in u or "NGHIA VU" in u:
        inn = "IN14"
    elif "1.4" in u:
        inn = "IN05"
    elif "2.5" in u or "BANG KE" in u:
        inn = "IN02"
    elif "2.2" in u or "KHO" in u:
        inn = "IN09"
    elif "2.3" in u or "NGAN HANG" in u:
        inn = "IN11"
    elif "2.1" in u:
        codes = re.findall(r"(?<!\d)(\d{3,5})(?!\d)", name)
        code = codes[-1] if codes else ""
        if code.startswith("111"):
            inn = "IN10"
        elif code.startswith("511") or code.startswith("632"):
            inn = "IN13"
        elif code.startswith("152") or code.startswith("155") or code.startswith("156"):
            inn = "IN09"
        else:
            inn = "IN08"
    elif "2.4" in u:
        inn = "—"
    elif "3.1" in u:
        inn = "IN07"
    elif "3.2" in u:
        inn = "IN07"
    elif "3.3" in u:
        inn = "IN05"
    elif "3.4" in u:
        inn = "—"
    return src, inn


def dinh_dang(ext: str) -> str:
    e = (ext or "").lower()
    if e == ".xml":
        return "XML"
    if e in {".xlsx", ".xls", ".xlsm"}:
        return "Cau_truc"
    if e == ".pdf":
        return "PDF-chua-trich"
    if e in {".jpg", ".jpeg", ".png", ".tif", ".tiff"}:
        return "Anh-chua-doc"
    if e in {".doc", ".docx"}:
        return "Word"
    return "Khac"


# ---------------------------------------------------------------------------
# 1. SCAN
# ---------------------------------------------------------------------------
print("SCAN")
manifest = []
skip_dir_names = {"sao-luu", "cau-thanh", "__pycache__"}
for dirpath, dirs, files in os.walk(lp(ROOT)):
    dirs[:] = [d for d in dirs if d.lower() not in skip_dir_names]
    for fn in files:
        if fn.startswith("KQ1.") and fn.endswith(".xlsx"):
            continue
        full = os.path.join(dirpath, fn)
        try:
            st = os.stat(full)
        except OSError:
            continue
        rel = os.path.relpath(full, lp(ROOT))
        src, inn = classify_file(rel, fn)
        ext = os.path.splitext(fn)[1].lower()
        h = ""
        if ext not in SKIP_HASH_EXT and st.st_size < 80_000_000:
            try:
                h = sha256(full)
            except Exception:
                h = ""
        manifest.append(
            {
                "rel": rel.replace("\\", "/"),
                "name": fn,
                "ext": ext,
                "size": st.st_size,
                "mtime": datetime.fromtimestamp(st.st_mtime).date(),
                "sha256": h,
                "nguon": src,
                "ma_in": inn,
            }
        )
print(" files", len(manifest))


# ---------------------------------------------------------------------------
# 2. XML
# ---------------------------------------------------------------------------
print("PARSE XML")


def tk_to_dict(tk) -> dict:
    return {
        "loai": tk.loai,
        "ma_tkhai": tk.ma_tkhai,
        "loai_tkhai": tk.loai_tkhai,
        "so_lan": tk.so_lan,
        "ky": tk.ky,
        "nam": tk.nam,
        "mst": tk.mst,
        "ten": tk.ten_nnt,
        "chi_tieu": tk.chi_tieu,
        "cdkt_cuoi": tk.bctc_cdkt_cuoi,
        "cdkt_dau": tk.bctc_cdkt_dau,
        "kqkd": tk.bctc_kqkd,
        "file": Path(tk.file).name if getattr(tk, "file", None) else "",
    }


def parse_any_xml(path: Path):
    tk = xt.parse_xml(path)
    if tk:
        return tk
    try:
        tree = ET.parse(lp(path))
        root = tree.getroot()
    except Exception:
        return None

    def ln(tag):
        return tag.split("}", 1)[-1] if tag else ""

    ma = xt._text(xt._find(root, "maTKhai"))
    ten_tk = xt._text(xt._find(root, "tenTKhai"))
    loai = xt.MA_LOAI.get(ma)
    fn = path.name.upper()
    if not loai:
        if fn.startswith("GTGT") or "GTGT" in (ten_tk or ""):
            loai = "gtgt"
        elif fn.startswith("TNDN") or "TNDN" in (ten_tk or ""):
            loai = "tndn"
        elif fn.startswith("BCTC") or "BCTC" in (ten_tk or "") or "B01" in (ten_tk or ""):
            loai = "bctc"
        elif fn.startswith("TNCN") or "TNCN" in (ten_tk or ""):
            loai = "tncn"
        elif fn.startswith("BC26") or fn.startswith("TB01") or "TB03" in fn:
            loai = "hd_ac"
        else:
            loai = "khac"
    ky_el = xt._find(root, "KyKKhaiThue")
    ky = xt._text(xt._child(ky_el, "kyKKhai")) if ky_el is not None else ""
    nam = xt._nam_from_ky(ky)
    nnt = xt._find(root, "NNT")
    tk = xt.ToKhai(
        loai=loai,
        ma_tkhai=ma,
        loai_tkhai=xt._text(xt._find(root, "loaiTKhai")) or "C",
        so_lan=int(xt._num(xt._text(xt._find(root, "soLan")))),
        ky=ky,
        nam=nam,
        tu_ngay=xt._text(xt._child(ky_el, "kyKKhaiTuNgay")) if ky_el is not None else "",
        den_ngay=xt._text(xt._child(ky_el, "kyKKhaiDenNgay")) if ky_el is not None else "",
        mst=xt._text(xt._child(nnt, "mst")) if nnt is not None else "",
        ten_nnt=xt._text(xt._child(nnt, "tenNNT")) if nnt is not None else "",
        chi_tieu={},
        file=str(path),
    )
    if loai == "bctc":
        tk.bctc_cdkt_cuoi = xt._ct_map(xt._find(root, "SoCuoiNam"))
        tk.bctc_cdkt_dau = xt._ct_map(xt._find(root, "SoDauNam"))
        kq = xt._find(root, "PL_KQHDSXKD")
        if kq is not None:
            tk.bctc_kqkd = xt._ct_map(xt._child(kq, "NamNay"))
            tk.chi_tieu = dict(tk.bctc_kqkd)
    elif loai == "gtgt":
        tk.chi_tieu = xt._ct_map(xt._find(root, "CTieuTKhaiChinh"))
        for code, val in xt._deep_ct(root).items():
            tk.chi_tieu.setdefault(code, val)
    elif loai == "tndn":
        pl = xt._find(root, "PLuc_03_1A_TNDN") or xt._find(root, "CTieuTKhaiChinh")
        tk.chi_tieu = xt._ct_map(pl) if pl is not None else xt._deep_ct(root)
        for code, val in xt._deep_ct(root).items():
            tk.chi_tieu.setdefault(code, val)
    elif loai == "tncn":
        tk.chi_tieu = xt._deep_ct(xt._find(root, "CTieuTKhaiChinh") or root)
    return tk


xml_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"
xml_rows = []
for fn in os.listdir(lp(xml_dir)):
    p = xml_dir / fn
    if not os.path.isfile(lp(p)):
        continue
    ext = p.suffix.lower()
    if ext == ".xml":
        tk = parse_any_xml(p)
        rec = {"file": fn, "parsed": tk is not None, "rel": f"1. DU LIEU THO NGANH THUE/1.1. To khai XML/{fn}", "size": os.path.getsize(lp(p))}
        if tk:
            rec.update(tk_to_dict(tk))
            rec["file"] = fn
        xml_rows.append(rec)
    elif ext == ".pdf":
        xml_rows.append({"file": fn, "parsed": False, "loai": "pdf", "rel": f"1. DU LIEU THO NGANH THUE/1.1. To khai XML/{fn}"})
print(" xml", len(xml_rows), "parsed", sum(1 for x in xml_rows if x.get("parsed")))
print(" loai", sorted({x.get("loai") for x in xml_rows if x.get("parsed")}))


def ky_norm(x: dict) -> str:
    ky = str(x.get("ky") or "").strip()
    nam = x.get("nam") or 0
    if re.fullmatch(r"[1-4]/20\d{2}", ky):
        q, y = ky.split("/")
        return f"Q{q}/{y}"
    if re.fullmatch(r"Q[1-4]/20\d{2}", ky, re.I):
        return ky.upper()
    if re.fullmatch(r"20\d{2}", ky):
        return ky
    if nam:
        return f"{ky}|{nam}" if ky else str(nam)
    return ky


def latest_xml(loai: str):
    groups = defaultdict(list)
    for x in xml_rows:
        if not x.get("parsed") or x.get("loai") != loai:
            continue
        groups[(x.get("nam"), ky_norm(x))].append(x)
    out = []
    for key, rows in groups.items():
        rows_sorted = sorted(
            rows,
            key=lambda r: (
                0 if str(r.get("loai_tkhai")).upper() in {"B", "2"} else 1,
                -int(r.get("so_lan") or 0),
                -int(r.get("size") or 0),
            ),
        )
        pick = dict(rows_sorted[0])
        pick["la_hieu_luc"] = True
        pick["so_ban"] = len(rows)
        pick["ky_norm"] = key[1]
        out.append(pick)
    return sorted(out, key=lambda r: (r.get("nam") or 0, str(r.get("ky_norm"))))


gtgt_all = latest_xml("gtgt")
tndn_eff = latest_xml("tndn")
bctc_eff = latest_xml("bctc")
tncn_qtt = latest_xml("tncn")
gtgt = [x for x in gtgt_all if (x.get("nam") or 0) in YEARS]
print(" GTGT ky", len(gtgt_all), "in ky", len(gtgt), [g.get("ky_norm") for g in gtgt])
print(" TNDN", [g.get("nam") for g in tndn_eff], "BCTC", [g.get("nam") for g in bctc_eff], "TNCN", [g.get("nam") for g in tncn_qtt])


# ---------------------------------------------------------------------------
# 3. HDDT
# ---------------------------------------------------------------------------
print("PARSE HDDT")
hddt = []
hddt_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu"
for fn in os.listdir(lp(hddt_dir)):
    if not fn.lower().endswith(".xlsx"):
        continue
    p = hddt_dir / fn
    chieu = "MUA" if "MUA" in fn.upper() else ("BAN" if "BAN" in fn.upper() else None)
    if chieu is None:
        continue
    wb = load_workbook(lp(p), read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, row in enumerate(rows[:15]):
            joined = " ".join(str(c or "").lower() for c in row)
            if ("số hóa đơn" in joined or "so hoa don" in joined) and ("ngày" in joined or "ngay" in joined):
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [str(c or "").strip() for c in rows[hdr_i]]

        def col(*keys):
            for i, h in enumerate(hdr):
                hl = h.lower()
                if all(k.lower() in hl for k in keys):
                    return i
            return None

        i_kh = col("ký hiệu") or col("ky hieu")
        i_so = col("số hóa đơn") or col("so hoa don")
        i_ng = col("ngày lập") or col("ngay lap") or col("ngày hđ")
        i_mst_b = col("mst người bán") or col("mst nguoi ban")
        i_ten_b = col("tên người bán") or col("ten nguoi ban")
        i_mst_m = col("mst người mua") or col("mst nguoi mua")
        i_ten_m = col("tên người mua") or col("ten nguoi mua")
        i_mh = col("mặt hàng") or col("ten hang")
        i_dvt = col("đơn vị tính") or col("dvt")
        i_sl = col("số lượng") or col("so luong")
        i_dg = col("đơn giá") or col("don gia")
        i_gt = col("tiền chưa thuế") or col("doanh thu") or col("giá trị")
        i_ts = col("thuế suất") or col("thue suat")
        i_thue = col("tiền thuế") or col("thue gtgt")
        i_tt = col("tổng thanh toán") or col("tong tien")
        i_tt_status = col("trạng thái")
        for row in rows[hdr_i + 1 :]:
            so_raw = row[i_so] if i_so is not None and i_so < len(row) else None
            so = so_hd_key(so_raw)
            if not so or so in {"0", "STT"}:
                continue
            ngay = ndate(row[i_ng] if i_ng is not None and i_ng < len(row) else None)
            nam = ngay.year if ngay else None
            mst_b = norm_mst(row[i_mst_b] if i_mst_b is not None and i_mst_b < len(row) else None)
            mst_m = norm_mst(row[i_mst_m] if i_mst_m is not None and i_mst_m < len(row) else None)
            stt = str(row[i_tt_status] or "").strip() if i_tt_status is not None and i_tt_status < len(row) else ""
            hddt.append(
                {
                    "chieu": chieu,
                    "nam": nam,
                    "quy": ((ngay.month - 1) // 3 + 1) if ngay else None,
                    "thang": ngay.month if ngay else None,
                    "ky_hieu": str(row[i_kh] or "").strip() if i_kh is not None and i_kh < len(row) else "",
                    "so_hd": str(so_raw or "").strip(),
                    "so_key": so,
                    "ngay": ngay,
                    "mst_nb": mst_b,
                    "ten_nb": str(row[i_ten_b] or "").strip() if i_ten_b is not None and i_ten_b < len(row) else "",
                    "mst_nm": mst_m,
                    "ten_nm": str(row[i_ten_m] or "").strip() if i_ten_m is not None and i_ten_m < len(row) else "",
                    "mst_dt": mst_b if chieu == "MUA" else mst_m,
                    "mat_hang": str(row[i_mh] or "").strip() if i_mh is not None and i_mh < len(row) else "",
                    "dvt": str(row[i_dvt] or "").strip() if i_dvt is not None and i_dvt < len(row) else "",
                    "sl": nnum(row[i_sl] if i_sl is not None and i_sl < len(row) else 0),
                    "don_gia": nnum(row[i_dg] if i_dg is not None and i_dg < len(row) else 0),
                    "gt": nnum(row[i_gt] if i_gt is not None and i_gt < len(row) else 0),
                    "thue_suat": row[i_ts] if i_ts is not None and i_ts < len(row) else None,
                    "thue": nnum(row[i_thue] if i_thue is not None and i_thue < len(row) else 0),
                    "tt": nnum(row[i_tt] if i_tt is not None and i_tt < len(row) else 0),
                    "trang_thai": stt,
                    "tinh_vao_tong": "Không" if "hủy" in stt.lower() or "huy" in stt.lower() else "Có",
                    "file": fn,
                }
            )
    wb.close()
n_mua = sum(1 for x in hddt if x["chieu"] == "MUA")
n_ban = sum(1 for x in hddt if x["chieu"] == "BAN")
miss_mst = sum(1 for x in hddt if not mst_base(x.get("mst_dt")))
print(" HDDT", len(hddt), "mua", n_mua, "ban", n_ban, "thieu MST", miss_mst)
print(" years mua", sorted({x["nam"] for x in hddt if x["chieu"] == "MUA"}))
print(" years ban", sorted({x["nam"] for x in hddt if x["chieu"] == "BAN"}))


# ---------------------------------------------------------------------------
# 4. BANG KE
# ---------------------------------------------------------------------------
print("PARSE BANG KE")
in02 = []
bk_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
bk_files = [f for f in os.listdir(lp(bk_dir)) if f.lower().endswith(".xlsx")]
bk_files.sort(key=lambda n: os.path.getsize(lp(bk_dir / n)), reverse=True)
primary_bk = bk_files[0] if bk_files else None
if primary_bk:
    p = bk_dir / primary_bk
    wb = load_workbook(lp(p), read_only=True, data_only=True)
    for sheet, chieu in (("MV", "MUA"), ("BR", "BAN")):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        year = None
        for row in ws.iter_rows(values_only=True):
            cells = list(row)
            joined = " ".join(str(c or "") for c in cells[:9])
            jl = joined.lower()
            m = re.search(r"năm\s+(20\d{2})", jl)
            if "kỳ tính thuế" in jl or "ky tinh thue" in jl:
                if m:
                    year = int(m.group(1))
                continue
            if "tổng cộng" in jl or jl.strip().startswith("tong cong"):
                continue
            stt = cells[0] if cells else None
            so = cells[1] if len(cells) > 1 else None
            ngay = ndate(cells[2] if len(cells) > 2 else None)
            if so_hd_key(so) in {"", "0", "STT"}:
                continue
            if not ngay and not (isinstance(stt, (int, float)) or str(stt or "").isdigit()):
                continue
            ten = str(cells[3] or "").strip() if len(cells) > 3 else ""
            mst_dt = norm_mst(cells[4] if len(cells) > 4 else None)
            mh = str(cells[5] or "").strip() if len(cells) > 5 else ""
            gt = nnum(cells[6] if len(cells) > 6 else 0)
            thue = nnum(cells[7] if len(cells) > 7 else 0)
            gc = str(cells[8] or "").strip() if len(cells) > 8 else ""
            nam = ngay.year if ngay else year
            in02.append(
                {
                    "chieu": chieu,
                    "nam": nam,
                    "quy": ((ngay.month - 1) // 3 + 1) if ngay else None,
                    "thang": ngay.month if ngay else None,
                    "ky_hieu": "",
                    "so_hd": str(so).strip() if so is not None else "",
                    "ngay": ngay,
                    "mst_goc": str(cells[4] or "").strip() if len(cells) > 4 else "",
                    "mst_chuan": mst_dt,
                    "ten": ten,
                    "mat_hang": mh,
                    "gt": gt,
                    "thue": thue,
                    "tt": gt + thue,
                    "ghi_chu": gc,
                    "file": primary_bk,
                }
            )
    wb.close()
n02_mua = sum(1 for x in in02 if x["chieu"] == "MUA")
n02_ban = sum(1 for x in in02 if x["chieu"] == "BAN")
miss02 = sum(1 for x in in02 if not mst_base(x.get("mst_chuan")))
print(" IN02", len(in02), "mua", n02_mua, "ban", n02_ban, "thieu MST", miss02, "file", primary_bk)


# ---------------------------------------------------------------------------
# 5. SO CHI TIET / NKC / KHO
# ---------------------------------------------------------------------------
print("PARSE BOOKS")


def iter_book_rows_xlsx(path):
    wb = load_workbook(lp(path), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdr_i = None
    hdr = []
    for i, row in enumerate(rows[:20]):
        joined = " ".join(str(c or "").lower() for c in row)
        if "phát sinh nợ" in joined or "phat sinh no" in joined:
            hdr_i = i
            hdr = [str(c or "").strip().lower() for c in row]
            break
    if hdr_i is None:
        return [], hdr, None
    idx = {}
    for i, h in enumerate(hdr):
        if "ngày chứng từ" in h or "ngay chung tu" in h:
            idx["ngay"] = i
        elif "ngày hạch toán" in h:
            idx.setdefault("ngay", i)
        elif h.startswith("số chứng từ") or "so chung tu" in h:
            idx["so"] = i
        elif "diễn giải" in h or "dien giai" in h:
            idx["dg"] = i
        elif h in {"tài khoản", "tai khoan"}:
            idx["tk"] = i
        elif "đối ứng" in h or "doi ung" in h:
            idx["du"] = i
        elif "phát sinh nợ" in h or "phat sinh no" in h:
            idx["no"] = i
        elif "phát sinh có" in h or "phat sinh co" in h:
            idx["co"] = i
        elif "tên đối tượng" in h:
            idx["ten"] = i
        elif "mã đối tượng" in h or "mst" in h:
            idx["mst"] = i
    out = []
    for row in rows[hdr_i + 1 :]:
        dg = str(row[idx["dg"]] or "").strip() if "dg" in idx and idx["dg"] < len(row) else ""
        if dg.lower() in {"số dư đầu kỳ", "so du dau ky", "số dư cuối kỳ", "so du cuoi ky"}:
            continue
        so = str(row[idx["so"]] or "").strip() if "so" in idx and idx["so"] < len(row) else ""
        if so.lower().startswith("số dòng"):
            continue
        ngay = ndate(row[idx["ngay"]] if "ngay" in idx and idx["ngay"] < len(row) else None)
        no = nnum(row[idx["no"]] if "no" in idx and idx["no"] < len(row) else 0)
        co = nnum(row[idx["co"]] if "co" in idx and idx["co"] < len(row) else 0)
        if not ngay and no == 0 and co == 0:
            continue
        out.append(
            {
                "ngay": ngay,
                "so": so,
                "dg": dg,
                "tk": str(row[idx["tk"]] or "").strip() if "tk" in idx and idx["tk"] < len(row) else "",
                "du": str(row[idx["du"]] or "").strip() if "du" in idx and idx["du"] < len(row) else "",
                "no": no,
                "co": co,
                "ten": str(row[idx["ten"]] or "").strip() if "ten" in idx and idx["ten"] < len(row) else "",
                "mst": norm_mst(row[idx["mst"]] if "mst" in idx and idx["mst"] < len(row) else None),
            }
        )
    return out, hdr, hdr_i


in08_agg = defaultdict(lambda: {"no": 0.0, "co": 0.0, "src": "", "file": "", "n": 0})
in10 = []
in13 = []
in09 = []
book_stats = defaultdict(int)

sct_root = ROOT / "2. HO SO NNT CUNG CAP" / "2.1. So ke toan tong hop va chi tiet"
for dirpath, dirs, files in os.walk(lp(sct_root)):
    for fn in files:
        ext = os.path.splitext(fn)[1].lower()
        full = os.path.join(dirpath, fn)
        rel = os.path.relpath(full, lp(ROOT)).replace("\\", "/")
        year_m = re.search(r"20\d{2}", dirpath + " " + fn)
        year_guess = int(year_m.group(0)) if year_m else None
        codes = re.findall(r"(?<!\d)(\d{3,5})(?!\d)", fn)
        code = codes[-1] if codes else ""
        if ext == ".xlsx" and ("NHAT_KY" in fn.upper() or "NHẬT KÝ" in fn.upper() or "NHAT KY" in fn.upper()):
            rows, _, _ = iter_book_rows_xlsx(full)
            book_stats["nkc_xlsx_rows"] += len(rows)
            for rec in rows:
                nam = rec["ngay"].year if rec["ngay"] else year_guess
                thang = rec["ngay"].month if rec["ngay"] else None
                tk = rec["tk"] or "NKC"
                key = (nam, thang, tk)
                in08_agg[key]["no"] += rec["no"]
                in08_agg[key]["co"] += rec["co"]
                in08_agg[key]["src"] = "NKC"
                in08_agg[key]["file"] = fn
                in08_agg[key]["n"] += 1
            continue
        if ext == ".xlsx" and code:
            rows, _, _ = iter_book_rows_xlsx(full)
            book_stats[f"sct_{code}"] += len(rows)
            for rec in rows:
                nam = rec["ngay"].year if rec["ngay"] else year_guess
                thang = rec["ngay"].month if rec["ngay"] else None
                tk = rec["tk"] or code
                key = (nam, thang, tk)
                in08_agg[key]["no"] += rec["no"]
                in08_agg[key]["co"] += rec["co"]
                in08_agg[key]["src"] = "SCTK"
                in08_agg[key]["file"] = fn
                in08_agg[key]["n"] += 1
                if str(tk).startswith("111") and rec["co"] > 0:
                    in10.append(
                        {
                            "ngay": rec["ngay"],
                            "nam": nam,
                            "so": rec["so"],
                            "du": rec["du"],
                            "mst": rec["mst"],
                            "ten": rec["ten"],
                            "dg": rec["dg"],
                            "tien": rec["co"],
                            "file": fn,
                        }
                    )
                if str(tk).startswith("511") or str(tk).startswith("632"):
                    in13.append(
                        {
                            "nam": nam,
                            "thang": thang,
                            "ngay": rec["ngay"],
                            "so": rec["so"],
                            "tk": tk,
                            "du": rec["du"],
                            "dg": rec["dg"],
                            "mst": rec["mst"],
                            "co": rec["co"],
                            "no": rec["no"],
                            "file": fn,
                        }
                    )
        elif ext == ".xls":
            book_stats["xls_files"] += 1

kho_root = ROOT / "2. HO SO NNT CUNG CAP" / "2.2. Kho - Chi phi - Gia thanh"
for dirpath, dirs, files in os.walk(lp(kho_root)):
    for fn in files:
        ext = os.path.splitext(fn)[1].lower()
        if ext != ".xlsx":
            continue
        fu = fn.upper()
        if "TỒN KHO" not in fn and "TON KHO" not in fu and "TONG HOP TON" not in fu.replace("Ồ", "O").replace("Ộ", "O"):
            if "TỔNG HỢP TỒN" not in fn and "CHI TIẾT VẬT TƯ" not in fn and "CHI TIET VAT TU" not in fu:
                continue
        full = os.path.join(dirpath, fn)
        year_m = re.search(r"20\d{2}", fn)
        year_guess = int(year_m.group(0)) if year_m else None
        try:
            wb = load_workbook(lp(full), read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
        except Exception:
            continue
        # take numeric totals from last non-empty row-ish
        ton_ck_tt = nhap_tt = xuat_tt = None
        for row in rows:
            joined = " ".join(str(c or "").lower() for c in row[:12])
            nums = [nnum(c) for c in row if isinstance(c, (int, float))]
            if "tổng" in joined or "cong" in joined:
                if nums:
                    ton_ck_tt = nums[-1]
            if not nums:
                continue
        # fallback: max money-like
        if ton_ck_tt is None:
            alln = []
            for row in rows:
                for c in row:
                    if isinstance(c, (int, float)) and c > 1000:
                        alln.append(float(c))
            if alln:
                ton_ck_tt = max(alln)
        in09.append({"nam": year_guess, "nhom": "HH/NVL", "ton_ck_tt": ton_ck_tt, "file": fn})

print(" IN08 keys", len(in08_agg), "IN10 chi", len(in10), "IN13", len(in13), "IN09", len(in09), "stats", dict(book_stats))


# ---------------------------------------------------------------------------
# 6. F0 / F1 / CA
# ---------------------------------------------------------------------------
print("PARSE CVCBRR")


def load_f0():
    out = []
    if not os.path.exists(lp(CVCBRR)):
        print(" CVCBRR missing")
        return out
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    if "F0" not in wb.sheetnames:
        wb.close()
        return out
    ws = wb["F0"]
    rows = list(ws.iter_rows(values_only=True))
    for row in rows[6:]:
        mst = norm_mst(row[1] if len(row) > 1 else None)
        if not mst or not mst[0].isdigit():
            continue
        out.append(
            {
                "mst": mst,
                "base": mst_base(mst),
                "ten": str(row[2] or "").strip() if len(row) > 2 else "",
                "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                "ngay_cv": ndate(row[4] if len(row) > 4 else None),
                "dvbh": str(row[5] or "").strip() if len(row) > 5 else "",
                "nd": str(row[6] or "").strip()[:240] if len(row) > 6 else "",
            }
        )
    wb.close()
    return out


def load_f1_nnt():
    hits = []
    if not os.path.exists(lp(CVCBRR)):
        return hits
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    if "F1" not in wb.sheetnames:
        wb.close()
        return hits
    ws = wb["F1"]
    for row in ws.iter_rows(values_only=True):
        mst_mua = mst_base(row[8] if len(row) > 8 else None)
        if mst_mua == MST:
            hits.append(
                {
                    "mst_ban": norm_mst(row[1] if len(row) > 1 else None),
                    "ten_ban": str(row[2] or "").strip() if len(row) > 2 else "",
                    "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                    "mst_mua": MST,
                }
            )
    wb.close()
    return hits


def load_ca_nnt():
    hits = []
    if not os.path.exists(lp(CA)):
        print(" CA file missing")
        return hits
    wb = load_workbook(lp(CA), read_only=True, data_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, row in enumerate(rows[:12]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "mst" in joined:
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [str(c or "").strip().lower() for c in rows[hdr_i]]
        i_mst = next((i for i, h in enumerate(hdr) if "mst" in h), None)
        if i_mst is None:
            continue
        for row in rows[hdr_i + 1 :]:
            if mst_base(row[i_mst] if i_mst < len(row) else None) == MST:
                hits.append({"sheet": name, "mst": MST})
    wb.close()
    return hits


f0 = load_f0()
f0_map = {}
for r in f0:
    if r["base"] and r["base"] not in f0_map:
        f0_map[r["base"]] = r
nnt_f0 = f0_map.get(MST)
nnt_f1 = load_f1_nnt()
nnt_ca = load_ca_nnt()
print(" F0", len(f0), "NNT F0", bool(nnt_f0), "NNT F1", len(nnt_f1), "NNT CA", len(nnt_ca))
cvcbrr_mtime = datetime.fromtimestamp(os.path.getmtime(lp(CVCBRR))).date() if os.path.exists(lp(CVCBRR)) else None
if cvcbrr_mtime and (TODAY - cvcbrr_mtime).days > 30:
    print(" WARN CVCBRR older than 30 days", cvcbrr_mtime)

# ---------------------------------------------------------------------------
# HAS / LEVEL
# ---------------------------------------------------------------------------
HAS = {
    "IN01": n_mua + n_ban > 0,
    "IN02": len(in02) > 0,
    "IN03": len(gtgt) > 0,
    "IN04": len(bctc_eff) + len(tndn_eff) > 0,
    "IN05": len(tncn_qtt) > 0,
    "IN07": True,  # nguon chi dinh CVCBRR + danh ba, se lap o GĐ2
    "IN08": len(in08_agg) > 0,
    "IN09": len(in09) > 0 or any(m["ma_in"] == "IN09" for m in manifest),
    "IN10": len(in10) > 0 or any(m["ma_in"] == "IN10" for m in manifest),
    "IN11": False,
    "IN12": False,
    "IN13": len(in13) > 0 or any(m["ma_in"] == "IN13" for m in manifest),
    "IN14": any(m["ma_in"] == "IN14" for m in manifest),
}
if HAS["IN01"] and HAS["IN03"] and HAS["IN04"]:
    LEVEL = "A"
if LEVEL == "A" and HAS["IN02"]:
    LEVEL = "B"
if LEVEL == "B" and HAS["IN08"] and HAS["IN13"]:
    LEVEL = "C"
if LEVEL == "C" and HAS["IN09"]:
    LEVEL = "D"
if LEVEL == "D" and HAS["IN10"] and HAS["IN11"]:
    LEVEL = "E"
DN_STATUS = "Một phần"
nnt_f0_txt = (
    f"NNT thuộc F0 - Công văn cảnh báo {nnt_f0.get('cv')}" if nnt_f0 else "NNT không có trên F0"
)
KY_SUY_RA = (
    "K12: kỳ 2015-2025 suy từ XML GTGT/TNDN/BCTC (2015-2025) và bảng kê NNT 2014-2025; "
    "HĐĐT portal chỉ 2022-2025; sổ DN 2018, 2020-2025 (thiếu 2019). "
    "Q4/2014 GTGT còn PDF chưa trích, không đưa vào kỳ chính. "
    f"{nnt_f0_txt}. NNT {'có' if nnt_f1 else 'không'} trên F1, {'có' if nnt_ca else 'không'} trên sổ Công an."
)
print(" LEVEL", LEVEL, "HAS", {k: v for k, v in HAS.items()})

# ---------------------------------------------------------------------------
# 7. COPY TEMPLATES
# ---------------------------------------------------------------------------
os.makedirs(lp(OUT), exist_ok=True)
os.makedirs(lp(VAULT_OUT), exist_ok=True)
files = {
    "kq11": (MAU / "KQ1.1_Kiem_ke_CQT.xlsx", OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_{VER}.xlsx"),
    "kq12": (MAU / "KQ1.2_Kiem_ke_ho_so_NNT.xlsx", OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_{VER}.xlsx"),
    "kq13": (MAU / "KQ1.3_Danh_gia_ho_so.xlsx", OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx"),
}
for src, dst in files.values():
    shutil.copy2(lp(src), lp(dst))
    print("copied", dst.name)


def banner_in00(ws):
    ws.cell(1, 2).value = f"{TEN} - MST {MST} - kỳ {KY} - {TODAY.strftime('%d/%m/%Y')} - {VER}"
    tnr(ws.cell(1, 2), bold=True)


# ---------------------------------------------------------------------------
# KQ1.1
# ---------------------------------------------------------------------------
print("FILL KQ1.1")
wb = load_workbook(lp(files["kq11"][1]))
ws = wb["IN00_Manifest"]
cqt_files = [m for m in manifest if m["nguon"] == "CQT"]
clear_from(ws, 3)
for i, m in enumerate(cqt_files, 1):
    n_dong = None
    if m.get("ma_in") == "IN01":
        n_dong = n_mua if "MUA" in m["name"].upper() else n_ban
    write_row(
        ws,
        2 + i,
        [
            f"N{i:03d}",
            m.get("ma_in"),
            "CQT",
            dinh_dang(m.get("ext")),
            m.get("mtime"),
            2014 if "2014" in m.get("name", "") else 2015,
            2025,
            m.get("sha256"),
            "Có" if "bo sung" in m["name"].lower() or "bổ sung" in m["name"].lower() else "Không",
            "Có",
            n_dong,
            0,
            "Giữ nguyên",
            "",
            m.get("rel"),
        ],
        dates={5},
        nums={11, 12},
    )
banner_in00(ws)
sub_placeholders(ws)

ws = wb["IN01_HDDT"]
clear_from(ws, 3)
for i, r in enumerate(hddt, 1):
    write_row(
        ws,
        2 + i,
        [
            r.get("chieu"),
            r.get("nam"),
            r.get("quy"),
            r.get("thang"),
            r.get("ky_hieu"),
            r.get("so_hd"),
            r.get("ngay"),
            r.get("mst_nb"),
            r.get("ten_nb"),
            r.get("mst_nm"),
            r.get("ten_nm"),
            r.get("mst_dt"),
            r.get("mat_hang") or "",
            r.get("dvt") or "",
            r.get("sl") or None,
            r.get("don_gia") or None,
            r.get("gt"),
            r.get("thue_suat"),
            r.get("thue"),
            r.get("tt"),
            r.get("trang_thai") or "Hợp lệ",
            "",
            r.get("tinh_vao_tong"),
            "",
            f"{r.get('ky_hieu')}|{r.get('so_key')}|{r.get('mst_dt')}",
            r.get("file"),
        ],
        nums={15, 16, 17, 19, 20},
        dates={7},
    )

ws = wb["IN03_ToKhai_GTGT"]
clear_from(ws, 3)
gtgt_write = [x for x in xml_rows if x.get("parsed") and x.get("loai") == "gtgt"]
gtgt_write.sort(key=lambda x: (x.get("nam") or 0, str(x.get("ky"))))
eff_keys = {(x.get("nam"), ky_norm(x)) for x in gtgt_all}
for i, x in enumerate(gtgt_write, 1):
    c = x.get("chi_tieu") or {}
    kn = ky_norm(x)
    hieu_luc = "Có" if (x.get("nam"), kn) in {(g.get("nam"), g.get("ky_norm")) for g in gtgt_all} and str(x.get("loai_tkhai")).upper() in {"B", "2", "C", "1", ""} else "Không"
    # mark latest of group
    is_eff = any(g.get("file") == x.get("file") for g in gtgt_all)
    write_row(
        ws,
        2 + i,
        [
            x.get("nam"),
            kn,
            "Bổ sung" if str(x.get("loai_tkhai")).upper() in {"B", "2"} else "Chính thức",
            x.get("so_lan") or 0,
            "Có" if is_eff else "Không",
            None,
            c.get("22"),
            c.get("23"),
            c.get("24"),
            c.get("25"),
            c.get("26"),
            c.get("29"),
            c.get("30"),
            c.get("31"),
            c.get("32"),
            c.get("33"),
            c.get("34"),
            c.get("35"),
            c.get("36"),
            c.get("37"),
            c.get("38"),
            c.get("39"),
            c.get("40"),
            c.get("41"),
            c.get("42"),
            c.get("43"),
            x.get("file"),
        ],
        nums=set(range(7, 27)),
    )

ws = wb["IN04_BCTC_TNDN"]
clear_from(ws, 3)
r = 3
pairs_def = [
    ("KQKD", "01", "Doanh thu bán hàng và cung cấp dịch vụ", "01"),
    ("KQKD", "11", "Giá vốn hàng bán", "11"),
    ("KQKD", "20", "Lợi nhuận gộp", "20"),
    ("KQKD", "21", "Doanh thu hoạt động tài chính", "21"),
    ("KQKD", "22", "Chi phí tài chính", "22"),
    ("KQKD", "24", "Chi phí quản lý doanh nghiệp", "24"),
    ("KQKD", "50", "Lợi nhuận trước thuế", "50"),
    ("KQKD", "60", "Lợi nhuận sau thuế", "60"),
    ("CDKT", "110", "Tiền và tương đương tiền", "110"),
    ("CDKT", "131", "Phải thu khách hàng", "131"),
    ("CDKT", "140", "Hàng tồn kho (ct140 TT133)", "140"),
    ("CDKT", "141", "Hàng tồn kho (ct141)", "141"),
    ("CDKT", "211", "TSCĐ hữu hình", "211"),
    ("CDKT", "311", "Nợ ngắn hạn", "311"),
    ("CDKT", "331", "Phải trả người bán", "331"),
    ("CDKT", "411", "Vốn góp chủ sở hữu", "411"),
]
bctc_by_year = {x.get("nam"): x for x in bctc_eff}
tndn_by_year = {x.get("nam"): x for x in tndn_eff}
for y in YEARS:
    b = bctc_by_year.get(y)
    if b:
        k = b.get("kqkd") or b.get("chi_tieu") or {}
        cd = b.get("cdkt_cuoi") or {}
        mau = f"maTKhai {b.get('ma_tkhai')}"
        for bao, ma, tenct, code in pairs_def:
            src = k if bao == "KQKD" else cd
            so = src.get(code)
            write_row(ws, r, [y, bao, ma, tenct, so, None, mau, b.get("file")], nums={5})
            r += 1
    tn = tndn_by_year.get(y)
    if tn:
        c = tn.get("chi_tieu") or {}
        for ma, tenct in [("04", "Doanh thu QTT TNDN [04]"), ("13", "Giá vốn QTT TNDN [13]"), ("16", "Thuế TNDN [16]"), ("18", "LNST [18]"), ("A1", "Doanh thu [A1] nếu có")]:
            if c.get(ma) is not None:
                write_row(ws, r, [y, "TNDN", ma, tenct, c.get(ma), None, "03/TNDN XML CQT", tn.get("file")], nums={5})
                r += 1
    if not b and not tn:
        write_row(ws, r, [y, "THIEU", "", "Chưa có XML BCTC/TNDN năm này trong thư mục 1.1", None, None, "", "KHÔNG_ĐỦ_DỮ_LIỆU"])
        r += 1

ws = wb["IN05_NLD_TNCN"]
clear_from(ws, 3)
r = 3
for x in tncn_qtt:
    y = x.get("nam")
    c = x.get("chi_tieu") or {}
    has_so = any(v not in (None, 0, 0.0, "") for v in c.values())
    write_row(
        ws,
        r,
        [
            "QTT",
            y,
            f"Quyết toán TNCN {y}",
            "",
            MST,
            "05/QTT-TNCN",
            "Tiền lương",
            c.get("20") or c.get("10") or c.get("26"),
            c.get("29") or c.get("27") or c.get("40"),
            "Không",
            None,
            "",
            "",
            "XML QTT; thư mục 1.4 trống, chưa có 05-1/05-2" if True else "",
            "Không",
            x.get("file") and "05/QTT-TNCN",
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            x.get("file"),
        ],
        nums={8, 9},
    )
    r += 1
write_row(ws, r, ["GHICHU", None, "Không có phụ lục 05-1/05-2 NLĐ trong thư mục 1.4", "", "", "", "", None, None, "", None, "", "", "IN05 chỉ có QTT XML", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])

ws = wb["IN14_TMS_Nghia_vu"]
clear_from(ws, 3)
r = 3
tms_files = [m for m in manifest if m.get("ma_in") == "IN14"]
for m in tms_files:
    write_row(
        ws,
        r,
        [
            MST,
            "FILE_TMS",
            "",
            KY,
            None,
            None,
            m.get("mtime"),
            None,
            None,
            None,
            None,
            "",
            None,
            m.get("name"),
            "TMS",
            m.get("rel"),
        ],
        dates={7},
    )
    r += 1

apply_tnr(wb["IN00_Manifest"])
autowidth(wb["IN00_Manifest"])
autowidth(wb["IN03_ToKhai_GTGT"])
autowidth(wb["IN04_BCTC_TNDN"])
autowidth(wb["IN05_NLD_TNCN"])
autowidth(wb["IN14_TMS_Nghia_vu"], max_row=200)
wb.save(lp(files["kq11"][1]))
wb.close()
print(" KQ1.1 saved")

# ---------------------------------------------------------------------------
# KQ1.2
# ---------------------------------------------------------------------------
print("FILL KQ1.2")
wb = load_workbook(lp(files["kq12"][1]))
ws = wb["IN00_Manifest"]
dn_files = [m for m in manifest if m["nguon"] in {"DN", "TRA_CUU", "BB"}]
clear_from(ws, 3)
for i, m in enumerate(dn_files, 1):
    write_row(
        ws,
        2 + i,
        [
            f"D{i:03d}",
            m.get("ma_in"),
            m.get("nguon"),
            dinh_dang(m.get("ext")),
            m.get("mtime"),
            2014 if "2014" in m.get("name", "") else 2018,
            2025,
            m.get("sha256"),
            "Không",
            "Có",
            None,
            0,
            "Giữ nguyên",
            "",
            m.get("rel"),
        ],
        dates={5},
        nums={11, 12},
    )
banner_in00(ws)
sub_placeholders(ws)

ws = wb["IN02_BangKe"]
clear_from(ws, 3)
for i, rec in enumerate(in02, 1):
    write_row(
        ws,
        2 + i,
        [
            rec["chieu"],
            rec["nam"],
            rec["quy"],
            rec["thang"],
            rec["ky_hieu"],
            rec["so_hd"],
            rec["ngay"],
            rec["mst_goc"],
            rec["mst_chuan"],
            rec["ten"],
            rec["mat_hang"],
            "",
            None,
            rec["gt"],
            None,
            rec["thue"],
            rec["tt"],
            "",
            "Có",
            "Có",
            None,
            None,
            rec.get("ghi_chu"),
            rec["file"],
        ],
        nums={14, 16, 17},
        dates={7},
    )

ws = wb["IN08_So_KeToan"]
clear_from(ws, 3)
items = sorted(in08_agg.items(), key=lambda x: (x[0][0] or 0, x[0][1] or 0, str(x[0][2])))
if items:
    for i, ((nam, thang, tk), v) in enumerate(items, 1):
        write_row(ws, 2 + i, [nam, thang, tk, v["no"], v["co"], None, None, None, v["src"], v["file"]], nums={4, 5})
else:
    put(ws, 3, 10, "Có file sổ 2018/2020-2025 nhưng chưa gộp được dòng (xls cũ). Không điền 0 giả.", wrap=True, fill=FILL_YEL)

ws = wb["IN09_Kho_GiaThanh"]
clear_from(ws, 3)
if in09:
    for i, rec in enumerate(in09, 1):
        write_row(ws, 2 + i, [rec["nam"], rec["nhom"], None, None, None, None, None, None, None, rec.get("ton_ck_tt"), None, None, None, None, None, rec["file"]], nums={10})
else:
    put(ws, 3, 16, "Có file tồn kho 2018-2025 tại 2.2; chưa bóc được dòng số lượng. Không điền 0 giả.", wrap=True, fill=FILL_YEL)

ws = wb["IN10_Chi_tien_mat"]
clear_from(ws, 3)
if in10:
    for i, rec in enumerate(in10, 1):
        write_row(
            ws,
            2 + i,
            [rec["ngay"], rec["nam"], rec["so"], rec["du"], rec["mst"], rec["ten"], rec["dg"], rec["tien"], "Chi 111", "Chưa rõ", "", rec["file"]],
            nums={8},
            dates={1},
        )
else:
    put(ws, 3, 12, "Có sổ 1111 (xlsx 2022-2025) nhưng chưa tách được phiếu chi. Không điền 0 giả.", wrap=True, fill=FILL_YEL)

ws = wb["IN11_Sao_ke_NH"]
clear_from(ws, 3)
put(ws, 3, 12, "Chưa nhận", wrap=True, fill=FILL_YEL)
put(ws, 3, 14, "Thư mục 2.3 trống. Có sổ 1121 (IN08) nhưng không phải sao kê/UNC.")

ws = wb["IN12_GDLK"]
clear_from(ws, 3)
put(ws, 3, 9, "Chưa nhận phụ lục GDLK NĐ 132/2020", wrap=True, fill=FILL_YEL)

ws = wb["IN13_So_511_632"]
clear_from(ws, 3)
if in13:
    for i, rec in enumerate(in13, 1):
        write_row(
            ws,
            2 + i,
            [
                rec["nam"],
                rec["thang"],
                rec["ngay"],
                rec["so"],
                rec["tk"],
                rec["du"],
                "",
                rec["dg"],
                rec["mst"],
                None,
                rec["co"],
                rec["no"],
                "",
                "",
                "",
                None,
                None,
                None,
                None,
                rec["file"],
            ],
            nums={11, 12},
            dates={3},
        )
else:
    put(ws, 3, 20, "Có file 5111/6321 nhưng chưa bóc dòng. Không điền 0 giả.", wrap=True, fill=FILL_YEL)

apply_tnr(wb["IN00_Manifest"])
autowidth(wb["IN00_Manifest"])
autowidth(wb["IN08_So_KeToan"])
autowidth(wb["IN09_Kho_GiaThanh"])
autowidth(wb["IN10_Chi_tien_mat"], cap=50)
autowidth(wb["IN13_So_511_632"], cap=45, max_row=2000)
wb.save(lp(files["kq12"][1]))
wb.close()
print(" KQ1.2 saved")

# ---------------------------------------------------------------------------
# KQ1.3
# ---------------------------------------------------------------------------
print("FILL KQ1.3")
wb = load_workbook(lp(files["kq13"][1]))
ws = wb["KQ1_Kiem_ke"]
ws["C2"] = MST
ws["J2"] = TEN
ws["C3"] = KY
ws["J3"] = LEVEL
ws["C5"] = VER
ws["J5"] = MAU_VER
ws["C6"] = TODAY
ws["C6"].number_format = DATE_FMT
ws["J6"] = DN_STATUS
ws["C7"] = "KQ2.1, KQ2.2, KQ3"
ws["J7"] = KY_SUY_RA
latest = max((m["mtime"] for m in manifest), default=TODAY)
ws["J4"] = latest
ws["J4"].number_format = DATE_FMT


def set_in_row(row, dinh, ky_phu, ky_thieu, so_dong, thieu_mst, bs, muc, lydo, tt, viec, ghichu):
    vals = {
        6: dinh,
        7: ky_phu,
        8: ky_thieu,
        9: so_dong,
        10: thieu_mst,
        11: bs,
        12: muc,
        13: lydo,
        14: tt,
        15: viec,
        17: ghichu,
    }
    for c, v in vals.items():
        ws.cell(row, c).value = v
        tnr(ws.cell(row, c), bold=(c == 12))
        if c == 12:
            if v == "Đủ":
                ws.cell(row, c).fill = FILL_GRN
            elif v in {"Mỏng", "PDF chưa trích"}:
                ws.cell(row, c).fill = FILL_YEL
            else:
                ws.cell(row, c).fill = FILL_RED
        if c == 9 and isinstance(v, (int, float)):
            ws.cell(row, c).number_format = NUM


gtgt_years = sorted({x.get("nam") for x in gtgt if x.get("nam")})
gtgt_missing = []
for y in YEARS:
    have = {str(x.get("ky_norm")) for x in gtgt if x.get("nam") == y}
    for q in range(1, 5):
        k = f"Q{q}/{y}"
        if k not in have and not any(str(x.get("ky_norm") or "").startswith(str(y)) and f"Q{q}" in str(x.get("ky_norm")) for x in gtgt):
            # also accept  q/y
            if not any(str(x.get("ky_norm")) == f"{q}/{y}" for x in gtgt):
                gtgt_missing.append(k)

bctc_years = sorted({x.get("nam") for x in bctc_eff if x.get("nam")})
tndn_years = sorted({x.get("nam") for x in tndn_eff if x.get("nam")})
tncn_years = sorted({x.get("nam") for x in tncn_qtt if x.get("nam")})
in08_years = sorted({k[0] for k in in08_agg if k[0]})
in13_years = sorted({r["nam"] for r in in13 if r.get("nam")})
in10_years = sorted({r["nam"] for r in in10 if r.get("nam")})
in09_years = sorted({r["nam"] for r in in09 if r.get("nam")})

set_in_row(10, "Excel", "2022-2025", "HĐĐT trước 2022 (portal không có)", n_mua + n_ban, miss_mst, "Không", "Đủ" if n_mua + n_ban > 12 else "Mỏng", f"Portal mua {n_mua} + bán {n_ban}", "Đã chuẩn hóa", "MST, khóa HĐ, Tinh_vao_tong, lọc hủy", "CQT 1.2")
set_in_row(11, "Excel", "2014-2025", "Ký hiệu HĐ thường trống (HĐ giấy)", len(in02), miss02, "Không", "Đủ" if len(in02) > 12 else "Mỏng", f"Bảng kê NNT MV {n02_mua} + BR {n02_ban}; đã bỏ Tổng cộng", "Đã chuẩn hóa", "MST chuẩn, lọc tổng cộng", f"DN 2.5 {primary_bk}")
set_in_row(12, "XML", "2015-2025 theo quý", "; ".join(gtgt_missing[:12]) or "", len(gtgt_write), 0, "Có", "Đủ" if len(gtgt) >= 36 else "Mỏng", f"{len(gtgt)} kỳ hiệu lực 2015-2025; PDF Q4/2014 chưa trích", "Đã chọn bản hiệu lực", "Bổ sung đánh dấu La_ban_hieu_luc", "CQT 1.1")
set_in_row(13, "XML", ",".join(str(y) for y in bctc_years) or "", ",".join(str(y) for y in YEARS if y not in bctc_years and y not in tndn_years) or "", len(bctc_eff) + len(tndn_eff), 0, "Có", "Đủ" if bctc_years else "Thiếu", f"BCTC {bctc_years}; TNDN {tndn_years}", "Đã trích chỉ tiêu", "HTK đọc ct140 và ct141, xác định mẫu từ XML", "CQT 1.1")
set_in_row(14, "XML", ",".join(str(y) for y in tncn_years) or "", "Phụ lục 05-1/05-2; thư mục 1.4 trống", len(tncn_qtt), 0, "Có", "Mỏng", f"{len(tncn_qtt)} QTT XML, chưa có danh sách NLĐ", "Đã nạp dòng QTT", "Không điền 0 giả cho NLĐ", "CQT 1.1")
set_in_row(15, "Excel sổ CVCBRR+CA+danh bạ", iso(cvcbrr_mtime) or "", "IN07 lập ở GĐ2 từ IN02", len(in02), 0, "Không", "Đủ" if os.path.exists(lp(CVCBRR)) else "Thiếu", f"{nnt_f0_txt}; F1 hits {len(nnt_f1)}; CA hits {len(nnt_ca)}; có Danh ba NNT.xls", "Chưa lập bảng IN07", "GĐ2 sẽ đối chiếu 100% NCC", "Nguồn chỉ định")
set_in_row(16, "Excel", ",".join(str(y) for y in in08_years) or "2018,2020-2025 file", "2019; xls 2018/2020/2021 chưa gộp dòng", len(in08_agg), 0, "Không", "Đủ" if len(in08_agg) > 12 else "Mỏng", f"{len(in08_agg)} khóa năm-tháng-TK từ NKC/SCTK xlsx", "Đã gộp tháng", "xls cũ ghi nhận file, không điền 0", "DN 2.1")
set_in_row(17, "Excel", ",".join(str(y) for y in in09_years) or "2018-2025 file", "2019; chưa có bảng giá thành 154/định mức", len(in09), 0, "Không", "Mỏng" if in09 else "Đủ", f"{len(in09)} file tổng hợp tồn kho; chưa bóc SL đầy đủ", "Nhận file", "GĐ3 sẽ đối chiếu TC3", "DN 2.2")
set_in_row(18, "Excel", ",".join(str(y) for y in in10_years) or "", "2019; xls 2018/2020/2021", len(in10), 0, "Không", "Đủ" if len(in10) > 12 else "Mỏng", f"{len(in10)} phiếu chi TK 1111 (PS Có)", "Đã lọc chi", "Ngưỡng >=20tr / >=5tr áp GĐ2/GĐ3", "DN 2.1")
set_in_row(19, "", "", "Chưa nhận sao kê NH / UNC", 0, 0, "Không", "Chưa nhận", "Thư mục 2.3 trống", "Để trống", "Không điền 0; sổ 1121 không thay sao kê", "DN")
set_in_row(20, "", "", "Chưa nhận PL GDLK", 0, 0, "Không", "Chưa nhận", "Không thấy phụ lục NĐ 132", "Để trống", "Không điền 0", "DN")
set_in_row(21, "Excel", ",".join(str(y) for y in in13_years) or "", "2019; xls cũ", len(in13), 0, "Không", "Đủ" if len(in13) > 12 else "Mỏng", f"{len(in13)} dòng 511/632 xlsx", "Đã nạp dòng", "Khóa ghép GĐ3 theo skill 511-632", "DN 2.1")
set_in_row(22, "Excel TMS", "2015-2026 file 1.3", "Chưa ánh xạ hết cột nợ thuế dòng", len(tms_files), 0, "Không", "Mỏng", f"{len(tms_files)} file TMS/PTDN/nghĩa vụ", "Nhận file", "LS01/LS02 đọc IN14", "CQT 1.3")

n_du = n_mong = n_thieu = n_chua = 0
for rr in range(10, 23):
    muc = str(ws.cell(rr, 12).value or "")
    if muc == "Đủ":
        n_du += 1
    elif muc == "Mỏng":
        n_mong += 1
    elif muc == "Thiếu":
        n_thieu += 1
    else:
        n_chua += 1
ws["C4"] = f"{n_du} Đủ | {n_mong} Mỏng | {n_thieu} Thiếu | {n_chua} Chưa nhận"

ws = wb["KQ1_Kha_nang_phan_tich"]
ws["C2"] = MST
ws["E2"] = KY
ws["C4"] = VER
ws["E4"] = MAU_VER
ws["C5"] = TODAY
ws["C5"].number_format = DATE_FMT
ws["E5"] = f"Mức {LEVEL} - có sổ DN một phần, thiếu sao kê NH (IN11) và GDLK (IN12)"

PART_OK = {"HD01", "HD03", "HD04", "HD08", "TT01", "GT01", "GT03", "GT04", "GT05", "TN02", "GT02", "DQ03"}
run_now = part = block = 0
for r in range(10, (ws.max_row or 10) + 1):
    ma = ws.cell(r, 1).value
    if not ma:
        continue
    ma = str(ma).strip()
    need = str(ws.cell(r, 4).value or "")
    need_ins = re.findall(r"IN\d{2}", need)
    missing = sorted({i for i in need_ins if not HAS.get(i, False)})
    if ma == "HD07":
        kn, thieu, inn, goi = "TREO", "Thiếu tham chiếu ngành (K13)", "K13", "Chạy sơ bộ khi có số ngành"
        block += 1
    elif ma in PART_OK and missing:
        kn, thieu, inn, goi = "Chạy một phần", ",".join(missing), ",".join(missing), "Chạy trên nguồn đã có, ghi chú thiếu " + ",".join(missing)
        part += 1
    elif not missing:
        kn, thieu, inn, goi = "Chạy được ngay", "", "", ""
        run_now += 1
    elif any(HAS.get(i, False) for i in need_ins) and missing:
        kn, thieu, inn, goi = "Chạy một phần", ",".join(missing), ",".join(missing), "Bổ sung " + ",".join(missing)
        part += 1
    else:
        kn, thieu, inn, goi = "Bị chặn", ",".join(missing) if missing else need, ",".join(missing), "Xin " + (",".join(missing) if missing else need)
        block += 1
    ws.cell(r, 8).value = kn
    ws.cell(r, 9).value = thieu
    ws.cell(r, 10).value = inn
    ws.cell(r, 11).value = goi
    for c in range(8, 12):
        tnr(ws.cell(r, c))
    if kn in {"Bị chặn", "TREO"}:
        ws.cell(r, 8).fill = FILL_RED
    elif kn == "Chạy một phần":
        ws.cell(r, 8).fill = FILL_YEL
    else:
        ws.cell(r, 8).fill = FILL_GRN
ws["C3"] = run_now
ws["E3"] = part
ws["G3"] = block

ws = wb["KQ1_De_nghi_bo_sung"]
ws["C2"] = MST
ws["E2"] = KY
ws["C3"] = VER
ws["E3"] = TODAY
ws["E3"].number_format = DATE_FMT
sug = [
    (1, "Sao kê ngân hàng Excel + UNC các món >= 20.000.000 (đến 30/06/2025) và >= 5.000.000 (từ 01/07/2025)", "IN11", "2.3", "File Excel từng tài khoản, không chỉ PDF scan", 5, "GT16, TT02, TT03, TT05, TT06", "Mở chuỗi thanh toán không dùng tiền mặt; mức E"),
    (2, "Sổ kế toán năm 2019 (NKC, SCTK 111/133/131/331/511/632) và xuất xls 2018/2020/2021 ra Excel dòng", "IN08, IN10, IN13", "2.1", "Excel phát sinh theo chứng từ", 8, "GT06, DT01-DT03, TN17, TT04", "Lấp lỗ 2019; xls cũ chưa gộp dòng"),
    (3, "Bảng tính giá thành 154 / định mức (nếu có sản xuất) và XNT đủ số lượng 152/155/156", "IN09", "2.2", "Tồn CK khớp BCTC ct140/ct141", 8, "KH01-KH07, TN04", "Hiện mới có tổng hợp tồn kho, chưa có định mức"),
    (4, "Phụ lục 05-1/05-2 NLĐ theo từng năm QTT và tra cứu TNCN đa nguồn (thư mục 1.4, 3.3)", "IN05", "1.4 / 3.3", "Danh sách người, CCCD, TNCT, thuế đã khấu trừ", 4, "LD01-LD04, TC22/TC23", "IN05 hiện chỉ có XML QTT"),
    (5, "Phụ lục giao dịch liên kết NĐ 132/2020 (nếu phát sinh) hoặc xác nhận không phát sinh", "IN12", "2.x", "Mẫu 01/02/03", 2, "LK01, LK02", "Chưa nhận PL GDLK"),
    (6, "XML GTGT các kỳ còn thiếu (nếu có) và PDF Q4/2014 đã OCR; HĐĐT bán ra trước 07/2022 nếu portal còn", "IN03, IN01", "1.1 / 1.2", "Bản hiệu lực CQT", 3, "GT01, GT13, HD04", f"Thiếu kỳ GTGT: {', '.join(gtgt_missing[:8]) or 'không'}"),
]
# clear example rows
if ws.max_row and ws.max_row >= 10:
    for r in range(10, ws.max_row + 1):
        for c in range(1, 9):
            ws.cell(r, c).value = None
for i, rowv in enumerate(sug):
    for c, v in enumerate(rowv, 1):
        put(ws, 10 + i, c, v, bold=(c == 1), num=isinstance(v, int) and c == 6, wrap=c in {2, 7, 8})

apply_tnr(wb["KQ1_Kiem_ke"])
wb.save(lp(files["kq13"][1]))
wb.close()
print(" KQ1.3 saved")

for src, dst in files.values():
    vault_dst = VAULT_OUT / dst.name
    shutil.copy2(lp(dst), lp(vault_dst))
    print("vault", vault_dst.name, os.path.getsize(lp(vault_dst)))

print("DONE LEVEL", LEVEL, "run_now", run_now, "part", part, "block", block)
print("OUT", OUT)
print("VAULT_OUT", VAULT_OUT)
