# -*- coding: utf-8 -*-
"""HUY KHANG 3603755214 - quy trinh 1.4 v1: KQ1.1-KQ3."""
from __future__ import annotations

import hashlib
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603755214"
TEN = "CÔNG TY TNHH THIẾT BỊ CÔNG NGHIỆP HUY KHANG"
KY = "2020-2026"
VER = "v1"
TODAY = date(2026, 11, 9)
LOAI_HINH = "Thương mại / thiết bị công nghiệp"

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Ke hoach nam\Bo sung\HUY_KHANG_3603755214"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
ENGINE = VAULT / "skills" / "phan-tich-ho-so-thue" / "engine"
sys.path.insert(0, str(ENGINE))
from ingest.xml_tkhai import parse_xml  # noqa: E402

OUT = ROOT / "4. KET QUA PHAN TICH"
XML_DIR = ROOT / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"
HDDT_DIR = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu"
TMS_DIR = ROOT / "1. DU LIEU THO NGANH THUE" / "1.3. Nghia vu nop thue & Thong bao"
EXTRACT = OUT / "cau-thanh" / "xml_extracted"
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
DATE_FMT = "dd/mm/yyyy"
NUM = "#,##0"
NUM2 = "#,##0.00"
PCT = "0.00%"
FILL_NAVY = PatternFill("solid", fgColor="1F4E78")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
FILL_GRN = PatternFill("solid", fgColor="E2EFDA")
FONT_W = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")

BCTC_MAP = {
    "01": ("KQKD", "Doanh thu bán hàng và cung cấp dịch vụ"),
    "02": ("KQKD", "Giá vốn hàng bán"),
    "10": ("KQKD", "Lợi nhuận gộp về bán hàng và CCDV"),
    "11": ("KQKD", "Doanh thu hoạt động tài chính"),
    "21": ("KQKD", "Chi phí bán hàng"),
    "22": ("KQKD", "Chi phí quản lý doanh nghiệp"),
    "23": ("KQKD", "Chi phí lãi vay"),
    "40": ("KQKD", "Lợi nhuận khác"),
    "50": ("KQKD", "Tổng lợi nhuận kế toán trước thuế"),
    "51": ("KQKD", "Chi phí thuế TNDN hiện hành"),
    "60": ("KQKD", "Lợi nhuận sau thuế TNDN"),
    "140": ("CDKT", "Hàng tồn kho (ct140)"),
    "141": ("CDKT", "Hàng tồn kho (ct141)"),
    "151": ("CDKT", "Thuế GTGT được khấu trừ"),
    "152": ("CDKT", "Thuế và các khoản phải thu Nhà nước"),
    "211": ("CDKT", "Tài sản cố định hữu hình"),
    "214": ("CDKT", "Hao mòn TSCĐ"),
    "242": ("CDKT", "Chi phí trả trước dài hạn"),
    "311": ("CDKT", "Vay và nợ thuê tài chính ngắn hạn"),
    "312": ("CDKT", "Người mua trả tiền trước ngắn hạn"),
    "331": ("CDKT", "Phải trả người bán ngắn hạn"),
    "333": ("CDKT", "Thuế và các khoản phải nộp Nhà nước"),
    "335": ("CDKT", "Chi phí phải trả ngắn hạn"),
    "338": ("CDKT", "Phải trả, phải nộp khác ngắn hạn"),
    "341": ("CDKT", "Vay và nợ thuê tài chính dài hạn"),
    "411": ("CDKT", "Vốn góp của chủ sở hữu"),
    "421": ("CDKT", "LN sau thuế chưa phân phối"),
}
TC28 = [
    ("TC1", "Chênh lệch DT GTGT [34] vs DT BCTC [01]", "Doanh thu", "lệch ≠ 0"),
    ("TC2", "Người mua trả tiền trước > 5% DT", "BCTC", "5%"),
    ("TC3", "HTK + XDCB dở dang > 30% DT", "Kho", "30%"),
    ("TC4", "DT chưa thực hiện > 0", "BCTC", "> 0"),
    ("TC5", "Lợi nhuận gộp âm", "Doanh thu - giá vốn", "< 0"),
    ("TC6", "Chi phí phải trả trích trước lớn", "BCTC", "[335]"),
    ("TC7", "Chi phí trả trước / lãi vay chưa vốn hóa", "BCTC", "[242]"),
    ("TC8", "Ưu đãi / miễn giảm TNDN", "TNDN", "C2/C11"),
    ("TC9", "VAT khấu trừ TK vs BCTC [152]", "GTGT", "[43] vs [152]"),
    ("TC10", "VAT khấu trừ vs 10%×(HTK+TSCĐ)", "GTGT", "×10%"),
    ("TC11", "Có DT KCT [26] nhưng chưa phân bổ thuế ĐV", "GTGT", "[26]>0"),
    ("TC12", "Điều chỉnh giảm LN trước thuế", "TNDN", "B7/B8"),
    ("TC13", "Trích lập dự phòng", "BCTC", "dự phòng"),
    ("TC14", "Chuyển nhượng vốn [22]>0 và [25]<0", "TNDN", ""),
    ("TC15", "Bù trừ lãi BĐS / lỗ SXKD", "TNDN", ""),
    ("TC16", "Dấu hiệu chuyển giá", "GDLK", "[251][252]"),
    ("TC17", "CP lãi vay vs 30% EBITDA", "TNDN", "30%"),
    ("TC18", "GDLK không có phụ lục", "GDLK", "IN12"),
    ("TC19", "Thuế TNDN hoãn lại", "TNDN", ""),
    ("TC20", "Phải trả người bán lớn", "Công nợ", "[331]"),
    ("TC21", "Lãi vay khi thiếu vốn điều lệ", "TNDN", "bỏ nếu thiếu TMS vốn"),
    ("TC22", "QTT TNCN vs tờ khai quý", "TNCN", ""),
    ("TC23", "Chi phí lương BCTC vs TNCT", "TNCN", "20%"),
    ("TC24", "KHBS vs hạch toán TMS", "TMS", "không chạy nếu thiếu cột"),
    ("TC25", "Bổ sung GTGT chưa điều chỉnh", "GTGT", ""),
    ("TC26", "Tỷ trọng HĐ mua NCC rủi ro", "Hóa đơn", "F0/F1"),
    ("TC27", "DT TK GTGT vs HĐĐT bán ra", "GTGT", "lệch ≠ 0"),
    ("TC28", "Giá vốn tăng nhanh hơn doanh thu", "Giá thành", "ngang/dọc BCTC"),
]


def lp(p) -> str:
    s = str(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def nnum(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("%", "")
    if not s:
        return None
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def nnum0(v) -> float:
    x = nnum(v)
    return 0.0 if x is None else x


def ndate(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
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


def norm_mst(v) -> str:
    if v is None:
        return ""
    s = str(v).strip().replace(" ", "").replace(".", "").upper()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    s = s.replace("_", "-")
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


def ln(tag: str) -> str:
    return tag.split("}", 1)[-1] if tag else ""


def put(ws, r, c, val, fmt=None, bold=False, fill=None):
    cell = ws.cell(r, c)
    if isinstance(cell, MergedCell):
        return
    if isinstance(val, str) and val.startswith("=") and not val.startswith("=SUM") and not val.startswith("=IF"):
        # avoid accidental formula from text
        if not val[1:2].isalpha() and val[1:2] != "(":
            pass
    cell.value = val
    cell.font = FONT_B if bold else FONT
    cell.border = THIN
    if fmt:
        cell.number_format = fmt
    elif isinstance(val, date) and not isinstance(val, datetime):
        cell.number_format = DATE_FMT
        cell.alignment = Alignment(horizontal="center")
    elif isinstance(val, datetime):
        cell.value = val.date()
        cell.number_format = DATE_FMT
        cell.alignment = Alignment(horizontal="center")
    elif isinstance(val, (int, float)) and fmt is None:
        cell.number_format = NUM
        cell.alignment = Alignment(horizontal="right")
    if fill:
        cell.fill = fill


def style_header_row(ws, r, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(r, c)
        if isinstance(cell, MergedCell):
            continue
        cell.font = FONT_W
        cell.fill = FILL_NAVY
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        cell.border = THIN


def autowidth(ws, cap=60, min_w=10, max_row=400):
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
            ln_ = min(len(s), cap)
            widths[cell.column] = max(widths.get(cell.column, min_w), ln_ + 2)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


def clear_example_from(ws, start_row=3, file_col=None):
    """Delete example data rows but keep header. If file_col given, drop VI_DU."""
    max_r = ws.max_row or start_row
    to_del = []
    for r in range(start_row, max_r + 1):
        vals = [ws.cell(r, c).value for c in range(1, (ws.max_column or 1) + 1)]
        joined = " ".join(str(v) for v in vals if v is not None)
        if "VI_DU" in joined or "0000000001" in joined or "CÔNG TY A" in joined or "CÔNG TY X" in joined:
            to_del.append(r)
            continue
        if file_col and str(ws.cell(r, file_col).value or "") == "VI_DU":
            to_del.append(r)
    for r in reversed(to_del):
        ws.delete_rows(r)


def apply_font_sheet(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            f = cell.font
            cell.font = Font(
                name="Times New Roman",
                size=f.size or 11,
                bold=bool(f.bold),
                italic=False,
                color=f.color,
            )
            if isinstance(cell.value, datetime):
                cell.value = cell.value.date()
                cell.number_format = DATE_FMT
            elif isinstance(cell.value, date):
                cell.number_format = DATE_FMT


# ---------------------------------------------------------------------------
# 1. Extract zips
# ---------------------------------------------------------------------------
EXTRACT.mkdir(parents=True, exist_ok=True)
extracted = []
for z in XML_DIR.glob("*.zip"):
    dest = EXTRACT / z.stem
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(lp(z)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = Path(info.filename).name
            if not name.lower().endswith(".xml"):
                continue
            # flatten + avoid .xml.xml
            while name.lower().endswith(".xml.xml"):
                name = name[:-4]
            out_p = dest / name
            with zf.open(info) as src, open(lp(out_p), "wb") as dst:
                dst.write(src.read())
            extracted.append(out_p)
print("extracted xml", len(extracted))

xml_files = [p for p in list(XML_DIR.glob("*.xml")) + extracted if p.is_file()]
print("xml total", len(xml_files))

# ---------------------------------------------------------------------------
# 2. Parse XML
# ---------------------------------------------------------------------------
tokhai = []
xml_fail = []
for p in xml_files:
    if not p.is_file():
        xml_fail.append(p.name)
        continue
    try:
        tk = parse_xml(p)
    except (FileNotFoundError, OSError):
        tk = None
    if tk is None:
        # fallback generic
        try:
            tree = ET.parse(lp(p))
            root = tree.getroot()
        except Exception:
            xml_fail.append(p.name)
            continue
        ma = ""
        ky = ""
        mst = ""
        ten = ""
        loai_tk = "C"
        so_lan = 0
        cts = {}
        for el in root.iter():
            n = ln(el.tag)
            t = (el.text or "").strip()
            if n == "maTKhai":
                ma = t
            elif n == "kyKKhai":
                ky = t
            elif n == "mst" and not mst:
                mst = t
            elif n == "tenNNT" and not ten:
                ten = t
            elif n == "loaiTKhai":
                loai_tk = t or "C"
            elif n == "soLan":
                try:
                    so_lan = int(float(t or 0))
                except ValueError:
                    so_lan = 0
            elif n.startswith("ct") and t and list(el) == []:
                code = n[2:]
                if code[:1].isdigit():
                    try:
                        cts.setdefault(code, float(t.replace(",", "")))
                    except ValueError:
                        pass
        loai_map = {"842": "gtgt", "892": "tndn", "953": "tncn", "402": "bctc", "200": "bctc"}
        loai = loai_map.get(ma)
        if not loai:
            name = p.name.upper()
            if "GTGT" in name:
                loai = "gtgt"
            elif "TNDN" in name:
                loai = "tndn"
            elif "TNCN" in name:
                loai = "tncn"
            elif "BCTC" in name:
                loai = "bctc"
        if not loai:
            xml_fail.append(p.name)
            continue

        class T:
            pass

        tk = T()
        tk.loai = loai
        tk.ma_tkhai = ma
        tk.loai_tkhai = loai_tk
        tk.so_lan = so_lan
        tk.ky = ky
        tk.nam = 0
        if "/" in ky:
            try:
                tk.nam = int(ky.split("/")[-1])
            except ValueError:
                tk.nam = 0
        else:
            m = re.search(r"(20\d{2})", ky or p.name)
            tk.nam = int(m.group(1)) if m else 0
        tk.tu_ngay = ""
        tk.den_ngay = ""
        tk.mst = mst
        tk.ten_nnt = ten
        tk.chi_tieu = cts
        tk.file = str(p)
        tk.bctc_cdkt_cuoi = {}
        tk.bctc_cdkt_dau = {}
        tk.bctc_kqkd = {}
        if loai == "bctc":
            # try engine-like split
            tk.chi_tieu = cts
    tokhai.append(tk)

print("parsed to khai", len(tokhai), "fail", xml_fail[:8], "nfail", len(xml_fail))
print(" by loai", Counter(t.loai for t in tokhai))

# NLD from TNCN XML
nld_rows = []


def parse_nld(path: Path, nam: int):
    try:
        tree = ET.parse(lp(path))
        root = tree.getroot()
    except Exception:
        return
    # 05-1
    for pl in root.iter():
        n = ln(pl.tag)
        if n not in {"PLuc_05_1_BK_QTT", "PLuc_05_2_BK_QTT", "BKeCTietCNhan"}:
            continue
        nguon = "05-1" if "05_1" in n else ("05-2" if "05_2" in n else "05-1")
        # each person node
        people = []
        if n == "BKeCTietCNhan":
            people = [pl]
        else:
            for ch in list(pl):
                cn = ln(ch.tag).lower()
                if cn in {"ctiet", "nld", "canhan", "dong"} or list(ch):
                    people.append(ch)
            if not people:
                people = [c for c in pl.iter() if ln(c.tag).lower() in {"ctietnld", "nld", "dongchitiet"}]
        # flatten: look for repeating sibling groups with hoTen
        recs = []
        cur = {}
        for el in pl.iter():
            nn = ln(el.tag)
            tt = (el.text or "").strip()
            if not tt or list(el):
                continue
            key = nn.lower()
            if key in {"hoten", "ten", "hovaten"}:
                if cur.get("ten"):
                    recs.append(cur)
                    cur = {}
                cur["ten"] = tt
            elif key in {"mst", "mstcanhan", "mst_cn"}:
                cur["mst"] = tt
            elif key in {"socccd", "cccd", "socmnd", "socmt", "soddcn", "so_ddcn"}:
                cur["cccd"] = tt
            elif key in {"tongtnct", "tnct", "ct22", "thunhap"} and nnum(tt) is not None:
                cur["tnct"] = nnum0(tt)
            elif "thue" in key and "kt" in key or key in {"thuedakt", "ct26", "thuekhautru"}:
                if nnum(tt) is not None:
                    cur["thue"] = nnum0(tt)
        if cur.get("ten"):
            recs.append(cur)
        for rec in recs:
            nld_rows.append(
                {
                    "nam": nam,
                    "nguon": nguon if n != "PLuc_05_2_BK_QTT" else "05-2",
                    "ten": rec.get("ten", ""),
                    "cccd": rec.get("cccd", ""),
                    "mst": rec.get("mst", ""),
                    "tnct": rec.get("tnct", 0),
                    "thue": rec.get("thue", 0),
                    "file": path.name,
                }
            )


for t in tokhai:
    if t.loai == "tncn":
        parse_nld(Path(t.file), t.nam)
print("nld rows", len(nld_rows))

# ---------------------------------------------------------------------------
# 3. HDDT
# ---------------------------------------------------------------------------
invoices = []
STATUS_OK = {"1", "2", "4", "hóa đơn mới", "hoa don moi", "mới", "thay thế", "điều chỉnh"}
STATUS_NO = {"3", "5", "6", "hủy", "đã hủy", "bị thay thế", "bị điều chỉnh", "huy"}


def status_map(v):
    s = str(v or "").strip()
    sl = s.lower()
    if s in {"1"} or "mới" in sl or sl == "hóa đơn mới":
        return "Mới", "Có"
    if s in {"2"} or "thay thế" in sl and "bị" not in sl:
        return "Thay thế", "Có"
    if s in {"4"} or ("điều chỉnh" in sl and "bị" not in sl):
        return "Điều chỉnh", "Có"
    if s in {"3"} or "bị thay thế" in sl:
        return "Bị thay thế", "Không"
    if s in {"5"} or "bị điều chỉnh" in sl:
        return "Bị điều chỉnh", "Không"
    if s in {"6"} or "hủy" in sl or "huy" in sl:
        return "Hủy", "Không"
    if not s:
        return "", ""
    return s, "Có"


def load_hddt(path: Path, chieu: str):
    wb = load_workbook(lp(path), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    n = 0
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i < 7:
            continue
        if not row or all(c is None or str(c).strip() == "" for c in row[:6]):
            continue
        stt = row[0]
        if str(stt).strip().lower() in {"tổng", "tong", "tổng cộng"}:
            continue
        ngay = ndate(row[4])
        if ngay is None and not row[3]:
            continue
        st_txt, tinh = status_map(row[20] if len(row) > 20 else "")
        ts = row[16] if len(row) > 16 else None
        if isinstance(ts, str) and ts.endswith("%"):
            ts = nnum(ts.replace("%", ""))
            if ts is not None and ts > 1:
                ts = ts / 100
        elif isinstance(ts, (int, float)) and ts > 1:
            ts = float(ts) / 100
        rec = {
            "chieu": chieu,
            "kyhieu_mau": row[1],
            "kyhieu": str(row[2] or "").strip(),
            "so": so_hd_key(row[3]),
            "so_goc": row[3],
            "ngay": ngay,
            "mst_nb": norm_mst(row[5]),
            "ten_nb": str(row[6] or "").strip(),
            "mst_nm": norm_mst(row[8]),
            "ten_nm": str(row[9] or "").strip(),
            "mh": str(row[11] or "").strip() if len(row) > 11 else "",
            "dvt": row[12] if len(row) > 12 else "",
            "sl": nnum(row[13]) if len(row) > 13 else None,
            "dg": nnum(row[14]) if len(row) > 14 else None,
            "gt": nnum0(row[15]) if len(row) > 15 else 0,
            "ts": nnum(ts),
            "thue": nnum0(row[17]) if len(row) > 17 else 0,
            "tong": nnum0(row[18]) if len(row) > 18 else 0,
            "status": st_txt,
            "tinh": tinh,
            "file": path.name,
        }
        invoices.append(rec)
        n += 1
    wb.close()
    print("HDDT", chieu, n, path.name)


for f in HDDT_DIR.glob("*.xlsx"):
    name = f.name.upper()
    if "MUA" in name:
        load_hddt(f, "MUA")
    elif "BAN" in name or "BÁN" in name:
        load_hddt(f, "BAN")
print("invoices", len(invoices), Counter(x["chieu"] for x in invoices))

# ---------------------------------------------------------------------------
# 4. TMS
# ---------------------------------------------------------------------------
tms_rows = []
nvkk = TMS_DIR / "Nghia_vu_ke_khai_HUY_KHANG_3603755214.xlsx"
if nvkk.exists():
    wb = load_workbook(lp(nvkk), data_only=True)
    if "Chi_tiet" in wb.sheetnames:
        ws = wb["Chi_tiet"]
        hdr = [str(c.value or "").strip() for c in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            rec = {hdr[i]: row[i] if i < len(row) else None for i in range(len(hdr))}
            tms_rows.append(rec)
    wb.close()
print("tms chi tiet", len(tms_rows))

sps_rows = []
sps = TMS_DIR / "SPS 2020-2021.XLSX"
if sps.exists():
    wb = load_workbook(lp(sps), data_only=True)
    ws = wb[wb.sheetnames[0]]
    hdr = [str(c.value or "").strip() for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        sps_rows.append({hdr[i]: row[i] if i < len(row) else None for i in range(len(hdr))})
    wb.close()
print("sps", len(sps_rows))

# ---------------------------------------------------------------------------
# 5. CVCBRR / CA
# ---------------------------------------------------------------------------


def _mst_pad(v) -> str:
    s = norm_mst(v)
    d = re.sub(r"\D", "", s)
    if len(d) == 9:
        d = "0" + d
        s = d
    return s


def _col(hdr, *need, default=0):
    for i, h in enumerate(hdr):
        for n in need:
            if n in h:
                return i
    return default


def load_mst_sheet(path, sheet_hint=None, mst_role=None):
    """mst_role: 'ban' | 'mua' | None (F1 -> mua)."""
    out = []
    if not path.exists():
        return out
    wb = load_workbook(lp(path), read_only=True, data_only=True)
    skip_mst = {"", "mstban", "mstmua", "mst bên bán", "mst bên mua", "mst"}
    for name in wb.sheetnames:
        if sheet_hint and name.strip().upper() != sheet_hint.strip().upper():
            continue
        ws = wb[name]
        role = mst_role or ("mua" if name.strip().upper() == "F1" else "ban")
        hdr = None
        i_mst = i_ten = i_cv = i_ngay = i_nd = 0
        for row in ws.iter_rows(values_only=True):
            if not row:
                continue
            joined = " ".join(str(c or "").lower() for c in row)
            if hdr is None:
                if "mst" not in joined and "mã số thuế" not in joined:
                    continue
                # skip the MSTBan code-row; wait for Vietnamese header if present later in first rows
                hdr = [str(c or "").strip().lower() for c in row]
                if all(h in {"stt", "mstban", "tenban", "socv", "ngaycv"} for h in hdr[:5] if h):
                    hdr = None
                    continue
                if role == "mua":
                    i_mst = _col(hdr, "mst bên mua", "mstmua", default=8)
                    i_ten = _col(hdr, "tên công ty mua", "tenmua", default=7)
                else:
                    i_mst = _col(hdr, "mst bên bán", "mst ban", "mstban", default=1)
                    i_ten = _col(hdr, "tên công ty bán", "tenban", default=2)
                i_cv = _col(hdr, "số công văn", "socv", default=3)
                i_ngay = _col(hdr, "ngày công văn", "ngaycv", default=4)
                i_nd = _col(hdr, "nội dung", "noidung", default=6)
                continue
            raw = row[i_mst] if i_mst < len(row) else None
            if raw is None or str(raw).strip().lower() in skip_mst:
                continue
            mst = _mst_pad(raw)
            d = re.sub(r"\D", "", mst)
            if len(d) < 10:
                continue
            out.append(
                {
                    "mst": mst,
                    "base": mst_base(mst) or d,
                    "ten": str(row[i_ten] or "").strip() if i_ten < len(row) else "",
                    "cv": str(row[i_cv] or "").strip() if i_cv < len(row) else "",
                    "ngay_cv": ndate(row[i_ngay] if i_ngay < len(row) else None),
                    "nd": str(row[i_nd] or "").strip()[:180] if i_nd < len(row) else "",
                    "sheet": name,
                }
            )
    wb.close()
    return out


f0 = load_mst_sheet(CVCBRR, "F0")
f1 = load_mst_sheet(CVCBRR, "F1")
ca_list = load_mst_sheet(CA, None)
print("F0", len(f0), "F1", len(f1), "CA", len(ca_list))
f0_set = {x["base"] for x in f0}
f1_set = {x["base"] for x in f1}
ca_set = {x["base"] for x in ca_list}
f0_map = {x["base"]: x for x in f0}
f1_map = {x["base"]: x for x in f1}
ca_map = {x["base"]: x for x in ca_list}

nnt_on_f1 = MST in f1_set
nnt_on_f0 = MST in f0_set
nnt_on_ca = MST in ca_set
print("NNT F0/F1/CA", nnt_on_f0, nnt_on_f1, nnt_on_ca)

# ---------------------------------------------------------------------------
# helpers for to khai
# ---------------------------------------------------------------------------


def ky_norm(ky: str, nam: int) -> str:
    ky = (ky or "").strip().replace(".", "/").replace("-", "/")
    m = re.fullmatch(r"Q([1-4])/?(20\d{2})", ky, re.I)
    if m:
        return f"Q{m.group(1)}/{m.group(2)}"
    m = re.fullmatch(r"0?([1-4])/(20\d{2})", ky)
    if m:
        return f"Q{m.group(1)}/{m.group(2)}"
    if re.fullmatch(r"20\d{2}", ky):
        return ky
    if nam and not ky:
        return str(nam)
    return ky or str(nam)


def is_bs(tk) -> bool:
    return str(tk.loai_tkhai or "").upper() in {"B", "BS"} or int(tk.so_lan or 0) > 0


def effective_gtgt():
    """latest return per ky."""
    groups = defaultdict(list)
    for t in tokhai:
        if t.loai != "gtgt":
            continue
        k = ky_norm(t.ky, t.nam)
        groups[k].append(t)
    out = {}
    for k, lst in groups.items():
        lst = sorted(lst, key=lambda x: (int(x.so_lan or 0), 1 if is_bs(x) else 0))
        out[k] = lst[-1]
    return out


eff_gtgt = effective_gtgt()
print("gtgt ky hieu luc", len(eff_gtgt), sorted(eff_gtgt))


def ct(tk, *codes):
    d = tk.chi_tieu or {}
    for c in codes:
        if c in d:
            return d[c]
        if c.lstrip("0") in d:
            return d[c.lstrip("0")]
    return 0.0


def year_of_inv(inv):
    return inv["ngay"].year if inv["ngay"] else 0


def q_of_inv(inv):
    if not inv["ngay"]:
        return 0
    return (inv["ngay"].month - 1) // 3 + 1


# HDDT yearly
hddt_year = defaultdict(lambda: {"ban_gt": 0, "ban_thue": 0, "mua_gt": 0, "mua_thue": 0, "ban_n": 0, "mua_n": 0})
for inv in invoices:
    if inv["tinh"] != "Có":
        continue
    y = year_of_inv(inv)
    if not y:
        continue
    if inv["chieu"] == "BAN":
        hddt_year[y]["ban_gt"] += inv["gt"]
        hddt_year[y]["ban_thue"] += inv["thue"]
        hddt_year[y]["ban_n"] += 1
    else:
        hddt_year[y]["mua_gt"] += inv["gt"]
        hddt_year[y]["mua_thue"] += inv["thue"]
        hddt_year[y]["mua_n"] += 1

gtgt_year = defaultdict(lambda: {"ct23": 0, "ct24": 0, "ct26": 0, "ct32": 0, "ct34": 0, "ct35": 0, "ct40": 0, "ct41": 0, "ct43": 0, "ct37": 0, "ct38": 0})
for k, t in eff_gtgt.items():
    y = t.nam
    for code in ("23", "24", "26", "32", "34", "35", "40", "41", "43", "37", "38", "22", "25"):
        gtgt_year[y][f"ct{code}"] = gtgt_year[y].get(f"ct{code}", 0) + ct(t, code)

# BCTC / TNDN by year
bctc_by = {}
tndn_by = {}
tncn_by = {}
for t in tokhai:
    if t.loai == "bctc":
        cur = bctc_by.get(t.nam)
        if cur is None or int(t.so_lan or 0) >= int(cur.so_lan or 0):
            bctc_by[t.nam] = t
    elif t.loai == "tndn":
        cur = tndn_by.get(t.nam)
        if cur is None or int(t.so_lan or 0) >= int(cur.so_lan or 0):
            tndn_by[t.nam] = t
    elif t.loai == "tncn":
        cur = tncn_by.get(t.nam)
        if cur is None or int(t.so_lan or 0) >= int(cur.so_lan or 0):
            tncn_by[t.nam] = t
print("bctc years", sorted(bctc_by), "tndn", sorted(tndn_by), "tncn", sorted(tncn_by))


def bctc_val(nam, *codes, which="cuoi"):
    t = bctc_by.get(nam)
    if not t:
        return None
    pools = []
    if which in {"cuoi", "any"}:
        pools.append(getattr(t, "bctc_cdkt_cuoi", {}) or {})
    if which in {"kqkd", "any"}:
        pools.append(getattr(t, "bctc_kqkd", {}) or {})
    pools.append(t.chi_tieu or {})
    for code in codes:
        for d in pools:
            if code in d and isinstance(d[code], (int, float)):
                return float(d[code])
    return None


# ---------------------------------------------------------------------------
# 6. Copy templates
# ---------------------------------------------------------------------------
names = {
    "kq11": f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_{VER}.xlsx",
    "kq12": f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_{VER}.xlsx",
    "kq13": f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx",
    "kq21": f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_{VER}.xlsx",
    "kq22": f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_{VER}.xlsx",
    "kq3": f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_{VER}.xlsx",
}
src = {
    "kq11": MAU / "KQ1.1_Kiem_ke_CQT.xlsx",
    "kq12": MAU / "KQ1.2_Kiem_ke_ho_so_NNT.xlsx",
    "kq13": MAU / "KQ1.3_Danh_gia_ho_so.xlsx",
    "kq21": MAU / "KQ2.1_Phan_tich_CQT.xlsx",
    "kq22": MAU / "KQ2.2_Phan_tich_ho_so_cty.xlsx",
    "kq3": MAU / "KQ3_Ket_luan_danh_gia.xlsx",
}
paths = {}
for k, fn in names.items():
    dst = OUT / fn
    shutil.copy2(lp(src[k]), lp(dst))
    paths[k] = dst
    print("copied", fn)

# ===========================================================================
# KQ1.1
# ===========================================================================
print("FILL KQ1.1")
wb = load_workbook(lp(paths["kq11"]))

# IN00
ws = wb["IN00_Manifest"]
clear_example_from(ws, 3)
cqt_files = []
for p in sorted(XML_DIR.glob("*")):
    if p.is_file():
        cqt_files.append(p)
for p in sorted(HDDT_DIR.glob("*.xlsx")):
    cqt_files.append(p)
for p in sorted(TMS_DIR.glob("*")):
    if p.is_file():
        cqt_files.append(p)

r = 3
for i, p in enumerate(cqt_files, 1):
    name = p.name.upper()
    if "GTGT" in name or p.suffix.lower() == ".xml" and "GTGT" in name:
        ma_in = "IN03_ToKhai_GTGT"
    elif "BCTC" in name:
        ma_in = "IN04_BCTC_TNDN"
    elif "TNDN" in name:
        ma_in = "IN04_BCTC_TNDN"
    elif "TNCN" in name:
        ma_in = "IN05_NLD_TNCN"
    elif "HDDT" in name or "HOADON" in name.upper() or "HÓA ĐƠN" in name:
        ma_in = "IN01_HDDT"
    else:
        ma_in = "IN14_TMS_Nghia_vu"
    dang = p.suffix.lower().lstrip(".")
    try:
        h = sha256(p)
    except Exception:
        h = ""
    put(ws, r, 1, f"N{i:03d}")
    put(ws, r, 2, ma_in)
    put(ws, r, 3, "CQT")
    put(ws, r, 4, dang)
    put(ws, r, 5, p.stat().st_mtime and datetime.fromtimestamp(p.stat().st_mtime).date())
    put(ws, r, 6, 2020)
    put(ws, r, 7, 2026)
    put(ws, r, 8, h)
    put(ws, r, 9, "Có" if "bo sung" in p.name.lower() or "bổ sung" in p.name.lower() else "Không")
    put(ws, r, 10, "Có")
    put(ws, r, 11, None)
    put(ws, r, 12, 0)
    put(ws, r, 13, "Giữ nguyên")
    put(ws, r, 14, "")
    put(ws, r, 15, p.name)
    r += 1
n_in00 = r - 3
autowidth(ws, max_row=80)

# IN01
ws = wb["IN01_HDDT"]
clear_example_from(ws, 3)
r = 3
missing_mst = 0
for inv in invoices:
    ngay = inv["ngay"]
    nam = ngay.year if ngay else None
    thang = ngay.month if ngay else None
    quy = q_of_inv(inv) if ngay else None
    if inv["chieu"] == "MUA":
        mst_dt = inv["mst_nb"]
    else:
        mst_dt = inv["mst_nm"]
    if not mst_dt:
        missing_mst += 1
    khoa = f"{inv['kyhieu']}|{inv['so']}|{mst_dt}"
    put(ws, r, 1, inv["chieu"])
    put(ws, r, 2, nam)
    put(ws, r, 3, quy)
    put(ws, r, 4, thang)
    put(ws, r, 5, inv["kyhieu"])
    put(ws, r, 6, inv["so_goc"] if inv["so_goc"] is not None else inv["so"])
    put(ws, r, 7, ngay)
    put(ws, r, 8, inv["mst_nb"])
    put(ws, r, 9, inv["ten_nb"])
    put(ws, r, 10, inv["mst_nm"])
    put(ws, r, 11, inv["ten_nm"])
    put(ws, r, 12, mst_dt)
    put(ws, r, 13, inv["mh"])
    put(ws, r, 14, inv["dvt"])
    put(ws, r, 15, inv["sl"], NUM2 if inv["sl"] is not None else None)
    put(ws, r, 16, inv["dg"], NUM2 if inv["dg"] is not None else None)
    put(ws, r, 17, inv["gt"], NUM)
    put(ws, r, 18, inv["ts"], "0%" if inv["ts"] is not None and inv["ts"] in (0, 0.0) else "0.00%")
    put(ws, r, 19, inv["thue"], NUM)
    put(ws, r, 20, inv["tong"], NUM)
    put(ws, r, 21, inv["status"])
    put(ws, r, 22, "")
    put(ws, r, 23, inv["tinh"])
    put(ws, r, 24, "")
    put(ws, r, 25, khoa)
    put(ws, r, 26, inv["file"])
    r += 1
n_in01 = r - 3
print("IN01 rows", n_in01, "thieu MST", missing_mst)
ws.auto_filter.ref = f"A2:Z{max(2, r-1)}"
autowidth(ws, max_row=30)

# IN03
ws = wb["IN03_ToKhai_GTGT"]
clear_example_from(ws, 3)
gtgt_all = [t for t in tokhai if t.loai == "gtgt"]
gtgt_all.sort(key=lambda x: (x.nam, ky_norm(x.ky, x.nam), int(x.so_lan or 0)))
r = 3
eff_ids = {id(t) for t in eff_gtgt.values()}
for t in gtgt_all:
    k = ky_norm(t.ky, t.nam)
    put(ws, r, 1, t.nam)
    put(ws, r, 2, k)
    put(ws, r, 3, "Bổ sung" if is_bs(t) else "Chính thức")
    put(ws, r, 4, int(t.so_lan or 0))
    put(ws, r, 5, "Có" if id(t) in eff_ids else "Không")
    put(ws, r, 6, None)
    for i, code in enumerate(["22", "23", "24", "25", "26", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43"]):
        put(ws, r, 7 + i, ct(t, code), NUM)
    put(ws, r, 27, Path(t.file).name)
    r += 1
n_in03 = r - 3
print("IN03", n_in03)
autowidth(ws, max_row=40)

# IN04
ws = wb["IN04_BCTC_TNDN"]
clear_example_from(ws, 3)
r = 3
n_in04 = 0
for nam in sorted(set(list(bctc_by) + list(tndn_by))):
    t = bctc_by.get(nam)
    if t:
        # detect che do
        chedo = "chưa xác định mẫu"
        blob = " ".join([str(t.ma_tkhai), Path(t.file).name])
        # try moTa from file name / ct
        if "DNN" in blob.upper():
            chedo = "TT133 (B01-DNN)"
        else:
            chedo = "TT200/TT99 (B01-DN) - sơ bộ"
        pools = [
            ("CDKT", getattr(t, "bctc_cdkt_cuoi", {}) or {}),
            ("KQKD", getattr(t, "bctc_kqkd", {}) or {}),
        ]
        used = set()
        for bao, d in pools:
            for code, val in d.items():
                if not isinstance(val, (int, float)):
                    continue
                key = (bao, code)
                if key in used:
                    continue
                used.add(key)
                ten_ct = BCTC_MAP.get(code, (bao, f"Chỉ tiêu {code}"))[1] if code in BCTC_MAP else f"Chỉ tiêu {code}"
                bao2 = BCTC_MAP.get(code, (bao, ""))[0]
                put(ws, r, 1, nam)
                put(ws, r, 2, bao2)
                put(ws, r, 3, code)
                put(ws, r, 4, ten_ct)
                put(ws, r, 5, float(val), NUM)
                put(ws, r, 6, None)
                put(ws, r, 7, chedo)
                put(ws, r, 8, Path(t.file).name)
                r += 1
                n_in04 += 1
        # also dump chi_tieu 01/02/10/50/60 if missing
        for code in ("01", "02", "10", "21", "22", "50", "60", "140", "141", "152", "211", "312", "331", "411"):
            if any(code == ws.cell(rr, 3).value and ws.cell(rr, 1).value == nam for rr in range(3, r)):
                continue
            val = None
            for d in (getattr(t, "bctc_kqkd", {}) or {}, getattr(t, "bctc_cdkt_cuoi", {}) or {}, t.chi_tieu or {}):
                if code in d and isinstance(d[code], (int, float)):
                    val = float(d[code])
                    break
            if val is None:
                continue
            bao2, ten_ct = BCTC_MAP.get(code, ("BCTC", f"Chỉ tiêu {code}"))
            put(ws, r, 1, nam)
            put(ws, r, 2, bao2)
            put(ws, r, 3, code)
            put(ws, r, 4, ten_ct)
            put(ws, r, 5, val, NUM)
            put(ws, r, 7, chedo)
            put(ws, r, 8, Path(t.file).name)
            r += 1
            n_in04 += 1
    t2 = tndn_by.get(nam)
    if t2:
        mapping = [
            ("A1", "Doanh thu QTT TNDN [A1]"),
            ("01", "Doanh thu QTT (ct01)"),
            ("A2", "Chi phí QTT"),
            ("C1", "Thu nhập chịu thuế"),
            ("C8", "Thuế TNDN phải nộp"),
            ("C14", "Thuế TNDN còn phải nộp"),
        ]
        d = t2.chi_tieu or {}
        wrote = False
        for code, ten_ct in mapping:
            val = d.get(code)
            if val is None:
                continue
            if not isinstance(val, (int, float)):
                continue
            put(ws, r, 1, nam)
            put(ws, r, 2, "TNDN")
            put(ws, r, 3, code)
            put(ws, r, 4, ten_ct)
            put(ws, r, 5, float(val), NUM)
            put(ws, r, 7, "03/TNDN")
            put(ws, r, 8, Path(t2.file).name)
            r += 1
            n_in04 += 1
            wrote = True
        if not wrote:
            # dump numeric cts of interest
            for code, val in d.items():
                if not isinstance(val, (int, float)):
                    continue
                if not str(code)[:1].isdigit() and not str(code)[:1] in "ABC":
                    continue
                put(ws, r, 1, nam)
                put(ws, r, 2, "TNDN")
                put(ws, r, 3, code)
                put(ws, r, 4, f"Chỉ tiêu {code}")
                put(ws, r, 5, float(val), NUM)
                put(ws, r, 7, "03/TNDN")
                put(ws, r, 8, Path(t2.file).name)
                r += 1
                n_in04 += 1
print("IN04", n_in04)
autowidth(ws, max_row=80)

# IN05
ws = wb["IN05_NLD_TNCN"]
# keep formulas? wipe data rows 3+ values but this sheet has 500 formula rows.
# Overwrite from row 3.
# First clear values in A-J, P-AC of data rows that look like example
clear_example_from(ws, 3)
r = 3
n_qtt = 0
for nam in sorted(tncn_by):
    t = tncn_by[nam]
    d = t.chi_tieu or {}
    so_ld = int(ct(t, "21") or ct(t, "22") or 0)
    put(ws, r, 1, "QTT")
    put(ws, r, 2, nam)
    put(ws, r, 16, "05/QTT-TNCN")
    put(ws, r, 17, so_ld)
    put(ws, r, 8, ct(t, "23"), NUM)
    put(ws, r, 9, ct(t, "32") or ct(t, "26") or ct(t, "33"), NUM)
    put(ws, r, 22, ct(t, "24"), NUM)
    put(ws, r, 25, ct(t, "40") or ct(t, "43"), NUM)
    put(ws, r, 29, Path(t.file).name)
    r += 1
    n_qtt += 1
n_nld = 0
# duplicate CCCD
cccd_count = Counter(x["cccd"] for x in nld_rows if x.get("cccd"))
for rec in nld_rows:
    put(ws, r, 1, "NLD")
    put(ws, r, 2, rec["nam"])
    put(ws, r, 3, rec["ten"])
    put(ws, r, 4, rec["cccd"])
    put(ws, r, 5, rec["mst"])
    put(ws, r, 6, rec["nguon"])
    put(ws, r, 8, rec["tnct"], NUM)
    put(ws, r, 9, rec["thue"], NUM)
    trich = "Không" if rec["nguon"] == "05-2" and rec["thue"] == 0 and rec["tnct"] > 0 else ("Có" if rec["thue"] > 0 else "")
    put(ws, r, 10, trich)
    if rec["tnct"]:
        put(ws, r, 11, rec["thue"] / rec["tnct"] if rec["tnct"] else None, PCT)
    put(ws, r, 13, "Có" if rec["nguon"] == "05-2" and trich == "Không" else "Không")
    put(ws, r, 15, "Có" if rec["cccd"] and cccd_count[rec["cccd"]] > 1 else "Không")
    put(ws, r, 29, rec["file"])
    r += 1
    n_nld += 1
print("IN05 QTT", n_qtt, "NLD", n_nld)
autowidth(ws, max_row=40)

# IN14
ws = wb["IN14_TMS_Nghia_vu"]
clear_example_from(ws, 3)
r = 3
n_in14 = 0
for rec in tms_rows:
    put(ws, r, 1, MST)
    put(ws, r, 2, "KE_KHAI")
    ten_nv = str(rec.get("Tên nghĩa vụ kê khai") or rec.get("Mẫu tờ khai") or "")
    sac = "GTGT" if "GTGT" in ten_nv else ("TNDN" if "TNDN" in ten_nv else ("TNCN" if "TNCN" in ten_nv else "Khác"))
    put(ws, r, 3, sac)
    put(ws, r, 4, rec.get("Tên kỳ tính thuế") or rec.get("Kỳ tính thuế"))
    ky = str(rec.get("Kỳ tính thuế") or "")
    nam = None
    m = re.search(r"(20\d{2})", str(rec.get("Tên kỳ tính thuế") or ""))
    if m:
        nam = int(m.group(1))
    else:
        m = re.match(r"(\d{2})Q", ky)
        if m:
            nam = 2000 + int(m.group(1))
    put(ws, r, 5, nam)
    put(ws, r, 6, ndate(rec.get("Hạn nộp")))
    put(ws, r, 7, ndate(rec.get("Ngày nộp tờ khai")))
    put(ws, r, 8, None)
    put(ws, r, 9, None)
    put(ws, r, 10, None)
    put(ws, r, 11, rec.get("Số ngày nộp quá hạn"))
    put(ws, r, 14, rec.get("Trạng thái nộp TK"))
    put(ws, r, 15, "Nghia_vu_ke_khai / Chi_tiet")
    put(ws, r, 16, nvkk.name)
    r += 1
    n_in14 += 1
for rec in sps_rows:
    put(ws, r, 1, MST)
    put(ws, r, 2, "SPS")
    put(ws, r, 3, "GTGT")
    put(ws, r, 4, rec.get("Kỳ tính thuế"))
    put(ws, r, 7, ndate(rec.get("Ngày hạch toán")))
    put(ws, r, 9, nnum(rec.get("Khấu trừ kỳ trước chuyển qua")), NUM)
    put(ws, r, 14, rec.get("Loại HSKT"))
    put(ws, r, 15, "SPS")
    put(ws, r, 16, sps.name)
    r += 1
    n_in14 += 1
print("IN14", n_in14)
autowidth(ws, max_row=40)
apply_font_sheet(wb["IN00_Manifest"])
wb.save(lp(paths["kq11"]))
wb.close()
print("saved KQ1.1")

# ===========================================================================
# KQ1.2 - company empty
# ===========================================================================
print("FILL KQ1.2")
wb = load_workbook(lp(paths["kq12"]))
for name in wb.sheetnames:
    ws = wb[name]
    clear_example_from(ws, 3)
    if name == "IN00_Manifest":
        put(ws, 3, 1, "N001")
        put(ws, 3, 2, "-")
        put(ws, 3, 3, "DN")
        put(ws, 3, 4, "-")
        put(ws, 3, 13, "Chưa nhận")
        put(ws, 3, 14, "Thư mục 2. HO SO NNT CUNG CAP trống toàn bộ IN02, IN08-IN13")
        put(ws, 3, 15, "(không có tệp)")
    apply_font_sheet(ws)
wb.save(lp(paths["kq12"]))
wb.close()
print("saved KQ1.2")

# ===========================================================================
# KQ1.3
# ===========================================================================
print("FILL KQ1.3")
wb = load_workbook(lp(paths["kq13"]))
ws = wb["KQ1_Kiem_ke"]
put(ws, 2, 2, MST, bold=True)
put(ws, 2, 8, TEN, bold=True)
put(ws, 3, 2, KY)
put(ws, 3, 8, "A", bold=True)
put(ws, 4, 2, "IN Đủ: IN01, IN03, IN04, IN05, IN14 | Mỏng: IN07 (lập từ HĐĐT) | Thiếu/Chưa nhận: IN02, IN08-IN13")
put(ws, 4, 8, TODAY)
put(ws, 5, 2, VER)
put(ws, 5, 8, "1.4 vòng 2")
put(ws, 6, 2, TODAY)
put(ws, 6, 8, "Chưa có")
put(ws, 7, 2, "KQ2.1 (phần CQT); KQ2.2 để trống; KQ3 chạy với KHÔNG_ĐỦ_DỮ_LIỆU cho quy tắc cần sổ DN")
put(ws, 7, 8, "Kỳ suy theo K12: checklist + XML 2020-2026, HĐĐT 12/2021-12/2025")

# rows 10-22 are IN01-IN14
in_eval = {
    10: ("IN01", "Excel", f"12/2021-12/2025", "2020, 2026 (HĐĐT chưa phủ)", n_in01, missing_mst, "Không", "Đủ", "HĐĐT portal đã gộp mua/bán", "Đã chuẩn hóa", "Gắn Tinh_vao_tong, MST chuẩn, khóa KH|So|MST", 0, "Ban ra 2022-2025; mua vào từ 12/2021"),
    11: ("IN02", "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Không có bảng kê NNT ở 2.5", "Chưa chuẩn hóa", "", None, "GT02/GT04/HD* bị chặn hoặc một phần"),
    12: ("IN03", "XML", "Q3/2020-Q3/2026", "Q1-Q2/2020", n_in03, 0, "Có", "Đủ", "XML GTGT chính thức + bổ sung + zip 2025-2026", "Đã chuẩn hóa", "Đánh dấu bản hiệu lực theo soLan/BS", 0, ""),
    13: ("IN04", "XML", "2020-2025", "2026 (năm chưa kết thúc BCTC)", n_in04, 0, "Có", "Đủ", "BCTC + QTT TNDN XML", "Đã chuẩn hóa", "Bảng dọc mã số; mẫu BCTC ghi sơ bộ", 0, "Mẫu B01 chưa đọc được chữ in trên XML"),
    14: ("IN05", "XML", "2020-2025", "Tờ khai TNCN quý; 2026", n_qtt + n_nld, 0, "Không", "Mỏng" if n_nld == 0 else "Đủ", "QTT TNCN XML; phụ lục 05-1/05-2", "Đã chuẩn hóa" if n_nld or n_qtt else "Một phần", "Loai_dong QTT/NLD", 0, "Thiếu file 1.4 tổng hợp NLĐ"),
    15: ("IN07", "Excel CVCBRR", "", "Danh bạ ĐKKD", None, None, "Không", "Mỏng", "Sẽ lập ở KQ2.1 từ HĐĐT + sổ F0/F1/CA", "Chưa chuẩn hóa", "Chưa có IN02", None, "Không có danh bạ 3.2"),
    16: ("IN08", "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Thiếu NKC/Sổ Cái", "Chưa chuẩn hóa", "", None, ""),
    17: ("IN09", "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Thiếu XNT/giá thành", "Chưa chuẩn hóa", "", None, ""),
    18: ("IN10", "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Thiếu sổ 111", "Chưa chuẩn hóa", "", None, ""),
    19: ("IN11", "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Thiếu sao kê/UNC", "Chưa chuẩn hóa", "", None, ""),
    20: ("IN12", "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Chưa thấy PL GDLK", "Chưa chuẩn hóa", "", None, ""),
    21: ("IN13", "—", "", "Toàn kỳ", 0, None, "Không", "Chưa nhận", "Thiếu sổ 511/632", "Chưa chuẩn hóa", "", None, ""),
    22: ("IN14", "Excel TMS", "2020-2026", "", n_in14, 0, "Không", "Đủ", "Nghĩa vụ kê khai + SPS + tổng quan", "Đã chuẩn hóa", "Ánh xạ cột thật Chi_tiet/SPS", 0, "Lịch sử nhật ký TMS không phải nợ thuế"),
}
for row, tup in in_eval.items():
    # cols 6-17
    put(ws, row, 6, tup[1])
    put(ws, row, 7, tup[2])
    put(ws, row, 8, tup[3])
    put(ws, row, 9, tup[4])
    put(ws, row, 10, tup[5])
    put(ws, row, 11, tup[6])
    put(ws, row, 12, tup[7])
    put(ws, row, 13, tup[8])
    put(ws, row, 14, tup[9])
    put(ws, row, 15, tup[10])
    put(ws, row, 16, tup[11])
    put(ws, row, 17, tup[12])
    muc = tup[7]
    if muc == "Đủ":
        ws.cell(row, 12).fill = FILL_GRN
    elif muc in {"Mỏng"}:
        ws.cell(row, 12).fill = FILL_YEL
    else:
        ws.cell(row, 12).fill = FILL_RED

# kha nang
ws = wb["KQ1_Kha_nang_phan_tich"]
put(ws, 2, 2, MST)
put(ws, 2, 6, KY)
put(ws, 4, 2, VER)
put(ws, 4, 6, "1.4")
put(ws, 5, 2, TODAY)
put(ws, 5, 6, "Mức A - chưa có sổ DN")

have = {"IN01", "IN03", "IN04", "IN05", "IN14"}
# IN07 will exist after KQ2.1 from HDDT
have_part = {"IN07"}

n_ngay = n_mot = n_chan = 0
for r in range(10, 90):
    ma = ws.cell(r, 1).value
    if not ma:
        continue
    need = str(ws.cell(r, 4).value or "")
    need_cty = str(ws.cell(r, 5).value or "")
    ins = set(re.findall(r"IN\d{2}", need.replace("–", "-").replace("—", "-")))
    # IN01–IN08 range
    if "IN01-IN08" in need.replace("–", "-") or "IN01–IN08" in need:
        ins |= {f"IN{i:02d}" for i in range(1, 9)}
    missing = [i for i in sorted(ins) if i not in have and i not in have_part]
    # special HD07 TREO
    if ma == "HD07":
        kha = "Bị chặn"
        thieu = "Tham chiếu ngành (K13)"
        goi = "TREO theo K13"
        n_chan += 1
    elif ma == "TN20":
        kha = "Bị chặn"
        thieu = "Chưa triển khai / thiếu khấu hao chi tiết"
        goi = "CHƯA_TRIỂN_KHAI"
        n_chan += 1
    elif not ins:
        kha = "Chạy ngay"
        thieu = ""
        goi = ""
        n_ngay += 1
    elif all(i in have for i in ins):
        kha = "Chạy ngay"
        thieu = ""
        goi = "Đủ nguồn CQT"
        n_ngay += 1
    elif any(i in have or i in have_part for i in ins) and missing:
        # if only company INs missing
        kha = "Chạy một phần"
        thieu = ", ".join(missing)
        goi = "Chạy phần CQT, ô công ty trống (T10)"
        n_mot += 1
    else:
        kha = "Bị chặn"
        thieu = ", ".join(missing) if missing else "Thiếu nguồn"
        goi = "Xin NNT / trích CQT"
        n_chan += 1
    put(ws, r, 8, kha)
    put(ws, r, 9, thieu)
    put(ws, r, 10, thieu)
    put(ws, r, 11, goi)
    if kha == "Chạy ngay":
        ws.cell(r, 8).fill = FILL_GRN
    elif kha == "Chạy một phần":
        ws.cell(r, 8).fill = FILL_YEL
    else:
        ws.cell(r, 8).fill = FILL_RED

put(ws, 3, 2, n_ngay)
put(ws, 3, 5, n_mot)
put(ws, 3, 6, n_chan)
print("rules ngay/mot/chan", n_ngay, n_mot, n_chan)

# de nghi
ws = wb["KQ1_De_nghi_bo_sung"]
put(ws, 2, 2, MST)
put(ws, 2, 6, KY)
put(ws, 3, 2, VER)
put(ws, 3, 6, TODAY)
items = [
    (1, "Sổ XNT kho 152/155/156, bảng giá thành 154, định mức", "IN09", "2.2", "NNT cung cấp Excel", 10, "KH01-KH07, TN04", "Mở khóa nhóm kho/giá thành"),
    (2, "Sao kê ngân hàng Excel + UNC các món ≥ 20tr/5tr", "IN11", "2.3", "NNT cung cấp", 5, "GT16, TT02, TT03, TT05, TT06", "K5 ngưỡng tiền mặt"),
    (3, "Sổ Nhật ký chung, Sổ Cái 133/111/112/131/331", "IN08", "2.1", "NNT cung cấp Excel", 4, "GT06, TN17, TN18, TT04", ""),
    (4, "Sổ chi tiết TK 511 và TK 632 theo dòng", "IN13", "2.1", "NNT cung cấp", 4, "DT01, DT02, DT03, KH07", "Skill doi-chieu-511-632"),
    (5, "Sổ chi tiết tiền mặt TK 111", "IN10", "2.1", "NNT cung cấp", 2, "TT02, TT03", ""),
    (6, "Bảng kê mua vào/bán ra do NNT xuất từ phần mềm", "IN02", "2.5", "NNT cung cấp, không trộn HĐĐT", 12, "GT02, GT04, HD01-HD05, TT01", "Mở NCC 100% và đối chiếu 4 nguồn"),
    (7, "Danh bạ trạng thái ĐKKD toàn bộ NCC", "IN07", "3.2", "CQT trích", 3, "HD02, HD05", ""),
    (8, "Tra cứu TNCN đa nguồn / trùng CCCD", "IN05", "3.3", "CQT trích", 1, "LD04", "Tùy chọn"),
]
for i, it in enumerate(items):
    rr = 10 + i
    for c, v in enumerate(it, 1):
        put(ws, rr, c, v)
autowidth(wb["KQ1_Kiem_ke"], max_row=30)
autowidth(wb["KQ1_Kha_nang_phan_tich"], max_row=90)
autowidth(wb["KQ1_De_nghi_bo_sung"], max_row=25)
wb.save(lp(paths["kq13"]))
wb.close()
print("saved KQ1.3")

print("GĐ1 done")
print("SUMMARY n_in01", n_in01, "n_in03", n_in03, "n_in04", n_in04, "nld", n_nld, "qtt", n_qtt)

_gd23 = Path(__file__).with_name("hk_fill_gd23.py")
exec(compile(_gd23.read_text(encoding="utf-8"), str(_gd23), "exec"))
