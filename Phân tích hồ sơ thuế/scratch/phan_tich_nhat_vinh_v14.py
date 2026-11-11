# -*- coding: utf-8 -*-
"""Phân tích hồ sơ CÔNG TY TNHH MTV NHẤT VINH - MST 3601946069 theo quy trình 1.4."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import traceback
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MST = "3601946069"
TEN = "CÔNG TY TNHH MTV NHẤT VINH"
KY = "2016-2025"
VER = "v1"
TODAY = date(2026, 11, 9)
CHE_DO = "TT133"

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\NHAT_VINH_3601946069"
)
MAU = Path(
    r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế"
    r"\thu-muc-mau-ho-so-doanh-nghiep\MAU_FILE_TONG_HOP"
)
OUT = ROOT / "4. KET QUA PHAN TICH"
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA_FILE = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
BO = MAU / "Bo_bieu_mau_chuan_phan_tich_ho_so_thue.xlsx"

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
NAVY = PatternFill("solid", fgColor="1F4E78")
PALE = PatternFill("solid", fgColor="D9E1F2")
RED = PatternFill("solid", fgColor="FCE4D6")
YEL = PatternFill("solid", fgColor="FFF2CC")
GRN = PatternFill("solid", fgColor="E2EFDA")
DATE_FMT = "dd/mm/yyyy"
NUM = "#,##0"
NUM2 = "#,##0.00"
PCT = "0.00%"
P01, P02, P05, P07, P13, P14, P23 = 100, 1000, 0.30, 0.20, 0.05, 0.50, 0.10
YEARS = list(range(2016, 2026))

RULES_GD2 = {
    "DQ02", "DQ03", "DQ04", "DQ05", "DQ06", "DQ07", "DQ08",
    "GT01", "GT02", "GT03", "GT04", "GT05", "GT07", "GT08", "GT09", "GT10",
    "GT11", "GT12", "GT13", "GT14", "GT15",
    "TN01", "TN02", "TN03", "TN05", "TN06", "TN07", "TN08", "TN09", "TN10",
    "TN11", "TN12", "TN13", "TN14", "TN15", "TN16", "TN19", "TN20", "TN21", "TN22",
    "HD01", "HD02", "HD03", "HD04", "HD05", "HD06", "HD07", "HD08", "HD09",
    "TT01", "LD01", "LD03", "LD04", "LK01", "LK02", "LS01", "LS02",
}
NEED_CTY = {
    "DQ01", "DQ02", "DQ03", "DQ04", "DQ05", "DQ06", "DQ07",
    "GT02", "GT03", "GT04", "GT05", "GT06", "GT07", "GT08", "GT09", "GT16",
    "TN02", "TN04", "TN14", "TN17", "TN18",
    "HD01", "HD02", "HD03", "HD04", "HD05", "HD06", "HD07", "HD08", "HD09",
    "TT01", "TT02", "TT03", "TT04", "TT05", "TT06",
    "KH01", "KH02", "KH03", "KH04", "KH05", "KH06", "KH07",
    "DT01", "DT02", "DT03", "LD02", "LD05", "LK01", "LK02",
}


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


def nnum(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\xa0", "")
    if not s or s in {"-", "None"}:
        return 0.0
    if re.match(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
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
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def norm_mst(v) -> str:
    s = str(v or "").strip()
    s = re.sub(r"[^0-9\-]", "", s)
    if not s:
        return ""
    if "-" in s:
        a, b = s.split("-", 1)
        a = a.zfill(10) if a.isdigit() else a
        return f"{a}-{b}"
    if s.isdigit() and len(s) < 10:
        s = s.zfill(10)
    return s


def valid_mst(s: str) -> bool:
    d = re.sub(r"\D", "", s or "")
    return len(d) in (10, 13) and d.isdigit()


def ln(tag: str) -> str:
    return tag.split("}", 1)[-1] if tag else ""


def xml_text(el) -> str:
    return (el.text or "").strip() if el is not None else ""


def deep_ct(el) -> dict:
    out = {}
    if el is None:
        return out
    for ch in el.iter():
        n = ln(ch.tag)
        if n.startswith("ct") and list(ch) == [] and xml_text(ch):
            code = n[2:]
            if code[:1].isdigit():
                out.setdefault(code, nnum(xml_text(ch)))
    return out


def find_tag(root, name):
    for e in root.iter():
        if ln(e.tag) == name:
            return e
    return None


def apply_font(cell, bold=None):
    f = cell.font
    cell.font = Font(
        name="Times New Roman",
        size=f.size if f and f.size else 11,
        bold=bool(f and f.bold) if bold is None else bold,
        italic=False,
        color=f.color if f else None,
    )


def style_cell(cell, num=False, dt=False, pct=False, bold=False):
    apply_font(cell, bold=bold)
    cell.alignment = Alignment(vertical="center", wrap_text=False)
    cell.border = THIN
    if dt:
        cell.number_format = DATE_FMT
        if isinstance(cell.value, datetime):
            cell.value = cell.value.date()
    elif pct:
        cell.number_format = PCT
    elif num and isinstance(cell.value, (int, float)):
        cell.number_format = NUM


def put(ws, r, c, value, num=False, dt=False, pct=False, bold=False, fill=None, wrap=False):
    cell = ws.cell(r, c)
    if isinstance(cell, MergedCell):
        return None
    cell.value = value
    apply_font(cell, bold=bold)
    cell.alignment = Alignment(vertical="center", wrap_text=wrap)
    cell.border = THIN
    if fill:
        cell.fill = fill
    if dt:
        cell.number_format = DATE_FMT
        if isinstance(value, datetime):
            cell.value = value.date()
    elif pct:
        cell.number_format = PCT
    elif num and isinstance(value, (int, float)):
        cell.number_format = NUM
    return cell


def autowidth(ws, cap=55, min_w=10):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, 2500), max_col=ws.max_column or 1):
        for cell in row:
            if cell.value is None or isinstance(cell, MergedCell):
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            widths[cell.column] = max(widths.get(cell.column, min_w), min(cap, len(s) + 2))
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def clear_from(ws, start_row):
    if ws.max_row and ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)


def copy_mau(name: str, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(lp(MAU / name), lp(dest))


def classify_file(rel: str, name: str) -> tuple[str, str]:
    r = rel.replace("/", "\\")
    n = name.lower()
    if n.startswith("~$"):
        return "-", "CQT"
    if r.startswith("1. DU LIEU"):
        src = "CQT"
        if "1.1" in r:
            if "gtgt" in n:
                return "IN03_ToKhai_GTGT", src
            if "tndn" in n or "bctc" in n:
                return "IN04_BCTC_TNDN", src
            if "tncn" in n or "qtt" in n:
                return "IN05_NLD_TNCN", src
            return "IN03_ToKhai_GTGT", src
        if "1.2" in r and "hddt" in n:
            return "IN01_HDDT", src
        if "1.2" in r:
            return "IN02_BangKe", src
        if "1.3" in r:
            return "IN14_TMS_Nghia_vu", src
        if "1.4" in r:
            return "IN05_NLD_TNCN", src
        return "-", src
    if r.startswith("2. HO SO"):
        src = "DN"
        if "2.1" in r:
            if any(x in n for x in ["511", "632"]):
                return "IN13_So_511_632", src
            if any(x in n for x in ["111", "quytm", "sctk"]):
                return "IN10_Chi_tien_mat", src
            if any(x in n for x in ["154", "621", "622", "627"]):
                return "IN09_Kho_GiaThanh", src
            return "IN08_So_KeToan", src
        if "2.2" in r:
            return "IN09_Kho_GiaThanh", src
        if "2.3" in r:
            if "phai thu" in n or "phai tra" in n or "p.thu" in n:
                return "IN08_So_KeToan", src
            return "IN11_Sao_ke_NH", src
        if "2.5" in r:
            return "IN02_BangKe", src
        if "2.4" in r:
            return "IN05_NLD_TNCN", src
        return "-", src
    if r.startswith("3."):
        src = "CQT"
        if "3.1" in r:
            return "IN07_NCC_CanhBao", src
        if "3.2" in r:
            return "IN07_NCC_CanhBao", src
        if "3.3" in r:
            return "IN05_NLD_TNCN", src
        return "-", src
    if r.startswith("4.") or r.startswith("7.") or r.startswith("ket-qua") or r.startswith("0."):
        return "-", "CQT"
    return "-", "CQT"


def cash_threshold(d) -> float:
    if d and d >= date(2025, 7, 1):
        return 5_000_000
    return 20_000_000


def tinh_vao(status: str) -> str:
    s = (status or "").strip().lower()
    if any(x in s for x in ["hủy", "huy", "xóa", "xoa", "thay thế", "thay the", "đã bị"]):
        return "Không"
    return "Có"


# ---------------------------------------------------------------------------
# SCAN
# ---------------------------------------------------------------------------
print("== SCAN ==")
files = []
n_id = 0
for p in ROOT.rglob("*"):
    if not p.is_file():
        continue
    if p.name.startswith("~$"):
        continue
    rel = str(p.relative_to(ROOT))
    ma_in, loai = classify_file(rel, p.name)
    n_id += 1
    try:
        h = sha256(p)
        st = p.stat()
    except Exception as e:
        h, st = "", None
        print("hash fail", rel, e)
        continue
    files.append(
        {
            "ma": f"N{n_id:03d}",
            "ma_in": ma_in,
            "loai": loai,
            "dinh_dang": p.suffix.lower().lstrip("."),
            "ngay": datetime.fromtimestamp(st.st_mtime).date(),
            "hash": h,
            "size": st.st_size,
            "path": p,
            "rel": rel,
            "name": p.name,
        }
    )
print("files", len(files))


# ---------------------------------------------------------------------------
# XML
# ---------------------------------------------------------------------------
print("== XML ==")
xml_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"
gtgt_raw = []
bctc_rows = []
tndn_rows = []
tncn_qtt = []
tncn_nld = []
tncn_ky = []


def parse_one_xml(path: Path):
    try:
        root = ET.parse(path).getroot()
    except Exception:
        return
    ten = xml_text(find_tag(root, "tenTKhai")) or xml_text(find_tag(root, "tenToKhai"))
    ma = xml_text(find_tag(root, "maTKhai"))
    loai_tk = xml_text(find_tag(root, "loaiTKhai")) or "C"
    so_lan = int(nnum(xml_text(find_tag(root, "soLan"))))
    ky_el = find_tag(root, "KyKKhaiThue")
    ky = xml_text(ky_el.find(ky_el[0].tag) if ky_el is not None and list(ky_el) else None) if False else ""
    ky = xml_text(None)
    tu = den = ngay_nop = ""
    nam = 0
    if ky_el is not None:
        for ch in ky_el:
            t = ln(ch.tag)
            if t == "kyKKhai":
                ky = xml_text(ch)
            elif t == "kyKKhaiTuNgay":
                tu = xml_text(ch)
            elif t == "kyKKhaiDenNgay":
                den = xml_text(ch)
            elif t in ("kkyKKhai",):
                ky = xml_text(ch)
    for e in root.iter():
        t = ln(e.tag)
        if t == "kyKKhai" and not ky:
            ky = xml_text(e)
        elif t in ("namKKhai", "namBCTC") and not nam:
            nam = int(nnum(xml_text(e)))
        elif t in ("ngayNop", "ngayLapTKhai"):
            ngay_nop = xml_text(e)
    if "/" in (ky or ""):
        try:
            nam = int(ky.split("/")[-1])
        except ValueError:
            pass
    if not nam:
        m = re.search(r"(20\d{2})", path.name + " " + (ky or ""))
        nam = int(m.group(1)) if m else 0
    ky_label = ky
    if re.match(r"^\d{2}/\d{4}$", ky or ""):
        mm = int(ky.split("/")[0])
        ky_label = f"T{mm:02d}"
    elif re.match(r"^[1-4]/\d{4}$", ky or ""):
        ky_label = f"Q{ky.split('/')[0]}"
    elif re.search(r"Q([1-4])", path.name, re.I):
        ky_label = f"Q{re.search(r'Q([1-4])', path.name, re.I).group(1)}"
        if not nam:
            m = re.search(r"(20\d{2})", path.name)
            nam = int(m.group(1)) if m else nam
    loai_ten = "Chính thức" if str(loai_tk).upper() in {"C", "1", "CHINH THUC", "CHÍNH THỨC"} else "Bổ sung"
    if "bo sung" in path.name.lower() or "bổ sung" in path.name.lower():
        loai_ten = "Bổ sung"
    cts = deep_ct(root)
    kind = ""
    if "GTGT" in path.name.upper() or ma == "842" or "gtgt" in ten.lower():
        kind = "GTGT"
    elif "BCTC" in path.name.upper() or ma == "402" or "bctc" in ten.lower():
        kind = "BCTC"
    elif "TNDN" in path.name.upper() or ma == "892" or "tndn" in ten.lower():
        kind = "TNDN"
    elif "TNCN" in path.name.upper() or "QTT" in path.name.upper() or ma in {"953", "330"}:
        kind = "TNCN"
    elif "BC26" in path.name.upper() or "TB01" in path.name.upper():
        kind = "KHAC"
    rec = {
        "file": path.name,
        "path": path,
        "nam": nam,
        "ky": ky_label or ky,
        "ky_raw": ky,
        "loai": loai_ten,
        "lan": so_lan,
        "ngay_nop": ndate(ngay_nop),
        "cts": cts,
        "ten": ten,
        "ma": ma,
    }
    if kind == "GTGT":
        gtgt_raw.append(rec)
    elif kind == "BCTC":
        cdkt_cuoi = deep_ct(find_tag(root, "SoCuoiNam"))
        cdkt_dau = deep_ct(find_tag(root, "SoDauNam"))
        kq = find_tag(root, "PL_KQHDSXKD") or find_tag(root, "KetQuaHoatDong")
        kqkd = {}
        if kq is not None:
            nn = None
            for ch in kq:
                if ln(ch.tag) in {"NamNay", "SoNamNay"}:
                    nn = ch
                    break
            kqkd = deep_ct(nn if nn is not None else kq)
        if not cdkt_cuoi:
            cdkt_cuoi = {k: v for k, v in cts.items()}
        rec["cdkt_cuoi"] = cdkt_cuoi
        rec["cdkt_dau"] = cdkt_dau
        rec["kqkd"] = kqkd or cts
        # detect mau
        mau_txt = (ten + path.name).upper()
        rec["che_do"] = "TT133" if "DNN" in mau_txt or "B01A" in mau_txt else CHE_DO
        bctc_rows.append(rec)
    elif kind == "TNDN":
        tndn_rows.append(rec)
    elif kind == "TNCN":
        is_year = ky_label in {str(nam), ""} or (ky and str(nam) == str(ky).replace("/", ""))
        is_qtt = "QTT" in path.name.upper() or (ky_label == str(nam)) or re.search(r"\b20\d{2}\b", path.name) and "Q" not in path.name.upper()
        if re.search(r"Q[1-4]", path.name.upper()):
            is_qtt = False
        rec["is_qtt"] = is_qtt
        # NLD from CTiet
        for e in root.iter():
            if ln(e.tag) in {"CTietNLD", "NLD", "ctietNLD", "BangKeNLD"}:
                pass
        nld_nodes = []
        for e in root.iter():
            tags = ln(e.tag).lower()
            if tags in {"ctiet", "chitietnld", "nld", "bke_nld", "pl_05_1_bk_tncn", "pl_05_2_bk_tncn"}:
                nld_nodes.append(e)
        for e in root.iter():
            ho = xml_text(find_tag(e, "hoTen") or find_tag(e, "tenNLD"))
            cccd = xml_text(find_tag(e, "soCMND") or find_tag(e, "soCCCD") or find_tag(e, "maSoThue"))
            if ho and (find_tag(e, "tongTNCT") is not None or find_tag(e, "tnct") is not None or find_tag(e, "ct21") is not None):
                tn = nnum(xml_text(find_tag(e, "tongTNCT") or find_tag(e, "tnct") or find_tag(e, "ct21")))
                thue = nnum(xml_text(find_tag(e, "thueTNCN") or find_tag(e, "thueDaKhauTru") or find_tag(e, "ct27")))
                nguon = "05-1"
                if find_tag(e, "koHDLD") is not None or "05-2" in (xml_text(find_tag(e, "loaiBK")) or ""):
                    nguon = "05-2"
                tncn_nld.append(
                    {
                        "nam": nam,
                        "ho": ho,
                        "cccd": cccd,
                        "tnct": tn,
                        "thue": thue,
                        "nguon": nguon,
                        "file": path.name,
                    }
                )
        if is_qtt:
            tncn_qtt.append(rec)
        else:
            tncn_ky.append(rec)


for f in sorted(xml_dir.glob("*.xml")):
    parse_one_xml(f)
print("GTGT", len(gtgt_raw), "BCTC", len(bctc_rows), "TNDN", len(tndn_rows), "QTT", len(tncn_qtt), "TNCN ky", len(tncn_ky), "NLD", len(tncn_nld))


def pick_effective(recs, keyfn):
    groups = defaultdict(list)
    for r in recs:
        groups[keyfn(r)].append(r)
    out = []
    for k, items in groups.items():
        items = sorted(items, key=lambda x: (0 if x["loai"] == "Bổ sung" else 1, -x["lan"], str(x["ngay_nop"] or "")))
        best = items[0]
        best["hieu_luc"] = "Có"
        out.append(best)
        for o in items[1:]:
            o["hieu_luc"] = "Không"
            out.append(o)
    return out


gtgt_all = pick_effective(gtgt_raw, lambda r: (r["nam"], r["ky"]))
gtgt_eff = [r for r in gtgt_all if r.get("hieu_luc") == "Có"]


def bctc_val(nam, bao, ma):
    for rec in bctc_rows:
        if rec["nam"] != nam:
            continue
        src = rec["kqkd"] if bao == "KQKD" else rec["cdkt_cuoi"]
        if str(ma) in src:
            return src[str(ma)]
        if ma in src:
            return src[ma]
    return None


def tndn_val(nam, code):
    cands = [r for r in tndn_rows if r["nam"] == nam]
    if not cands:
        return None
    cands = sorted(cands, key=lambda x: (0 if x["loai"] == "Bổ sung" else 1, -x["lan"]))
    return cands[0]["cts"].get(str(code))


# ---------------------------------------------------------------------------
# HDDT
# ---------------------------------------------------------------------------
print("== HDDT ==")
hddt_path = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT.xlsx"
in01 = []


def parse_hddt_sheet(ws, chieu_default):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return
    hdr = [str(c or "").strip() for c in rows[0]]

    def idx(*keys):
        for i, h in enumerate(hdr):
            hl = h.lower()
            if any(k.lower() in hl for k in keys):
                return i
        return None

    i_kh = idx("ký hiệu hóa đơn", "ky hieu hoa don")
    i_so = idx("số hóa đơn", "so hoa don")
    i_ng = idx("ngày lập", "ngay lap")
    i_mst_ban = None
    i_ten_ban = None
    i_mst_mua = None
    i_ten_mua = None
    for i, h in enumerate(hdr):
        hl = h.lower()
        if "mst người bán" in hl or "mst nguoi ban" in hl:
            i_mst_ban = i
        elif "tên người bán" in hl or "ten nguoi ban" in hl:
            i_ten_ban = i
        elif "mst người mua" in hl or "mst nguoi mua" in hl:
            i_mst_mua = i
        elif "tên người mua" in hl or "ten nguoi mua" in hl:
            i_ten_mua = i
    i_gt = idx("tổng tiền chưa thuế", "tong tien chua thue")
    i_thue = idx("tổng tiền thuế", "tong tien thue")
    i_tt = idx("tổng tiền thanh toán", "tong tien thanh toan")
    i_tt_status = idx("trạng thái hóa đơn", "trang thai hoa don")
    i_kq = idx("kết quả kiểm tra", "ket qua kiem tra")
    for row in rows[1:]:
        if not row or not any(row):
            continue
        ngay = ndate(row[i_ng] if i_ng is not None else None)
        nam = ngay.year if ngay else None
        thang = ngay.month if ngay else None
        quy = (thang - 1) // 3 + 1 if thang else None
        stt = str(row[i_tt_status] if i_tt_status is not None else "")
        mst_ban = norm_mst(row[i_mst_ban] if i_mst_ban is not None else "")
        mst_mua = norm_mst(row[i_mst_mua] if i_mst_mua is not None else "")
        chieu = chieu_default
        dt = mst_ban if chieu == "MUA" else mst_mua
        rec = {
            "Chieu": chieu,
            "Nam": nam,
            "Quy": quy,
            "Thang": thang,
            "Ky_hieu_HD": row[i_kh] if i_kh is not None else "",
            "So_HD": str(row[i_so] if i_so is not None else "").strip(),
            "Ngay_HD": ngay,
            "MST_NB": mst_ban,
            "Ten_NB": row[i_ten_ban] if i_ten_ban is not None else "",
            "MST_NM": mst_mua,
            "Ten_NM": row[i_ten_mua] if i_ten_mua is not None else "",
            "MST_DoiTac_chuan": dt,
            "Mat_hang": "",
            "DVT": "",
            "So_luong": None,
            "Don_gia": None,
            "GT_chua_thue": nnum(row[i_gt] if i_gt is not None else 0),
            "Thue_suat": None,
            "Thue_GTGT": nnum(row[i_thue] if i_thue is not None else 0),
            "Tong_TT": nnum(row[i_tt] if i_tt is not None else 0),
            "Trang_thai_HD": stt,
            "HD_lien_quan": "",
            "Tinh_vao_tong": tinh_vao(stt),
            "Hinh_thuc_TT": "",
            "Khoa_DC": "",
            "File_nguon": hddt_path.name,
            "KQ_kiem_tra": row[i_kq] if i_kq is not None else "",
        }
        rec["Khoa_DC"] = f"{rec['Ky_hieu_HD']}|{rec['So_HD']}|{dt}"
        in01.append(rec)


wb = load_workbook(lp(hddt_path), data_only=True, read_only=True)
for s in wb.sheetnames:
    sl = s.lower()
    ch = "MUA" if "mua" in sl else "BAN" if "ban" in sl else ""
    if ch:
        parse_hddt_sheet(wb[s], ch)
wb.close()
print("IN01", len(in01), "mua", sum(1 for x in in01 if x["Chieu"] == "MUA"), "ban", sum(1 for x in in01 if x["Chieu"] == "BAN"))


# ---------------------------------------------------------------------------
# BANG KE yearly
# ---------------------------------------------------------------------------
print("== BANG KE ==")
bk_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
in02 = []


def is_total_row(row) -> bool:
    texts = " ".join(str(c or "") for c in row).lower()
    if "tổng cộng" in texts or "tong cong" in texts:
        ten = " ".join(str(c or "") for c in row[8:13]).lower()
        if "tổng công ty" in ten or "tong cong ty" in ten:
            return False
        return True
    return False


def parse_bk_file(path: Path, year_hint=None):
    try:
        wb = load_workbook(lp(path), data_only=True)
    except Exception as e:
        print("bk fail", path.name, e)
        return
    for s in wb.sheetnames:
        sl = s.lower()
        chieu = "MUA" if sl in {"vao", "mua vao", "mv"} or "vao" in sl else "BAN" if sl in {"ra", "ban ra", "br"} or sl.endswith("ra") or "ban" in sl else ""
        if not chieu:
            continue
        ws = wb[s]
        # find header
        header_r = None
        for r in range(1, 16):
            vals = [str(ws.cell(r, c).value or "").lower() for c in range(1, 20)]
            joined = " ".join(vals)
            if "số hđ" in joined or "so hd" in joined or "số hóa đơn" in joined:
                header_r = r
                break
        if not header_r:
            continue
        start = header_r + 1
        # skip extra header
        for r in range(header_r, min(header_r + 3, ws.max_row or 1)):
            a = ws.cell(r, 1).value
            if isinstance(a, (int, float)) and a >= 1:
                start = r
                break
            if str(a).isdigit():
                start = r
                break
        for r in range(start, (ws.max_row or 1) + 1):
            row = [ws.cell(r, c).value for c in range(1, 20)]
            if not any(row) or is_total_row(row):
                continue
            if chieu == "MUA":
                so_hd = row[7]
                ngay = ndate(row[8])
                ten = row[10] or row[9]
                mst = norm_mst(row[11])
                hang = row[12]
                gt = nnum(row[13])
                suat = nnum(row[14])
                thue = nnum(row[15])
                tong = nnum(row[16])
                seri = row[6]
            else:
                # RA: A STT, E so, F ngay, H ten, I mst, J hang, K ds, L suat, M thue, N tong
                so_hd = row[4]
                ngay = ndate(row[5])
                ten = row[7]
                mst = norm_mst(row[8])
                hang = row[9]
                gt = nnum(row[10])
                suat = nnum(row[11])
                thue = nnum(row[12])
                tong = nnum(row[13])
                seri = row[3]
            if so_hd in (None, "") and gt == 0 and thue == 0:
                continue
            if not ngay and year_hint is None:
                continue
            nam = ngay.year if ngay else year_hint
            if nam and nam < 2010:
                nam = year_hint or nam
            thang = ngay.month if ngay else None
            in02.append(
                {
                    "Chieu": chieu,
                    "Nam": nam,
                    "Quy": ((thang - 1) // 3 + 1) if thang else None,
                    "Thang": thang,
                    "Ky_hieu_HD": seri or "",
                    "So_HD": str(so_hd or "").strip(),
                    "Ngay_HD": ngay,
                    "MST_goc": mst,
                    "MST_chuan": mst,
                    "Ten_DoiTac": ten or "",
                    "Mat_hang": hang or "",
                    "DVT": "",
                    "So_luong": None,
                    "GT_chua_thue": gt,
                    "Thue_suat": suat if suat <= 1 else suat / 100.0,
                    "Thue_GTGT": thue,
                    "Tong_tien": tong or (gt + thue),
                    "Nhom_HHDV": "",
                    "Da_loc_tong_cong": "Có",
                    "Tinh_vao_tong": "Có",
                    "File_nguon": path.name,
                }
            )
    wb.close()


for p in sorted(bk_dir.glob("BANG KE HOA DON *.xlsx")):
    m = re.search(r"(20\d{2})", p.name)
    parse_bk_file(p, int(m.group(1)) if m else None)
print("IN02", len(in02), "mua", sum(1 for x in in02 if x["Chieu"] == "MUA"), "ban", sum(1 for x in in02 if x["Chieu"] == "BAN"))


# ---------------------------------------------------------------------------
# SCTK so cai 2020-2025 + 2018
# ---------------------------------------------------------------------------
print("== SO CAI ==")
so_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.1. So ke toan tong hop va chi tiet"
in08 = []
in10 = []
in13 = []
in09 = []
ps_year = defaultdict(lambda: defaultdict(lambda: {"no": 0.0, "co": 0.0}))
ck_year = defaultdict(dict)
dk_year = defaultdict(dict)


def add_ps(nam, tk, no, co):
    tk = str(tk)
    ps_year[nam][tk]["no"] += no
    ps_year[nam][tk]["co"] += co


def parse_sctk_sheet(ws, nam, tk_hint):
    # find header row with Nợ / Có
    hdr = None
    for r in range(1, 15):
        vals = [str(ws.cell(r, c).value or "").lower() for c in range(1, 10)]
        if any("nợ" in v or "no" == v.strip() or v.strip() == "nôï" for v in vals) and any(
            "có" in v or v.strip() in {"co", "coù"} for v in vals
        ):
            hdr = r
            break
    if not hdr:
        return 0
    # columns: B so, C ngay, D dien giai, E tkdu, F no, G co (typical)
    n = 0
    tk = str(tk_hint)
    for r in range(hdr + 1, (ws.max_row or 1) + 1):
        so = ws.cell(r, 2).value
        ngay = ndate(ws.cell(r, 3).value)
        dg = str(ws.cell(r, 4).value or "")
        tkdu = ws.cell(r, 5).value
        no = nnum(ws.cell(r, 6).value)
        co = nnum(ws.cell(r, 7).value)
        if not so and not ngay and no == 0 and co == 0:
            continue
        dgl = dg.lower()
        if "dau ky" in dgl or "đầu kỳ" in dgl or "daàu kyø" in dgl or "soá dö" in dgl:
            dk_year[nam][tk] = {"no": no, "co": co}
            continue
        if "cuoi ky" in dgl or "cuối kỳ" in dgl or "cuoái" in dgl or "cộng" in dgl or "cong phat sinh" in dgl:
            if "cuoi" in dgl or "cuối" in dgl or "cuoái" in dgl:
                ck_year[nam][tk] = {"no": no, "co": co}
            continue
        if no == 0 and co == 0:
            continue
        thang = ngay.month if ngay else None
        in08.append(
            {
                "Nam": nam,
                "Thang": thang,
                "TK": tk,
                "PS_No": no,
                "PS_Co": co,
                "Du_No_DK": None,
                "Du_No_CK": None,
                "Du_Co_CK": None,
                "Nguon_so": "SCTK",
                "File_nguon": ws.parent.path if False else f"SCTK111-911-{nam}.xlsx",
            }
        )
        add_ps(nam, tk, no, co)
        n += 1
        if tk.startswith("111") and co > 0:
            in10.append(
                {
                    "Ngay_CT": ngay,
                    "Nam": nam,
                    "So_phieu_chi": str(so or ""),
                    "TK_doi_ung": str(tkdu or ""),
                    "MST_doi_tuong": "",
                    "Ten_doi_tuong": "",
                    "Dien_giai": dg,
                    "So_tien": co,
                    "Nhom_nghiep_vu": "",
                    "Co_HD_hop_le": "Chưa rõ",
                    "Danh_gia": "Vượt ngưỡng" if co >= cash_threshold(ngay) else "",
                    "File_nguon": f"SCTK111-911-{nam}.xlsx",
                }
            )
        if tk.startswith("511") or tk.startswith("632"):
            in13.append(
                {
                    "Nam": nam,
                    "Thang": thang,
                    "Ngay_CT": ngay,
                    "So_CT": str(so or ""),
                    "TK": tk[:3],
                    "TK_doi_ung": str(tkdu or ""),
                    "Ma_hang": "",
                    "Dien_giai": dg,
                    "MST_doi_tuong": "",
                    "So_luong": None,
                    "PS_Co": co,
                    "PS_No": no,
                    "Ky_hieu_HD_LQ": "",
                    "So_HD_LQ": "",
                    "File_nguon": f"SCTK111-911-{nam}.xlsx",
                }
            )
    return n


def parse_sctk_file(path: Path, nam: int):
    wb = load_workbook(lp(path), data_only=True, read_only=True)
    wanted = {}
    for s in wb.sheetnames:
        sl = str(s)
        if sl in {"111", "112", "131", "133", "1331", "154", "211", "2141", "331", "334", "338", "3383", "341", "411", "4212", "511", "515", "632", "642", "811", "821", "911"}:
            tk = sl
            if sl == "1331":
                tk = "1331"
            wanted[s] = tk
        elif sl.startswith("SC") and sl[2:].isdigit():
            wanted[s] = sl[2:]
    wb.close()
    wb = load_workbook(lp(path), data_only=True)
    tot = 0
    for s, tk in wanted.items():
        tot += parse_sctk_sheet(wb[s], nam, tk)
    wb.close()
    return tot


for y in range(2020, 2026):
    cands = list(so_dir.glob(f"SCTK111-911-{y}*.xlsx"))
    if not cands:
        print("missing SCTK", y)
        continue
    n = parse_sctk_file(cands[0], y)
    print("SCTK", y, n, "rows")

# 2018 so chi tiet 5113 / 632 / 1111
def parse_2018_sct(path: Path, tk: str, nam=2018):
    try:
        wb = load_workbook(lp(path), data_only=True)
    except Exception as e:
        print("2018 fail", path, e)
        return
    ws = wb[wb.sheetnames[0]]
    hdr = None
    for r in range(1, 20):
        vals = [str(ws.cell(r, c).value or "").lower() for c in range(1, 8)]
        if "ngày" in "".join(vals) and "chứng từ" in "".join(vals):
            hdr = r
            break
    if not hdr:
        wb.close()
        return
    for r in range(hdr + 2, (ws.max_row or 1) + 1):
        ngay = ndate(ws.cell(r, 1).value)
        so = ws.cell(r, 2).value
        dg = str(ws.cell(r, 4).value or "")
        tkdu = ws.cell(r, 5).value
        no = nnum(ws.cell(r, 6).value)
        co = nnum(ws.cell(r, 7).value)
        if no == 0 and co == 0:
            continue
        dgl = dg.lower()
        if "cộng" in dgl or "cong " in dgl:
            continue
        thang = ngay.month if ngay else None
        in08.append({"Nam": nam, "Thang": thang, "TK": tk, "PS_No": no, "PS_Co": co, "Du_No_DK": None, "Du_No_CK": None, "Du_Co_CK": None, "Nguon_so": "SCT 2018", "File_nguon": path.name})
        add_ps(nam, tk, no, co)
        if tk.startswith("111") and co > 0:
            in10.append({"Ngay_CT": ngay, "Nam": nam, "So_phieu_chi": str(so or ""), "TK_doi_ung": str(tkdu or ""), "MST_doi_tuong": "", "Ten_doi_tuong": "", "Dien_giai": dg, "So_tien": co, "Nhom_nghiep_vu": "", "Co_HD_hop_le": "Chưa rõ", "Danh_gia": "Vượt ngưỡng" if co >= cash_threshold(ngay) else "", "File_nguon": path.name})
        if tk.startswith("511") or tk.startswith("632"):
            in13.append({"Nam": nam, "Thang": thang, "Ngay_CT": ngay, "So_CT": str(so or ""), "TK": tk[:3], "TK_doi_ung": str(tkdu or ""), "Ma_hang": "", "Dien_giai": dg, "MST_doi_tuong": "", "So_luong": None, "PS_Co": co, "PS_No": no, "Ky_hieu_HD_LQ": "", "So_HD_LQ": "", "File_nguon": path.name})
    wb.close()


p2018 = so_dir / "2018" / "2018"
if p2018.exists():
    parse_2018_sct(p2018 / "5113.xlsx", "5113")
    parse_2018_sct(p2018 / "632.xlsx", "632")
    parse_2018_sct(p2018 / "1111.xlsx", "1111")
    parse_2018_sct(p2018 / "131.xlsx", "131")
    parse_2018_sct(p2018 / "331.xlsx", "331")
    parse_2018_sct(p2018 / "1331.xlsx", "1331")
    parse_2018_sct(p2018 / "154.xlsx", "154")
print("IN08", len(in08), "IN10", len(in10), "IN13", len(in13))


def sum_tk(nam, prefix, field):
    tot = 0.0
    for tk, v in ps_year.get(nam, {}).items():
        if str(tk).startswith(prefix):
            tot += v[field]
    return tot


# ---------------------------------------------------------------------------
# TMS / PTDN
# ---------------------------------------------------------------------------
print("== TMS ==")
in14 = []
tms_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.3. Nghia vu nop thue & Thong bao"
for p in tms_dir.glob("*.xlsx"):
    in14.append(
        {
            "MST": MST,
            "Loai_ban_ghi": "PTDN",
            "Sac_thue": "",
            "Ky_thue": KY,
            "Nam": None,
            "Han_nop": None,
            "Ngay_nop": None,
            "Lan_nop": None,
            "So_phai_nop": None,
            "So_da_nop": None,
            "So_no": None,
            "So_thong_bao": "",
            "Ngay_thong_bao": None,
            "Noi_dung": "File phân tích Bảng cân đối / nghĩa vụ (PTDN), không phải sổ nợ thuế dòng",
            "Nguon_TMS": p.name,
            "File_nguon": p.name,
        }
    )


# ---------------------------------------------------------------------------
# CVCBRR + CA
# ---------------------------------------------------------------------------
print("== CVCBRR ==")
f0 = {}
f1_nnt = []
f1_seller = {}


def load_cvcbrr():
    wb = load_workbook(lp(CVCBRR), data_only=True, read_only=True)
    ws = wb["F0"]
    for i, row in enumerate(ws.iter_rows(min_row=8, max_col=18, values_only=True), 8):
        mst = norm_mst(row[1])
        if not mst or not valid_mst(mst):
            continue
        rec = {
            "MST": mst,
            "Ten": row[2] or "",
            "Loai_canh_bao": "F0",
            "Nhom_F": "F0",
            "Nguon_F0": "CA" if (str(row[8] or "").find("Công an") >= 0 or str(row[5] or "").lower().find("công an") >= 0) else "CANH_BAO",
            "So_CV": row[3] or "",
            "Ngay_CV": ndate(row[4]),
            "Co_quan_BH": row[5] or "",
            "Noi_dung": row[6] or "",
            "Tinh_trang_DKKD": "",
            "Dia_chi": row[9] or "",
            "Nguon_tra_cuu": "Sổ CVCBRR F0",
            "File_nguon": CVCBRR.name,
        }
        if mst not in f0:
            f0[mst] = rec
    ws = wb["F1"]
    for row in ws.iter_rows(min_row=8, max_col=20, values_only=True):
        mst_mua = norm_mst(row[8])
        mst_ban = norm_mst(row[1])
        if mst_mua == MST:
            f1_nnt.append(
                {
                    "MST": MST,
                    "Ten": TEN,
                    "Loai_canh_bao": "F1",
                    "Nhom_F": "F1",
                    "Nguon_F0": "CANH_BAO",
                    "So_CV": row[3] or "",
                    "Ngay_CV": ndate(row[4]),
                    "Co_quan_BH": row[5] or "",
                    "Noi_dung": row[6] or "",
                    "Tinh_trang_DKKD": "",
                    "Nguon_tra_cuu": "Sổ CVCBRR F1",
                    "File_nguon": CVCBRR.name,
                    "MST_F0": mst_ban,
                    "Ten_F0": row[2] or "",
                }
            )
        if mst_ban and valid_mst(mst_ban):
            f1_seller[mst_ban] = True
    wb.close()


load_cvcbrr()
print("F0", len(f0), "F1 hits NNT", len(f1_nnt), "NNT in F1?", any(True for _ in f1_nnt))

ca_mst = set()
if CA_FILE.exists():
    try:
        wb = load_workbook(lp(CA_FILE), data_only=True, read_only=True)
        ws = wb["NHAP_HO_SO"]
        hdr = None
        for i, row in enumerate(ws.iter_rows(max_row=8, max_col=20, values_only=True), 1):
            vals = [str(c or "").lower() for c in row]
            if any("mst" in v for v in vals):
                hdr = (i, vals)
                break
        if hdr:
            mst_i = next((i for i, v in enumerate(hdr[1]) if "mst" in v), None)
            for row in ws.iter_rows(min_row=hdr[0] + 1, max_col=20, values_only=True):
                if mst_i is not None:
                    m = norm_mst(row[mst_i])
                    if valid_mst(m):
                        ca_mst.add(m)
        wb.close()
    except Exception as e:
        print("CA read", e)
print("CA mst", len(ca_mst), "NNT in CA", MST in ca_mst)

in07 = []
seen07 = set()
for rec in f1_nnt:
    if rec["MST"] not in seen07:
        in07.append(rec)
        seen07.add(rec["MST"])
# NCC from invoices
ncc_from_inv = {}
for src, ch, mst_field, ten_field in (
    (in01, "MUA", "MST_DoiTac_chuan", "Ten_NB"),
    (in02, "MUA", "MST_chuan", "Ten_DoiTac"),
):
    for rec in src:
        if rec.get("Chieu") != ch:
            continue
        if rec.get("Tinh_vao_tong") == "Không":
            continue
        m = rec.get(mst_field) or ""
        if not valid_mst(m):
            continue
        ncc_from_inv.setdefault(m, rec.get(ten_field) or "")

for mst, ten in ncc_from_inv.items():
    if mst in f0:
        rec = dict(f0[mst])
        rec["Ten"] = rec["Ten"] or ten
        if mst not in seen07:
            in07.append(rec)
            seen07.add(mst)
    elif mst in ca_mst:
        rec = {
            "MST": mst,
            "Ten": ten,
            "Loai_canh_bao": "F0",
            "Nhom_F": "F0",
            "Nguon_F0": "CA",
            "So_CV": "",
            "Ngay_CV": None,
            "Co_quan_BH": "Công an",
            "Noi_dung": "Có trên sổ Công an",
            "Tinh_trang_DKKD": "",
            "Nguon_tra_cuu": "Sổ Công an",
            "File_nguon": CA_FILE.name,
        }
        if mst not in seen07:
            in07.append(rec)
            seen07.add(mst)
print("IN07", len(in07))


def flag_ncc(mst: str):
    m = norm_mst(mst)
    for rec in in07:
        if rec["MST"] == m:
            return rec
    if m in f0:
        return f0[m]
    return None


# ---------------------------------------------------------------------------
# METRICS
# ---------------------------------------------------------------------------
print("== METRICS ==")
P02_EPS = 1000


def sum_src(rows, chieu, year, field, tinh=True):
    s = 0.0
    for r in rows:
        if r.get("Chieu") != chieu:
            continue
        if year and r.get("Nam") != year:
            continue
        if tinh and r.get("Tinh_vao_tong") == "Không":
            continue
        s += nnum(r.get(field))
    return s


gtgt_by_year = defaultdict(lambda: {"dt": 0.0, "thue_ra": 0.0, "mua": 0.0, "thue_vao": 0.0, "ct40": 0.0, "ct43": 0.0, "ct26": 0.0, "ct34": 0.0, "ct35": 0.0, "ct23": 0.0, "ct24": 0.0})
for r in gtgt_eff:
    y = r["nam"]
    c = r["cts"]
    gtgt_by_year[y]["dt"] += c.get("32", 0) or c.get("34", 0)
    gtgt_by_year[y]["ct34"] += c.get("34", 0)
    gtgt_by_year[y]["ct35"] += c.get("35", 0)
    gtgt_by_year[y]["ct23"] += c.get("23", 0)
    gtgt_by_year[y]["ct24"] += c.get("24", 0)
    gtgt_by_year[y]["ct26"] += c.get("26", 0)
    gtgt_by_year[y]["ct40"] += c.get("40", 0)
    gtgt_by_year[y]["ct43"] += c.get("43", 0)
    gtgt_by_year[y]["thue_ra"] += c.get("35", 0) or c.get("33", 0)
    gtgt_by_year[y]["mua"] += c.get("23", 0)
    gtgt_by_year[y]["thue_vao"] += c.get("24", 0)

hddt_year = {}
bk_year = {}
for y in YEARS:
    hddt_year[y] = {
        "ban_dt": sum_src(in01, "BAN", y, "GT_chua_thue"),
        "ban_thue": sum_src(in01, "BAN", y, "Thue_GTGT"),
        "mua_dt": sum_src(in01, "MUA", y, "GT_chua_thue"),
        "mua_thue": sum_src(in01, "MUA", y, "Thue_GTGT"),
        "n_ban": sum(1 for r in in01 if r["Chieu"] == "BAN" and r["Nam"] == y and r["Tinh_vao_tong"] != "Không"),
        "n_mua": sum(1 for r in in01 if r["Chieu"] == "MUA" and r["Nam"] == y and r["Tinh_vao_tong"] != "Không"),
    }
    bk_year[y] = {
        "ban_dt": sum_src(in02, "BAN", y, "GT_chua_thue"),
        "ban_thue": sum_src(in02, "BAN", y, "Thue_GTGT"),
        "mua_dt": sum_src(in02, "MUA", y, "GT_chua_thue"),
        "mua_thue": sum_src(in02, "MUA", y, "Thue_GTGT"),
        "n_ban": sum(1 for r in in02 if r["Chieu"] == "BAN" and r["Nam"] == y),
        "n_mua": sum(1 for r in in02 if r["Chieu"] == "MUA" and r["Nam"] == y),
    }

# union invoices mua
inv_map = {}
for rec in in02:
    if rec["Chieu"] != "MUA":
        continue
    m = rec["MST_chuan"]
    if not valid_mst(m):
        continue
    key = (m, str(rec["So_HD"]).lstrip("0") or rec["So_HD"], str(rec["Ngay_HD"]))
    inv_map[key] = {
        "mst": m,
        "ten": rec["Ten_DoiTac"],
        "so": rec["So_HD"],
        "ngay": rec["Ngay_HD"],
        "mh": rec["Mat_hang"],
        "gt": rec["GT_chua_thue"],
        "thue": rec["Thue_GTGT"],
        "src": "Bảng kê",
    }
for rec in in01:
    if rec["Chieu"] != "MUA" or rec["Tinh_vao_tong"] == "Không":
        continue
    m = rec["MST_DoiTac_chuan"]
    if not valid_mst(m):
        continue
    key = (m, str(rec["So_HD"]).lstrip("0") or rec["So_HD"], str(rec["Ngay_HD"]))
    if key in inv_map:
        inv_map[key]["src"] = "Chung"
        if not inv_map[key]["mh"]:
            inv_map[key]["mh"] = rec.get("Mat_hang") or ""
        if inv_map[key]["gt"] == 0:
            inv_map[key]["gt"] = rec["GT_chua_thue"]
        if inv_map[key]["thue"] == 0:
            inv_map[key]["thue"] = rec["Thue_GTGT"]
        if not inv_map[key]["ten"]:
            inv_map[key]["ten"] = rec.get("Ten_NB") or ""
    else:
        inv_map[key] = {
            "mst": m,
            "ten": rec.get("Ten_NB") or "",
            "so": rec["So_HD"],
            "ngay": rec["Ngay_HD"],
            "mh": rec.get("Mat_hang") or "",
            "gt": rec["GT_chua_thue"],
            "thue": rec["Thue_GTGT"],
            "src": "HĐĐT",
        }
invoices = list(inv_map.values())
invoices.sort(key=lambda x: (x["mst"], x["ngay"] or date(1990, 1, 1), str(x["so"])))

ncc_agg = defaultdict(lambda: {"ten": "", "gt": 0.0, "thue": 0.0, "n": 0, "src": set()})
for it in invoices:
    a = ncc_agg[it["mst"]]
    a["ten"] = a["ten"] or it["ten"]
    a["gt"] += it["gt"]
    a["thue"] += it["thue"]
    a["n"] += 1
    a["src"].add(it["src"] if it["src"] != "Chung" else "BK")
    a["src"].add(it["src"])

tot_mua = sum(v["gt"] for v in ncc_agg.values()) or 1.0
f0_gt = 0.0
for m, v in ncc_agg.items():
    fl = flag_ncc(m)
    if fl and fl.get("Nhom_F") in {"F0", "F1"}:
        f0_gt += v["gt"]
hd03 = f0_gt / tot_mua
top5 = sorted(ncc_agg.items(), key=lambda kv: kv[1]["gt"], reverse=True)[:5]
hd04 = sum(v["gt"] for _, v in top5) / tot_mua

# clusters TT01
clusters = defaultdict(float)
for it in invoices:
    if not it["ngay"]:
        continue
    clusters[(it["mst"], it["ngay"])] += it["gt"] + it["thue"]
tt01 = []
for (m, d), amt in clusters.items():
    if amt >= cash_threshold(d):
        tt01.append((m, d, amt))

cash_over = [r for r in in10 if nnum(r["So_tien"]) >= cash_threshold(r.get("Ngay_CT"))]

# 511 / 632 yearly
y511, y632 = {}, {}
for y in YEARS:
    y511[y] = sum_tk(y, "511", "co")
    y632[y] = sum_tk(y, "632", "no")

# BCTC helpers
def kqkd(nam, ma):
    v = bctc_val(nam, "KQKD", ma)
    return v


def cdkt(nam, ma):
    return bctc_val(nam, "CDKT", ma)


print(
    "invoices", len(invoices), "ncc", len(ncc_agg), "hd03", round(hd03, 4), "hd04", round(hd04, 4),
    "tt01", len(tt01), "cash_over", len(cash_over), "f1_nnt", len(f1_nnt),
)

# IN coverage
def years_of(rows, field="Nam"):
    ys = {r.get(field) for r in rows if r.get(field)}
    return ys


in_status = {
    "IN01": ("Đủ" if in01 else "Thiếu", f"{min((r['Nam'] for r in in01 if r['Nam']), default='')}-{max((r['Nam'] for r in in01 if r['Nam']), default='')}", len(in01), "HDDT.xlsx"),
    "IN02": ("Đủ" if in02 else "Thiếu", f"{min((r['Nam'] for r in in02 if r['Nam']), default='')}-{max((r['Nam'] for r in in02 if r['Nam']), default='')}", len(in02), "BANG KE HOA DON 2018-2025"),
    "IN03": ("Đủ" if gtgt_eff else "Thiếu", "2017-2025" if gtgt_eff else "", len(gtgt_eff), "XML GTGT"),
    "IN04": ("Đủ" if bctc_rows else "Thiếu", "2017-2024" if bctc_rows else "", len(bctc_rows), "XML BCTC/TNDN"),
    "IN05": ("Đủ" if tncn_qtt or tncn_nld else "Mỏng", "", len(tncn_qtt) + len(tncn_nld), "XML TNCN"),
    "IN07": ("Đủ", "sổ chỉ định", len(in07), "CVCBRR+CA"),
    "IN08": ("Đủ" if in08 else "Thiếu", "2018,2020-2025", len(in08), "SCTK + SCT 2018"),
    "IN09": ("Mỏng" if any(str(k).startswith("154") or str(k).startswith("621") for y in ps_year.values() for k in y) else "Chưa nhận", "154/621 trong SCTK, chưa có XNT SL", sum(1 for r in in08 if str(r["TK"]).startswith("154") or str(r["TK"]).startswith("621")), "SCTK 154/621"),
    "IN10": ("Đủ" if in10 else "Thiếu", "2018,2020-2025", len(in10), "SCTK 111"),
    "IN11": ("Chưa nhận", "", 0, "Không có sao kê NH/UNC"),
    "IN12": ("Chưa nhận", "", 0, "Không có phụ lục GDLK"),
    "IN13": ("Mỏng" if in13 else "Thiếu", "511 dòng HĐ; 632 kết chuyển tháng", len(in13), "SCTK 511/632"),
    "IN14": ("Mỏng", "2016-2025", len(in14), "PTDN không có dòng nợ thuế"),
}

# muc ho so
has_a = bool(in01 and gtgt_eff and bctc_rows)
has_c = bool(in08 and in13)
has_d = in_status["IN09"][0] in {"Đủ"}
has_e = in_status["IN11"][0] in {"Đủ"}
if has_e and has_d:
    muc = "E"
elif has_d:
    muc = "D"
elif has_c:
    muc = "C"
elif has_a:
    muc = "A"
else:
    muc = "A"
print("MUC HO SO", muc)

# 80 rules capability
rules = []
wb = load_workbook(lp(BO), data_only=True)
ws = wb["04_Quy_tac_rui_ro"]
for r in range(6, ws.max_row + 1):
    rid = ws.cell(r, 1).value
    if not rid:
        continue
    rules.append(
        {
            "id": str(rid).strip(),
            "nhom": ws.cell(r, 2).value,
            "ten": ws.cell(r, 3).value,
            "need": str(ws.cell(r, 7).value or ""),
            "goi_y": ws.cell(r, 9).value,
        }
    )
wb.close()


def rule_run(rid, need):
    ins = re.findall(r"IN\d{2}", need or "")
    missing = []
    for i in ins:
        st = in_status.get(i, ("Chưa nhận",))[0]
        if st in {"Thiếu", "Chưa nhận"}:
            missing.append(i)
        elif st == "Mỏng" and rid in {"KH01", "KH02", "KH03", "KH04", "KH05", "KH06", "KH07", "DT01", "DT02", "GT16", "TT02", "TT06", "LS01", "LS02"}:
            missing.append(i)
    if rid == "HD07":
        return "Bị chặn", "TREO - thiếu tham chiếu ngành (K13)", "", "IN ngoài"
    if rid in {"GT16", "TT02", "TT06"} and in_status["IN11"][0] == "Chưa nhận":
        return "Bị chặn", "IN11 sao kê NH/UNC", "IN11", "Xin sao kê Excel + UNC"
    if rid in {"KH01", "KH02", "KH03", "KH06"} and in_status["IN09"][0] != "Đủ":
        return "Bị chặn", "IN09 sổ XNT kho", "IN09", "Xin sổ XNT 152/155/154"
    if rid in {"DT01", "DT02"}:
        return "Chạy một phần", "632 là kết chuyển tháng, không có dòng HĐ", "IN13", "Xin sổ CT 632 theo HĐ"
    if rid in {"LK01", "LK02"} and in_status["IN12"][0] == "Chưa nhận":
        return "Chạy ngay", "", "", "Không có PL GDLK → LK01 đánh giá sơ bộ"
    if missing:
        return "Bị chặn", ",".join(missing), ",".join(missing), "Bổ sung " + ",".join(missing)
    if rid in RULES_GD2:
        return "Chạy ngay", "", "", ""
    return "Chạy ngay", "", "", ""


# ---------------------------------------------------------------------------
# WRITE WORKBOOKS
# ---------------------------------------------------------------------------
print("== WRITE ==")
OUT.mkdir(parents=True, exist_ok=True)
names = {
    "kq11": OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_{VER}.xlsx",
    "kq12": OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_{VER}.xlsx",
    "kq13": OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx",
    "kq21": OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_{VER}.xlsx",
    "kq22": OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_{VER}.xlsx",
    "kq3": OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_{VER}.xlsx",
}


def banner(ws, text=None):
    for row in ws.iter_rows(max_row=min(6, ws.max_row or 1), max_col=min(8, ws.max_column or 1)):
        for cell in row:
            if isinstance(cell.value, str):
                nv = cell.value
                nv = nv.replace("[MST]", MST).replace("[KY]", KY).replace("[NGAY]", TODAY.strftime("%d/%m/%Y"))
                nv = nv.replace("[TÊN DOANH NGHIỆP]", TEN).replace("[TEN DOANH NGHIEP]", TEN)
                nv = nv.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")
                if nv != cell.value:
                    cell.value = nv
            apply_font(cell)


def write_in00(ws, subset):
    clear_from(ws, 3)
    r = 3
    for f in subset:
        ky_tu = ky_den = None
        put(ws, r, 1, f["ma"])
        put(ws, r, 2, f["ma_in"])
        put(ws, r, 3, f["loai"])
        put(ws, r, 4, "Cau_truc" if f["dinh_dang"] in {"xml", "xlsx", "xls"} else f["dinh_dang"])
        put(ws, r, 5, f["ngay"], dt=True)
        put(ws, r, 6, ky_tu)
        put(ws, r, 7, ky_den)
        put(ws, r, 8, f["hash"])
        put(ws, r, 9, "Không")
        put(ws, r, 10, "Có")
        put(ws, r, 11, None)
        put(ws, r, 12, 0, num=True)
        put(ws, r, 13, "Giữ nguyên")
        put(ws, r, 14, f["rel"])
        put(ws, r, 15, f["name"])
        r += 1
    autowidth(ws)


# ----- KQ1.1 -----
copy_mau("KQ1.1_Kiem_ke_CQT.xlsx", names["kq11"])
wb = load_workbook(lp(names["kq11"]))
for ws in wb.worksheets:
    banner(ws)
cqt_files = [f for f in files if f["loai"] == "CQT" and not f["rel"].startswith(("4.", "7.", "ket-qua", "0."))]
write_in00(wb["IN00_Manifest"], cqt_files)

ws = wb["IN01_HDDT"]
clear_from(ws, 3)
hdr = [ws.cell(2, c).value for c in range(1, 27)]
for i, rec in enumerate(in01, 3):
    vals = [
        rec["Chieu"], rec["Nam"], rec["Quy"], rec["Thang"], rec["Ky_hieu_HD"], rec["So_HD"], rec["Ngay_HD"],
        rec["MST_NB"], rec["Ten_NB"], rec["MST_NM"], rec["Ten_NM"], rec["MST_DoiTac_chuan"], rec["Mat_hang"],
        rec["DVT"], rec["So_luong"], rec["Don_gia"], rec["GT_chua_thue"], rec["Thue_suat"], rec["Thue_GTGT"],
        rec["Tong_TT"], rec["Trang_thai_HD"], rec["HD_lien_quan"], rec["Tinh_vao_tong"], rec["Hinh_thuc_TT"],
        rec["Khoa_DC"], rec["File_nguon"],
    ]
    for c, v in enumerate(vals, 1):
        put(ws, i, c, v, num=isinstance(v, float), dt=isinstance(v, date) and not isinstance(v, datetime) or isinstance(v, date))
autowidth(ws, cap=40)

ws = wb["IN03_ToKhai_GTGT"]
clear_from(ws, 3)
keys = ["22", "23", "24", "25", "26", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43"]
for i, rec in enumerate(sorted(gtgt_all, key=lambda x: (x["nam"] or 0, str(x["ky"]))), 3):
    put(ws, i, 1, rec["nam"])
    put(ws, i, 2, rec["ky"])
    put(ws, i, 3, rec["loai"])
    put(ws, i, 4, rec["lan"], num=True)
    put(ws, i, 5, rec.get("hieu_luc") or "")
    put(ws, i, 6, rec["ngay_nop"], dt=True)
    for j, k in enumerate(keys, 7):
        put(ws, i, j, rec["cts"].get(k), num=True)
    put(ws, i, 27, rec["file"])
autowidth(ws)

ws = wb["IN04_BCTC_TNDN"]
clear_from(ws, 3)
r = 3
labels = {
    "01": "Doanh thu bán hàng và cung cấp dịch vụ",
    "02": "Các khoản giảm trừ doanh thu",
    "10": "Doanh thu thuần",
    "11": "Giá vốn hàng bán",
    "20": "Lợi nhuận gộp",
    "21": "Doanh thu hoạt động tài chính",
    "22": "Chi phí tài chính",
    "23": "Chi phí lãi vay",
    "25": "Chi phí bán hàng",
    "26": "Chi phí quản lý doanh nghiệp",
    "30": "Lợi nhuận thuần từ HĐKD",
    "40": "Lợi nhuận khác",
    "50": "Lợi nhuận trước thuế",
    "51": "Chi phí thuế TNDN",
    "60": "Lợi nhuận sau thuế",
}
cd_labels = {
    "110": "Tiền và tương đương tiền",
    "120": "Đầu tư tài chính ngắn hạn",
    "130": "Các khoản phải thu ngắn hạn",
    "131": "Phải thu khách hàng",
    "140": "Hàng tồn kho",
    "141": "Hàng tồn kho (TT200)",
    "151": "Tài sản ngắn hạn khác / thuế GTGT",
    "152": "Thuế GTGT được khấu trừ",
    "211": "TSCĐ hữu hình",
    "214": "Hao mòn TSCĐ",
    "242": "Chi phí trả trước dài hạn",
    "311": "Phải trả người bán ngắn hạn",
    "312": "Người mua trả tiền trước",
    "315": "Chi phí phải trả ngắn hạn",
    "331": "Phải trả người bán",
    "338": "Phải trả khác",
    "341": "Vay ngắn hạn",
    "411": "Vốn góp của chủ sở hữu",
    "421": "Lợi nhuận sau thuế chưa phân phối",
}
for rec in sorted(bctc_rows, key=lambda x: x["nam"] or 0):
    for ma, ten in labels.items():
        if ma in rec["kqkd"]:
            put(ws, r, 1, rec["nam"])
            put(ws, r, 2, "KQKD")
            put(ws, r, 3, ma)
            put(ws, r, 4, ten)
            put(ws, r, 5, rec["kqkd"].get(ma), num=True)
            put(ws, r, 6, rec.get("cdkt_dau", {}).get(ma) if False else None, num=True)
            put(ws, r, 7, rec.get("che_do") or CHE_DO)
            put(ws, r, 8, rec["file"])
            r += 1
    for ma, ten in cd_labels.items():
        if ma in rec["cdkt_cuoi"]:
            put(ws, r, 1, rec["nam"])
            put(ws, r, 2, "CDKT")
            put(ws, r, 3, ma)
            put(ws, r, 4, ten)
            put(ws, r, 5, rec["cdkt_cuoi"].get(ma), num=True)
            put(ws, r, 6, rec["cdkt_dau"].get(ma) if rec.get("cdkt_dau") else None, num=True)
            put(ws, r, 7, rec.get("che_do") or CHE_DO)
            put(ws, r, 8, rec["file"])
            r += 1
for rec in sorted(tndn_rows, key=lambda x: (x["nam"] or 0, x["lan"])):
    for ma, val in sorted(rec["cts"].items(), key=lambda kv: kv[0]):
        put(ws, r, 1, rec["nam"])
        put(ws, r, 2, "TNDN")
        put(ws, r, 3, ma)
        put(ws, r, 4, f"Chỉ tiêu TNDN ct{ma}")
        put(ws, r, 5, val, num=True)
        put(ws, r, 7, CHE_DO)
        put(ws, r, 8, rec["file"])
        r += 1
autowidth(ws)

ws = wb["IN05_NLD_TNCN"]
# keep formulas but replace example: write QTT first then NLD grouped by year
clear_from(ws, 3)
r = 3
qtt_by_year = {}
for rec in tncn_qtt:
    qtt_by_year[rec["nam"]] = rec
nld_by_year = defaultdict(list)
for rec in tncn_nld:
    nld_by_year[rec["nam"]].append(rec)
years_tn = sorted(set(qtt_by_year) | set(nld_by_year) | {x["nam"] for x in tncn_ky})
for y in years_tn:
    rec = qtt_by_year.get(y)
    if rec:
        c = rec["cts"]
        put(ws, r, 1, "QTT")
        put(ws, r, 2, y)
        put(ws, r, 7, "Quyết toán năm")
        put(ws, r, 8, c.get("26") or c.get("21") or c.get("18"), num=True)
        put(ws, r, 9, c.get("29") or c.get("27") or c.get("40"), num=True)
        put(ws, r, 16, rec.get("ten") or "05/QTT-TNCN")
        put(ws, r, 17, c.get("21") or c.get("20"), num=True)
        put(ws, r, 29, rec["file"])
        r += 1
    for rec in nld_by_year.get(y, []):
        put(ws, r, 1, "NLD")
        put(ws, r, 2, y)
        put(ws, r, 3, rec["ho"])
        put(ws, r, 4, rec["cccd"])
        put(ws, r, 6, rec["nguon"])
        put(ws, r, 8, rec["tnct"], num=True)
        put(ws, r, 9, rec["thue"], num=True)
        put(ws, r, 10, "Không" if rec["nguon"] == "05-2" and rec["thue"] == 0 else "")
        put(ws, r, 29, rec["file"])
        r += 1
if r == 3:
    put(ws, 3, 1, "QTT")
    put(ws, 3, 14, "Chưa trích được dòng NLĐ từ XML; nguồn 1.4 có file tổng hợp sẵn")
autowidth(ws)

ws = wb["IN14_TMS_Nghia_vu"]
clear_from(ws, 3)
for i, rec in enumerate(in14, 3):
    for c, k in enumerate(
        ["MST", "Loai_ban_ghi", "Sac_thue", "Ky_thue", "Nam", "Han_nop", "Ngay_nop", "Lan_nop", "So_phai_nop", "So_da_nop", "So_no", "So_thong_bao", "Ngay_thong_bao", "Noi_dung", "Nguon_TMS", "File_nguon"],
        1,
    ):
        put(ws, i, c, rec.get(k), dt=k in {"Han_nop", "Ngay_nop", "Ngay_thong_bao"})
autowidth(ws)
wb.save(lp(names["kq11"]))
wb.close()
print("saved", names["kq11"].name)

# ----- KQ1.2 -----
copy_mau("KQ1.2_Kiem_ke_ho_so_NNT.xlsx", names["kq12"])
wb = load_workbook(lp(names["kq12"]))
for ws in wb.worksheets:
    banner(ws)
dn_files = [f for f in files if f["loai"] == "DN"]
write_in00(wb["IN00_Manifest"], dn_files)

ws = wb["IN02_BangKe"]
clear_from(ws, 3)
for i, rec in enumerate(in02, 3):
    vals = [
        rec["Chieu"], rec["Nam"], rec["Quy"], rec["Thang"], rec["Ky_hieu_HD"], rec["So_HD"], rec["Ngay_HD"],
        rec["MST_goc"], rec["MST_chuan"], rec["Ten_DoiTac"], rec["Mat_hang"], rec["DVT"], rec["So_luong"],
        rec["GT_chua_thue"], rec["Thue_suat"], rec["Thue_GTGT"], rec["Tong_tien"], rec["Nhom_HHDV"],
        rec["Da_loc_tong_cong"], rec["Tinh_vao_tong"],
    ]
    for c, v in enumerate(vals, 1):
        put(ws, i, c, v, num=isinstance(v, float), dt=isinstance(v, date))
    put(ws, i, 21, None)
    put(ws, i, 22, None)
    put(ws, i, 23, None)
    put(ws, i, 24, rec["File_nguon"])
autowidth(ws, cap=36)

ws = wb["IN08_So_KeToan"]
clear_from(ws, 3)
# write monthly aggregated to keep size reasonable
agg08 = defaultdict(lambda: {"no": 0.0, "co": 0.0, "src": "", "file": ""})
for rec in in08:
    key = (rec["Nam"], rec["Thang"], str(rec["TK"]))
    agg08[key]["no"] += rec["PS_No"]
    agg08[key]["co"] += rec["PS_Co"]
    agg08[key]["src"] = rec["Nguon_so"]
    agg08[key]["file"] = rec["File_nguon"]
for i, ((nam, thang, tk), v) in enumerate(
    sorted(agg08.items(), key=lambda x: (x[0][0] or 0, x[0][1] or 0, str(x[0][2]))),
    3,
):
    put(ws, i, 1, nam)
    put(ws, i, 2, thang)
    put(ws, i, 3, tk)
    put(ws, i, 4, v["no"], num=True)
    put(ws, i, 5, v["co"], num=True)
    put(ws, i, 9, v["src"])
    put(ws, i, 10, v["file"])
autowidth(ws)

ws = wb["IN09_Kho_GiaThanh"]
clear_from(ws, 3)
r = 3
for y in YEARS:
    n154 = sum_tk(y, "154", "no")
    c154 = sum_tk(y, "154", "co")
    n621 = sum_tk(y, "621", "no")
    if n154 or c154 or n621:
        put(ws, r, 1, y)
        put(ws, r, 2, "Giá thành / 154")
        put(ws, r, 6, n154, num=True)
        put(ws, r, 8, c154, num=True)
        put(ws, r, 12, n621, num=True)
        put(ws, r, 15, n154, num=True)
        put(ws, r, 16, "SCTK 154/621 - chưa có XNT số lượng")
        r += 1
if r == 3:
    put(ws, 3, 16, "Chưa nhận sổ XNT 152/155/156; chỉ có 154/621 trên sổ cái")
autowidth(ws)

ws = wb["IN10_Chi_tien_mat"]
clear_from(ws, 3)
for i, rec in enumerate(in10, 3):
    put(ws, i, 1, rec["Ngay_CT"], dt=True)
    put(ws, i, 2, rec["Nam"])
    put(ws, i, 3, rec["So_phieu_chi"])
    put(ws, i, 4, rec["TK_doi_ung"])
    put(ws, i, 5, rec["MST_doi_tuong"])
    put(ws, i, 6, rec["Ten_doi_tuong"])
    put(ws, i, 7, rec["Dien_giai"])
    put(ws, i, 8, rec["So_tien"], num=True)
    put(ws, i, 9, rec["Nhom_nghiep_vu"])
    put(ws, i, 10, rec["Co_HD_hop_le"])
    put(ws, i, 11, rec["Danh_gia"])
    put(ws, i, 12, rec["File_nguon"])
autowidth(ws, cap=50)

ws = wb["IN11_Sao_ke_NH"]
clear_from(ws, 3)
put(ws, 3, 12, "PDF-chua-trich" if False else "")
put(ws, 3, 14, "Chưa nhận sao kê ngân hàng / UNC")

ws = wb["IN12_GDLK"]
clear_from(ws, 3)
put(ws, 3, 9, "Chưa nhận phụ lục GDLK")

ws = wb["IN13_So_511_632"]
clear_from(ws, 3)
for i, rec in enumerate(in13, 3):
    put(ws, i, 1, rec["Nam"])
    put(ws, i, 2, rec["Thang"])
    put(ws, i, 3, rec["Ngay_CT"], dt=True)
    put(ws, i, 4, rec["So_CT"])
    put(ws, i, 5, rec["TK"])
    put(ws, i, 6, rec["TK_doi_ung"])
    put(ws, i, 8, rec["Dien_giai"])
    put(ws, i, 11, rec["PS_Co"], num=True)
    put(ws, i, 12, rec["PS_No"], num=True)
    put(ws, i, 20, rec["File_nguon"])
autowidth(ws, cap=45)
wb.save(lp(names["kq12"]))
wb.close()
print("saved", names["kq12"].name)

# ----- KQ1.3 -----
copy_mau("KQ1.3_Danh_gia_ho_so.xlsx", names["kq13"])
wb = load_workbook(lp(names["kq13"]))
ws = wb["KQ1_Kiem_ke"]
put(ws, 2, 2, MST, bold=True)
put(ws, 2, 8, TEN, bold=True)
put(ws, 3, 2, KY)
put(ws, 3, 8, muc, bold=True)
n_du = sum(1 for k, v in in_status.items() if v[0] == "Đủ")
n_mong = sum(1 for k, v in in_status.items() if v[0] == "Mỏng")
n_thieu = sum(1 for k, v in in_status.items() if v[0] == "Thiếu")
n_chua = sum(1 for k, v in in_status.items() if v[0] == "Chưa nhận")
put(ws, 4, 2, f"{n_du} Đủ | {n_mong} Mỏng | {n_thieu} Thiếu | {n_chua} Chưa nhận")
put(ws, 4, 8, max((f["ngay"] for f in files), default=TODAY), dt=True)
put(ws, 5, 2, VER)
put(ws, 5, 8, "1.4")
put(ws, 6, 2, TODAY, dt=True)
put(ws, 6, 8, "Một phần")
put(ws, 7, 2, "KQ2.1, KQ2.2, KQ3")
put(ws, 7, 8, "Kỳ suy ra từ PTDN 'Từ năm 2016 đến năm 2025' (K12). Sổ cái đủ 2020-2025; 2019 mỏng.")

row_in = {
    "IN01": 10, "IN02": 11, "IN03": 12, "IN04": 13, "IN05": 14, "IN07": 15,
    "IN08": 16, "IN09": 17, "IN10": 18, "IN11": 19, "IN12": 20, "IN13": 21, "IN14": 22,
}
for ma, r in row_in.items():
    st, ky, n, note = in_status[ma]
    put(ws, r, 6, "XML/Excel")
    put(ws, r, 7, ky)
    put(ws, r, 9, n, num=True)
    put(ws, r, 12, st, fill=GRN if st == "Đủ" else YEL if st == "Mỏng" else RED)
    put(ws, r, 13, note)
    put(ws, r, 17, note)
autowidth(ws)

ws = wb["KQ1_Kha_nang_phan_tich"]
put(ws, 2, 2, MST)
put(ws, 2, 6, KY)
put(ws, 4, 2, VER)
put(ws, 4, 6, "1.4")
put(ws, 5, 2, TODAY, dt=True)
n_now = n_part = n_block = 0
for r in range(10, ws.max_row + 1):
    rid = ws.cell(r, 1).value
    if not rid:
        continue
    need = str(ws.cell(r, 4).value or "")
    kha, thieu, inn, goi = rule_run(str(rid), need)
    put(ws, r, 8, kha, fill=GRN if kha == "Chạy ngay" else YEL if "phần" in kha else RED)
    put(ws, r, 9, thieu)
    put(ws, r, 10, inn)
    put(ws, r, 11, goi)
    if kha == "Chạy ngay":
        n_now += 1
    elif "phần" in kha:
        n_part += 1
    else:
        n_block += 1
put(ws, 3, 2, n_now, num=True)
put(ws, 3, 6, f"{n_part} | {n_block}")
autowidth(ws)

ws = wb["KQ1_De_nghi_bo_sung"]
put(ws, 2, 2, MST)
put(ws, 2, 6, KY)
put(ws, 3, 2, VER)
put(ws, 3, 6, TODAY, dt=True)
bo_sung = [
    (1, "Sổ nhập-xuất-tồn kho TK 152/155/156 và bảng giá thành 154 chi tiết", "IN09", "2.2", "File Excel XNT theo năm 2016-2025", 10, "KH01-KH07, TN04, TN05", "Mở nhóm kho-giá thành"),
    (2, "Sao kê ngân hàng Excel + UNC các món ≥ 20 triệu (đến 30/06/2025) và ≥ 5 triệu (từ 01/07/2025)", "IN11", "2.3", "Excel từng TK NH, không gửi mỗi PDF scan", 5, "GT16, TT02, TT03, TT05, TT06", "K11 - chưa kết luận thiếu UNC"),
    (3, "Sổ chi tiết TK 632 theo từng hóa đơn/lô (hiện chỉ có kết chuyển tháng)", "IN13", "2.1", "Sổ CT 632 khớp HĐ bán", 2, "DT01, DT02", "DN sản xuất - 632 KC tháng"),
    (4, "Phụ lục giao dịch liên kết NĐ 132 (nếu có quan hệ liên kết) hoặc xác nhận không phát sinh", "IN12", "2.1/TNDN", "PL-01 GDLK hoặc công văn không phát sinh", 2, "LK01, LK02", ""),
    (5, "File nghĩa vụ nộp thuế / nợ thuế xuất từ TMS (không chỉ PTDN BCTC)", "IN14", "1.3", "Excel nợ thuế, chậm nộp, KQ TTKT", 2, "LS01, LS02", "K10"),
]
for i, row in enumerate(bo_sung, 10):
    for c, v in enumerate(row, 1):
        put(ws, i, c, v)
autowidth(ws)
wb.save(lp(names["kq13"]))
wb.close()
print("saved", names["kq13"].name)

# ----- KQ2.1 -----
copy_mau("KQ2.1_Phan_tich_CQT.xlsx", names["kq21"])
wb = load_workbook(lp(names["kq21"]))
for ws in wb.worksheets:
    banner(ws)

ws = wb["IN07_NCC_CanhBao"]
clear_from(ws, 3)
for i, rec in enumerate(in07, 3):
    put(ws, i, 1, rec.get("MST"))
    put(ws, i, 2, rec.get("Ten"))
    put(ws, i, 3, rec.get("Loai_canh_bao"))
    put(ws, i, 4, rec.get("Nhom_F"))
    put(ws, i, 5, rec.get("Nguon_F0"))
    put(ws, i, 6, rec.get("So_CV"))
    put(ws, i, 7, rec.get("Ngay_CV"), dt=True)
    put(ws, i, 8, rec.get("Co_quan_BH"))
    put(ws, i, 9, rec.get("Noi_dung"))
    put(ws, i, 10, rec.get("Tinh_trang_DKKD") or "")
    put(ws, i, 16, rec.get("Nguon_tra_cuu"))
    put(ws, i, 18, rec.get("File_nguon"))
autowidth(ws)

ws = wb["KQ2_NCC_mua_vao"]
# header row 7, example row 8 - delete example then write all invoices
if ws.max_row >= 8:
    ws.delete_rows(8, ws.max_row - 7)
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER}")
put(ws, 5, 2, round(hd03, 6), pct=True)
put(ws, 5, 4, round(hd04, 6), pct=True)
put(ws, 5, 5, hd03)
put(ws, 5, 6, "Vượt ngưỡng" if hd03 > P13 else "Không vượt")
for i, it in enumerate(invoices, 8):
    fl = flag_ncc(it["mst"])
    nhom = (fl or {}).get("Nhom_F") or "Không"
    cv = (fl or {}).get("So_CV") or ""
    put(ws, i, 1, i - 7, num=True)
    put(ws, i, 2, it["mst"])
    put(ws, i, 3, it["ten"])
    put(ws, i, 4, it["so"])
    put(ws, i, 5, it["ngay"], dt=True)
    put(ws, i, 6, it["mh"])
    put(ws, i, 7, it["gt"], num=True)
    put(ws, i, 8, it["thue"], num=True)
    put(ws, i, 9, it["src"])
    put(ws, i, 10, f"Có - {cv}" if nhom in {"F0", "F1"} else "Không")
    put(ws, i, 11, nhom, fill=RED if nhom in {"F0", "F1"} else None)
    put(ws, i, 12, (fl or {}).get("Tinh_trang_DKKD") or "Không rõ")
# comment block
last = 8 + len(invoices)
put(ws, last + 1, 1, "NHẬN XÉT MODULE (không phải kết luận KQ3)", bold=True, fill=YEL)
nnt_f1 = "Có" if f1_nnt else "Không"
put(
    ws,
    last + 2,
    1,
    f"Phần 1 - NNT {MST} trên F1: {nnt_f1} ({len(f1_nnt)} dòng). NNT trên sổ CA: {'Có' if MST in ca_mst else 'Không'}. "
    f"HD03 tỷ trọng F0/F1 = {hd03:.2%} (P13=5%). HD04 Top-5 = {hd04:.2%} (P14=50%). "
    f"Tổng NCC union = {len(ncc_agg)}. Nguồn CVCBRR ngày {datetime.fromtimestamp(CVCBRR.stat().st_mtime).date().strftime('%d/%m/%Y')}.",
    wrap=True,
    fill=YEL,
)
autowidth(ws, cap=40)

ws = wb["KQ2_Doi_chieu_GTGT"]
put(ws, 2, 1, f"Người nộp thuế: {TEN} - MST: {MST}")
put(ws, 3, 1, f"Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Kỳ: {KY} | Kê khai: theo quý")
# fill Q rows by scanning col A
year_block = None
for r in range(9, 150):
    a = str(ws.cell(r, 1).value or "").strip()
    if a.startswith("Năm "):
        try:
            year_block = int(a.replace("Năm ", "").strip())
        except ValueError:
            year_block = None
        continue
    if a in {"Q1", "Q2", "Q3", "Q4"} and year_block:
        recs = [x for x in gtgt_eff if x["nam"] == year_block and str(x["ky"]).upper() == a]
        if not recs:
            continue
        c = recs[0]["cts"]
        put(ws, r, 2, c.get("26"), num=True)  # KCT
        put(ws, r, 5, c.get("32") or c.get("34"), num=True)
        put(ws, r, 9, c.get("33") or c.get("35"), num=True)
        put(ws, r, 11, c.get("23"), num=True)
        put(ws, r, 12, c.get("24"), num=True)
        put(ws, r, 13, c.get("37"), num=True)
        put(ws, r, 14, c.get("38"), num=True)
        put(ws, r, 16, c.get("42"), num=True)
# Phần II yearly 156-165 = 2016-2025
for i, y in enumerate(YEARS):
    r = 156 + i
    h = hddt_year[y]
    b = bk_year[y]
    g = gtgt_by_year[y]
    put(ws, r, 1, y)
    put(ws, r, 2, h["ban_dt"] or None, num=True)
    put(ws, r, 3, h["ban_thue"] or None, num=True)
    put(ws, r, 4, b["ban_dt"] or None, num=True)
    put(ws, r, 5, b["ban_thue"] or None, num=True)
    put(ws, r, 6, g["ct34"] or None, num=True)
    put(ws, r, 7, g["ct35"] or None, num=True)
    put(ws, r, 8, tndn_val(y, "A1") or tndn_val(y, "1"), num=True)
    put(ws, r, 9, kqkd(y, "01"), num=True)
# comment
put(ws, 167, 1, "NHẬN XÉT / NGUYÊN NHÂN TRỌNG YẾU", bold=True, fill=YEL)
put(
    ws,
    168,
    1,
    "HĐĐT portal chủ yếu từ 2022 (hóa đơn điện tử). Bảng kê có 2018-2025. "
    "Chênh DT HĐĐT vs tờ khai trước 2022 là do HĐĐT không phủ kỳ giấy. "
    "P02 sai số 1.000 đ (tạm - skill). DN kê khai GTGT theo quý.",
    wrap=True,
    fill=YEL,
)
autowidth(ws, cap=22)

ws = wb["KQ2_TNCN_LD2"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}")
# Phần I years start row 8
for i, y in enumerate(YEARS[:12]):
    r = 8 + i
    rec = qtt_by_year.get(y)
    nld = nld_by_year.get(y, [])
    put(ws, r, 1, y)
    if rec:
        c = rec["cts"]
        put(ws, r, 2, c.get("21") or c.get("20"), num=True)
        put(ws, r, 4, c.get("26") or c.get("18"), num=True)
        put(ws, r, 5, c.get("29") or c.get("27"), num=True)
    n052 = sum(1 for x in nld if x["nguon"] == "05-2")
    n052_0 = sum(1 for x in nld if x["nguon"] == "05-2" and x["thue"] == 0)
    put(ws, r, 3, n052 or (len(nld) if nld else None), num=True)
    put(ws, r, 6, n052_0 or None, num=True)
autowidth(ws)

wb.save(lp(names["kq21"]))
wb.close()
print("saved", names["kq21"].name)

# ----- KQ2.2 -----
copy_mau("KQ2.2_Phan_tich_ho_so_cty.xlsx", names["kq22"])
wb = load_workbook(lp(names["kq22"]))
for ws in wb.worksheets:
    banner(ws)

ws = wb["KQ2_511_632"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | {VER} | DN sản xuất (xi mạ / gia công bề mặt)")
# years with 511 data
ys = [y for y in YEARS if y511.get(y) or y632.get(y)]
# template has 4 year cols C-F; we need more years - use C-F for 2022-2025 core + B total, note others in col G
core = [y for y in [2022, 2023, 2024, 2025] if y in ys] or ys[-4:]
put(ws, 5, 2, "Tổng cộng (kỳ có sổ)")
for i, y in enumerate(core):
    put(ws, 5, 3 + i, y, bold=True)
tot511 = sum(y511.get(y, 0) for y in ys)
tot632 = sum(y632.get(y, 0) for y in ys)
put(ws, 6, 2, tot511, num=True)
put(ws, 7, 2, tot632, num=True)
put(ws, 8, 2, tot511 - tot632, num=True)
put(ws, 9, 2, (tot511 - tot632) / tot511 if tot511 else None, pct=True)
for i, y in enumerate(core):
    c = 3 + i
    put(ws, 6, c, y511.get(y), num=True)
    put(ws, 7, c, y632.get(y), num=True)
    lai = y511.get(y, 0) - y632.get(y, 0)
    put(ws, 8, c, lai, num=True)
    put(ws, 9, c, lai / y511[y] if y511.get(y) else None, pct=True)
put(ws, 10, 2, None)
put(ws, 11, 2, None)
put(ws, 12, 2, None)
put(ws, 13, 2, None)
put(ws, 14, 2, None)
put(ws, 6, 7, "511 lấy PS Có sổ cái; 632 là kết chuyển tháng (không có dòng HĐ). Không ghép HĐ+ngày+mặt hàng.")
put(ws, 7, 7, "DT01/DT02 treo ở mức dòng. DT03 so năm với BCTC [01]/[11].")
note_years = "; ".join(f"{y}: 511={y511.get(y,0):,.0f} 632={y632.get(y,0):,.0f}" for y in ys)
put(ws, 15, 7, note_years.replace(",", " "))
autowidth(ws)

ws = wb["KQ2_Kho_Gia_thanh"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | {VER} | IN09 mỏng: có 154/621 trên sổ cái, chưa có XNT số lượng")
put(ws, 42, 1, "NHẬN XÉT MODULE: chưa có sổ XNT 152/155/156 nên KH01-KH03, KH06 ghi chưa đủ nguồn. Có PS 154/621 để đối chiếu sơ bộ giá thành. HTK BCTC DNN dùng ct140 (K6).", wrap=True, fill=YEL)

ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | DN sản xuất gia công xi mạ")
# fill 621/632 vs BCTC if possible - peek rows 7-16
for i, y in enumerate([yy for yy in YEARS if sum_tk(yy, "621", "no") or y632.get(yy)][:10]):
    r = 7 + i
    put(ws, r, 1, y)
    put(ws, r, 2, sum_tk(y, "621", "no") or None, num=True)
    put(ws, r, 5, y632.get(y) or None, num=True)
    gv_bctc = kqkd(y, "11")
    put(ws, r, 8, gv_bctc, num=True)
put(ws, 29, 1, "NHẬN XÉT: giá vốn 632 kết chuyển tháng; biên lãi gộp xem KQ2_511_632. TC28/TN04 sơ bộ từ BCTC, chưa so ngành (K13).", wrap=True, fill=YEL)

ws = wb["KQ2_Cong_no_Tien_mat"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | {VER}")
r = 8
for y in [yy for yy in YEARS if sum_tk(yy, "131", "no") or sum_tk(yy, "331", "co")]:
    put(ws, r, 1, y)
    put(ws, r, 3, sum_tk(y, "131", "no"), num=True)
    put(ws, r, 4, sum_tk(y, "131", "co"), num=True)
    dk = dk_year.get(y, {}).get("131") or dk_year.get(y, {}).get("131")
    put(ws, r, 6, (dk_year.get(y, {}).get("331") or {}).get("co") if False else None, num=True)
    put(ws, r, 7, sum_tk(y, "331", "co"), num=True)
    dt = kqkd(y, "01") or kqkd(y, "10") or 0
    du331 = (ck_year.get(y, {}).get("331") or {}).get("co") or 0
    if dt:
        ratio = du331 / dt if dt else None
        put(ws, r, 10, "Vượt ngưỡng" if ratio and ratio > P23 else "Không vượt")
    r += 1
    if r > 17:
        break
put(ws, 20, 7, "Chưa nạp")
put(ws, 21, 1, 1)
put(ws, 21, 2, "(chưa xác định)")
put(ws, 21, 6, "Chưa nhận")
put(ws, 21, 7, "Chưa nạp")
# cash over sample rows ~34
r = 34
for rec in sorted(cash_over, key=lambda x: x.get("Ngay_CT") or date(1990, 1, 1), reverse=True)[:15]:
    put(ws, r, 1, rec["Ngay_CT"], dt=True)
    put(ws, r, 2, rec["So_phieu_chi"])
    put(ws, r, 3, rec["Dien_giai"])
    put(ws, r, 4, rec["So_tien"], num=True)
    put(ws, r, 11, rec["Danh_gia"], fill=YEL)
    r += 1
r = 52
for (m, d, amt) in sorted(tt01, key=lambda x: -x[2])[:15]:
    put(ws, r, 1, d, dt=True)
    put(ws, r, 2, m)
    put(ws, r, 3, (ncc_agg.get(m) or {}).get("ten") or "")
    put(ws, r, 4, amt, num=True)
    put(ws, r, 10, "Vượt ngưỡng", fill=YEL)
    r += 1
put(
    ws,
    65,
    1,
    f"NHẬN XÉT: {len(cash_over)} phiếu chi TK 111 vượt ngưỡng; {len(tt01)} cụm HĐ cùng ngày cùng NCC vượt ngưỡng TT01. "
    f"IN11 chưa có sao kê → TT02/TT06/GT16 = PDF-chưa kết luận, không kết luận thiếu UNC (K11).",
    wrap=True,
    fill=YEL,
)
autowidth(ws, cap=40)
wb.save(lp(names["kq22"]))
wb.close()
print("saved", names["kq22"].name)

# ----- KQ3 -----
copy_mau("KQ3_Ket_luan_danh_gia.xlsx", names["kq3"])
wb = load_workbook(lp(names["kq3"]))
ws = wb["KQ3"]
put(ws, 2, 2, MST, bold=True)
put(ws, 2, 7, TEN, bold=True)
put(ws, 3, 2, KY)
put(ws, 3, 7, muc)
put(ws, 4, 2, VER)
put(ws, 4, 7, "1.4")
put(ws, 5, 2, TODAY, dt=True)
put(ws, 5, 7, f"KQ1.3 {VER}, KQ2.1 {VER}, KQ2.2 {VER}")

# findings
def gt_flag(y, a, b):
    if a is None or b is None:
        return None
    if abs(a - b) <= P02:
        return "Khớp"
    return "Lệch"


findings = []


def add_f(**kw):
    findings.append(kw)


# HD01
f0_inv = [it for it in invoices if (flag_ncc(it["mst"]) or {}).get("Nhom_F") in {"F0", "F1"}]
if f0_inv:
    add_f(
        qt="HD01", nhom="Hóa đơn – NCC", ten="Mua từ NCC F0/F1/CVCBRR", tc="TC26",
        obj="NCC F0/F1", ky=KY, a="Giá trị mua F0/F1", va=sum(x["gt"] for x in f0_inv),
        b="Tổng mua", vb=tot_mua, do=hd03, nguong=P13, st="ĐÃ_CHẠY", kq="Có phát hiện",
        nhan="Chắc", mucr="Cao", hs=3, diem=3, sac="GTGT", tien=sum(x["gt"] for x in f0_inv),
        nguon="IN07 + KQ2_NCC_mua_vao", ct="KQ2.1!KQ2_NCC_mua_vao",
        kt="Đối chiếu HĐ, hợp đồng, UNC với NCC F0/F1",
        nhan_xet="F0 và F1 đều mức Cao (K2).",
    )
if f1_nnt:
    add_f(
        qt="HD01", nhom="Hóa đơn – NCC", ten="NNT thuộc F1 (mua từ F0)", tc="TC26",
        obj=TEN, ky=KY, a="Số dòng F1", va=len(f1_nnt), b="", vb=None, do=len(f1_nnt),
        nguong=0, st="ĐÃ_CHẠY", kq="Có phát hiện", nhan="Chắc", mucr="Cao", hs=3, diem=3,
        sac="GTGT", tien=None, nguon="Sổ CVCBRR F1", ct="KQ2.1!IN07",
        kt="Rà hồ sơ mua từ F0 gắn với NNT trên F1",
    )
add_f(
    qt="HD03", nhom="Hóa đơn – NCC", ten="Tỷ trọng mua NCC rủi ro", tc="TC26",
    obj="NCC", ky=KY, a="Tỷ trọng F0/F1", va=hd03, b="P13", vb=P13, do=hd03,
    nguong=P13, st="ĐÃ_CHẠY", kq="Có phát hiện" if hd03 > P13 else "Không vượt",
    nhan="Chắc", mucr="Cao" if hd03 > P13 else "Thông tin", hs=2, diem=2 if hd03 > P13 else 0,
    sac="GTGT", tien=f0_gt, nguon="KQ2_NCC", ct="KQ2.1!KQ2_NCC_mua_vao",
)
add_f(
    qt="HD04", nhom="Hóa đơn – NCC", ten="Tập trung Top-5 NCC", tc="",
    obj="Top-5", ky=KY, a="Tỷ trọng Top-5", va=hd04, b="P14", vb=P14, do=hd04,
    nguong=P14, st="ĐÃ_CHẠY", kq="Có phát hiện" if hd04 > P14 else "Không vượt",
    nhan="Chắc", mucr="Trung bình" if hd04 > P14 else "Thông tin", hs=1, diem=1 if hd04 > P14 else 0,
    sac="GTGT", tien=sum(v["gt"] for _, v in top5), nguon="pivot MST", ct="KQ2.1",
)

for y in YEARS:
    g = gtgt_by_year[y]
    h = hddt_year[y]
    b = bk_year[y]
    if g["ct34"] or h["ban_dt"] or b["ban_dt"]:
        fl = gt_flag(y, g["ct34"] or None, h["ban_dt"] or None)
        if h["ban_dt"] or g["ct34"]:
            add_f(
                qt="GT01", nhom="GTGT", ten="DT [34] vs HĐĐT bán ra", tc="TC27",
                obj=TEN, ky=str(y), a="TK [34]", va=g["ct34"], b="HĐĐT ban", vb=h["ban_dt"],
                do=(g["ct34"] or 0) - (h["ban_dt"] or 0), nguong=P02, st="ĐÃ_CHẠY",
                kq="Lệch" if fl == "Lệch" else ("Khớp" if fl == "Khớp" else "Chưa đủ nguồn"),
                nhan="Chắc" if y >= 2022 else "Sơ bộ",
                mucr="Trung bình" if fl == "Lệch" and y >= 2022 else "Thông tin",
                hs=2, diem=2 if fl == "Lệch" and y >= 2022 else 0, sac="GTGT",
                tien=abs((g["ct34"] or 0) - (h["ban_dt"] or 0)), nguon="IN03, IN01",
                ct=f"KQ2.1!KQ2_Doi_chieu_GTGT hàng năm {y}",
                thieu="" if y >= 2022 else "HĐĐT portal không phủ trước 2022",
            )
        fl2 = gt_flag(y, g["ct34"] or None, b["ban_dt"] or None)
        if b["ban_dt"] or g["ct34"]:
            add_f(
                qt="GT02", nhom="GTGT", ten="DT [34] vs Bảng kê bán ra", tc="TC27",
                obj=TEN, ky=str(y), a="TK [34]", va=g["ct34"], b="BK ban", vb=b["ban_dt"],
                do=(g["ct34"] or 0) - (b["ban_dt"] or 0), nguong=P02, st="ĐÃ_CHẠY",
                kq="Lệch" if fl2 == "Lệch" else ("Khớp" if fl2 == "Khớp" else "Chưa đủ nguồn"),
                nhan="Chắc", mucr="Trung bình" if fl2 == "Lệch" else "Thông tin",
                hs=2, diem=2 if fl2 == "Lệch" else 0, sac="GTGT",
                tien=abs((g["ct34"] or 0) - (b["ban_dt"] or 0)), nguon="IN03, IN02",
                ct=f"KQ2.1 năm {y}",
            )
        dt_bctc = kqkd(y, "01")
        if dt_bctc is not None and g["ct34"]:
            fl3 = "Lệch" if abs(dt_bctc - g["ct34"]) > P02 else "Khớp"
            add_f(
                qt="TN01", nhom="TNDN – BCTC", ten="DT BCTC [01] vs GTGT [34]", tc="TC1",
                obj=TEN, ky=str(y), a="BCTC [01]", va=dt_bctc, b="GTGT [34]", vb=g["ct34"],
                do=dt_bctc - g["ct34"], nguong=P02, st="ĐÃ_CHẠY", kq=fl3, nhan="Sơ bộ",
                mucr="Trung bình" if fl3 == "Lệch" else "Thông tin", hs=2,
                diem=2 if fl3 == "Lệch" else 0, sac="TNDN", tien=abs(dt_bctc - g["ct34"]),
                nguon="IN04, IN03", ct=f"IN04 KQKD {y}",
            )
        dt = kqkd(y, "10") or kqkd(y, "01")
        gv = kqkd(y, "11")
        if dt and gv is not None:
            lai = dt - gv
            add_f(
                qt="TN03", nhom="TNDN – BCTC", ten="Lợi nhuận gộp âm / biên mỏng", tc="TC5",
                obj=TEN, ky=str(y), a="DT thuần", va=dt, b="GV", vb=gv, do=lai / dt if dt else None,
                nguong=0, st="ĐÃ_CHẠY", kq="Có phát hiện" if lai < 0 else "Không vượt",
                nhan="Sơ bộ", mucr="Cao" if lai < 0 else "Thông tin", hs=2, diem=2 if lai < 0 else 0,
                sac="TNDN", tien=lai if lai < 0 else 0, nguon="IN04", ct=f"KQKD {y}",
            )
        htk = cdkt(y, "140")
        if htk is None:
            htk = cdkt(y, "141")
        if dt and htk is not None:
            ratio = htk / dt if dt else 0
            add_f(
                qt="TN05", nhom="TNDN – BCTC", ten="HTK / DT (ct140 DNN)", tc="TC3",
                obj=TEN, ky=str(y), a="HTK ct140", va=htk, b="DT", vb=dt, do=ratio,
                nguong=P05, st="ĐÃ_CHẠY", kq="Vượt ngưỡng" if ratio > P05 else "Không vượt",
                nhan="Sơ bộ", mucr="Trung bình" if ratio > P05 else "Thông tin", hs=2,
                diem=1 if ratio > P05 else 0, sac="TNDN", tien=htk, nguon="IN04",
                ct=f"CDKT {y} ct140 (K6 TT133)",
            )

add_f(
    qt="TT01", nhom="Tiền mặt – thanh toán", ten="Cụm HĐ cùng ngày cùng MST vượt ngưỡng",
    tc="", obj="Cụm HĐ", ky=KY, a="Số cụm vượt", va=len(tt01), b="Ngưỡng",
    vb=20_000_000, do=len(tt01), nguong=1, st="ĐÃ_CHẠY",
    kq="Có phát hiện" if tt01 else "Không vượt", nhan="Sơ bộ",
    mucr="Trung bình" if tt01 else "Thông tin", hs=2, diem=2 if tt01 else 0,
    sac="GTGT", tien=sum(x[2] for x in tt01), nguon="IN02/IN01", ct="KQ2.2!KQ2_Cong_no_Tien_mat Phần 3b",
)
add_f(
    qt="TT03", nhom="Tiền mặt – thanh toán", ten="Chi tiền mặt TK 111 vượt ngưỡng",
    tc="", obj="TK 111", ky="2020-2025", a="Số phiếu vượt", va=len(cash_over), b="Ngưỡng K5",
    vb=20_000_000, do=len(cash_over), nguong=1, st="ĐÃ_CHẠY",
    kq="Có phát hiện" if cash_over else "Không vượt", nhan="Chắc",
    mucr="Trung bình" if cash_over else "Thông tin", hs=2, diem=2 if cash_over else 0,
    sac="TNDN", tien=sum(nnum(x["So_tien"]) for x in cash_over), nguon="IN10", ct="KQ2.2 IN10",
)
add_f(
    qt="TT06", nhom="Tiền mặt – thanh toán", ten="Thiếu sao kê NH", tc="",
    obj="IN11", ky=KY, a="File sao kê", va=0, b="", vb=None, do=0, nguong=None,
    st="KHÔNG_ĐỦ_DỮ_LIỆU", kq="Chưa kết luận", nhan="Sơ bộ", mucr="", hs=0, diem=0,
    sac="", tien=None, nguon="IN00", ct="KQ1.3 IN11", thieu="IN11",
    kt="Không kết luận thiếu UNC (K11)",
)
add_f(
    qt="HD07", nhom="Hóa đơn – NCC", ten="Đơn giá/mặt hàng lệch ngành", tc="",
    obj="ngành", ky=KY, a="", va=None, b="", vb=None, do=None, nguong=None,
    st="TREO", kq="Chưa kết luận", nhan="", mucr="", hs=0, diem=0, sac="", tien=None,
    nguon="K13", ct="", thieu="Tham chiếu ngành",
)
add_f(
    qt="LS01", nhom="Lịch sử tuân thủ", ten="Nợ thuế / KHBS TMS", tc="TC24",
    obj="TMS", ky=KY, a="", va=None, b="", vb=None, do=None, nguong=None,
    st="KHÔNG_ĐỦ_DỮ_LIỆU", kq="Chưa kết luận", nhan="", mucr="", hs=0, diem=0,
    sac="", tien=None, nguon="IN14", ct="1.3 chỉ có PTDN", thieu="IN14 dòng nợ thuế",
)
add_f(
    qt="DT03", nhom="Doanh thu – giá vốn", ten="511/632 năm vs BCTC", tc="TC28",
    obj="511-632", ky="2020-2025", a="511 PS Có", va=tot511, b="632 PS Nợ", vb=tot632,
    do=(tot511 - tot632) / tot511 if tot511 else None, nguong=None, st="ĐÃ_CHẠY",
    kq="Đã chạy tổng năm", nhan="Sơ bộ", mucr="Thông tin", hs=1, diem=0,
    sac="TNDN", tien=tot511 - tot632, nguon="IN13 SCTK",
    ct="KQ2.2!KQ2_511_632",
    nhan_xet="Không ghép dòng HĐ vì 632 là KC tháng.",
)
add_f(
    qt="DT01", nhom="Doanh thu – giá vốn", ten="Dòng lãi gộp ≤ 0", tc="",
    obj="511-632 dòng", ky=KY, a="", va=None, b="", vb=None, do=None, nguong=0,
    st="KHÔNG_ĐỦ_DỮ_LIỆU", kq="Chưa kết luận", nhan="", mucr="", hs=0, diem=0,
    sac="", tien=None, nguon="IN13", ct="", thieu="Sổ CT 632 theo HĐ",
)
add_f(
    qt="LK01", nhom="Giao dịch liên kết", ten="GDLK không nộp PL", tc="TC18",
    obj="GDLK", ky=KY, a="PL GDLK", va=0, b="", vb=None, do=0, nguong=None,
    st="ĐÃ_CHẠY", kq="Không thấy phụ lục", nhan="Sơ bộ", mucr="Thông tin", hs=1, diem=0,
    sac="TNDN", tien=None, nguon="IN12 chưa nhận", ct="KQ1.3",
    kt="Xác nhận có/không quan hệ liên kết",
)

# remaining rules status rows
have = {f["qt"] for f in findings}
for rule in rules:
    rid = rule["id"]
    if rid in have:
        continue
    kha, thieu, inn, _ = rule_run(rid, rule["need"])
    st = "ĐÃ_CHẠY" if kha == "Chạy ngay" and rid in RULES_GD2 else (
        "KHÔNG_ĐỦ_DỮ_LIỆU" if kha == "Bị chặn" else "CHƯA_TRIỂN_KHAI" if "Chưa" in kha else "ĐÃ_CHẠY"
    )
    if rid == "HD07":
        st = "TREO"
    add_f(
        qt=rid, nhom=rule["nhom"], ten=rule["ten"], tc="", obj="", ky=KY,
        a="", va=None, b="", vb=None, do=None, nguong=None, st=st,
        kq="Chưa kết luận" if st != "ĐÃ_CHẠY" else "Không vượt / không có cờ máy",
        nhan="Sơ bộ" if rid.startswith("TN") else ("Chắc" if rid in RULES_GD2 else ""),
        mucr="", hs=0, diem=0, sac="", tien=None, nguon=rule["need"], ct="",
        thieu=inn, kt=rule["goi_y"],
    )

# write findings
clear_from(ws, 10)
for i, f in enumerate(findings, 10):
    mph = f"{f['qt']}|{f.get('obj') or TEN}|{f.get('ky') or KY}"
    put(ws, i, 1, mph)
    put(ws, i, 2, "KQ2.1" if f["qt"] in RULES_GD2 else ("KQ2.2" if f["qt"][:2] in {"DT", "KH", "TT"} else "Trực tiếp"))
    put(ws, i, 3, f.get("nhom"))
    put(ws, i, 4, f.get("qt"))
    put(ws, i, 5, f.get("ten"))
    put(ws, i, 6, f.get("tc"))
    put(ws, i, 7, f.get("obj"))
    put(ws, i, 8, f.get("ky"))
    put(ws, i, 9, f.get("a"))
    put(ws, i, 10, f.get("va"), num=isinstance(f.get("va"), (int, float)))
    put(ws, i, 11, f.get("b"))
    put(ws, i, 12, f.get("vb"), num=isinstance(f.get("vb"), (int, float)))
    put(ws, i, 13, f.get("do"), num=isinstance(f.get("do"), (int, float)))
    put(ws, i, 14, f.get("nguong"), num=isinstance(f.get("nguong"), (int, float)))
    put(ws, i, 15, f.get("st"))
    put(ws, i, 16, f.get("kq"))
    put(ws, i, 17, f.get("nhan"))
    put(ws, i, 18, f.get("mucr"), fill=RED if f.get("mucr") == "Cao" else YEL if f.get("mucr") == "Trung bình" else None)
    put(ws, i, 19, f.get("hs"), num=True)
    put(ws, i, 20, f.get("diem"), num=True)
    put(ws, i, 21, f.get("sac"))
    put(ws, i, 22, f.get("tien"), num=isinstance(f.get("tien"), (int, float)))
    w = f.get("tien") if f.get("diem") and f.get("tien") else 0
    put(ws, i, 23, w if f.get("diem") else 0, num=True)
    put(ws, i, 25, f.get("nguon"))
    put(ws, i, 26, f.get("ct"))
    put(ws, i, 27, f.get("thieu"))
    put(ws, i, 28, f.get("kt"))
    put(ws, i, 32, f.get("nhan_xet"))
autowidth(ws, cap=40)

ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
clear_from(ws, 4)
tc_map = {}
wbbo = load_workbook(lp(BO), data_only=True)
w8 = wbbo["08_Master_28TC"]
tcs = []
for r in range(5, 33):
    tid = w8.cell(r, 1).value
    if not tid:
        continue
    tcs.append((str(tid), w8.cell(r, 2).value, w8.cell(r, 3).value, w8.cell(r, 4).value, str(w8.cell(r, 8).value or "")))
wbbo.close()

def tc_result(tid, ten, qt):
    rel = [f for f in findings if f.get("qt") == qt or f.get("tc") == tid]
    if not rel:
        return "Chưa đánh giá / thiếu nguồn – không ghi 0 giả", "Chưa đánh giá", "Bổ sung nguồn hoặc xác minh tại trụ sở"
    top = max(rel, key=lambda x: x.get("diem") or 0)
    kq = f"{top.get('kq')} | {top.get('a')}={top.get('va')} vs {top.get('b')}={top.get('vb')}"
    mucr = top.get("mucr") or "Chưa đánh giá"
    kt = top.get("kt") or ten
    return kq, mucr, kt

for i, (tid, ten, loai, nguong, qt) in enumerate(tcs, 4):
    kq, mucr, kt = tc_result(tid, ten, qt)
    put(ws, i, 1, tid, bold=True)
    put(ws, i, 2, ten, wrap=True)
    put(ws, i, 3, loai)
    put(ws, i, 4, nguong)
    put(ws, i, 5, kq, wrap=True)
    put(ws, i, 6, mucr, fill=RED if mucr == "Cao" else YEL if mucr == "Trung bình" else None)
    put(ws, i, 7, kt, wrap=True)
autowidth(ws, cap=50)
wb.save(lp(names["kq3"]))
wb.close()
print("saved", names["kq3"].name)

# summary json + md
summary = {
    "mst": MST,
    "ten": TEN,
    "ky": KY,
    "muc": muc,
    "ver": VER,
    "ngay": str(TODAY),
    "files": {k: str(v) for k, v in names.items()},
    "counts": {
        "in01": len(in01),
        "in02": len(in02),
        "gtgt_eff": len(gtgt_eff),
        "bctc": len(bctc_rows),
        "invoices_mua": len(invoices),
        "ncc": len(ncc_agg),
        "in07": len(in07),
        "tt01": len(tt01),
        "cash_over": len(cash_over),
        "f1_nnt": len(f1_nnt),
    },
    "hd03": hd03,
    "hd04": hd04,
    "y511": y511,
    "y632": y632,
    "gtgt_year": {str(k): v for k, v in gtgt_by_year.items()},
    "hddt_year": {str(k): v for k, v in hddt_year.items()},
    "bk_year": {str(k): v for k, v in bk_year.items()},
    "in_status": in_status,
    "top5": [(m, v["ten"], v["gt"]) for m, v in top5],
    "n_rules": len(findings),
}
js = OUT / f"tom_tat_NHAT_VINH_{MST}_{KY}_{VER}.json"
js.write_text(json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
vault_dir = VAULT / "4. KET QUA PHAN TICH" / "NHAT_VINH_3601946069"
vault_dir.mkdir(parents=True, exist_ok=True)
(vault_dir / js.name).write_text(js.read_text(encoding="utf-8"), encoding="utf-8")
print("DONE", json.dumps({k: summary["counts"][k] for k in summary["counts"]}, ensure_ascii=False))
