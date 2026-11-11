# -*- coding: utf-8 -*-
"""Fill v1.4 KQ1.1-KQ3 for THANH THANH PHUONG 3600840461."""
from __future__ import annotations

import json
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

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3600840461"
TEN = "CÔNG TY TNHH THANH THANH PHƯƠNG"
KY = "2020-2025"
YEARS = [2020, 2021, 2022, 2023, 2024, 2025]
Y511 = [2022, 2023, 2024, 2025]
TODAY = date(2026, 11, 9)
VER = "v1"
MAU_VER = "1.4"
LEVEL = "D"
DN_STATUS = "Một phần"

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"
    r"\THANH_THANH_PHUONG_3600840461"
)
OUT = ROOT / "4. KET QUA PHAN TICH"
SAO = ROOT / "7. Sao luu tong hop"
VAULT_KQ = VAULT / "ket-qua"
CACHE = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\scratch\ttp_3600840461\cache")

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FONT_H = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
FILL_HDR = PatternFill("solid", fgColor="1F4E78")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
FILL_GRN = PatternFill("solid", fgColor="E2EFDA")
FILL_NOTE = PatternFill("solid", fgColor="D9E1F2")
RED_F = Font(name="Times New Roman", size=11, bold=True, color="C00000")
ORG_F = Font(name="Times New Roman", size=11, bold=True, color="B25900")
GRN_F = Font(name="Times New Roman", size=11, color="385723")
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
NUM = "#,##0"
PCT = "0.00%"
DATE_FMT = "dd/mm/yyyy"
P13, P14, P05, P23, P03 = 0.05, 0.50, 0.30, 0.10, 0.02


def lp(p) -> str:
    s = str(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s


def dparse(s):
    if not s:
        return None
    if isinstance(s, date) and not isinstance(s, datetime):
        return s
    if isinstance(s, datetime):
        return s.date()
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def mst_base(s: str) -> str:
    s = str(s or "").split("-")[0]
    return re.sub(r"\D", "", s)


def so_key(v) -> str:
    s = str(v or "").strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.lstrip("0") or "0"


def tnr(cell, bold=None, color=None):
    f = cell.font
    cell.font = Font(
        name="Times New Roman",
        size=f.size if f and f.size else 11,
        bold=bool(bold) if bold is not None else (f.bold if f else False),
        italic=False,
        color=color or (f.color if f else None),
    )


def apply_tnr(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            tnr(cell)


def autowidth(ws, cap=60, min_w=10):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, 4000), max_col=ws.max_column or 1):
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


def unmerge_origin(ws, r, c):
    cell = ws.cell(r, c)
    if not isinstance(cell, MergedCell):
        return r, c
    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            return rng.min_row, rng.min_col
    return r, c


def put(ws, r, c, v, num=False, dt=False, pct=False, bold=False, fill=None, wrap=False, center=False, font=None):
    r, c = unmerge_origin(ws, r, c)
    cell = ws.cell(r, c, v)
    if font:
        cell.font = font
    else:
        tnr(cell, bold=bold)
    cell.border = THIN
    al = Alignment(vertical="center", wrap_text=wrap, horizontal="center" if center else None)
    cell.alignment = al
    if fill:
        cell.fill = fill
    if num and isinstance(v, (int, float)):
        cell.number_format = NUM
    if pct and isinstance(v, (int, float)):
        cell.number_format = PCT
    if dt and v not in (None, ""):
        cell.number_format = DATE_FMT
    return cell


def write_row(ws, r, values, nums=(), dates=(), fill=None):
    for c, v in enumerate(values, 1):
        put(ws, r, c, v, num=c in nums and isinstance(v, (int, float)), dt=c in dates, fill=fill, wrap=c >= 8)


def sub_all(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str):
                v = cell.value
                v = v.replace("[MST]", MST).replace("[KY]", KY).replace("[NGAY]", TODAY.strftime("%d/%m/%Y"))
                v = v.replace("[TÊN DOANH NGHIỆP]", TEN).replace("[DN]", TEN)
                v = v.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")
                v = v.replace("[Đủ | Một phần | Chưa có]", DN_STATUS)
                if v != cell.value:
                    cell.value = v
            tnr(cell)


def fnum(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def header_map(ws, row=9):
    m = {}
    for c in range(1, (ws.max_column or 1) + 1):
        v = ws.cell(row, c).value
        if v:
            m[str(v).strip()] = c
    return m


print("LOAD CACHE")
manifest = json.loads((CACHE / "manifest.json").read_text(encoding="utf-8"))
hddt = json.loads((CACHE / "hddt.json").read_text(encoding="utf-8"))
bk = json.loads((CACHE / "bangke.json").read_text(encoding="utf-8"))
nkc = json.loads((CACHE / "nkc.json").read_text(encoding="utf-8"))
congno = json.loads((CACHE / "congno.json").read_text(encoding="utf-8"))
nld = json.loads((CACHE / "nld.json").read_text(encoding="utf-8"))
danhba = json.loads((CACHE / "danhba.json").read_text(encoding="utf-8"))
f0 = json.loads((CACHE / "f0.json").read_text(encoding="utf-8"))
nguon = json.loads((CACHE / "nguon_rui_ro.json").read_text(encoding="utf-8"))
summary = json.loads((CACHE / "summary.json").read_text(encoding="utf-8"))
eff_gtgt = json.loads((CACHE / "eff_gtgt.json").read_text(encoding="utf-8"))
eff_tndn = json.loads((CACHE / "eff_tndn.json").read_text(encoding="utf-8"))
eff_bctc = json.loads((CACHE / "eff_bctc.json").read_text(encoding="utf-8"))
eff_tncn = json.loads((CACHE / "eff_tncn.json").read_text(encoding="utf-8"))

f0m = {mst_base(k): v for k, v in f0.items()}
dbm = {mst_base(k): v for k, v in danhba.items()}
nnt_f1 = bool(nguon.get("f1_nnt"))
nnt_ca = bool((nguon.get("ca") or {}).get("nnt_hits"))
nnt_f0 = MST in f0m

# bang ke unique + skip dummy year
bk_u = []
seen_bk = {}
for r in bk:
    if r.get("nam") and r["nam"] < 2016:
        continue
    k = (r.get("chieu"), so_key(r.get("so")), mst_base(r.get("mst")), r.get("ngay"))
    if k in seen_bk:
        i = seen_bk[k]
        bk_u[i]["gt"] = (bk_u[i].get("gt") or 0) + (r.get("gt") or 0)
        bk_u[i]["thue"] = (bk_u[i].get("thue") or 0) + (r.get("thue") or 0)
    else:
        seen_bk[k] = len(bk_u)
        rec = dict(r)
        rec["so_key"] = so_key(r.get("so"))
        rec["mst_dt"] = r.get("mst")
        rec["thue_suat"] = r.get("ts")
        rec["so_hd"] = r.get("so")
        bk_u.append(rec)

hddt_ok = [r for r in hddt if r.get("tinh") == "Có"]


def sum_hd(chieu, nam=None, field="gt", tinh_only=True):
    s = n = 0
    rows = hddt_ok if tinh_only else hddt
    for r in rows:
        if r.get("chieu") != chieu:
            continue
        if nam is not None and r.get("nam") != nam:
            continue
        s += r.get(field) or 0
        n += 1
    return s, n


def sum_bk(chieu, nam=None, field="gt"):
    s = n = 0
    for r in bk_u:
        if r.get("chieu") != chieu:
            continue
        if nam is not None and r.get("nam") != nam:
            continue
        s += r.get(field) or 0
        n += 1
    return s, n


def ct(rec, *codes):
    m = (rec or {}).get("chi_tieu") or {}
    for c in codes:
        if m.get(c) not in (None, ""):
            return fnum(m.get(c))
    return None


def bctc_y(y):
    rec = eff_bctc.get(str(y)) or {}
    k = rec.get("kqkd") or {}
    cd = rec.get("cdkt_cuoi") or {}
    return {
        "dt": fnum(k.get("01")),
        "gv": fnum(k.get("11")),
        "ln_gop": fnum(k.get("20")),
        "dt_tc": fnum(k.get("21")),
        "cp_tc": fnum(k.get("23")),
        "thue_tndn": fnum(k.get("51") or k.get("22")),
        "lnst": fnum(k.get("60") or k.get("50")),
        "ln_tt": fnum(k.get("50") or k.get("30")),
        "htk": fnum(cd.get("141") or cd.get("140")),
        "htk140": fnum(cd.get("140")),
        "htk141": fnum(cd.get("141")),
        "pt": fnum(cd.get("131")),
        "tr": fnum(cd.get("311") or cd.get("331")),
        "tm": fnum(cd.get("111")),
        "nh": fnum(cd.get("112")),
        "von": fnum(cd.get("411")),
        "thue_kt": fnum(cd.get("152")),
        "tscd": fnum(cd.get("211")),
        "file": rec.get("file"),
        "cdkt": cd,
        "kqkd": k,
    }


BM = {y: bctc_y(y) for y in YEARS}


def tndn_dt(y):
    rec = eff_tndn.get(str(y)) or {}
    m = rec.get("chi_tieu") or {}
    if y == 2020:
        return fnum(m.get("01"))  # mẫu cũ: [01] = DT
    return fnum(m.get("04"))


def gtgt_year(y):
    d = {"34": 0.0, "35": 0.0, "23": 0.0, "24": 0.0, "32": 0.0, "40": 0.0, "43": 0.0, "kys": []}
    n = 0
    last43 = 0.0
    for ky, rec in eff_gtgt.items():
        if rec.get("nam") != y:
            continue
        m = rec.get("chi_tieu") or {}
        d["34"] += fnum(m.get("34")) or 0
        d["35"] += fnum(m.get("35")) or 0
        d["23"] += fnum(m.get("23")) or 0
        d["24"] += fnum(m.get("24")) or 0
        d["32"] += fnum(m.get("32")) or 0
        d["40"] += fnum(m.get("40")) or 0
        last43 = fnum(m.get("43")) or last43
        d["kys"].append(ky)
        n += 1
    d["43"] = last43
    d["n"] = n
    return d


GY = {y: gtgt_year(y) for y in YEARS}

# NKC 511/632
y511 = {y: None for y in YEARS}
y632 = {y: None for y in YEARS}
y133 = {y: None for y in YEARS}
y156_xuat = {y: None for y in YEARS}
y111_chi = {y: 0.0 for y in YEARS}
lines_511 = []
lines_632 = []
cash_chi = []
nkc_mong = set()
for pack in nkc:
    y = pack.get("nam")
    tk = pack.get("tk") or {}
    t511 = tk.get("5111") or tk.get("511") or {}
    t632 = tk.get("632") or {}
    t133 = tk.get("1331") or tk.get("133") or {}
    t156 = tk.get("1561") or tk.get("156") or {}
    t111 = tk.get("1111") or tk.get("111") or {}
    co511 = t511.get("co") or 0
    no632 = t632.get("no") or 0
    if no632 == 0:
        no632 = t632.get("co") or 0
    y511[y] = co511
    y632[y] = no632
    y133[y] = t133.get("no") or 0
    y156_xuat[y] = t156.get("co") or 0
    y111_chi[y] = t111.get("co") or 0
    if y in (2021, 2022):
        nkc_mong.add(y)
    for a in pack.get("lines_511") or []:
        lines_511.append(a)
    for a in pack.get("lines_632") or []:
        lines_632.append(a)
    for a in pack.get("cash_chi") or []:
        a = dict(a)
        a["nam"] = y
        cash_chi.append(a)

# NCC
ncc = defaultdict(lambda: {"ten": "", "gt": 0.0, "thue": 0.0, "n_hd": 0, "n_bk": 0, "src": set()})
for r in hddt_ok:
    if r["chieu"] != "MUA":
        continue
    m = mst_base(r.get("mst_dt"))
    if not m:
        continue
    rec = ncc[m]
    rec["ten"] = rec["ten"] or r.get("ten_ban") or ""
    rec["gt"] += r.get("gt") or 0
    rec["thue"] += r.get("thue") or 0
    rec["n_hd"] += 1
    rec["src"].add("HĐĐT")
for r in bk_u:
    if r["chieu"] != "MUA":
        continue
    m = mst_base(r.get("mst_dt") or r.get("mst"))
    if not m:
        continue
    rec = ncc[m]
    rec["ten"] = rec["ten"] or r.get("ten") or ""
    rec["n_bk"] += 1
    rec["src"].add("Bảng kê")
    if rec["n_hd"] == 0:
        rec["gt"] += r.get("gt") or 0
        rec["thue"] += r.get("thue") or 0

tot_mua = sum(v["gt"] for v in ncc.values())
top5 = sorted(ncc.items(), key=lambda x: -x[1]["gt"])[:5]
top5_gt = sum(v["gt"] for _, v in top5)
top5_ty = (top5_gt / tot_mua) if tot_mua else 0

f0_hits = []
for m, rec in ncc.items():
    info = f0m.get(m)
    if info:
        f0_hits.append((m, rec, info))
f0_gt = sum(r["gt"] for _, r, _ in f0_hits)
f0_thue = sum(r["thue"] for _, r, _ in f0_hits)
f0_ty = (f0_gt / tot_mua) if tot_mua else 0

# TT01 clusters
clusters = defaultdict(lambda: {"n": 0, "gt": 0.0, "tt": 0.0, "ten": ""})
for r in hddt_ok:
    if r["chieu"] != "MUA" or not r.get("ngay"):
        continue
    k = (r.get("mst_dt"), r.get("ngay"))
    clusters[k]["n"] += 1
    clusters[k]["gt"] += r.get("gt") or 0
    clusters[k]["tt"] += r.get("tt") or 0
    clusters[k]["ten"] = r.get("ten_ban") or ""
tt01 = []
for (mst, ngay), v in clusters.items():
    d = dparse(ngay)
    nguong = 5_000_000 if d and d >= date(2025, 7, 1) else 20_000_000
    if v["tt"] >= nguong or v["gt"] >= nguong:
        tt01.append({"mst": mst, "ngay": ngay, "n": v["n"], "gt": v["gt"], "tt": v["tt"], "ten": v["ten"], "nguong": nguong})
tt01.sort(key=lambda x: -x["tt"])

cn_map = {(x.get("nam"), x.get("tk")): x for x in congno}
nld_ds = [x for x in nld if x.get("sheet", "").startswith("2.") and x.get("Họ và tên")]
nld_qtt = {}
for y in YEARS:
    rec = eff_tncn.get(str(y)) or {}
    m = rec.get("chi_tieu") or {}
    if y == 2020:
        nld_qtt[y] = {"so_ld": m.get("21") or m.get("22"), "tnct": m.get("28") or m.get("29"), "thue_kt": m.get("36") or 0, "file": rec.get("file")}
    else:
        nld_qtt[y] = {"so_ld": m.get("16") or m.get("17"), "tnct": m.get("23") or m.get("24"), "thue_kt": m.get("36") or m.get("37") or 0, "file": rec.get("file")}


def ncc_flag(m):
    m = mst_base(m)
    info = f0m.get(m)
    db = dbm.get(m) or {}
    tt = db.get("tt") or ""
    if info:
        return "F0", info.get("nguon") or "CANH_BAO", info.get("so_cv") or "", tt or info.get("ten") or ""
    return "", "", "", tt


# XNT 156
def parse_nxt(path: Path, nam: int):
    try:
        wb = xlrd.open_workbook(str(path), formatting_info=False)
    except Exception as e:
        print("  nxt fail", path.name, e)
        return None
    dk_sl = dk_tt = nh_sl = nh_tt = x_sl = x_tt = ck_sl = ck_tt = 0.0
    first = True
    last_ck = (0.0, 0.0)
    for si in range(wb.nsheets):
        sh = wb.sheet_by_index(si)
        hdr = None
        for r in range(min(12, sh.nrows)):
            row = [str(sh.cell_value(r, c) or "").strip().upper() for c in range(min(12, sh.ncols))]
            if any("TỒN ĐẦU" in x or "TON DAU" in x for x in row) or (len(row) > 4 and row[4] in {"SL", "SỐ LƯỢNG"}):
                hdr = r
                break
        start = (hdr + 2) if hdr is not None else 7
        s_dk_sl = s_dk_tt = s_nh_sl = s_nh_tt = s_x_sl = s_x_tt = s_ck_sl = s_ck_tt = 0.0
        for r in range(start, sh.nrows):
            ma = str(sh.cell_value(r, 1) or "").strip()
            tenh = str(sh.cell_value(r, 2) or "").strip().lower()
            if not ma and not tenh:
                continue
            if "tổng" in tenh or "tong cong" in tenh or "cộng" in tenh:
                continue

            def celln(c):
                try:
                    v = sh.cell_value(r, c)
                    return float(v or 0)
                except Exception:
                    return 0.0

            s_dk_sl += celln(4)
            s_dk_tt += celln(5)
            s_nh_sl += celln(6)
            s_nh_tt += celln(7)
            s_x_sl += celln(8)
            s_x_tt += celln(9)
            s_ck_sl += celln(10)
            s_ck_tt += celln(11)
        if first:
            dk_sl, dk_tt = s_dk_sl, s_dk_tt
            first = False
        nh_sl += s_nh_sl
        nh_tt += s_nh_tt
        x_sl += s_x_sl
        x_tt += s_x_tt
        last_ck = (s_ck_sl, s_ck_tt)
    ck_sl, ck_tt = last_ck
    return {
        "nam": nam,
        "dk_sl": dk_sl,
        "dk_tt": dk_tt,
        "nhap_sl": nh_sl,
        "nhap_tt": nh_tt,
        "xuat_sl": x_sl,
        "xuat_tt": x_tt,
        "ck_sl": ck_sl,
        "ck_tt": ck_tt,
        "file": path.name,
    }


kho = {}
kho_dir = ROOT / r"2. HO SO NNT CUNG CAP\2.2. Kho - Chi phi - Gia thanh"
if kho_dir.exists():
    for p in kho_dir.glob("*.xls"):
        m = re.search(r"20(2[0-5])", p.name)
        if not m:
            continue
        y = int("20" + m.group(1))
        print("parse XNT", p.name)
        rec = parse_nxt(p, y)
        if rec:
            kho[y] = rec
            print(" ", y, "DK", round(rec["dk_tt"]), "NH", round(rec["nhap_tt"]), "XU", round(rec["xuat_tt"]), "CK", round(rec["ck_tt"]))

print("511", {y: y511.get(y) for y in YEARS})
print("632", {y: y632.get(y) for y in YEARS})
print("GT34", {y: GY[y]["34"] for y in YEARS})
print("top5", round(top5_ty, 4), "f0", len(f0_hits), "tt01", len(tt01), "ncc", len(ncc))
print("NNT F1/F0/CA", nnt_f1, nnt_f0, nnt_ca)

has = {
    "IN01": True,
    "IN02": True,
    "IN03": True,
    "IN04": True,
    "IN05": True,
    "IN07": True,
    "IN08": True,
    "IN09": bool(kho),
    "IN10": bool(cash_chi),
    "IN11": False,
    "IN12": False,
    "IN13": bool(lines_511),
    "IN14": True,
}

OUT.mkdir(parents=True, exist_ok=True)
SAO.mkdir(parents=True, exist_ok=True)
VAULT_KQ.mkdir(parents=True, exist_ok=True)

files = {
    "kq11": (MAU / "KQ1.1_Kiem_ke_CQT.xlsx", OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_{VER}.xlsx"),
    "kq12": (MAU / "KQ1.2_Kiem_ke_ho_so_NNT.xlsx", OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_{VER}.xlsx"),
    "kq13": (MAU / "KQ1.3_Danh_gia_ho_so.xlsx", OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx"),
    "kq21": (MAU / "KQ2.1_Phan_tich_CQT.xlsx", OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_{VER}.xlsx"),
    "kq22": (MAU / "KQ2.2_Phan_tich_ho_so_cty.xlsx", OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_{VER}.xlsx"),
    "kq3": (MAU / "KQ3_Ket_luan_danh_gia.xlsx", OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_{VER}.xlsx"),
}
for src, dst in files.values():
    shutil.copy2(lp(src), lp(dst))
    print("copied", dst.name)

# =====================================================================
print("FILL KQ1.1")
wb = load_workbook(lp(files["kq11"][1]))
ws = wb["IN00_Manifest"]
cqt = [m for m in manifest if m["goc"] == "CQT"]
clear_from(ws, 3)
for i, m in enumerate(cqt, 1):
    write_row(
        ws,
        2 + i,
        [
            f"N{i:03d}",
            m.get("ma_in"),
            "CQT",
            (m.get("ext") or "").lstrip(".").upper() or "KHAC",
            dparse(m.get("mtime")),
            2020,
            2025,
            m.get("sha256"),
            "Có" if "bo" in (m.get("name") or "").lower() and "sung" in (m.get("name") or "").lower() else "Không",
            "Có",
            None,
            0,
            "Giữ nguyên",
            "",
            m.get("rel"),
        ],
        dates={5},
    )
ws.cell(1, 2).value = f"{TEN} - MST {MST} - kỳ {KY} - {TODAY.strftime('%d/%m/%Y')} - {VER}"

ws = wb["IN01_HDDT"]
clear_from(ws, 3)
for i, r in enumerate(hddt, 1):
    ngay = dparse(r.get("ngay"))
    q = ((r.get("thang") or 1) - 1) // 3 + 1 if r.get("thang") else None
    write_row(
        ws,
        2 + i,
        [
            r.get("chieu"),
            r.get("nam"),
            q,
            r.get("thang"),
            r.get("ky_hieu"),
            r.get("so"),
            ngay,
            r.get("mst_ban"),
            r.get("ten_ban"),
            r.get("mst_mua"),
            r.get("ten_mua"),
            r.get("mst_dt"),
            "",
            "",
            None,
            None,
            r.get("gt"),
            None,
            r.get("thue"),
            r.get("tt"),
            r.get("trang_thai") or "Hợp lệ",
            r.get("kq_kt") or "",
            r.get("tinh"),
            "",
            f"{r.get('ky_hieu')}|{so_key(r.get('so'))}|{r.get('mst_dt')}",
            r.get("file"),
        ],
        nums={17, 19, 20},
        dates={7},
    )

ws = wb["IN03_ToKhai_GTGT"]
clear_from(ws, 3)
gtgt_list = []
for ky, rec in sorted(eff_gtgt.items(), key=lambda z: (z[1].get("nam") or 0, z[0])):
    if rec.get("nam") in YEARS:
        gtgt_list.append(rec)
for i, x in enumerate(gtgt_list, 1):
    m = x.get("chi_tieu") or {}
    ky = x.get("ky") or ""
    write_row(
        ws,
        2 + i,
        [
            x.get("nam"),
            ky,
            "Bổ sung" if str(x.get("loai_tkhai") or "").upper().startswith("B") else "Chính thức",
            x.get("so_lan") or 0,
            "Có",
            None,
            m.get("22"),
            m.get("23"),
            m.get("24"),
            m.get("25"),
            m.get("26"),
            m.get("29"),
            m.get("30"),
            m.get("31"),
            m.get("32"),
            m.get("33"),
            m.get("34"),
            m.get("35"),
            m.get("36"),
            m.get("37"),
            m.get("38"),
            m.get("39"),
            m.get("40"),
            m.get("41"),
            m.get("42"),
            m.get("43"),
            x.get("file"),
        ],
        nums=set(range(7, 27)),
    )

ws = wb["IN04_BCTC_TNDN"]
clear_from(ws, 3)
r = 3
for y in YEARS:
    b = BM[y]
    pairs = [
        ("KQKD", "01", "Doanh thu bán hàng và cung cấp dịch vụ", b["dt"]),
        ("KQKD", "11", "Giá vốn hàng bán", b["gv"]),
        ("KQKD", "20", "Lợi nhuận gộp", b["ln_gop"]),
        ("KQKD", "21", "Doanh thu hoạt động tài chính", b["dt_tc"]),
        ("KQKD", "50", "Lợi nhuận kế toán trước thuế", b["ln_tt"]),
        ("CDKT", "141", "Hàng tồn kho [141]", b["htk141"]),
        ("CDKT", "140", "Hàng tồn kho [140]", b["htk140"]),
        ("CDKT", "131", "Phải thu khách hàng", b["pt"]),
        ("CDKT", "111", "Tiền mặt", b["tm"]),
        ("CDKT", "112", "Tiền gửi ngân hàng", b["nh"]),
        ("CDKT", "411", "Vốn góp chủ sở hữu", b["von"]),
        ("CDKT", "152", "Thuế GTGT được khấu trừ", b["thue_kt"]),
        ("CDKT", "211", "TSCĐ hữu hình", b["tscd"]),
    ]
    tn = (eff_tndn.get(str(y)) or {}).get("chi_tieu") or {}
    dt = tndn_dt(y)
    if dt is not None:
        pairs.append(("TNDN", "04" if y != 2020 else "01", "Doanh thu QTT TNDN (2020 mẫu cũ=[01])", dt))
    if tn.get("13") is not None:
        pairs.append(("TNDN", "13", "Giá vốn QTT TNDN [13]", tn.get("13")))
    if tn.get("18") is not None:
        pairs.append(("TNDN", "18", "Lợi nhuận trước thuế [18]", tn.get("18")))
    for bao, ma, tenct, so in pairs:
        write_row(ws, r, [y, bao, ma, tenct, so, None, "TT200 - B01-DN đọc từ XML CQT", b.get("file")], nums={5})
        r += 1

ws = wb["IN05_NLD_TNCN"]
clear_from(ws, 3)
rr = 3
for y in YEARS:
    q = nld_qtt.get(y) or {}
    if q.get("tnct") is None and q.get("so_ld") is None:
        continue
    write_row(
        ws,
        rr,
        ["QTT", y, f"Quyết toán TNCN {y}", "", MST, "05/QTT-TNCN", "Tiền lương", q.get("tnct"), q.get("thue_kt"), "Không", None, "", "", f"Số LĐ {int(q.get('so_ld') or 0)}", "Không", q.get("file")],
        nums={8, 9},
    )
    rr += 1
    for row in nld_ds:
        if row.get("Năm") != y:
            continue
        write_row(
            ws,
            rr,
            [
                "NLD",
                y,
                row.get("Họ và tên"),
                row.get("CCCD / CMND"),
                row.get("MST cá nhân"),
                "05-1",
                "Tiền lương HĐLĐ",
                row.get("Tổng TNCT (VNĐ)"),
                row.get("Thuế TNCN đã khấu trừ"),
                "Không",
                None,
                row.get("Ủy quyền QTT"),
                "",
                row.get("Hợp đồng LĐ"),
                "Không",
                row.get("Ghi chú"),
            ],
            nums={8, 9},
        )
        rr += 1

ws = wb["IN14_TMS_Nghia_vu"]
clear_from(ws, 3)
tms = [m for m in manifest if m.get("ma_in") == "IN14_TMS_Nghia_vu"]
for i, m in enumerate(tms, 1):
    write_row(
        ws,
        2 + i,
        [MST, "PHAN_TICH_BCTC_TMS", "BCTC", "2016-2025", None, None, None, None, None, None, None, "", None, "File PTDN TMS - chưa bóc dòng nợ thuế. LS01/LS02: chưa đủ lịch sử nợ/chậm nộp.", "TMS_PTDN", m.get("rel")],
        dates={6},
    )
apply_tnr(wb["IN00_Manifest"])
wb.save(lp(files["kq11"][1]))
wb.close()
print(" KQ1.1 saved")

# =====================================================================
print("FILL KQ1.2")
wb = load_workbook(lp(files["kq12"][1]))
ws = wb["IN00_Manifest"]
dn = [m for m in manifest if m["goc"] in {"DN", "TRA_CUU", "BIEN_BAN"}]
clear_from(ws, 3)
for i, m in enumerate(dn, 1):
    write_row(
        ws,
        2 + i,
        [f"D{i:03d}", m.get("ma_in"), m.get("goc"), (m.get("ext") or "").lstrip(".").upper(), dparse(m.get("mtime")), 2020, 2025, m.get("sha256"), "Không", "Có", None, 0, "Giữ nguyên", "", m.get("rel")],
        dates={5},
    )
ws.cell(1, 2).value = f"{TEN} - MST {MST} - kỳ {KY} - {TODAY.strftime('%d/%m/%Y')} - {VER}"

ws = wb["IN02_BangKe"]
clear_from(ws, 3)
hd_keys = {(so_key(r.get("so")), mst_base(r.get("mst_dt"))) for r in hddt}
for i, r in enumerate(bk_u, 1):
    k = (so_key(r.get("so")), mst_base(r.get("mst_dt") or r.get("mst")))
    st = "KHOP" if k in hd_keys else "CHI_CO_BK"
    q = ((r.get("thang") or 1) - 1) // 3 + 1 if r.get("thang") else None
    write_row(
        ws,
        2 + i,
        [
            r.get("chieu"),
            r.get("nam"),
            q,
            r.get("thang"),
            r.get("ky_hieu"),
            r.get("so"),
            dparse(r.get("ngay")),
            r.get("mst_goc"),
            r.get("mst") or r.get("mst_dt"),
            r.get("ten"),
            r.get("hang"),
            "",
            None,
            r.get("gt"),
            r.get("ts"),
            r.get("thue"),
            (r.get("gt") or 0) + (r.get("thue") or 0),
            "",
            "Có",
            "Có",
            st,
            "",
            f"|{so_key(r.get('so'))}|{r.get('mst')}",
            r.get("file"),
        ],
        nums={14, 16, 17},
        dates={7},
    )

ws = wb["IN08_So_KeToan"]
clear_from(ws, 3)
rr = 3
keep = {"111", "1111", "112", "1121", "131", "1311", "133", "1331", "156", "1561", "331", "3311", "511", "5111", "632"}
for pack in nkc:
    y = pack.get("nam")
    for tk, v in (pack.get("tk") or {}).items():
        base = tk[:3] if tk[:3] in {"111", "112", "131", "133", "156", "331", "511", "632"} else tk
        if tk not in keep and base not in keep:
            continue
        write_row(ws, rr, [y, None, tk, v.get("no"), v.get("co"), None, None, None, "NKC", pack.get("file")], nums={4, 5})
        rr += 1

ws = wb["IN09_Kho_GiaThanh"]
clear_from(ws, 3)
rr = 3
for y in YEARS:
    k = kho.get(y)
    if not k:
        continue
    write_row(
        ws,
        rr,
        [y, "Hàng hóa TK 156", k["dk_sl"], k["dk_tt"], k["nhap_sl"], k["nhap_tt"], k["xuat_sl"], k["xuat_tt"], k["ck_sl"], k["ck_tt"], None, None, None, None, None, k["file"]],
        nums=set(range(3, 15)),
    )
    rr += 1

ws = wb["IN10_Chi_tien_mat"]
clear_from(ws, 3)
rr = 3
cash_ok = sorted(cash_chi, key=lambda x: -(x.get("tien") or 0))
for x in cash_ok[:800]:
    write_row(
        ws,
        rr,
        [dparse(x.get("ngay")), x.get("nam"), x.get("so"), "331" if x.get("co_331") else "", "", "", x.get("dg"), x.get("tien"), "Chi TK111", "Chưa rõ", "", x.get("file")],
        nums={8},
        dates={1},
    )
    rr += 1

ws = wb["IN11_Sao_ke_NH"]
clear_from(ws, 3)
put(ws, 3, 1, None)
put(ws, 3, 13, "Thư mục 2.3 trống - IN11 chưa nhận. Không điền 0.")

ws = wb["IN12_GDLK"]
clear_from(ws, 3)
put(ws, 3, 9, "Chưa nhận phụ lục GDLK NĐ 132.")

ws = wb["IN13_So_511_632"]
clear_from(ws, 3)
rr = 3
for x in lines_511:
    write_row(
        ws,
        rr,
        [x.get("nam"), x.get("thang"), dparse(x.get("ngay")), x.get("so"), x.get("tk"), "", "", x.get("dg"), "", None, x.get("co"), x.get("no") or 0, "", "", x.get("so"), None, None, None, "", x.get("file")],
        nums={11, 12},
        dates={3},
    )
    rr += 1
for x in lines_632:
    write_row(
        ws,
        rr,
        [x.get("nam"), x.get("thang"), dparse(x.get("ngay")), x.get("so"), x.get("tk"), "", "", x.get("dg"), "", None, x.get("co") or 0, x.get("no") or 0, "", "", x.get("so"), None, None, None, "", x.get("file")],
        nums={11, 12},
        dates={3},
    )
    rr += 1
wb.save(lp(files["kq12"][1]))
wb.close()
print(" KQ1.2 saved")

# =====================================================================
print("FILL KQ1.3")
wb = load_workbook(lp(files["kq13"][1]))
ws = wb["KQ1_Kiem_ke"]
put(ws, 2, 3, MST, bold=True)
put(ws, 2, 10, TEN, bold=True)
put(ws, 3, 3, KY, bold=True)
put(ws, 3, 10, LEVEL, bold=True)
put(ws, 4, 3, "Đủ 8 | Mỏng 3 | Thiếu 0 | Chưa nhận 2")
put(ws, 5, 3, VER)
put(ws, 5, 10, MAU_VER)
put(ws, 6, 3, TODAY, dt=True)
put(ws, 6, 10, DN_STATUS)
put(ws, 7, 3, "KQ2.1, KQ2.2, KQ3")
put(
    ws,
    7,
    10,
    "Kỳ suy ra theo K12 từ BCTC/TNDN 2020-2025, GTGT XML 12/2020 + 2021-2025, bảng kê 2021-2025, HĐĐT bán từ 18/05/2022. NNT không có trên F1/F0/CA. DN thương mại (TK 156, không 621/622). P03/P04 tạm 2% mặc định.",
    wrap=True,
)


def set_in_row(row, dinh, ky_phu, ky_thieu, so_dong, thieu_mst, bs, muc, lydo, tt, viec, ghichu):
    put(ws, row, 6, dinh)
    put(ws, row, 7, ky_phu)
    put(ws, row, 8, ky_thieu)
    put(ws, row, 9, so_dong, num=True)
    put(ws, row, 10, thieu_mst, num=True)
    put(ws, row, 11, bs)
    put(ws, row, 12, muc, bold=True)
    put(ws, row, 13, lydo, wrap=True)
    put(ws, row, 14, tt)
    put(ws, row, 15, viec, wrap=True)
    put(ws, row, 17, ghichu, wrap=True)


set_in_row(10, "Excel", "HĐĐT bán 2022-2025; mua 2021-2025", "Không có HĐĐT bán 2020-2021 (portal)", len(hddt), 0, "Không", "Đủ", "Portal HĐĐT mua+bán; bán bắt đầu 18/05/2022", "Đã chuẩn hóa", "MST, Tinh_vao_tong", "CQT")
set_in_row(11, "Excel", "2021-2025 đủ kỳ tháng; 2020 chỉ 8 HĐ mua", "Bán ra 2020 không có BK", len(bk_u), sum(1 for x in bk_u if not x.get("mst")), "Không", "Đủ", "Bảng kê gốc 01_GTGT theo tháng", "Đã gộp, lọc dummy", "BT6 MST", "DN")
set_in_row(12, "XML", "2021-2025 đủ 12 kỳ/năm; 12/2020", "GTGT T01-T11/2020 thiếu XML", len(gtgt_list), 0, "Có", "Mỏng", "Kê khai tháng; 2020 chỉ còn 12/2020", "Đã chọn bản hiệu lực", "Mã 01 cũ 2020-2021 đã reparse", "CQT")
set_in_row(13, "XML", "BCTC+TNDN 2020-2025", "TNDN 2020 mẫu cũ [01] thay [04]", 6 + 6, 0, "Có TNDN 2024 BS", "Đủ", "B01-DN TT200", "Đã trích chỉ tiêu", "Không suy thông tư", "CQT")
set_in_row(14, "XML/Excel", "QTT 2020-2025; 4 LĐ/năm", "Đa nguồn 3.3 đã có file tổng hợp", len(nld_ds), 0, "Không", "Đủ", "4 LĐ cố định, thuế KT=0", "QTT trước NLD", "Đối chiếu QTT vs danh sách LĐ", "CQT")
set_in_row(15, "Excel sổ CVCBRR+CA", "cập nhật 09/11/2026", "", len(f0m), 0, "Không", "Đủ", "Nguồn chỉ định F0/F1/CA cùng ngày chạy", "Đã khớp MST", "NNT không thuộc F1/CA; 0 NCC F0", "CQT")
set_in_row(16, "Excel NKC", "2023-2025 đủ; 2021-2022 .xls mỏng", "NKC 2021-2022 lệch BCTC (chỉ khoảng T12)", sum(p.get("n_line") or 0 for p in nkc), 0, "Không", "Mỏng", "2023-2025 511 khớp BCTC; 2021-2022 không dùng làm năm đủ", "Đã gom TK", "VNI trên .xls 2021-2022", "DN")
set_in_row(17, "Excel NXT 156", "2021-2025", "Không có 152/155 (DN thương mại)", len(kho), 0, "Không", "Đủ", "5 file AT-NXT 156", "Đã cộng XNT", "Ghi vào Phần 1 kèm chú TK 156", "DN")
set_in_row(18, "Excel (từ NKC 111)", "2021-2025", "Không có sổ 111 tách file", len(cash_ok), 0, "Không", "Đủ", "Chi TK 1111 từ NKC", "Đã lọc", "Lọc chi >= ngưỡng", "DN")
set_in_row(19, "", "", "Thư mục 2.3 trống", 0, 0, "Không", "Chưa nhận", "Không có sao kê/UNC", "Để trống", "Không điền 0", "DN")
set_in_row(20, "", "", "Chưa nhận PL GDLK", 0, 0, "Không", "Chưa nhận", "Không thấy phụ lục NĐ 132", "Để trống", "Không điền 0", "DN")
set_in_row(21, "Excel NKC", "2023-2025 khớp HĐ+ngày sơ bộ; 2021-2022 mỏng", "632 là bút toán kết chuyển tháng, không theo từng HĐ", len(lines_511) + len(lines_632), 0, "Không", "Mỏng", "511 2023-2025 = BCTC [01]", "Đã tách TK", "DN thương mại", "DN")
set_in_row(22, "Excel TMS PTDN", "2016-2025", "Không có bảng nợ thuế/chậm nộp", len(tms), 0, "Không", "Mỏng", "PTDN01/02", "Nhận file", "LS01/LS02 chưa đủ", "CQT")

ws = wb["KQ1_Kha_nang_phan_tich"]
put(ws, 2, 3, MST)
put(ws, 2, 7, KY)
put(ws, 4, 3, VER)
put(ws, 4, 7, MAU_VER)
put(ws, 5, 3, TODAY, dt=True)
put(ws, 5, 7, "Mức D: thiếu IN11/IN12; NKC 2021-2022 mỏng; GTGT 2020 thiếu 11 kỳ")

g3_na = {"KH03", "KH05", "LD05"}
n_now = n_part = n_block = 0
for r in range(10, (ws.max_row or 10) + 1):
    ma = ws.cell(r, 1).value
    if not ma:
        continue
    ma = str(ma).strip()
    need = str(ws.cell(r, 4).value or "")
    if ma in g3_na:
        put(ws, r, 8, "Không áp dụng")
        put(ws, r, 9, "DN thương mại - không SX (không TK 152/155/621)")
        put(ws, r, 10, "")
        put(ws, r, 11, "Dùng KH01/KH04/KH06 trên TK 156")
        continue
    if ma == "HD07":
        put(ws, r, 8, "TREO")
        put(ws, r, 9, "Thiếu tham chiếu ngành (K13)")
        put(ws, r, 10, "K13")
        put(ws, r, 11, "Chạy sơ bộ khi có số ngành")
        n_block += 1
        continue
    missing = []
    for i in re.findall(r"IN\d{2}", need):
        if i == "IN07":
            continue
        if not has.get(i, False):
            missing.append(i)
    if ma in {"LS01", "LS02"}:
        put(ws, r, 8, "Chạy một phần")
        put(ws, r, 9, "TMS PTDN có file nhưng chưa có nợ thuế")
        put(ws, r, 10, "IN14")
        put(ws, r, 11, "Xuất nghĩa vụ nộp thuế TMS")
        n_part += 1
    elif missing:
        kn = "Chạy một phần" if any(has.get(i, False) for i in re.findall(r"IN\d{2}", need)) else "Bị chặn"
        put(ws, r, 8, kn)
        put(ws, r, 9, ",".join(sorted(set(missing))))
        put(ws, r, 10, ",".join(sorted(set(missing))))
        put(ws, r, 11, "Đề nghị NNT bổ sung")
        if kn == "Chạy một phần":
            n_part += 1
        else:
            n_block += 1
    else:
        put(ws, r, 8, "Chạy được ngay")
        put(ws, r, 9, "")
        put(ws, r, 10, "")
        put(ws, r, 11, "")
        n_now += 1
put(ws, 3, 3, n_now, num=True)
put(ws, 3, 5, f"Một phần {n_part}")
put(ws, 3, 7, f"Bị chặn {n_block}")

ws = wb["KQ1_De_nghi_bo_sung"]
put(ws, 2, 3, MST)
put(ws, 2, 7, KY)
put(ws, 3, 3, VER)
put(ws, 3, 7, TODAY, dt=True)
sug = [
    (1, "Sao kê ngân hàng Excel + UNC các món >= 20tr (đến 30/06/2025) và >= 5tr (từ 01/07/2025)", "IN11", "2.3", "File Excel dòng tiền, không chỉ PDF scan", "TT02 TT03 TT05 TT06 GT16", "Thư mục 2.3 trống"),
    (2, "Tờ khai GTGT XML T01-T11/2020 (nếu còn trên TMS)", "IN03", "1.1", "Bổ sung 11 kỳ còn thiếu", "GT01 TN01", "Hiện chỉ có 12/2020"),
    (3, "NKC/sổ cái đầy đủ năm 2021-2022 (file hiện tại chỉ bóc được khoảng T12, lệch BCTC)", "IN08", "2.1", "File Excel đủ phát sinh Nợ/Có cả năm", "DT01 DT02 TN01", "nkc 2021.xls / nkc 2022.xls mỏng"),
    (4, "Phụ lục GDLK NĐ 132 (nếu có quan hệ liên kết) hoặc xác nhận không phát sinh", "IN12", "2.x", "PL 01", "LK01 LK02", "Chưa thấy trên hồ sơ"),
    (5, "File TMS nghĩa vụ nộp thuế / nợ / chậm nộp (không chỉ PTDN CĐKT-KQKD)", "IN14", "1.3", "Bảng nợ thuế", "LS01 LS02", "PTDN01/02 không thay bảng nợ"),
]
for i, rowv in enumerate(sug):
    for c, v in enumerate(rowv, 1):
        put(ws, 10 + i, c, v, wrap=True)
wb.save(lp(files["kq13"][1]))
wb.close()
print(" KQ1.3 saved")

# =====================================================================
print("FILL KQ2.1")
wb = load_workbook(lp(files["kq21"][1]))
ws = wb["KQ2_Doi_chieu_GTGT"]
sub_all(ws)
put(ws, 2, 1, f"Người nộp thuế: {TEN} - MST: {MST} - Kỳ: {KY}", bold=True)

def year_start(y):
    return 9 + (y - 2016) * 14

for y in YEARS:
    base = year_start(y)
    for ky, rec in eff_gtgt.items():
        if rec.get("nam") != y:
            continue
        kyu = (ky or "").replace(" ", "")
        mt = re.search(r"(?:^|/)(0?[1-9]|1[0-2])(?:/|$)", kyu)
        if not mt:
            continue
        mth = int(mt.group(1))
        idx = mth - 1  # T01->Q1 ... T04->Q4, T05..T12
        rr = base + 1 + idx
        m = rec.get("chi_tieu") or {}
        dt = fnum(m.get("34"))
        if dt is None:
            dt = fnum(m.get("32"))
        th = fnum(m.get("35"))
        if th is None:
            th = fnum(m.get("33"))
        put(ws, rr, 5, dt, num=True)
        put(ws, rr, 9, th, num=True)
        put(ws, rr, 11, fnum(m.get("23")), num=True)
        put(ws, rr, 12, fnum(m.get("24")), num=True)
        put(ws, rr, 17, fnum(m.get("40")), num=True)
        put(ws, rr, 18, fnum(m.get("43")), num=True)

# Phần II / III yearly 2020-2025 = rows 160-165
for y in YEARS:
    rr = 156 + (y - 2016)
    hd_b, _ = sum_hd("BAN", y)
    hd_bt, _ = sum_hd("BAN", y, "thue")
    bk_b, _ = sum_bk("BAN", y)
    bk_bt, _ = sum_bk("BAN", y, "thue")
    tk34 = GY[y]["34"]
    tk35 = GY[y]["35"]
    # HDDT 2020-2021 bán: portal không có HĐ -> 0 là số thật
    put(ws, rr, 2, hd_b, num=True)
    put(ws, rr, 3, hd_bt, num=True)
    put(ws, rr, 4, bk_b if (bk_b or y >= 2021) else None, num=True)
    put(ws, rr, 5, bk_bt if (bk_b or y >= 2021) else None, num=True)
    put(ws, rr, 6, tk34 if GY[y]["n"] else None, num=True)
    put(ws, rr, 7, tk35 if GY[y]["n"] else None, num=True)
    put(ws, rr, 8, tndn_dt(y), num=True)
    put(ws, rr, 9, BM[y]["dt"], num=True)
    hd_m, _ = sum_hd("MUA", y)
    hd_mt, _ = sum_hd("MUA", y, "thue")
    bk_m, _ = sum_bk("MUA", y)
    bk_mt, _ = sum_bk("MUA", y, "thue")
    put(ws, rr, 13, hd_m, num=True)
    put(ws, rr, 14, bk_m if bk_m else None, num=True)
    put(ws, rr, 15, GY[y]["23"] if GY[y]["n"] else None, num=True)
    put(ws, rr, 19, hd_mt, num=True)
    put(ws, rr, 20, bk_mt if bk_m else None, num=True)
    put(ws, rr, 21, GY[y]["24"] if GY[y]["n"] else None, num=True)
    if y133.get(y):
        put(ws, rr, 25, y133[y], num=True)
    note = []
    if y == 2020:
        note.append("GTGT XML chỉ 12/2020; TNDN mẫu cũ DT=[01]; HĐĐT bán chưa có trên portal")
    if y in (2020, 2021):
        note.append("HĐĐT bán portal = 0 HĐ (bắt đầu 18/05/2022)")
    if y in nkc_mong:
        note.append("NKC 133 năm này mỏng, không dùng đối chiếu")
    if abs((tk34 or 0) - (BM[y]["dt"] or 0)) > 1 and GY[y]["n"]:
        note.append(f"Lệch TK[34]-BCTC[01]={tk34 - (BM[y]['dt'] or 0):,.0f}")
    put(ws, rr, 28, "; ".join(note), wrap=True)

put(ws, 170, 1, "Nhận xét / nguyên nhân trọng yếu", bold=True, fill=FILL_NOTE)
notes = []
for y in YEARS:
    tk34 = GY[y]["34"] if GY[y]["n"] else None
    hd = sum_hd("BAN", y)[0]
    bkv = sum_bk("BAN", y)[0]
    bctc01 = BM[y]["dt"]
    tndn04 = tndn_dt(y)
    notes.append(
        f"{y}: TK[34]={tk34 if tk34 is not None else 'thiếu kỳ'} | HĐĐT bán={hd:,.0f} | BK bán={bkv:,.0f} | BCTC[01]={bctc01} | TNDN={tndn04} | n_ky={GY[y]['n']}."
    )
notes.append("2022 HĐĐT bán 7,34 tỷ vs tờ khai/BK ~10,82 tỷ: portal HĐĐT chưa đủ năm 2022 (bắt đầu 18/05). 2024 khớp gần 100%. 2025 TK[34] 8,55 tỷ vs BCTC/TNDN 8,09 tỷ.")
notes.append("P03/P04 tạm 2% (mặc định K8, cán bộ chưa chỉ định). Nguồn CVCBRR/CA mtime 09/11/2026.")
for i, line in enumerate(notes):
    put(ws, 171 + i, 1, line, wrap=True, fill=FILL_NOTE)
    try:
        ws.merge_cells(start_row=171 + i, start_column=1, end_row=171 + i, end_column=10)
    except Exception:
        pass

# IN07
ws = wb["IN07_NCC_CanhBao"]
clear_from(ws, 3)
put(ws, 3, 1, MST)
put(ws, 3, 2, TEN)
put(ws, 3, 3, "NNT đang kiểm tra")
put(ws, 3, 4, "Không")
put(ws, 3, 5, "")
put(ws, 3, 9, "Không có trên F1/F0/CA tại sổ chỉ định 09/11/2026")
put(ws, 3, 10, "Đang hoạt động")
rr = 4
for m, rec, info in f0_hits:
    write_row(
        ws,
        rr,
        [m, rec["ten"], "F0", "F0", info.get("nguon"), info.get("so_cv"), dparse(info.get("ngay_cv")), info.get("dv"), info.get("nd"), (dbm.get(m) or {}).get("tt"), None, None, "", rec["gt"], rec["thue"], rec["n_hd"], "IN02+IN01", ""],
        nums={14, 15},
        dates={7},
        fill=FILL_RED,
    )
    rr += 1
if not f0_hits:
    put(ws, 4, 9, "Không có NCC mua vào trùng F0/CA trên kỳ kiểm tra.")

# NCC 100% HĐ mua
ws = wb["KQ2_NCC_mua_vao"]
sub_all(ws)
clear_from(ws, 8)
r = 8
bk_keys = {(so_key(x.get("so")), mst_base(x.get("mst"))) for x in bk_u if x["chieu"] == "MUA"}
hd_keys2 = {(so_key(x.get("so")), mst_base(x.get("mst_dt"))) for x in hddt_ok if x["chieu"] == "MUA"}
inv = []
for rec in hddt_ok:
    if rec["chieu"] != "MUA":
        continue
    k = (so_key(rec.get("so")), mst_base(rec.get("mst_dt")))
    src = "Chung" if k in bk_keys else "HĐĐT"
    inv.append(
        {
            "mst": rec.get("mst_dt"),
            "ten": rec.get("ten_ban"),
            "so": rec.get("so"),
            "ngay": rec.get("ngay"),
            "hang": "",
            "gt": rec.get("gt"),
            "thue": rec.get("thue"),
            "src": src,
        }
    )
for rec in bk_u:
    if rec["chieu"] != "MUA":
        continue
    k = (so_key(rec.get("so")), mst_base(rec.get("mst")))
    if k in hd_keys2:
        continue
    inv.append(
        {
            "mst": rec.get("mst"),
            "ten": rec.get("ten"),
            "so": rec.get("so"),
            "ngay": rec.get("ngay"),
            "hang": rec.get("hang"),
            "gt": rec.get("gt"),
            "thue": rec.get("thue"),
            "src": "Bảng kê",
        }
    )
for rec in inv:
    nhom, nguon_f, cv, tt = ncc_flag(rec["mst"])
    fill = FILL_RED if nhom == "F0" else None
    write_row(
        ws,
        r,
        [
            r - 7,
            rec["mst"],
            rec["ten"],
            rec["so"],
            dparse(rec["ngay"]),
            rec["hang"],
            rec["gt"] or 0,
            rec["thue"] or 0,
            rec["src"],
            f"Có - {cv}" if cv else ("Có" if nhom else "Không"),
            nhom or "Không",
            tt,
        ],
        nums={1, 7, 8},
        dates={5},
        fill=fill,
    )
    r += 1
put(ws, r + 1, 1, "Nhận xét module (3.4)", bold=True, fill=FILL_NOTE)
put(
    ws,
    r + 2,
    1,
    f"HD04 Top-5 NCC / tổng mua = {top5_ty:.2%} (ngưỡng P14=50%). Số NCC={len(ncc)}. F0={len(f0_hits)} doanh số {f0_gt:,.0f} thuế {f0_thue:,.0f}. NNT trên F1: {'Có' if nnt_f1 else 'Không'}; trên CA: {'Có' if nnt_ca else 'Không'}. Nguồn CVCBRR mtime {nguon.get('cvcbrr_mtime')}.",
    wrap=True,
    fill=FILL_NOTE,
)

# TNCN
ws = wb["KQ2_TNCN_LD2"]
sub_all(ws)
rr = 8
for y in YEARS:
    q = nld_qtt.get(y) or {}
    ds = [x for x in nld_ds if x.get("Năm") == y]
    cccd = [x.get("CCCD / CMND") for x in ds if x.get("CCCD / CMND")]
    dup = len(cccd) - len(set(cccd)) if cccd else 0
    put(ws, rr, 1, y, num=True)
    put(ws, rr, 2, q.get("so_ld") or len(ds), num=True)
    put(ws, rr, 3, 0, num=True)
    put(ws, rr, 4, q.get("tnct"), num=True)
    put(ws, rr, 5, q.get("thue_kt") or 0, num=True)
    put(ws, rr, 6, 0, num=True)
    put(ws, rr, 7, dup, num=True)
    rr += 1
put(ws, rr + 1, 1, "Nhận xét: 4 LĐ cố định 2020-2025; toàn bộ TNCT dưới ngưỡng chịu thuế, thuế đã khấu trừ = 0. Không thấy 05-2 vãng lai. Trùng CCCD = 0 trên danh sách đã bóc.", wrap=True, fill=FILL_NOTE)

if "KQ2_EBITDA_Phan_tich" not in wb.sheetnames:
    ews = wb.create_sheet("KQ2_EBITDA_Phan_tich", 1)
else:
    ews = wb["KQ2_EBITDA_Phan_tich"]
ews.sheet_view.showGridLines = True
put(ews, 1, 1, "KQ2_EBITDA_Phan_tich - NĐ 132/2020 (30% EBITDA)", bold=True, fill=FILL_HDR, font=FONT_H)
headers = ["Chỉ tiêu"] + [str(y) for y in YEARS] + ["Ghi chú"]
for i, h in enumerate(headers, 1):
    put(ews, 3, i, h, bold=True, fill=FILL_HDR, font=FONT_H, center=True)
names_e = [
    "Lợi nhuận trước thuế [50]",
    "Chi phí lãi vay (BCTC [23])",
    "Khấu hao TSCD",
    "EBITDA",
    "Trần 30% EBITDA",
    "Vượt ngưỡng lãi vay",
    "Có GDLK?",
]
for i, name in enumerate(names_e):
    put(ews, 4 + i, 1, name)
    for j, y in enumerate(YEARS):
        if i == 0:
            put(ews, 4 + i, 2 + j, BM[y]["ln_tt"], num=True)
        elif i == 1:
            put(ews, 4 + i, 2 + j, BM[y]["cp_tc"], num=True)
        elif i == 6:
            put(ews, 4 + i, 2 + j, "Chưa đủ nguồn IN12")
put(ews, 12, 1, "Nhận xét: thiếu số khấu hao chi tiết và IN12 nên EBITDA chưa chốt. Chi phí lãi vay lấy BCTC [23] nếu có.", wrap=True, fill=FILL_NOTE)

ws = wb["KQ2_Tong_hop_co"]
sub_all(ws)
wb.save(lp(files["kq21"][1]))
wb.close()
print(" KQ2.1 saved")

# =====================================================================
print("FILL KQ2.2")
wb = load_workbook(lp(files["kq22"][1]))
ws = wb["KQ2_511_632"]
sub_all(ws)
# 4 year columns = 2022-2025
for j, y in enumerate(Y511):
    put(ws, 5, 3 + j, y, bold=True, center=True)
    put(ws, 6, 3 + j, y511.get(y), num=True)
    put(ws, 7, 3 + j, y632.get(y), num=True)
    a, b = y511.get(y) or 0, y632.get(y) or 0
    lai = a - b
    put(ws, 8, 3 + j, lai, num=True)
    put(ws, 9, 3 + j, (lai / a) if a else None, pct=True)
    n511 = sum(1 for x in lines_511 if x.get("nam") == y)
    n632 = sum(1 for x in lines_632 if x.get("nam") == y)
    put(ws, 13, 3 + j, n511, num=True)
    put(ws, 14, 3 + j, n632, num=True)
    if y in nkc_mong:
        put(ws, 6, 7, "2022 NKC mỏng (khoảng T12), không dùng làm năm đủ.")
    else:
        put(ws, 6, 7, "2023-2025: 511 PS Có khớp BCTC [01]; 632 là KC tháng nên không ghép từng HĐ.")
tot_a = sum(y511.get(y) or 0 for y in (2023, 2024, 2025))
tot_b = sum(y632.get(y) or 0 for y in (2023, 2024, 2025))
put(ws, 6, 2, tot_a, num=True)
put(ws, 7, 2, tot_b, num=True)
put(ws, 8, 2, tot_a - tot_b, num=True)
put(ws, 9, 2, (tot_a - tot_b) / tot_a if tot_a else None, pct=True)
put(ws, 16, 1, "Tổng cộng cột B chỉ cộng 2023-2025 (NKC đủ). 2021-2022 file .xls mỏng, 511 lệch BCTC nên không cộng. 632 là bút toán kết chuyển, không kết luận bán dưới giá theo từng HĐ.", wrap=True, fill=FILL_NOTE)

if "KQ2_511_632_Chi_tiet" not in wb.sheetnames:
    cws = wb.create_sheet("KQ2_511_632_Chi_tiet")
else:
    cws = wb["KQ2_511_632_Chi_tiet"]
heads = ["Nam", "Ngay_CT", "So_CT", "TK", "Dien_giai", "DT_511_PS_Co", "GV_632", "Ghi_chu"]
for i, h in enumerate(heads, 1):
    put(cws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_H, center=True)
rr = 2
for x in lines_511:
    put(cws, rr, 1, x.get("nam"), num=True)
    put(cws, rr, 2, dparse(x.get("ngay")), dt=True, center=True)
    put(cws, rr, 3, x.get("so"))
    put(cws, rr, 4, x.get("tk"))
    put(cws, rr, 5, x.get("dg"), wrap=True)
    put(cws, rr, 6, x.get("co") or 0, num=True)
    put(cws, rr, 8, "NKC 2021-2022 mỏng" if x.get("nam") in nkc_mong else "")
    rr += 1
    if rr > 2500:
        break
apply_tnr(cws)

ws = wb["KQ2_Kho_Gia_thanh"]
sub_all(ws)
put(ws, 5, 1, "PHẦN 1. NHẬP - XUẤT - TỒN KHO HÀNG HÓA (TK 156) - DN thương mại, không dùng TK 152", bold=True, fill=FILL_SEC)
rr = 7
for y in YEARS:
    k = kho.get(y)
    put(ws, rr, 1, y, num=True)
    if not k:
        put(ws, rr, 12, "Chưa đọc được XNT năm này")
        rr += 1
        continue
    put(ws, rr, 2, k["dk_sl"], num=True)
    put(ws, rr, 3, k["dk_tt"], num=True)
    put(ws, rr, 4, k["nhap_sl"], num=True)
    put(ws, rr, 5, k["nhap_tt"], num=True)
    put(ws, rr, 6, k["xuat_sl"], num=True)
    put(ws, rr, 7, k["xuat_tt"], num=True)
    put(ws, rr, 8, k["ck_sl"], num=True)
    put(ws, rr, 9, k["ck_tt"], num=True)
    htk = BM[y]["htk"]
    note = f"BCTC [141]={htk}; xuất NKC 156={y156_xuat.get(y)}; 632={y632.get(y)}"
    put(ws, rr, 12, note, wrap=True)
    rr += 1
put(ws, 20, 1, "Nhận xét TC3: BCTC B01-DN TT200 dùng [141]=[140]. DN thương mại - XNT TK 156. Phần 2 TK 155 để trống (không SX). Đối chiếu tồn CK XNT vs BCTC [141] trên từng năm.", wrap=True, fill=FILL_NOTE)

ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
sub_all(ws)
put(ws, 5, 1, "PHẦN 1. CƠ CẤU GIÁ VỐN THƯƠNG MẠI (không có 621/622/627) - KHÔNG_ÁP_DỤNG giá thành SX", bold=True, fill=FILL_SEC)
for j, y in enumerate(YEARS):
    put(ws, 7 + j, 1, y, num=True)
    put(ws, 7 + j, 8, y632.get(y), num=True)
    dt, gv = BM[y]["dt"], BM[y]["gv"]
    tl = (dt - gv) / dt if dt else None
    put(ws, 7 + j, 11, f"BCTC [01]={dt} [11]={gv} lãi gộp {tl:.2%}" if dt and tl is not None else "Thiếu BCTC")
put(ws, 20, 1, "Nhận xét TC28: DN thương mại thiết bị vệ sinh. Lãi gộp BCTC 2020-2025 khoảng 2,5%-6,6%; LNTT âm 2020-2024, 2025 [50] lãi nhỏ. Không tách 621/622.", wrap=True, fill=FILL_NOTE)

ws = wb["KQ2_Cong_no_Tien_mat"]
sub_all(ws)
cn_map = {(x.get("nam"), x.get("tk")): x for x in congno}
rr = 8
for y in YEARS:
    c131 = cn_map.get((y, "131")) or {}
    c331 = cn_map.get((y, "331")) or {}
    put(ws, rr, 1, y, num=True)
    put(ws, rr, 2, c131.get("dk_no"), num=True)
    put(ws, rr, 3, c131.get("ps_no"), num=True)
    put(ws, rr, 4, c131.get("ps_co"), num=True)
    put(ws, rr, 5, c131.get("ck_no"), num=True)
    put(ws, rr, 6, c331.get("dk_co") or c331.get("dk_no"), num=True)
    put(ws, rr, 7, c331.get("ps_co"), num=True)
    put(ws, rr, 8, c331.get("ps_no"), num=True)
    put(ws, rr, 9, c331.get("ck_co") or c331.get("ck_no"), num=True)
    dt = BM[y]["dt"] or 0
    ck331 = c331.get("ck_co") or c331.get("ck_no") or 0
    put(ws, rr, 11, dt, num=True)
    put(ws, rr, 12, (ck331 / dt) if dt else None, pct=True)
    if y in (2021, 2022) and (c131.get("ps_co") in (0, None)):
        put(ws, rr, 10, "Sổ .xls 2021-2022 ps_co=0 - không tin tổng")
    rr += 1
put(ws, 40, 1, f"TT01: {len(tt01)} cụm HĐ mua cùng ngày >= ngưỡng tiền mặt (20tr đến 30/06/2025, 5tr từ 01/07/2025). TT02/GT16 chờ sao kê IN11.", wrap=True, fill=FILL_NOTE)
put(ws, 41, 1, f"IN10 chi TK 111: {len(cash_ok)} dòng. IN11: Chưa nhận - không kết luận thiếu UNC.", wrap=True, fill=FILL_NOTE)

ws = wb["KQ2_Tong_hop_co"]
sub_all(ws)
wb.save(lp(files["kq22"][1]))
wb.close()
print(" KQ2.2 saved")

# =====================================================================
print("FILL KQ3")
wb = load_workbook(lp(files["kq3"][1]))
ws = wb["KQ3"]
put(ws, 2, 3, MST, bold=True)
put(ws, 2, 10, TEN, bold=True)
put(ws, 3, 3, KY)
put(ws, 3, 10, LEVEL)
put(ws, 4, 3, VER)
put(ws, 4, 10, MAU_VER)
put(ws, 5, 3, TODAY, dt=True)
put(ws, 5, 10, f"KQ1.3 {VER}; KQ2.1 {VER}; KQ2.2 {VER}")

wb13 = load_workbook(lp(files["kq13"][1]), data_only=True)
wsk = wb13["KQ1_Kha_nang_phan_tich"]
cat = []
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

findings = []
rows_kq3 = []
for rd in cat:
    ma = rd["ma"]
    need = rd["need"]
    missing = [x for x in re.findall(r"IN\d{2}", need) if x not in {"IN07"} and not has.get(x, False)]
    status = "ĐÃ_CHẠY"
    ket = "Không"
    nhan = "Chắc"
    muc_rr = "Thấp"
    diem = 0
    va = vb = do = None
    ca_a = cb_b = ""
    v_tien = w_tien = None
    sac = ""
    thieu_txt = ""
    nguon_txt = ""
    if ma in {"KH03", "KH05", "LD05"}:
        status = "KHÔNG_ÁP_DỤNG"
        ket = "Không"
        nhan = ""
        muc_rr = ""
        thieu_txt = "DN thương mại - không SX"
    elif ma == "HD07":
        status = "TREO"
        ket = "Chưa kết luận"
        nhan = ""
        muc_rr = ""
        thieu_txt = "thiếu số liệu ngành (K13)"
    elif ma == "TN20":
        status = "CHƯA_TRIỂN_KHAI"
        ket = "Chưa kết luận"
        nhan = ""
        muc_rr = ""
    elif missing:
        status = "KHÔNG_ĐỦ_DỮ_LIỆU"
        ket = "Chưa kết luận"
        nhan = ""
        muc_rr = ""
        thieu_txt = ",".join(sorted(set(missing)))

    if ma == "GT01" and status == "ĐÃ_CHẠY":
        diffs = []
        for y in YEARS:
            a = GY[y]["34"] if GY[y]["n"] else None
            b = sum_hd("BAN", y)[0]
            if a is None:
                continue
            diffs.append((y, a, b, a - b))
        if diffs:
            y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
            ca_a, va, cb_b, vb, do = f"TK[34] {y}", a, "HĐĐT bán", b, d
            nguon_txt = "KQ2.1 KQ2_Doi_chieu_GTGT"
            if abs(d) > 1:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
                v_tien = abs(d)
                w_tien = abs(d)
                findings.append(("GT01", y, d, a, b))
    if ma == "GT02" and status == "ĐÃ_CHẠY":
        diffs = []
        for y in YEARS:
            a = GY[y]["34"] if GY[y]["n"] else None
            b = sum_bk("BAN", y)[0]
            if a is None:
                continue
            diffs.append((y, a, b, a - b))
        if diffs:
            y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
            ca_a, va, cb_b, vb, do = f"TK[34] {y}", a, "BK bán", b, d
            nguon_txt = "KQ2.1"
            if abs(d) > 1:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Trung bình", 6, "GTGT"
                v_tien = abs(d)
                w_tien = 0
                findings.append(("GT02", y, d, a, b))
    if ma == "GT04" and status == "ĐÃ_CHẠY":
        diffs = []
        for y in YEARS:
            a = GY[y]["23"] if GY[y]["n"] else None
            b = sum_hd("MUA", y)[0]
            if a is None:
                continue
            diffs.append((y, a, b, a - b))
        if diffs:
            y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
            ca_a, va, cb_b, vb, do = f"TK[23] {y}", a, "HĐĐT mua", b, d
            if abs(d) > 1:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 8, "GTGT"
                v_tien = abs(d)
                w_tien = abs(d)
                findings.append(("GT04", y, d, a, b))
    if ma == "HD01" and status == "ĐÃ_CHẠY":
        ca_a, va, cb_b, vb, do = "NCC F0", len(f0_hits), "Doanh số F0", f0_gt, f0_gt
        nguon_txt = "KQ2.1 IN07/NCC"
        if f0_hits:
            ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
            v_tien = f0_thue
            w_tien = f0_thue
        if nnt_f1:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Cao", 9
    if ma == "HD04" and status == "ĐÃ_CHẠY":
        ca_a, va, cb_b, vb, do = "Top5/tổng mua", top5_ty, "Ngưỡng P14", 0.5, top5_ty - 0.5
        if top5_ty > 0.5:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Trung bình", 6
    if ma == "HD08" and status == "ĐÃ_CHẠY":
        ca_a, va = "NNT trên F1", 1 if nnt_f1 else 0
        if nnt_f1:
            ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
    if ma == "TT01" and status == "ĐÃ_CHẠY":
        ca_a, va = "HĐ mua >= ngưỡng TM", len(tt01)
        do = len(tt01)
        nhan = "Sơ bộ"
        if tt01:
            ket, muc_rr, diem = "Có phát hiện", "Trung bình", 5
    if ma == "TN01" and status == "ĐÃ_CHẠY":
        diffs = []
        for y in YEARS:
            a = GY[y]["34"] if GY[y]["n"] else None
            b = BM[y]["dt"]
            if a is None or b is None:
                continue
            diffs.append((y, a, b, a - b))
        if diffs:
            y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
            ca_a, va, cb_b, vb, do = f"GTGT[34] {y}", a, "BCTC[01]", b, d
            if abs(d) > (b or 0) * P03:
                ket, muc_rr, diem, sac = "Có phát hiện", "Cao", 7, "TNDN"
                v_tien = abs(d)
                w_tien = 0
                nhan = "Chắc"
    if ma == "TN03" and status == "ĐÃ_CHẠY":
        nhan = "Sơ bộ - chưa so ngành"
        for y in YEARS:
            dt = BM[y]["dt"]
            gv = BM[y]["gv"]
            if dt and gv is not None and dt - gv < 0:
                ket, muc_rr, diem = "Có phát hiện", "Cao", 7
                ca_a, va, cb_b, vb, do = f"DT {y}", dt, "GV", gv, dt - gv
                break
    if ma == "TN17" and status == "ĐÃ_CHẠY":
        hit = False
        for y in YEARS:
            dt = BM[y]["dt"] or 0
            c331 = cn_map.get((y, "331")) or {}
            ck = c331.get("ck_co") or c331.get("ck_no") or 0
            if dt and ck / dt > P23:
                ket, nhan, muc_rr, diem = "Có phát hiện", "Sơ bộ", "Trung bình", 5
                ca_a, va, cb_b, vb, do = f"Dư Có 331 {y}", ck, "DT", dt, ck / dt
                hit = True
                break
        if not hit:
            nhan = "Sơ bộ"
    if ma in {"DT01", "DT02", "DT03"} and not missing:
        status = "ĐÃ_CHẠY"
        ket = "Chưa kết luận"
        nhan = "Sơ bộ"
        thieu_txt = "632 là KC tháng, không ghép từng HĐ; NKC 2021-2022 mỏng"
        ca_a, va = "Dòng 511", len(lines_511)
    if ma == "KH04" and has["IN09"] and has["IN04"]:
        status = "ĐÃ_CHẠY"
        ket = "Chưa kết luận"
        nhan = "Sơ bộ"
        thieu_txt = "Đã cộng XNT 156; cần đối chiếu từng năm với BCTC [141] tại trụ sở"
    if ma in {"LS01", "LS02"}:
        status = "ĐÃ_CHẠY" if has["IN14"] else "KHÔNG_ĐỦ_DỮ_LIỆU"
        ket = "Chưa kết luận"
        thieu_txt = "PTDN01/02 là BCTC TMS, chưa có bảng nợ thuế"
        nhan = ""
        muc_rr = ""
    if ma == "DQ02" and status == "ĐÃ_CHẠY":
        ket = "Không"
        nhan = "Chắc"
        ca_a, va = "Dòng Tổng cộng đã lọc", 0

    hs = {"Cao": 3, "Trung bình": 2, "Thấp": 1}.get(muc_rr, 0)
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
            "Ngưỡng": "P03 tạm 2%" if ma in {"GT01", "TN01"} else rd.get("need"),
            "Trạng thái chạy": status,
            "Kết quả": ket,
            "Nhãn": nhan,
            "Mức rủi ro": muc_rr,
            "Hệ số trọng yếu": hs,
            "Điểm": diem,
            "Sắc thuế": sac,
            "Số tiền liên quan (đồng)": v_tien,
            "Số tiền tính vào tổng (đồng)": w_tien if ket == "Có phát hiện" else (0 if status == "ĐÃ_CHẠY" else None),
            "Nguồn": nguon_txt,
            "Còn thiếu để kết luận": thieu_txt,
            "Nội dung kiểm tra tại trụ sở": "Xem chứng từ gốc" if ket == "Có phát hiện" else "",
            "Chứng từ cần xem": "",
            "Ghi chú": "",
        }
    )

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
            put(ws, rr, c, v, wrap=True)
    if rec.get("Mức rủi ro") == "Cao":
        ws.cell(rr, hm["Mức rủi ro"]).fill = FILL_RED
        ws.cell(rr, hm["Mức rủi ro"]).font = RED_F
    elif rec.get("Mức rủi ro") == "Trung bình":
        ws.cell(rr, hm["Mức rủi ro"]).fill = FILL_YEL
        ws.cell(rr, hm["Mức rủi ro"]).font = ORG_F
    if rec.get("Trạng thái chạy") == "KHÔNG_ĐỦ_DỮ_LIỆU":
        ws.cell(rr, hm["Trạng thái chạy"]).fill = FILL_YEL

w_gtgt = sum((r.get("Số tiền tính vào tổng (đồng)") or 0) for r in rows_kq3 if r.get("Sắc thuế") == "GTGT")
put(ws, 7, 3, w_gtgt, num=True)
n_cao = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Cao" and r.get("Kết quả") == "Có phát hiện")
n_tb = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Trung bình" and r.get("Kết quả") == "Có phát hiện")
put(ws, 6, 10, f"{n_cao} Cao / {n_tb} TB / 0 Thấp", bold=True)

ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
if str(ws.cell(4, 1).value or "").startswith("TC"):
    ws.delete_rows(4, 1)
tc = [
    ("TC01", "Chênh lệch DT GTGT [34] vs DT BCTC [01]", "Doanh thu", "lệch ≠ 0"),
    ("TC02", "Người mua trả tiền trước > 5% DT", "Công nợ", "5%"),
    ("TC03", "HTK / DT", "Hàng tồn kho", "30% - B01-DN dùng [141]"),
    ("TC04", "DT chưa thực hiện > 0", "Doanh thu", "> 0"),
    ("TC05", "DT thuần - giá vốn < 0", "Giá vốn", "< 0"),
    ("TC06", "Chi phí phải trả trích trước lớn", "Chi phí", "[315]"),
    ("TC07", "Lãi vay / CP trả trước", "Chi phí", "[242]"),
    ("TC08", "Ưu đãi / miễn giảm TNDN", "TNDN", "C2/C11"),
    ("TC09", "VAT khấu trừ vs BCTC [152]", "GTGT", "[43] vs [152]"),
    ("TC10", "VAT khấu trừ vs 10%×(HTK+TSCĐ)", "GTGT", "×10%"),
    ("TC11", "Có DT KCT [26] nhưng [25]=[24]", "GTGT", ""),
    ("TC12", "Điều chỉnh giảm LN trước thuế", "TNDN", "B7/B8"),
    ("TC13", "Trích lập dự phòng", "BCTC", ""),
    ("TC14", "CN vốn [22]>0 và [25]<0", "BCTC", ""),
    ("TC15", "Bù trừ lãi BĐS / lỗ SXKD", "TNDN", ""),
    ("TC16", "Chuyển giá [251][252]>0 và [421]<0", "GDLK", ""),
    ("TC17", "CP lãi vay vs 30% EBITDA", "EBITDA", "NĐ 132"),
    ("TC18", "B11>0 không có PL GDLK", "GDLK", "IN12"),
    ("TC19", "Thuế TNDN hoãn lại", "TNDN", ""),
    ("TC20", "Phải trả người bán lớn", "Công nợ", "[311]+[331]"),
    ("TC21", "Lãi vay khi thiếu VĐL", "Tài chính", ""),
    ("TC22", "QTT TNCN vs TK tháng/quý", "TNCN", ""),
    ("TC23", "Dòng tiền chi NLĐ vs TNCT", "TNCN", "20%"),
    ("TC24", "KHBS vs hạch toán TMS", "TMS", "không chạy"),
    ("TC25", "Bổ sung GTGT chưa điều chỉnh", "GTGT", ""),
    ("TC26", "Tỷ trọng HĐ mua NNT TT 03/06", "Hóa đơn", "cần ĐKKD"),
    ("TC27", "DT TK GTGT vs HĐĐT bán ra", "Hóa đơn", "lệch ≠ 0"),
    ("TC28", "Phân tích ngang/dọc BCTC - GV vs DT", "Giá vốn", ""),
]
for i, (id_, ten, phan, nguong) in enumerate(tc):
    kq = "Chưa đủ dữ liệu để chốt"
    muc_tc = "Thấp"
    kn = "Bổ sung chứng từ nếu thiếu nguồn"
    if id_ == "TC01":
        kq = "2020 GTGT XML chỉ 12/2020 (0,84 tỷ) vs BCTC 6,81 tỷ. 2025 TK[34] 8,55 tỷ vs BCTC 8,09 tỷ. 2021-2024 gần khớp."
        muc_tc = "Cao"
        kn = "Xin XML GTGT T01-T11/2020; đối chiếu cut-off 2025."
    elif id_ == "TC03":
        bits = []
        for y in YEARS:
            dt = BM[y]["dt"] or 0
            htk = BM[y]["htk"] or 0
            if dt:
                bits.append(f"{y}:{htk/dt:.0%}")
        kq = "HTK/DT " + ", ".join(bits) + ". B01-DN dùng [141]."
        muc_tc = "Trung bình"
        kn = "Đối chiếu XNT 156 vs BCTC [141] từng năm."
    elif id_ == "TC05":
        kq = "Lãi gộp BCTC dương mọi năm (khoảng 2,5%-6,6%). LNTT âm 2020-2024."
        muc_tc = "Trung bình"
        kn = "Rà chi phí bán hàng/QLDN và lãi vay."
    elif id_ == "TC17":
        kq = "EBITDA: thiếu khấu hao chi tiết và IN12, chưa chốt 30%."
        muc_tc = "Thấp"
        kn = "Lấy số 635 và khấu hao từ sổ cái."
    elif id_ == "TC18":
        kq = "IN12 chưa nhận - KHÔNG_ĐỦ_DỮ_LIỆU."
        muc_tc = "Thấp"
        kn = "Xin PL GDLK nếu có bên liên kết."
    elif id_ == "TC22":
        kq = "QTT TNCN 2020-2025: 4 LĐ, thuế KT=0. File đối chiếu quý vs năm đã có ở 3.3."
        muc_tc = "Thấp"
        kn = "Đối chiếu bảng lương 2.4 vs QTT."
    elif id_ == "TC24":
        kq = "Không chạy (cấm TMS dòng)."
        muc_tc = "Thấp"
        kn = ""
    elif id_ == "TC26":
        kq = f"Đã tra danh bạ {len(dbm)} MST. 0 NCC trùng F0/CA."
        muc_tc = "Thấp"
        kn = "Chọn mẫu NCC ngừng/đóng MST trên danh bạ."
    elif id_ == "TC27":
        kq = "Lệch DT tờ khai vs HĐĐT bán ra 2020-2022 (portal HĐĐT bán từ 18/05/2022). 2024 khớp."
        muc_tc = "Cao"
        kn = "Chọn mẫu HĐ điều chỉnh/hủy/sai kỳ năm 2022 và 2025."
    elif id_ == "TC28":
        kq = "Cơ cấu GV thương mại; 632 2023-2025 khớp BCTC [11]. 2021-2022 NKC mỏng."
        muc_tc = "Trung bình"
        kn = "Xin NKC đủ 2021-2022."
    put(ws, 4 + i, 1, id_, bold=True)
    put(ws, 4 + i, 2, ten)
    put(ws, 4 + i, 3, phan)
    put(ws, 4 + i, 4, nguong)
    put(ws, 4 + i, 5, kq, wrap=True)
    put(ws, 4 + i, 6, muc_tc, bold=True)
    put(ws, 4 + i, 7, kn, wrap=True)
    if muc_tc == "Cao":
        ws.cell(4 + i, 6).fill = FILL_RED
        ws.cell(4 + i, 6).font = RED_F
    elif muc_tc == "Trung bình":
        ws.cell(4 + i, 6).fill = FILL_YEL
        ws.cell(4 + i, 6).font = ORG_F
apply_tnr(ws)
autowidth(ws)
apply_tnr(wb["KQ3"])
wb.save(lp(files["kq3"][1]))
wb.close()
print(" KQ3 saved")

# extra GĐ3
print("FILL extra GĐ3")
dau_path = OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}.xlsx"
giai_path = OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}.xlsx"
wb = Workbook()
ws = wb.active
ws.title = "Dau_hieu"
heads = ["STT", "Ma", "Dau hieu", "Muc", "So lieu", "Sheet KQ", "Huong kiem tra", "Tai lieu can", "Trang thai"]
for i, h in enumerate(heads, 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_H, center=True)
dau = [
    (1, "GT01", "Lệch DT tờ khai GTGT vs HĐĐT bán ra (2020-2022, 2025)", "Cao", "Xem GT01 theo năm", "KQ2.1/KQ3", "1. HĐ điều chỉnh 2. Sai kỳ 3. HĐ hủy 4. XML 2020 thiếu", "HĐĐT portal, BK bán, sổ 511", "Chưa kiểm tra tại trụ sở"),
    (2, "GT04", "Lệch GT mua vào tờ khai vs HĐĐT", "Cao", "Xem GT04", "KQ2.1/KQ3", "1. HĐ mua không kê 2. Kê không có HĐ", "HĐĐT mua, BK mua, sổ 133", "Chưa kiểm tra tại trụ sở"),
    (3, "TN01", "GTGT[34] vs BCTC[01] 2020 (thiếu XML) và 2025", "Cao", "2020 và 2025", "KQ2.1", "Xin XML T01-T11/2020; cut-off 2025", "XML GTGT, BCTC", "Thiếu 11 kỳ 2020"),
    (4, "HD04", f"Top-5 NCC {top5_ty:.2%}", "Trung bình" if top5_ty > 0.5 else "Thấp", f"{len(ncc)} NCC", "KQ2.1", "Chọn mẫu hợp đồng Top-5", "Hợp đồng, BK", "Chưa kiểm tra"),
    (5, "TT01", f"{len(tt01)} cụm HĐ mua vượt ngưỡng tiền mặt", "Trung bình", "P10/P11", "KQ2.2", "Đối UNC; không kết luận thiếu chứng từ khi chưa có sao kê", "Sao kê, UNC", "Thiếu IN11"),
    (6, "IN11", "Thiếu sao kê ngân hàng", "Trung bình", "2.3 trống", "KQ1.3", "Đề nghị NNT bổ sung", "Excel sao kê", "Chưa nhận"),
    (7, "IN08", "NKC 2021-2022 mỏng so với BCTC", "Trung bình", "511 2021=1,03 tỷ vs BCTC 6,76 tỷ", "KQ2.2", "Xin sổ đủ năm", "NKC/sổ cái 2021-2022", "Đã ghi nhận"),
]
for i, row in enumerate(dau):
    for c, v in enumerate(row, 1):
        put(ws, 2 + i, c, v, wrap=True)
    if row[3] == "Cao":
        ws.cell(2 + i, 4).fill = FILL_RED
        ws.cell(2 + i, 4).font = RED_F
    elif row[3] == "Trung bình":
        ws.cell(2 + i, 4).fill = FILL_YEL
        ws.cell(2 + i, 4).font = ORG_F
apply_tnr(ws)
autowidth(ws)
wb.save(lp(dau_path))
wb.close()

wb = Workbook()
sheets = [
    "00_Tong_hop_chi_tiet",
    "1_NCC_F0_CanhBao",
    "2_NCC_DKKD_bat_thuong",
    "3_Ty_le_gia_von",
    "4_HD_ban_duoi_gia",
    "5_Khong_vay_no_EBITDA",
    "6_Top5_NCC",
    "7_Lech_HDDT_BangKe",
    "8_Cong_no_TK131_TK331",
    "9_TNCN_NLD_nhieu_nguon",
]
wb.active.title = sheets[0]
for s in sheets[1:]:
    wb.create_sheet(s)
ws = wb["00_Tong_hop_chi_tiet"]
put(ws, 1, 1, f"{TEN} - {MST} - {KY} - Tổng hợp chi tiết dấu hiệu", bold=True, fill=FILL_HDR, font=FONT_H)
ws.merge_cells("A1:D1")
for i, h in enumerate(["Phần", "Chỉ tiêu", "Giá trị", "Nguồn"], 1):
    put(ws, 3, i, h, bold=True, fill=FILL_HDR, font=FONT_H, center=True)
summary_rows = [
    ("I", "Số HĐĐT mua (tính vào tổng)", sum(1 for x in hddt_ok if x["chieu"] == "MUA"), "IN01"),
    ("I", "Số HĐĐT bán (tính vào tổng)", sum(1 for x in hddt_ok if x["chieu"] == "BAN"), "IN01"),
    ("I", "Số dòng bảng kê mua", sum(1 for x in bk_u if x["chieu"] == "MUA"), "IN02"),
    ("I", "Số dòng bảng kê bán", sum(1 for x in bk_u if x["chieu"] == "BAN"), "IN02"),
    ("I", "Số NCC mua", len(ncc), "IN02+IN01"),
    ("II", "NCC F0", len(f0_hits), "IN07"),
    ("II", "Doanh số F0", f0_gt, "IN07"),
    ("II", "Top5 %", top5_ty, "HD04"),
    ("II", "NNT trên F1", int(nnt_f1), "IN07"),
    ("III", "TT01 cụm HĐ >= ngưỡng TM", len(tt01), "IN01"),
    ("III", "IN11 sao kê", "Chưa nhận", "2.3"),
]
for i, row in enumerate(summary_rows):
    for c, v in enumerate(row, 1):
        put(ws, 4 + i, c, v, num=isinstance(v, (int, float)))

ws = wb["1_NCC_F0_CanhBao"]
put(ws, 1, 1, "Không có NCC mua vào trùng F0/CA trên kỳ 2020-2025.", wrap=True)
ws = wb["6_Top5_NCC"]
put(ws, 1, 1, "MST", bold=True, fill=FILL_HDR, font=FONT_H)
put(ws, 1, 2, "Tên", bold=True, fill=FILL_HDR, font=FONT_H)
put(ws, 1, 3, "GT mua", bold=True, fill=FILL_HDR, font=FONT_H)
put(ws, 1, 4, "Tỷ trọng", bold=True, fill=FILL_HDR, font=FONT_H)
for i, (m, rec) in enumerate(top5):
    put(ws, 2 + i, 1, m)
    put(ws, 2 + i, 2, rec["ten"])
    put(ws, 2 + i, 3, rec["gt"], num=True)
    put(ws, 2 + i, 4, rec["gt"] / tot_mua if tot_mua else None, pct=True)
ws = wb["7_Lech_HDDT_BangKe"]
put(ws, 1, 1, "Năm", bold=True, fill=FILL_HDR, font=FONT_H)
put(ws, 1, 2, "TK[34]", bold=True, fill=FILL_HDR, font=FONT_H)
put(ws, 1, 3, "HĐĐT bán", bold=True, fill=FILL_HDR, font=FONT_H)
put(ws, 1, 4, "BK bán", bold=True, fill=FILL_HDR, font=FONT_H)
put(ws, 1, 5, "BCTC[01]", bold=True, fill=FILL_HDR, font=FONT_H)
for i, y in enumerate(YEARS):
    put(ws, 2 + i, 1, y, num=True)
    put(ws, 2 + i, 2, GY[y]["34"] if GY[y]["n"] else None, num=True)
    put(ws, 2 + i, 3, sum_hd("BAN", y)[0], num=True)
    put(ws, 2 + i, 4, sum_bk("BAN", y)[0], num=True)
    put(ws, 2 + i, 5, BM[y]["dt"], num=True)
ws = wb["9_TNCN_NLD_nhieu_nguon"]
put(ws, 1, 1, "4 LĐ cố định, thuế khấu trừ = 0 mọi năm. Không thấy thu nhập nhiều nguồn trên XML QTT.")
for s in wb.sheetnames:
    apply_tnr(wb[s])
    autowidth(wb[s])
wb.save(lp(giai_path))
wb.close()

# copy to sao luu + vault
copied = []
for _, dst in files.values():
    shutil.copy2(lp(dst), lp(SAO / dst.name))
    shutil.copy2(lp(dst), lp(VAULT_KQ / dst.name))
    copied.append(dst.name)
for p in (dau_path, giai_path):
    shutil.copy2(lp(p), lp(SAO / p.name))
    shutil.copy2(lp(p), lp(VAULT_KQ / p.name))
    copied.append(p.name)
print("COPIED", copied)
print("DONE")
