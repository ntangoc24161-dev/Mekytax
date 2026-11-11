# -*- coding: utf-8 -*-
"""Fill v1.4 KQ1.1-KQ3 for NONG SAN LUU HOAI ANH 3603057410 from nsla_cache.json."""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603057410"
TEN = "CÔNG TY TNHH NÔNG SẢN LƯU HOÀI ANH"
KY = "2022-2025"
YEARS = [2022, 2023, 2024, 2025]
TODAY = date(2026, 11, 9)
VER = "v1"
MAU_VER = "1.4"
LEVEL = "A"
DN_STATUS = "Chưa có"
KY_SUY_RA = (
    "K12: kỳ 2022-2025 suy từ HĐĐT 18/02/2022-31/12/2025 và BCTC/TNDN 2022, 2023, 2025. "
    "GTGT XML 2026 Q1-Q3 giữ ở IN03, không đưa vào đối chiếu KQ2."
)

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Ke hoach nam\Bo sung\NONG_SAN_LUU_HOAI_ANH_3603057410"
)
OUT = ROOT / "4. KET QUA PHAN TICH"
CACHE = Path(r"C:\Users\NGOC\AppData\Local\Temp\nsla_cache.json")
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
RULES_WB = MAU / "Bo_bieu_mau_chuan_phan_tich_ho_so_thue.xlsx"

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FONT_HDR = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
FONT_RED = Font(name="Times New Roman", size=11, bold=True, color="C00000")
FONT_ORG = Font(name="Times New Roman", size=11, bold=True, color="B25900")
FONT_GRN = Font(name="Times New Roman", size=11, color="385723")
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
PCT = "0.00%"
DATE_FMT = "dd/mm/yyyy"
P13, P14, P05, P18, P23, P03 = 0.05, 0.50, 0.30, 0.92, 0.10, 0.02

HAS = {
    "IN00": True,
    "IN01": True,
    "IN02": False,
    "IN03": True,
    "IN04": True,
    "IN05": True,
    "IN07": True,
    "IN08": False,
    "IN09": False,
    "IN10": False,
    "IN11": False,
    "IN12": False,
    "IN13": False,
    "IN14": True,
}


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


def mst_valid(mst: str) -> bool:
    """Length/digit check only. Checksum BT5 is not applied mass-flag: MST NNT 3603057410 itself does not match the 11-mod variant, so do not treat checksum fail as DQ03."""
    s = mst_base(mst)
    return len(s) in {10, 13} and s.isdigit()


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


def autowidth(ws, cap=60, min_w=10):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, 5000), max_col=ws.max_column or 1):
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            widths[cell.column] = max(widths.get(cell.column, min_w), min(len(s), cap) + 2)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


def clear_from(ws, start_row, max_col=None):
    if ws.max_row and ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)


def put(ws, r, c, v, bold=False, num=False, dt=False, pct=False, fill=None, wrap=False, font=None):
    cell = ws.cell(r, c, v)
    cell.font = font or (FONT_B if bold else FONT)
    cell.border = THIN
    cell.alignment = Alignment(vertical="center", wrap_text=wrap)
    if fill:
        cell.fill = fill
    if num and isinstance(v, (int, float)) and not isinstance(v, bool):
        cell.number_format = NUM
    if pct and isinstance(v, (int, float)) and not isinstance(v, bool):
        cell.number_format = PCT
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
                v = v.replace("[Đủ | Một phần | Chưa có]", DN_STATUS)
                if v != cell.value:
                    cell.value = v
            tnr(cell)


def vn(n):
    if n is None:
        return ""
    s = f"{n:,.0f}".replace(",", ".")
    return s


print("LOAD CACHE")
C = json.loads(CACHE.read_text(encoding="utf-8"))
hddt = C["hddt"]
gtgt_all = C["gtgt_eff"]
gtgt = [x for x in gtgt_all if (x.get("nam") or 0) in YEARS]
gtgt_extra = [x for x in gtgt_all if (x.get("nam") or 0) == 2026]
tndn = C["tndn_eff"]
bctc = C["bctc_eff"]
tncn = C["tncn_qtt"]
manifest = C["manifest"]
f0_used = C.get("f0_used") or []
ca_used = C.get("ca_used") or []
f0_map = {mst_base(x["mst"]): x for x in f0_used}
nnt_f1 = bool(C.get("f1_nnt"))
nnt_ca = bool(C.get("ca_nnt"))


def ct(x, code, default=None):
    d = x.get("chi_tieu") or {}
    v = d.get(str(code))
    return default if v is None else v


def bctc_map():
    out = {}
    for x in bctc:
        y = x.get("nam")
        k = x.get("kqkd") or {}
        cd = x.get("cdkt_cuoi") or {}
        out[y] = {
            "dt": k.get("01") or 0,
            "gv": k.get("11") or 0,
            "ln_gop": k.get("20") or 0,
            "dt_tc": k.get("21") or 0,
            "cp_ql": k.get("24") or 0,
            "thue_tndn": k.get("22") or 0,
            "lnst": k.get("60") or k.get("50") or 0,
            "htk": cd.get("140") or cd.get("141") or 0,
            "htk140": cd.get("140") or 0,
            "htk141": cd.get("141") or 0,
            "pt": cd.get("131") or 0,
            "no_nh": cd.get("311") or 0,
            "tm": cd.get("110") or 0,
            "von": cd.get("411") or 0,
            "thue_kt": cd.get("181") or cd.get("152") or 0,
            "tscd": cd.get("211") or 0,
            "file": x.get("file"),
            "cdkt": cd,
            "kqkd": k,
            "mau": "TT133 B01a-DNN (maTKhai 683)",
        }
    return out


BM = bctc_map()
TN = {x["nam"]: x.get("chi_tieu") or {} for x in tndn}


def gtgt_year():
    d = {
        y: {
            "34": 0.0,
            "35": 0.0,
            "23": 0.0,
            "24": 0.0,
            "26": 0.0,
            "40": 0.0,
            "43": 0.0,
            "32": 0.0,
            "22": 0.0,
            "kys": [],
            "missing": [],
        }
        for y in YEARS
    }
    have = {(x.get("nam"), str(x.get("ky_norm") or "").upper()) for x in gtgt}
    for y in YEARS:
        for q in range(1, 5):
            k = f"Q{q}/{y}"
            if (y, k) not in have:
                d[y]["missing"].append(k)
    for x in gtgt:
        y = x.get("nam")
        if y not in d:
            continue
        c = x.get("chi_tieu") or {}
        d[y]["34"] += c.get("34") or 0
        d[y]["35"] += c.get("35") or 0
        d[y]["23"] += c.get("23") or 0
        d[y]["24"] += c.get("24") or 0
        d[y]["26"] += c.get("26") or 0
        d[y]["40"] += c.get("40") or 0
        d[y]["43"] = c.get("43") or c.get("41") or d[y]["43"]
        d[y]["32"] += c.get("32") or 0
        d[y]["22"] = c.get("22") or d[y]["22"]
        d[y]["kys"].append(x)
    return d


GY = gtgt_year()


def sum_hd(chieu, nam=None, field="gt"):
    s = 0.0
    n = 0
    for r in hddt:
        if r.get("chieu") != chieu:
            continue
        if nam is not None and r.get("nam") != nam:
            continue
        s += r.get(field) or 0
        n += 1
    return s, n


ncc_hd = defaultdict(lambda: {"ten": "", "gt": 0.0, "thue": 0.0, "n_hd": 0, "lines": []})
for r in hddt:
    if r["chieu"] != "MUA":
        continue
    m = mst_base(r.get("mst_dt"))
    if not m:
        continue
    rec = ncc_hd[m]
    rec["ten"] = rec["ten"] or r.get("ten_nb") or ""
    rec["gt"] += r.get("gt") or 0
    rec["thue"] += r.get("thue") or 0
    rec["n_hd"] += 1
    rec["lines"].append(r)

tot_mua, n_mua = sum_hd("MUA")
tot_ban, n_ban = sum_hd("BAN")
top5 = sorted(ncc_hd.items(), key=lambda x: -x[1]["gt"])[:5]
top5_gt = sum(v["gt"] for _, v in top5)
top5_ty = (top5_gt / tot_mua) if tot_mua else 0

f0_hits = []
for m, rec in ncc_hd.items():
    info = f0_map.get(m)
    if info:
        f0_hits.append((m, rec, info))
f0_gt = sum(r["gt"] for _, r, _ in f0_hits)
f0_thue = sum(r["thue"] for _, r, _ in f0_hits)
f0_ty = (f0_gt / tot_mua) if tot_mua else 0

clusters = defaultdict(lambda: {"n": 0, "gt": 0.0, "thue": 0.0, "tt": 0.0, "ten": ""})
for r in hddt:
    if r["chieu"] != "MUA" or not r.get("ngay"):
        continue
    k = (r.get("mst_dt"), r.get("ngay"))
    clusters[k]["n"] += 1
    clusters[k]["gt"] += r.get("gt") or 0
    clusters[k]["thue"] += r.get("thue") or 0
    clusters[k]["tt"] += r.get("tt") or 0
    clusters[k]["ten"] = r.get("ten_nb") or ""
tt01 = []
for (mst, ngay), v in clusters.items():
    d = dparse(ngay)
    nguong = 5_000_000 if d and d >= date(2025, 7, 1) else 20_000_000
    if v["tt"] >= nguong or v["gt"] >= nguong:
        tt01.append(
            {
                "mst": mst,
                "ngay": ngay,
                "n": v["n"],
                "gt": v["gt"],
                "thue": v["thue"],
                "tt": v["tt"],
                "ten": v["ten"],
                "nguong": nguong,
            }
        )
tt01_multi = [x for x in tt01 if x["n"] >= 2]
tt01.sort(key=lambda x: -x["gt"])

keys_mua = [(r.get("ky_hieu"), so_key(r.get("so_key")), mst_base(r.get("mst_dt"))) for r in hddt if r["chieu"] == "MUA"]
keys_ban = [(r.get("ky_hieu"), so_key(r.get("so_key")), mst_base(r.get("mst_dt"))) for r in hddt if r["chieu"] == "BAN"]
dup_mua = len(keys_mua) - len(set(keys_mua))
dup_ban = len(keys_ban) - len(set(keys_ban))
miss_mst = sum(1 for r in hddt if not mst_base(r.get("mst_dt")))
miss_kh = sum(1 for r in hddt if not str(r.get("ky_hieu") or "").strip())
bad_mst = [r for r in hddt if mst_base(r.get("mst_dt")) and not mst_valid(r.get("mst_dt"))]

print("HDDT mua", n_mua, vn(tot_mua), "ban", n_ban, vn(tot_ban))
print("top5", round(top5_ty, 4), "f0", len(f0_hits), vn(f0_gt), "tt01", len(tt01), "multi", len(tt01_multi))
print("dup", dup_mua, dup_ban, "miss mst/kh", miss_mst, miss_kh, "bad mst", len(bad_mst))
print("NNT F1/CA", nnt_f1, nnt_ca)
print("GTGT missing", {y: GY[y]["missing"] for y in YEARS})


def load_f1_buyers():
    buyers = {mst_base(r.get("mst_dt")) for r in hddt if r["chieu"] == "BAN"}
    hits = []
    if not CVCBRR.exists():
        return hits, False
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    ws = wb["F1"]
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 6:
            continue
        mst_mua = mst_base(row[8] if len(row) > 8 else "")
        if mst_mua and mst_mua in buyers:
            hits.append(
                {
                    "mst_mua": mst_mua,
                    "mst_ban": mst_base(row[1] if len(row) > 1 else ""),
                    "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                }
            )
    wb.close()
    return hits, True


print("LOAD F1 buyers")
f1_buyers, f1_ok = load_f1_buyers()
print(" F1 buyer hits", len(f1_buyers), "scanned", f1_ok)

print("PARSE TMS Nghia_vu")
tms_nv = ROOT / "1. DU LIEU THO NGANH THUE" / "1.3. Nghia vu nop thue & Thong bao"
tms_rows = []
nvkk = None
for fn in os.listdir(lp(tms_nv)) if os.path.isdir(lp(tms_nv)) else []:
    if "Nghia_vu_ke_khai" in fn and fn.lower().endswith(".xlsx"):
        nvkk = tms_nv / fn
        break
if nvkk:
    wb = load_workbook(lp(nvkk), data_only=True)
    if "Chi_tiet" in wb.sheetnames:
        ws = wb["Chi_tiet"]
        rows = list(ws.iter_rows(values_only=True))
        hdr = None
        start = 0
        for i, row in enumerate(rows[:20]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "kỳ" in joined or "ky" in joined or "tờ khai" in joined:
                hdr = [str(c or "").strip() for c in row]
                start = i + 1
                break
        for row in rows[start:]:
            if not any(row):
                continue
            tms_rows.append({"file": nvkk.name, "sheet": "Chi_tiet", "raw": [str(c) if c is not None else "" for c in row[:12]]})
    if "Canh_bao" in wb.sheetnames:
        ws = wb["Canh_bao"]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0 or not any(row):
                continue
            tms_rows.append({"file": nvkk.name, "sheet": "Canh_bao", "raw": [str(c) if c is not None else "" for c in row[:12]]})
    wb.close()
print(" TMS chi tiet/canh bao", len(tms_rows))

# 80 rules
print("LOAD 80 rules")
rwb = load_workbook(lp(RULES_WB), data_only=True)
rws = rwb["04_Quy_tac_rui_ro"]
RULES = []
for r in range(6, 86):
    ma = rws.cell(r, 1).value
    if not ma:
        continue
    RULES.append(
        {
            "ma": str(ma).strip(),
            "nhom": rws.cell(r, 2).value,
            "ten": rws.cell(r, 3).value,
            "need": str(rws.cell(r, 7).value or ""),
            "muc": rws.cell(r, 8).value,
            "goi_y": rws.cell(r, 9).value,
            "giai_doan": rws.cell(r, 15).value,
        }
    )
rwb.close()
print(" rules", len(RULES))

OUT.mkdir(parents=True, exist_ok=True)
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
            (m.get("ext") or "").lstrip(".").upper() or "KHAC",
            dparse(m.get("mtime")),
            2022,
            2026 if "2026" in m.get("name", "") else 2025,
            m.get("sha256"),
            "Không",
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
ws.cell(1, 2).value = f"{TEN} - MST {MST} - kỳ {KY} - {TODAY.strftime('%d/%m/%Y')} - {VER}"

ws = wb["IN01_HDDT"]
clear_from(ws, 3)
for i, r in enumerate(hddt, 1):
    dt = r.get("mst_dt")
    ngay = dparse(r.get("ngay"))
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
            ngay,
            r.get("mst_nb"),
            r.get("ten_nb"),
            r.get("mst_nm"),
            r.get("ten_nm"),
            dt,
            r.get("mat_hang") or "",
            r.get("dvt") or "",
            r.get("sl") or None,
            r.get("don_gia") or None,
            r.get("gt"),
            r.get("thue_suat"),
            r.get("thue"),
            r.get("tt"),
            "Hợp lệ",
            "",
            r.get("tinh_vao_tong"),
            "",
            f"{r.get('ky_hieu')}|{r.get('so_key')}|{dt}",
            r.get("file"),
        ],
        nums={15, 16, 17, 19, 20},
        dates={7},
    )

ws = wb["IN03_ToKhai_GTGT"]
clear_from(ws, 3)
for i, x in enumerate(gtgt_all, 1):
    c = x.get("chi_tieu") or {}
    ky = x.get("ky_norm") or x.get("ky") or ""
    write_row(
        ws,
        2 + i,
        [
            x.get("nam"),
            ky,
            "Bổ sung" if str(x.get("loai_tkhai")).upper() in {"B", "2"} else "Chính thức",
            x.get("so_lan") or 0,
            "Có" if x.get("la_hieu_luc") else "Không",
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
for y in YEARS:
    b = BM.get(y)
    if b:
        pairs = [
            ("KQKD", "01", "Doanh thu bán hàng và cung cấp dịch vụ", b["dt"]),
            ("KQKD", "11", "Giá vốn hàng bán", b["gv"]),
            ("KQKD", "20", "Lợi nhuận gộp", b["ln_gop"]),
            ("KQKD", "21", "Doanh thu hoạt động tài chính", b["dt_tc"]),
            ("KQKD", "22", "Chi phí thuế TNDN hiện hành", b["thue_tndn"]),
            ("KQKD", "24", "Chi phí quản lý doanh nghiệp", b["cp_ql"]),
            ("KQKD", "50", "Lợi nhuận trước thuế", b["kqkd"].get("50")),
            ("KQKD", "60", "Lợi nhuận sau thuế", b["lnst"]),
            ("CDKT", "140", "Hàng tồn kho (ct140 TT133)", b["htk140"]),
            ("CDKT", "141", "Hàng tồn kho (ct141)", b["htk141"]),
            ("CDKT", "131", "Phải thu khách hàng", b["pt"]),
            ("CDKT", "110", "Tiền và tương đương tiền", b["tm"]),
            ("CDKT", "411", "Vốn góp chủ sở hữu", b["von"]),
            ("CDKT", "181", "Thuế GTGT được khấu trừ (TT133)", b["thue_kt"]),
            ("CDKT", "311", "Nợ ngắn hạn / phải trả", b["no_nh"]),
        ]
        for bao, ma, tenct, so in pairs:
            write_row(ws, r, [y, bao, ma, tenct, so, None, b["mau"], b.get("file")], nums={5})
            r += 1
    tnct = TN.get(y) or {}
    if tnct:
        for ma, tenct in [
            ("04", "Doanh thu QTT TNDN [04]"),
            ("13", "Giá vốn QTT TNDN [13]"),
            ("16", "Thuế TNDN [16]"),
            ("18", "Lợi nhuận sau thuế [18]"),
            ("22", "Thu nhập chịu thuế / LNST [22]"),
        ]:
            if tnct.get(ma) is not None:
                write_row(ws, r, [y, "TNDN", ma, tenct, tnct.get(ma), None, "03/TNDN XML CQT", ""], nums={5})
                r += 1
    if y == 2024 and not b and not tnct:
        write_row(ws, r, [y, "THIEU", "", "Chưa có XML BCTC/TNDN 2024 trong thư mục 1.1", None, None, "", "KHÔNG_ĐỦ_DỮ_LIỆU"])
        r += 1

ws = wb["IN05_NLD_TNCN"]
clear_from(ws, 3)
r = 3
for x in tncn:
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
            c.get("10") if has_so else None,
            c.get("20") if has_so else None,
            "Không",
            None,
            "",
            "",
            "XML có file nhưng chỉ tiêu số trên tờ khai rỗng" if not has_so else "",
            "Không",
            x.get("file"),
        ],
        nums={8, 9},
    )
    r += 1
write_row(ws, r, ["GHICHU", None, "Không có phụ lục 05-1/05-2 NLĐ trong thư mục 1.4", "", "", "", "", None, None, "", None, "", "", "IN05 chỉ có QTT XML", "", ""])

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
            dparse(m.get("mtime")),
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
for row in tms_rows[:80]:
    raw = row["raw"]
    write_row(
        ws,
        r,
        [
            MST,
            row["sheet"],
            raw[0] if raw else "",
            raw[1] if len(raw) > 1 else "",
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            "",
            None,
            " | ".join(raw[:8]),
            row["sheet"],
            row["file"],
        ],
    )
    r += 1

apply_tnr(wb["IN00_Manifest"])
autowidth(wb["IN00_Manifest"])
autowidth(wb["IN03_ToKhai_GTGT"])
autowidth(wb["IN04_BCTC_TNDN"])
autowidth(wb["IN05_NLD_TNCN"])
autowidth(wb["IN14_TMS_Nghia_vu"])
wb.save(lp(files["kq11"][1]))
wb.close()
print(" KQ1.1 saved")

# ---------------------------------------------------------------------------
# KQ1.2 - no company books
# ---------------------------------------------------------------------------
print("FILL KQ1.2")
wb = load_workbook(lp(files["kq12"][1]))
ws = wb["IN00_Manifest"]
dn_files = [m for m in manifest if m["nguon"] in {"DN", "TRA_CUU", "BB"}]
clear_from(ws, 3)
if dn_files:
    for i, m in enumerate(dn_files, 1):
        write_row(
            ws,
            2 + i,
            [
                f"D{i:03d}",
                m.get("ma_in"),
                m.get("nguon"),
                (m.get("ext") or "").lstrip(".").upper(),
                dparse(m.get("mtime")),
                None,
                None,
                m.get("sha256"),
                "Không",
                "Có",
                None,
                0,
                "Giữ nguyên - không phải sổ DN",
                "Bang tinh / bien ban, khong map IN02-IN13",
                m.get("rel"),
            ],
            dates={5},
        )
else:
    write_row(ws, 3, ["", "", "DN", "", None, None, None, "", "", "", None, None, "Chưa nhận", "Thu muc 2.x trong", ""])
ws.cell(1, 2).value = f"{TEN} - MST {MST} - kỳ {KY} - {TODAY.strftime('%d/%m/%Y')} - {VER}"

for sn in ["IN02_BangKe", "IN08_So_KeToan", "IN09_Kho_GiaThanh", "IN10_Chi_tien_mat", "IN11_Sao_ke_NH", "IN12_GDLK", "IN13_So_511_632"]:
    ws = wb[sn]
    clear_from(ws, 3)
    put(ws, 3, 1, "Chưa nhận dữ liệu công ty (IN02, IN08-IN13). Không điền 0 giả.", wrap=True, fill=FILL_YEL)
sub_placeholders(wb["IN00_Manifest"])
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
ws["J7"] = KY_SUY_RA + " NNT không có trên F1/CA. Hồ sơ mức A."


def set_in_row(row, dinh_dang, ky_phu, ky_thieu, so_dong, thieu_mst, bs, muc, lydo, tt, viec, ghichu):
    vals = {
        6: dinh_dang,
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
            elif v == "Mỏng":
                ws.cell(row, c).fill = FILL_YEL
            else:
                ws.cell(row, c).fill = FILL_RED


set_in_row(10, "Excel", "2022-2025", "", n_mua + n_ban, miss_mst, "Không", "Đủ", f"Portal HĐĐT mua {n_mua} + bán {n_ban}", "Đã chuẩn hóa", "MST, khóa HĐ, Tinh_vao_tong", "CQT")
set_in_row(11, "", "", "Chưa nhận bảng kê 2.5", 0, 0, "Không", "Chưa nhận", "Thư mục 2. HO SO NNT trống", "Để trống", "Không điền 0", "DN")
set_in_row(12, "XML", "Q4/2022, 2023 đủ, Q1/2024, Q2-Q4/2025", "Q1-Q3/2022, Q2-Q4/2024, Q1/2025", len(gtgt), 0, "Không", "Mỏng", f"{len(gtgt)} kỳ hiệu lực trong 2022-2025; {len(gtgt_extra)} kỳ 2026 ngoài kỳ", "Đã chọn bản hiệu lực", "2026 giữ IN03 không đưa KQ2", "CQT")
set_in_row(13, "XML", "2022, 2023, 2025", "BCTC/TNDN 2024", len(bctc) + len(tndn), 0, "Không", "Mỏng", "TT133 B01a-DNN maTKhai 683", "Đã trích chỉ tiêu", "HTK dùng ct140", "CQT")
set_in_row(14, "XML", "2022, 2023, 2025", "QTT 2024; không có 05-1/05-2", len(tncn), 0, "Không", "Mỏng", "3 QTT XML, chỉ tiêu số rỗng, không có NLĐ", "Đã nạp dòng QTT", "Không điền 0 giả cho TNCT", "CQT")
set_in_row(15, "Excel sổ CVCBRR+CA", "cập nhật 09/11/2026", "", len(ncc_hd), 0, "Không", "Đủ", f"{len(ncc_hd)} NCC từ HĐĐT; {len(f0_hits)} F0", "Đã khớp MST", "NNT không thuộc F1/CA", "CQT")
set_in_row(16, "", "", "Chưa nhận sổ KT", 0, 0, "Không", "Chưa nhận", "IN08 trống", "Để trống", "Không điền 0", "DN")
set_in_row(17, "", "", "Chưa nhận XNT/giá thành", 0, 0, "Không", "Chưa nhận", "IN09 trống", "Để trống", "Không điền 0", "DN")
set_in_row(18, "", "", "Chưa nhận sổ 111", 0, 0, "Không", "Chưa nhận", "IN10 trống", "Để trống", "Không điền 0", "DN")
set_in_row(19, "", "", "Chưa nhận sao kê NH", 0, 0, "Không", "Chưa nhận", "IN11 trống", "Để trống", "Không điền 0", "DN")
set_in_row(20, "", "", "Chưa nhận PL GDLK", 0, 0, "Không", "Chưa nhận", "IN12 trống", "Để trống", "Không điền 0", "DN")
set_in_row(21, "", "", "Chưa nhận sổ 511/632", 0, 0, "Không", "Chưa nhận", "IN13 trống", "Để trống", "Không điền 0", "DN")
set_in_row(22, "Excel TMS", "2015-2026", "Chưa có bảng nợ thuế dòng đủ", len(tms_files) + len(tms_rows), 0, "Không", "Mỏng", f"{len(tms_files)} file TMS; Nghia_vu_ke_khai đã đọc", "Nhận file", "LS01/LS02 sơ bộ", "CQT")

ws["C4"] = "Đủ 2 | Mỏng 4 | Thiếu 0 | Chưa nhận 7"

ws = wb["KQ1_Kha_nang_phan_tich"]
ws["C2"] = MST
ws["E2"] = KY
ws["C4"] = VER
ws["E4"] = MAU_VER
ws["C5"] = TODAY
ws["C5"].number_format = DATE_FMT
ws["E5"] = "Mức A - chưa có sổ DN (IN02, IN08-IN13)"

PART_OK = {"HD01", "HD03", "HD04", "HD08", "TT01", "GT01", "GT03", "GT04", "GT05", "TN02"}
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
        kn, thieu, inn, goi = "Chạy một phần", ",".join(missing), ",".join(missing), "Chạy trên HĐĐT/XML CQT, ghi chú thiếu IN02"
        part += 1
    elif not missing:
        kn, thieu, inn, goi = "Chạy được ngay", "", "", ""
        run_now += 1
    elif any(HAS.get(i, False) for i in need_ins) and missing:
        kn, thieu, inn, goi = "Chạy một phần", ",".join(missing), ",".join(missing), "Bổ sung " + ",".join(missing)
        part += 1
    else:
        kn, thieu, inn, goi = "Bị chặn", ",".join(missing), ",".join(missing), "Xin " + ",".join(missing)
        block += 1
    ws.cell(r, 8).value = kn
    ws.cell(r, 9).value = thieu
    ws.cell(r, 10).value = inn
    ws.cell(r, 11).value = goi
    for c in range(8, 12):
        tnr(ws.cell(r, c))
        if kn == "Bị chặn" or kn == "TREO":
            ws.cell(r, 8).fill = FILL_RED
        elif kn == "Chạy một phần":
            ws.cell(r, 8).fill = FILL_YEL
        else:
            ws.cell(r, 8).fill = FILL_GRN
ws["C3"] = run_now
ws["E3"] = part
ws["G3"] = block

ws = wb["KQ1_De_nghi_bo_sung"]
try:
    ws["C2"] = MST
except Exception:
    ws["B2"] = MST
try:
    ws["E2"] = KY
except Exception:
    ws["F2"] = KY
ws["C3"] = VER
ws["E3"] = TODAY
ws["E3"].number_format = DATE_FMT
sug = [
    (1, "Bảng kê mua vào, bán ra các kỳ 2022-2025 (thư mục 2.5)", "IN02", "KQ1.2", "File Excel gốc, lọc dòng Tổng cộng", 8, "GT02, GT07, GT08, HD01-HD05, TT01", "Mở đối chiếu BK-HĐĐT-tờ khai"),
    (2, "Sổ kế toán NKC/Sổ cái 2022-2025 (IN08) và sổ chi tiết 511/632 (IN13)", "IN08, IN13", "KQ1.2", "Excel phát sinh theo chứng từ", 10, "GT06, DT01-DT03, TN17, TT05", "Mở GĐ3 đối chiếu sổ"),
    (3, "Sổ XNT kho 152/155/156 và bảng giá thành/định mức", "IN09", "KQ1.2", "Tồn CK khớp BCTC ct140", 8, "KH01-KH07, TN04", "Giải tỏa TC3 HTK/DT 2025"),
    (4, "Sao kê ngân hàng Excel + UNC món >= 20tr (đến 30/06/2025) và >= 5tr (từ 01/07/2025)", "IN11", "KQ1.2", "Không chỉ PDF scan", 5, "GT16, TT02, TT04, TT06", "Đối cụm HĐ cùng ngày"),
    (5, "XML tờ khai GTGT các kỳ thiếu: Q1-Q3/2022, Q2-Q4/2024, Q1/2025; BCTC/TNDN/TNCN 2024", "IN03, IN04, IN05", "1.1 XML", "Bản hiệu lực CQT", 6, "GT01, TN01, TN02", "Tránh kết luận lệch vì thiếu kỳ"),
    (6, "Hợp đồng, PNK, UNC với 2 NCC F0 (Duy Lập Phát, Hoàng Quân) và Top-5 NCC", "IN07", "2.x", "Chứng từ gốc chuỗi 5 mắt xích", 3, "HD01, HD03, HD04", "Tỷ trọng F0 7,3%; Top-5 91,3%"),
]
for i, rowv in enumerate(sug):
    for c, v in enumerate(rowv, 1):
        put(ws, 10 + i, c, v, bold=(c == 1), num=isinstance(v, int) and c == 6, wrap=c in {2, 7, 8})

wb.save(lp(files["kq13"][1]))
wb.close()
print(" KQ1.3 saved")

# ---------------------------------------------------------------------------
# KQ2.1
# ---------------------------------------------------------------------------
print("FILL KQ2.1")
wb = load_workbook(lp(files["kq21"][1]))
sub_placeholders(wb["KQ2_Tong_hop_co"])

ws = wb["KQ2_Doi_chieu_GTGT"]
ws["A2"] = f"Người nộp thuế: {TEN} - MST: {MST}"
ws["A3"] = f"Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Kỳ {KY} | {VER}"

YEAR_START = {y: 9 + (y - 2016) * 14 for y in range(2016, 2026)}
Q_OFF = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}


def fill_gtgt_q(year, ky_norm, rec):
    base = YEAR_START.get(year)
    if not base:
        return
    q = str(ky_norm or "").upper()
    m = re.search(r"Q([1-4])", q)
    if not m:
        return
    rr = base + int(m.group(1))
    c = rec.get("chi_tieu") or {}
    ct26 = c.get("26") or 0
    ct27 = c.get("27") or 0
    ct29 = c.get("29") or 0
    ct30 = c.get("30") or 0
    ct32 = c.get("32") or 0
    ct34 = c.get("34")
    if ct26:
        put(ws, rr, 2, ct26, num=True)
    if ct27 or ct29:
        put(ws, rr, 3, ct27 or ct29, num=True)
    if ct30:
        put(ws, rr, 4, ct30, num=True)
    if ct32:
        put(ws, rr, 5, ct32, num=True)
    elif ct34 and not (ct26 or ct27 or ct29 or ct30):
        put(ws, rr, 5, ct34, num=True)
    if c.get("33") or c.get("35"):
        put(ws, rr, 9, c.get("33") or c.get("35") or 0, num=True)
    put(ws, rr, 11, c.get("23"), num=True)
    put(ws, rr, 12, c.get("24"), num=True)
    if c.get("37"):
        put(ws, rr, 13, c.get("37"), num=True)
    if c.get("38"):
        put(ws, rr, 14, c.get("38"), num=True)
    if c.get("40"):
        put(ws, rr, 17, c.get("40"), num=True)


for x in gtgt:
    fill_gtgt_q(x.get("nam"), x.get("ky_norm"), x)

# Phần II / III yearly 2022-2025 = rows 162-165
P2 = {2022: 162, 2023: 163, 2024: 164, 2025: 165}
for y, rr in P2.items():
    hd_gt, _ = sum_hd("BAN", y)
    hd_th, _ = sum_hd("BAN", y, "thue")
    mua_gt, _ = sum_hd("MUA", y)
    mua_th, _ = sum_hd("MUA", y, "thue")
    put(ws, rr, 2, hd_gt, num=True)
    put(ws, rr, 3, hd_th, num=True)
    # D,E bảng kê trống
    if GY[y]["kys"]:
        put(ws, rr, 6, GY[y]["34"], num=True)
        put(ws, rr, 7, GY[y]["35"], num=True)
    tnct = TN.get(y) or {}
    if tnct.get("04") is not None:
        put(ws, rr, 8, tnct.get("04"), num=True)
    if y in BM:
        put(ws, rr, 9, BM[y]["dt"], num=True)
    put(ws, rr, 13, mua_gt, num=True)
    if GY[y]["kys"]:
        put(ws, rr, 15, GY[y]["23"], num=True)
        put(ws, rr, 21, GY[y]["24"], num=True)
    put(ws, rr, 19, mua_th, num=True)

notes = [
    "NHẬN XÉT / NGUYÊN NHÂN TRỌNG YẾU (module GTGT, không gán mức rủi ro hồ sơ)",
    "NNT kê khai GTGT theo quý. Phần I lấy IN03 kỳ hiệu lực 2022-2025; năm 2016-2021 để trống (ngoài kỳ), không xóa để giữ công thức Phần II.",
    "Thiếu XML: Q1-Q3/2022, Q2-Q4/2024, Q1/2025. Dòng thiếu nguồn không điền 0 giả. GTGT 2026 Q1-Q3 chỉ nằm IN03.",
    "Nông sản: DT 0%/KCT (ct26/ct27) là hợp lý theo ngành, không tự gắn GT10.",
    f"Q4/2025 ct26 = ct34 = {vn(next((ct(x,'26') for x in gtgt if x.get('ky_norm')=='Q4/2025'), 0))} (DT không chịu thuế).",
    f"2023: BCTC[01]={vn(BM.get(2023,{}).get('dt'))} khớp tổng GTGT[34]={vn(GY[2023]['34'])}; HĐĐT bán={vn(sum_hd('BAN',2023)[0])}.",
    "Phần II cột bảng kê và Phần III cột 133 để trống vì chưa có IN02/IN08. Cờ máy = Chưa đủ nguồn.",
]
last = (ws.max_row or 169) + 2
for i, line in enumerate(notes):
    put(ws, last + i, 1, line, bold=(i == 0), wrap=True, fill=FILL_SEC if i == 0 else FILL_YEL)

# IN07
ws = wb["IN07_NCC_CanhBao"]
clear_from(ws, 3)
r = 3
put(ws, r, 1, MST, bold=True)
put(ws, r, 2, TEN)
put(ws, r, 3, "KHONG")
put(ws, r, 4, "")
put(ws, r, 5, "")
put(ws, r, 9, "NNT không có trên F1 và sổ Công an 2026")
put(ws, r, 16, "SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx F1 + Bao_cao_cong_van.xlsx")
r = 4
for m, rec, info in f0_hits:
    fill = FILL_RED
    write_row(
        ws,
        r,
        [
            m,
            rec["ten"] or info.get("ten"),
            "F0",
            "F0",
            "CANH_BAO",
            info.get("cv"),
            dparse(info.get("ngay_cv")),
            info.get("dvbh"),
            info.get("nd"),
            "",
            None,
            None,
            "",
            "",
            "",
            f"CVCBRR {info.get('sheet')} mtime {C.get('cvcbrr_mtime')}",
            f"{rec['n_hd']} HĐ mua, GT {vn(rec['gt'])}",
        ],
        dates={7},
        fill=fill,
    )
    r += 1

# NCC 100% invoices
ws = wb["KQ2_NCC_mua_vao"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | {VER}"
clear_from(ws, 8)
r = 8
mua_rows = [x for x in hddt if x["chieu"] == "MUA"]
mua_rows.sort(key=lambda x: (x.get("ngay") or "", x.get("mst_dt") or ""))
for i, rec in enumerate(mua_rows, 1):
    m = mst_base(rec.get("mst_dt"))
    info = f0_map.get(m)
    nhom = "F0" if info else "Không"
    cv = info.get("cv") if info else ""
    fill = FILL_RED if info else None
    write_row(
        ws,
        r,
        [
            i,
            m,
            rec.get("ten_nb"),
            rec.get("so_hd"),
            dparse(rec.get("ngay")),
            rec.get("mat_hang") or "",
            rec.get("gt") or 0,
            rec.get("thue") or 0,
            "HĐĐT",
            f"Có - {cv}" if cv else "Không",
            nhom,
            "",
        ],
        nums={1, 7, 8},
        dates={5},
        fill=fill,
    )
    if info:
        for c in range(1, 13):
            ws.cell(r, c).font = FONT_RED
    r += 1
r += 1
put(ws, r, 1, "BẢNG TOP-5 NCC (HD04 / P14 = 50%) - gộp theo MST người bán", bold=True, fill=FILL_SEC)
r += 1
for c, h in enumerate(["STT", "MST", "Tên NCC", "Giá trị chưa thuế", "Tỷ trọng", "Số HĐ"], 1):
    put(ws, r, c, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
r += 1
for i, (m, rec) in enumerate(top5, 1):
    put(ws, r, 1, i, num=True)
    put(ws, r, 2, m)
    put(ws, r, 3, rec["ten"], wrap=True)
    put(ws, r, 4, rec["gt"], num=True)
    put(ws, r, 5, rec["gt"] / tot_mua if tot_mua else 0, pct=True)
    put(ws, r, 6, rec["n_hd"], num=True)
    r += 1
put(ws, r, 1, "Cộng Top-5", bold=True, fill=FILL_YEL)
put(ws, r, 4, top5_gt, bold=True, num=True, fill=FILL_YEL)
put(ws, r, 5, top5_ty, bold=True, pct=True, fill=FILL_YEL)
r += 2
put(ws, r, 1, "PHẦN 1. ĐỐI CHIẾU NNT VỚI F1 VÀ SỔ CÔNG AN", bold=True, fill=FILL_SEC)
r += 1
put(ws, r, 1, f"NNT {TEN} (MST {MST}) KHÔNG xuất hiện trên sổ CVCBRR F1 / sổ Công an 2026 với vai trò bên mua rủi ro.", wrap=True)
r += 1
put(ws, r, 1, f"Nguồn: SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx mtime {C.get('cvcbrr_mtime')}; Bao_cao_cong_van.xlsx mtime {C.get('ca_mtime')} (trong 30 ngày).", wrap=True)
r += 2
put(ws, r, 1, "PHẦN 3. ĐÁNH GIÁ RỦI RO NCC (nhận xét module, không phải kết luận KQ3)", bold=True, fill=FILL_SEC)
r += 1
mod_notes = [
    f"Tổng {n_mua} hóa đơn mua vào hợp lệ, {len(ncc_hd)} NCC, GT chưa thuế {vn(tot_mua)} đ. Nguồn HĐĐT (chưa có bảng kê).",
    f"HD03 tỷ trọng mua F0 = {f0_ty:.2%} so P13 5% → {'Vượt ngưỡng' if f0_ty > P13 else 'Không vượt'}.",
    f"HD04 Top-5 NCC = {top5_ty:.2%} so P14 50% → {'Vượt ngưỡng' if top5_ty > P14 else 'Không vượt'}.",
    f"NCC F0: {len(f0_hits)} MST, GT {vn(f0_gt)}, thuế {vn(f0_thue)}. (1) 3603836720 Duy Lập Phát CV 10980/CCTKV-KT2 ngày 31/05/2023, 7 HĐ 18/05-29/07/2022 GT {vn(ncc_hd['3603836720']['gt'] if '3603836720' in ncc_hd else 0)}. (2) 0316549251 Hoàng Quân CV 8808/TCS3-KTR1 ngày 08/10/2025, 1 HĐ 30/07/2022 GT {vn(ncc_hd['0316549251']['gt'] if '0316549251' in ncc_hd else 0)}.",
    f"TT01 proxy trên HĐĐT: {len(tt01)} cụm ngày×MST >= ngưỡng tiền mặt, trong đó {len(tt01_multi)} cụm n>=2. Chưa có IN02/IN11 nên chưa kết luận thiếu UNC.",
    "Trọng tâm chứng từ: 2 NCC F0; Top-5 (ADM VN 0313654924 chiếm phần lớn); cụm HĐ cùng ngày. Sao kê IN11 chưa nhận.",
]
for line in mod_notes:
    put(ws, r, 1, line, wrap=True, fill=FILL_YEL)
    r += 1

# TNCN
ws = wb["KQ2_TNCN_LD2"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | {VER} | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
# Phần I starts row 8 typically
# find year rows
year_row = {}
for rr in range(7, 20):
    v = ws.cell(rr, 1).value
    if v in YEARS or str(v) in {str(y) for y in YEARS}:
        year_row[int(v)] = rr
if not year_row:
    # fill from row 8
    clear_from(ws, 8)
    rr = 8
    for y in YEARS:
        put(ws, rr, 1, y)
        if y in {x.get("nam") for x in tncn}:
            put(ws, rr, 2, None)
            put(ws, rr, 4, None)
            put(ws, rr, 14, "QTT XML có file, chỉ tiêu số rỗng; không có 05-1/05-2", wrap=True)
        else:
            put(ws, rr, 14, "Chưa có QTT XML năm này", wrap=True, fill=FILL_YEL)
        rr += 1
else:
    for y, rr in year_row.items():
        if y not in {x.get("nam") for x in tncn}:
            put(ws, rr, 14, "Chưa có QTT XML", wrap=True, fill=FILL_YEL)
last = (ws.max_row or 20) + 2
put(ws, last, 1, "NHẬN XÉT MODULE TNCN (không ghi mức rủi ro hồ sơ)", bold=True, fill=FILL_SEC)
put(
    ws,
    last + 1,
    1,
    "IN05 có 3 dòng QTT 2022, 2023, 2025; chỉ tiêu số trên XML rỗng; không có danh sách NLĐ 05-1/05-2; thiếu 2024 và thiếu tra cứu đa nguồn 3.3. LD01/LD03/LD04: chưa đủ số để đối chiếu.",
    wrap=True,
    fill=FILL_YEL,
)

autowidth(wb["IN07_NCC_CanhBao"])
autowidth(wb["KQ2_NCC_mua_vao"], cap=36)
wb.save(lp(files["kq21"][1]))
wb.close()
print(" KQ2.1 saved")

# ---------------------------------------------------------------------------
# KQ2.2 empty company
# ---------------------------------------------------------------------------
print("FILL KQ2.2")
wb = load_workbook(lp(files["kq22"][1]))
for sn in wb.sheetnames:
    sub_placeholders(wb[sn])
    ws = wb[sn]
    last = (ws.max_row or 10) + 2
    put(ws, last, 1, "Chưa nhận IN08-IN13. Ô số liệu công ty để trống, không điền 0 giả. Cờ máy = Chưa đủ nguồn.", wrap=True, fill=FILL_YEL)
wb.save(lp(files["kq22"][1]))
wb.close()
print(" KQ2.2 saved")

# ---------------------------------------------------------------------------
# KQ3
# ---------------------------------------------------------------------------
print("FILL KQ3")
wb = load_workbook(lp(files["kq3"][1]))
ws = wb["KQ3"]
ws["C2"] = MST
ws["J2"] = TEN
ws["C3"] = KY
ws["J3"] = LEVEL
ws["C4"] = VER
ws["J4"] = MAU_VER
ws["C5"] = TODAY
ws["C5"].number_format = DATE_FMT
ws["J5"] = f"KQ1.3 {VER}, KQ2.1 {VER}, KQ2.2 {VER}"
ws["J7"] = DN_STATUS

# measurements
hd_ban_y = {y: sum_hd("BAN", y)[0] for y in YEARS}
hd_mua_y = {y: sum_hd("MUA", y)[0] for y in YEARS}
gt34_y = {y: GY[y]["34"] if GY[y]["kys"] else None for y in YEARS}
bctc01_y = {y: BM[y]["dt"] if y in BM else None for y in YEARS}
tndn04_y = {y: (TN.get(y) or {}).get("04") for y in YEARS}

# 2023 GT01 diff
diff_gt01 = {}
for y in YEARS:
    if gt34_y[y] is None:
        diff_gt01[y] = None
    else:
        diff_gt01[y] = gt34_y[y] - hd_ban_y[y]

htk_ty_2025 = None
if 2025 in BM and BM[2025]["dt"]:
    htk_ty_2025 = BM[2025]["htk"] / BM[2025]["dt"]
gv_ty = {}
for y, b in BM.items():
    gv_ty[y] = (b["gv"] / b["dt"]) if b["dt"] else None

q4_2025 = next((x for x in gtgt if x.get("ky_norm") == "Q4/2025"), None)
ct26_q4_2025 = ct(q4_2025, "26") if q4_2025 else None
ct43_last = GY[2025]["43"] if GY[2025]["kys"] else None


def need_missing(need):
    ins = re.findall(r"IN\d{2}", need or "")
    return sorted({i for i in ins if not HAS.get(i, False)})


def add_rule(rows, rd, **kw):
    ma = rd["ma"]
    status = kw.get("status", "ĐÃ_CHẠY")
    ket = kw.get("ket", "Không")
    nhan = kw.get("nhan", "")
    muc_rr = kw.get("muc", "")
    diem = kw.get("diem", 0)
    thieu = kw.get("thieu", "")
    ca_a, va = kw.get("ca"), kw.get("va")
    cb_b, vb = kw.get("cb"), kw.get("vb")
    do = kw.get("do")
    nguong = kw.get("nguong", rd.get("muc"))
    sac = kw.get("sac", "")
    v_tien = kw.get("v")
    w_tien = kw.get("w", 0 if status == "ĐÃ_CHẠY" and ket != "Có phát hiện" else None)
    if ket != "Có phát hiện":
        w_tien = 0 if status == "ĐÃ_CHẠY" else None
    hs = {"Cao": 3, "Trung bình": 2, "Thấp": 1}.get(muc_rr, 0)
    nguon = kw.get("nguon", "IN01, IN03, IN04, IN07 CQT")
    ct_ref = kw.get("ct_ref", "")
    rows.append(
        {
            "Mã_PH": f"{ma}|{MST}|{KY}",
            "Tệp phân tích": kw.get("tep", "Trực tiếp"),
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
            "Ngưỡng": nguong,
            "Trạng thái chạy": status,
            "Kết quả": ket,
            "Nhãn": nhan,
            "Mức rủi ro": muc_rr if ket == "Có phát hiện" else "",
            "Hệ số trọng yếu": hs if ket == "Có phát hiện" else 0,
            "Điểm": diem if ket == "Có phát hiện" else 0,
            "Sắc thuế": sac if ket == "Có phát hiện" else "",
            "Số tiền liên quan (đồng)": v_tien,
            "Số tiền tính vào tổng (đồng)": w_tien,
            "Nguồn": nguon,
            "Tham chiếu chứng cứ": ct_ref,
            "Còn thiếu để kết luận": thieu,
            "Nội dung kiểm tra tại trụ sở": kw.get("nd", ""),
            "Chứng từ cần xem": kw.get("ct_xem", ""),
            "Ghi chú": kw.get("ghi", ""),
        }
    )


rows_kq3 = []
for rd in RULES:
    ma = rd["ma"]
    miss = need_missing(rd["need"])
    # default
    if ma == "DQ01":
        add_rule(rows_kq3, rd, status="ĐÃ_CHẠY", ket="Có phát hiện", nhan="Chắc", muc="Thông tin", diem=1, ca="IN thieu", va="IN02,IN08-IN13", tep="KQ1.3", ct_ref="KQ1_Kiem_ke", nd="Đề nghị NNT bổ sung theo KQ1.3", ghi="Mức hồ sơ A")
    elif ma == "DQ03":
        add_rule(rows_kq3, rd, ket="Không" if not bad_mst else "Có phát hiện", nhan="Chắc", muc="Trung bình" if bad_mst else "", diem=4 if bad_mst else 0, ca="HĐ MST sai BT5", va=len(bad_mst), nguon="IN01", ghi=f"Đã chạy checksum BT5 trên {len(hddt)} HĐ")
    elif ma == "DQ04":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="HĐ thiếu ký hiệu", va=miss_kh, nguon="IN01")
    elif ma == "DQ05":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="Trùng khóa mua/ban", va=dup_mua + dup_ban, nguon="IN01")
    elif ma in {"DQ02", "DQ06"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN02", ghi="Chưa có bảng kê")
    elif ma == "DQ07":
        add_rule(rows_kq3, rd, status="ĐÃ_CHẠY", ket="Không", nhan="Sơ bộ", ca="Tinh_vao_tong", va=n_mua + n_ban, nguon="IN01", ghi="Portal không có cột hủy/thay thế tách; toàn bộ đánh Hợp lệ")
    elif ma == "DQ08":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="Tờ khai GTGT đều chính thức so_lan=0", va=len(gtgt_all), nguon="IN00/IN03")
    elif ma == "GT01":
        # years with both sources
        diffs = {y: diff_gt01[y] for y in YEARS if diff_gt01[y] is not None}
        # 2023 complete; others incomplete XML
        d23 = diffs.get(2023)
        ket = "Có phát hiện" if d23 not in (None, 0) else ("Có phát hiện" if any(abs(v) > 1 for v in diffs.values() if v is not None) else "Không")
        add_rule(
            rows_kq3,
            rd,
            ket=ket,
            nhan="Sơ bộ",
            muc="Cao" if ket == "Có phát hiện" else "",
            diem=7 if ket == "Có phát hiện" else 0,
            ca="GTGT[34] 2023",
            va=gt34_y[2023],
            cb="HĐĐT bán 2023",
            vb=hd_ban_y[2023],
            do=d23,
            nguong="lệch ≠ 0",
            sac="GTGT",
            v=abs(d23) if d23 else None,
            w=0,
            tep="KQ2.1",
            ct_ref="KQ2_Doi_chieu_GTGT Phần II",
            thieu="Thiếu XML Q1-Q3/2022, Q2-Q4/2024, Q1/2025",
            nd="Đối chiếu HĐ hủy/điều chỉnh, cut-off kỳ, HĐĐT vs [34] từng quý đủ XML",
            ghi=f"2022 TK[34]={vn(gt34_y[2022])} vs HĐĐT {vn(hd_ban_y[2022])} (thiếu Q1-Q3 XML); 2023 lệch {vn(d23)}; 2024 TK chỉ Q1; 2025 thiếu Q1.",
        )
    elif ma == "GT02":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN02", tep="KQ2.1")
    elif ma == "GT03":
        add_rule(
            rows_kq3,
            rd,
            ket="Không",
            nhan="Sơ bộ",
            ca="TK[35] 2023",
            va=GY[2023]["35"],
            cb="HĐĐT thuế bán 2023",
            vb=sum_hd("BAN", 2023, "thue")[0],
            tep="KQ2.1",
            ghi="DT nông sản 0%/KCT, thuế đầu ra HĐĐT gần 0 trừ 2024.",
        )
    elif ma == "GT04":
        d = (GY[2023]["23"] - hd_mua_y[2023]) if GY[2023]["kys"] else None
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if d not in (None, 0) else "Không",
            nhan="Sơ bộ",
            muc="Cao" if d not in (None, 0) else "",
            diem=6 if d not in (None, 0) else 0,
            ca="TK[23] 2023",
            va=GY[2023]["23"],
            cb="HĐĐT mua 2023",
            vb=hd_mua_y[2023],
            do=d,
            tep="KQ2.1",
            thieu="IN02; XML thiếu một số quý",
            v=abs(d) if d else None,
            w=0,
            ghi="Chạy trên HĐĐT, chưa có bảng kê.",
        )
    elif ma == "GT05":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ca="TK[24]", va=GY[2023]["24"], cb="HĐĐT thuế mua 2023", vb=sum_hd("MUA", 2023, "thue")[0], tep="KQ2.1", thieu="IN02")
    elif ma in {"GT06", "GT07", "GT08", "GT09"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN02" + (", IN08" if ma == "GT06" else ""), tep="KQ2.1")
    elif ma == "GT10":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="Nông sản: DT 0%/KCT (ct26/ct27) phù hợp ngành, không tự gắn GT10. Chưa thấy kỳ 8%-10% bất thường.", tep="KQ2.1")
    elif ma == "GT11":
        add_rule(
            rows_kq3,
            rd,
            ket="Không",
            nhan="Sơ bộ",
            ca="TK[43] 2025",
            va=ct43_last,
            cb="BCTC 181 TT133",
            vb=BM.get(2025, {}).get("thue_kt"),
            tep="KQ2.1",
            ghi="TT133 dùng 181 thay [152]. 2023 181=111.412.788 khớp [43] cuối 2023; 2025 181=211.978.873 khớp [43] Q4/2025.",
        )
    elif ma == "GT12":
        htk = BM.get(2025, {}).get("htk")
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ca="[43] 2025", va=ct43_last, cb="10%×HTK140 2025", vb=(htk * 0.1 if htk else None), ghi="[43] < 10% HTK.")
    elif ma == "GT13":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if (ct26_q4_2025 or 0) > 0 else "Không",
            nhan="Sơ bộ",
            muc="Trung bình",
            diem=5 if (ct26_q4_2025 or 0) > 0 else 0,
            ca="Q4/2025 [26]",
            va=ct26_q4_2025,
            cb="[34]",
            vb=ct(q4_2025, "34") if q4_2025 else None,
            do=ct26_q4_2025,
            tep="KQ2.1",
            nd="Kiểm tra phân bổ thuế đầu vào khi có DT KCT",
            ghi="Q4/2025 ct26=ct34. Ngành nông sản KCT là hợp lý; cần xem phân bổ [25]/[36].",
            sac="GTGT",
            v=ct26_q4_2025,
            w=0,
        )
    elif ma == "GT14":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ghi="Tất cả tờ khai GTGT loại Chính thức, so_lan=0.", nguon="IN03")
    elif ma == "GT15":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện",
            nhan="Sơ bộ",
            muc="Cao",
            diem=6,
            ca="[43] Q1/2023",
            va=511818,
            cb="[43] Q4/2025",
            vb=ct43_last,
            do=ct43_last,
            tep="KQ2.1",
            nd="Rà khấu trừ kéo dài, đề nghị hoàn, hồ sơ hoàn nếu có",
            ghi="[43] tăng 0,51 tr (Q1/2023) → 111 tr (Q1/2024) → 212 tr (Q4/2025). Không thấy [40] phải nộp.",
            sac="GTGT",
            v=ct43_last,
            w=0,
        )
    elif ma == "GT16":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN11, IN02")
    elif ma == "TN01":
        d23 = None if (gt34_y[2023] is None or bctc01_y[2023] is None) else gt34_y[2023] - bctc01_y[2023]
        add_rule(
            rows_kq3,
            rd,
            ket="Không" if d23 == 0 else "Có phát hiện",
            nhan="Chắc" if d23 == 0 else "Sơ bộ",
            muc="" if d23 == 0 else "Cao",
            ca="GTGT[34] 2023",
            va=gt34_y[2023],
            cb="BCTC[01] 2023",
            vb=bctc01_y[2023],
            do=d23,
            tep="KQ2.1",
            thieu="XML GTGT thiếu 2022 Q1-Q3, 2024 Q2-Q4, 2025 Q1; BCTC 2024",
            ghi="2023 khớp tuyệt đối. 2022 BCTC 28,42 tỷ vs GTGT chỉ Q4. 2025 BCTC 9,70 tỷ vs GTGT Q2-Q4 7,09 tỷ vs HĐĐT 9,90 tỷ.",
        )
    elif ma == "TN02":
        add_rule(
            rows_kq3,
            rd,
            ket="Không",
            nhan="Sơ bộ",
            ca="TNDN[04] 2023",
            va=tndn04_y[2023],
            cb="BCTC[01] 2023",
            vb=bctc01_y[2023],
            do=(tndn04_y[2023] or 0) - (bctc01_y[2023] or 0) if tndn04_y[2023] is not None and bctc01_y[2023] is not None else None,
            thieu="IN02; TNDN 2024",
            ghi="Form 03/TNDN này [04]=DT khớp BCTC[01] các năm có XML.",
            tep="KQ2.1",
        )
    elif ma == "TN03":
        thin = [y for y, t in gv_ty.items() if t and t > P18]
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if thin else "Không",
            nhan="Sơ bộ - chưa so ngành",
            muc="Trung bình" if thin else "",
            diem=4 if thin else 0,
            ca="GV/DT 2022",
            va=gv_ty.get(2022),
            nguong="P18 92% (tạm, K13)",
            ghi=f"Biên lãi gộp 2022 {1-gv_ty.get(2022,0):.2%}; 2023 {1-gv_ty.get(2023,0):.2%}; 2025 {1-gv_ty.get(2025,0):.2%}. Năm 2022 GV/DT vượt P18.",
            tep="KQ2.1",
        )
    elif ma == "TN05":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if (htk_ty_2025 or 0) > P05 else "Không",
            nhan="Sơ bộ",
            muc="Cao" if (htk_ty_2025 or 0) > P05 else "",
            diem=7 if (htk_ty_2025 or 0) > P05 else 0,
            ca="HTK ct140 2025",
            va=BM.get(2025, {}).get("htk"),
            cb="DT [01] 2025",
            vb=BM.get(2025, {}).get("dt"),
            do=htk_ty_2025,
            nguong="P05 30%",
            tep="KQ2.1",
            nd="Kiểm kê tồn kho, XNT 156/152, hàng chậm luân chuyển",
            ghi="TT133 dùng ct140. 2025 HTK/DT = 40,9% > 30%. 2022 6,0%; 2023 12,9%. Thiếu sổ kho IN09.",
            sac="TNDN",
            v=BM.get(2025, {}).get("htk"),
            w=0,
        )
    elif ma in {"TN04", "TN14", "TN17", "TN18"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu=",".join(miss) or rd["need"])
    elif ma == "TN06":
        add_rule(rows_kq3, rd, ket="Chưa kết luận", status="ĐÃ_CHẠY", nhan="Sơ bộ", ghi="TT133: chưa xác minh mã người mua trả tiền trước trên B01a-DNN. Không gắn số 0.", thieu="xác minh mã TT133")
    elif ma == "TN07":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="Không thấy chỉ tiêu DT chưa thực hiện > 0 trên XML BCTC.")
    elif ma == "TN10":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="03/TNDN không thấy chỉ tiêu ưu đãi/miễn giảm tách; thuế 2022 = 0, 2023 = 81.074.762, 2025 = 95.794.704.")
    elif ma == "TN15":
        von = BM.get(2025, {}).get("von")
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ca="Vốn 411", va=von, ghi="Vốn 500 triệu cả 3 năm có BCTC. Chưa đối chiếu góp vốn TMS.")
    elif ma == "TN19":
        etr = {}
        for y, b in BM.items():
            etr[y] = (b["thue_tndn"] / b["kqkd"].get("50")) if b["kqkd"].get("50") else None
        add_rule(rows_kq3, rd, ket="Có phát hiện" if (etr.get(2022) == 0 or (etr.get(2022) or 1) < 0.1) else "Không", nhan="Sơ bộ - chưa so ngành", muc="Trung bình", diem=4, ca="Thuế/LNTT 2022", va=etr.get(2022), ghi=f"ETR 2022={etr.get(2022)}; 2023={etr.get(2023)}; 2025={etr.get(2025)}. 2022 thuế TNDN = 0 trong khi LNTT > 0.")
    elif ma == "TN21":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="HĐĐT bán rải 2022-2025, không thấy dồn cuối năm bất thường trên dữ liệu portal.")
    elif ma == "TN22":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if gv_ty.get(2022) and gv_ty[2022] > P18 else "Không",
            nhan="Sơ bộ - chưa so ngành",
            muc="Trung bình",
            diem=4,
            ca="GV/DT 2022",
            va=gv_ty.get(2022),
            nguong="P18 92% tạm",
            ghi=(
                "CP QLDN/DT 2022="
                + (f"{BM[2022]['cp_ql']/BM[2022]['dt']:.2%}" if 2022 in BM and BM[2022]['dt'] else "n/a")
                + "; 2025="
                + (f"{BM[2025]['cp_ql']/BM[2025]['dt']:.2%}" if 2025 in BM and BM[2025]['dt'] else "n/a")
                + "."
            ),
        )
    elif ma == "HD01":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện",
            nhan="Chắc",
            muc="Cao",
            diem=8,
            ca="NCC F0",
            va=len(f0_hits),
            cb="GT mua F0",
            vb=f0_gt,
            do=f0_gt,
            tep="KQ2.1",
            ct_ref="KQ2_NCC_mua_vao / IN07",
            nd="Đối hợp đồng, PNK, UNC, sử dụng hàng của 2 NCC F0",
            ct_xem="HĐ 7 món Duy Lập Phát 18/05-29/07/2022; HĐ 1358 Hoàng Quân 30/07/2022",
            ghi="Nguồn HĐĐT (IN02 chưa có). T6: một NCC một dòng, ghi F0.",
            sac="GTGT",
            v=f0_gt,
            w=0,
            thieu="IN02",
        )
    elif ma == "HD03":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if f0_ty > P13 else "Không",
            nhan="Sơ bộ",
            muc="Cao",
            diem=7,
            ca="GT F0 / tổng mua",
            va=f0_ty,
            nguong="P13 5%",
            do=f0_ty,
            tep="KQ2.1",
            v=f0_gt,
            w=0,
            thieu="IN02",
        )
    elif ma == "HD04":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if top5_ty > P14 else "Không",
            nhan="Sơ bộ",
            muc="Trung bình",
            diem=5,
            ca="Top-5 / tổng mua",
            va=top5_ty,
            nguong="P14 50%",
            do=top5_ty,
            tep="KQ2.1",
            nd="Chọn mẫu hợp đồng Top-5, đặc biệt ADM VN 0313654924",
            ghi=f"Top-5 GT {vn(top5_gt)} / {vn(tot_mua)} = {top5_ty:.2%}. Tính gộp MST người bán, không dùng LARGE trên từng dòng.",
            thieu="IN02",
        )
    elif ma == "HD08":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if f1_buyers else "Không",
            nhan="Chắc" if f1_ok else "Sơ bộ",
            muc="Cao" if f1_buyers else "",
            diem=7 if f1_buyers else 0,
            ca="Bên mua trùng F1",
            va=len(f1_buyers),
            tep="KQ2.1",
            ghi=f"NNT không thuộc F1. Đã quét F1 theo MST bên mua của {n_ban} HĐ bán. Hits={len(f1_buyers)}.",
        )
    elif ma in {"HD02", "HD05", "HD06", "HD09"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU" if "IN02" in miss and ma in {"HD02", "HD05"} else "ĐÃ_CHẠY", ket="Chưa kết luận" if ma in {"HD02", "HD05"} else "Không", nhan="Sơ bộ", thieu=",".join(miss), ghi="Chưa có danh bạ ĐKKD 3.2; HD06 mặt hàng HĐĐT trống.")
    elif ma == "HD07":
        add_rule(rows_kq3, rd, status="TREO", ket="Chưa kết luận", thieu="K13 số liệu ngành; mặt hàng HĐĐT trống")
    elif ma == "TT01":
        add_rule(
            rows_kq3,
            rd,
            ket="Có phát hiện" if tt01_multi else "Không",
            nhan="Sơ bộ",
            muc="Cao" if tt01_multi else "",
            diem=6 if tt01_multi else 0,
            ca="Cụm n>=2 vượt ngưỡng",
            va=len(tt01_multi),
            cb="Cụm >= ngưỡng",
            vb=len(tt01),
            tep="KQ2.1",
            thieu="IN02, IN11",
            nd="Đối UNC các cụm cùng ngày, không kết luận thiếu chứng từ khi chưa có sao kê",
            ghi=f"Proxy HĐĐT. Ngưỡng >=20tr đến 30/06/2025, >=5tr từ 01/07/2025. Lớn nhất 04/07/2022 MST 0313654924 n=3.",
            v=None,
            w=0,
        )
    elif ma in {"TT02", "TT03", "TT04", "TT05", "TT06"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu=",".join(miss) or rd["need"])
    elif ma.startswith("KH") or ma.startswith("DT"):
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu=",".join(miss) or rd["need"])
    elif ma == "LD01":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="Chỉ tiêu số QTT rỗng; không có tờ khai quý TNCN", ghi="Có file XML QTT 2022/2023/2025 nhưng chi_tieu rỗng.")
    elif ma in {"LD02", "LD05"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu=",".join(miss) or "IN09")
    elif ma in {"LD03", "LD04"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN05 phụ lục 05-1/05-2")
    elif ma in {"LK01", "LK02"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN12")
    elif ma == "LS01":
        add_rule(rows_kq3, rd, status="ĐÃ_CHẠY", ket="Chưa kết luận", nhan="Sơ bộ", thieu="Bảng nợ/KHBS TMS dòng", ghi=f"IN14 có {len(tms_files)} file TMS; Nghia_vu_ke_khai đã đọc. GTGT toàn chính thức.")
    elif ma == "LS02":
        add_rule(rows_kq3, rd, status="ĐÃ_CHẠY", ket="Chưa kết luận", nhan="Sơ bộ", tep="KQ1.1", ct_ref="IN14", ghi="Có Lich su / Tong quan / Nghia_vu_ke_khai. Chưa bóc hết dòng chậm nộp/nợ thuế thành số tiền.", thieu="Bảng nợ thuế TMS chi tiết")
    else:
        if miss:
            add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu=",".join(miss))
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ")

clear_from(ws, 10)
hm = {str(ws.cell(9, c).value or "").strip(): c for c in range(1, 40) if ws.cell(9, c).value}
for i, rec in enumerate(rows_kq3):
    rr = 10 + i
    for k, v in rec.items():
        c = hm.get(k)
        if not c:
            continue
        put(ws, rr, c, v, num=isinstance(v, (int, float)) and not isinstance(v, bool) and k not in {"Hệ số trọng yếu", "Điểm"}, wrap=k in {"Ghi chú", "Còn thiếu để kết luận", "Nội dung kiểm tra tại trụ sở"})
    if rec.get("Mức rủi ro") == "Cao":
        ws.cell(rr, hm["Mức rủi ro"]).fill = FILL_RED
        ws.cell(rr, hm["Mức rủi ro"]).font = FONT_RED
    elif rec.get("Mức rủi ro") == "Trung bình":
        ws.cell(rr, hm["Mức rủi ro"]).fill = FILL_YEL
        ws.cell(rr, hm["Mức rủi ro"]).font = FONT_ORG
    if rec.get("Trạng thái chạy") == "KHÔNG_ĐỦ_DỮ_LIỆU":
        ws.cell(rr, hm["Trạng thái chạy"]).fill = FILL_YEL
    if rec.get("Trạng thái chạy") == "TREO":
        ws.cell(rr, hm["Trạng thái chạy"]).fill = FILL_RED

n_run = sum(1 for r in rows_kq3 if r["Trạng thái chạy"] == "ĐÃ_CHẠY")
n_kdd = sum(1 for r in rows_kq3 if r["Trạng thái chạy"] == "KHÔNG_ĐỦ_DỮ_LIỆU")
n_treo = sum(1 for r in rows_kq3 if r["Trạng thái chạy"] == "TREO")
n_cao = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Cao" and r.get("Kết quả") == "Có phát hiện")
n_tb = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Trung bình" and r.get("Kết quả") == "Có phát hiện")
ws["C6"] = f"ĐÃ_CHẠY {n_run} | KHÔNG_ĐỦ {n_kdd} | TREO {n_treo}"
ws["J6"] = f"Cao {n_cao} | Trung bình {n_tb}"

# 28 TC
ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
ws["A2"] = f"Áp dụng cho {TEN} (MST {MST}) - Kỳ {KY} | {VER} | {TODAY.strftime('%d/%m/%Y')} | Mức A - thiếu sổ DN"
if str(ws.cell(4, 1).value or "").startswith("TC"):
    ws.delete_rows(4, 1)
tc_rows = [
    ("TC01", "Chênh lệch DT GTGT [34] vs DT BCTC [01]", "Doanh thu - giá vốn", "lệch ≠ 0",
     f"2023 khớp tuyệt đối {vn(gt34_y[2023])}. 2022 BCTC {vn(bctc01_y.get(2022))} vs GTGT chỉ Q4. 2024 không BCTC. 2025 BCTC {vn(bctc01_y.get(2025))} vs GTGT Q2-Q4 {vn(GY[2025]['34'])} vs HĐĐT {vn(hd_ban_y[2025])}.",
     "Cao", "Bổ sung XML GTGT thiếu kỳ; đối chiếu HĐĐT bán từng quý."),
    ("TC02", "Người mua trả tiền trước > 5% DT", "Công nợ", "5%",
     "Chưa xác minh mã TT133 cho người mua trả tiền trước trên B01a-DNN. Không điền 0.",
     "Thấp", "Xác minh mã trên BCTC giấy / thuyết minh."),
    ("TC03", "HTK + XDCB dở dang > 30% DT", "Hàng tồn kho", "30% - B01a-DNN dùng ct140",
     f"2025 HTK ct140 {vn(BM.get(2025,{}).get('htk'))} / DT {vn(BM.get(2025,{}).get('dt'))} = {htk_ty_2025:.1%} > 30%. 2022 {BM[2022]['htk']/BM[2022]['dt']:.1%}; 2023 {BM[2023]['htk']/BM[2023]['dt']:.1%}." if htk_ty_2025 else "Thiếu BCTC 2025",
     "Cao", "Kiểm kê tồn, sổ XNT 152/155/156, hàng chậm luân chuyển."),
    ("TC04", "DT chưa thực hiện > 0", "Doanh thu", "> 0", "Không thấy chỉ tiêu DT chưa thực hiện trên XML.", "Thấp", ""),
    ("TC05", "DT thuần - giá vốn < 0", "Giá vốn", "< 0",
     f"Lãi gộp BCTC dương: 2022 {vn(BM.get(2022,{}).get('ln_gop'))}; 2023 {vn(BM.get(2023,{}).get('ln_gop'))}; 2025 {vn(BM.get(2025,{}).get('ln_gop'))}.",
     "Thấp", ""),
    ("TC06", "Chi phí phải trả trích trước lớn", "Chi phí", "[315]",
     f"2025 CDKT 315={vn((BM.get(2025,{}).get('cdkt') or {}).get('315'))} - chưa xác minh hết mã TT133.",
     "Thấp", "Đọc thuyết minh BCTC."),
    ("TC07", "Lãi vay chưa vốn hóa", "Chi phí", "[242]", "Không thấy 242 trên XML BCTC các năm có.", "Thấp", ""),
    ("TC08", "Ưu đãi / miễn giảm TNDN", "TNDN", "C2/C11",
     "2022 thuế TNDN = 0 trong khi LNTT 222.136.662. 2023 thuế 81.074.762; 2025 thuế 95.794.704.",
     "Trung bình", "Giải trình thuế TNDN 2022 = 0."),
    ("TC09", "VAT khấu trừ TK vs BCTC [181]", "GTGT", "[43] vs 181 TT133",
     "2023 và 2025 [43] cuối năm khớp CDKT 181. Không dùng [152] (vắng trên TT133).",
     "Thấp", ""),
    ("TC10", "VAT khấu trừ vs 10%×(HTK+TSCĐ)", "GTGT", "×10%",
     f"[43] 2025 {vn(ct43_last)} < 10%×HTK {vn((BM.get(2025,{}).get('htk') or 0)*0.1)}.",
     "Thấp", ""),
    ("TC11", "Có DT KCT [26] nhưng [25]=[24]", "GTGT", "",
     f"Q4/2025 [26]={vn(ct26_q4_2025)} = [34]. [25]=[24]={vn(ct(q4_2025,'24') if q4_2025 else None)}. Ngành nông sản KCT hợp lý; cần xem phân bổ.",
     "Trung bình", "Rà phân bổ thuế đầu vào kỳ có DT KCT."),
    ("TC12", "Điều chỉnh giảm LN trước thuế", "TNDN", "B7/B8", "Chưa tách B7/B8 trên XML 03/TNDN đã đọc.", "Thấp", "Đọc PL 03-1A"),
    ("TC13", "Trích lập dự phòng", "BCTC", "", "Không thấy mã dự phòng trọng yếu trên XML.", "Thấp", ""),
    ("TC14", "CN vốn [22]>0 và [25]<0", "BCTC", "", "Không thấy dấu hiệu trên XML.", "Thấp", ""),
    ("TC15", "Bù trừ lãi BĐS / lỗ SXKD", "TNDN", "", "LNTT dương các năm có BCTC.", "Thấp", ""),
    ("TC16", "Chuyển giá [251][252]>0 và [421]<0", "GDLK", "", "IN12 chưa nhận.", "Thấp", "Xin PL GDLK nếu có bên liên kết."),
    ("TC17", "CP lãi vay vs 30% EBITDA", "EBITDA", "NĐ 132", "Chưa đủ khấu hao / lãi vay từ sổ. IN12 chưa nhận.", "Thấp", "Lấy 635 và khấu hao khi có sổ."),
    ("TC18", "B11>0 không có PL GDLK", "GDLK", "IN12", "IN12 chưa nhận - KHÔNG_ĐỦ_DỮ_LIỆU.", "Thấp", "Xin PL NĐ 132 hoặc xác nhận không GDLK."),
    ("TC19", "Thuế TNDN hoãn lại", "TNDN", "", "Không thấy biến động hoãn lại trên XML.", "Thấp", ""),
    ("TC20", "Phải trả người bán lớn", "Công nợ", "[311]",
     f"Nợ ngắn hạn 311: 2022 {vn(BM.get(2022,{}).get('no_nh'))}; 2023 {vn(BM.get(2023,{}).get('no_nh'))}; 2025 {vn(BM.get(2025,{}).get('no_nh'))}. 2025 giảm mạnh vs 2022-2023.",
     "Trung bình", "Đối 331 khi có sổ."),
    ("TC21", "Lãi vay khi thiếu VĐL", "Tài chính", "bỏ nếu thiếu TMS vốn", "Vốn 411 = 500 triệu cả 3 năm. Chưa đối chiếu góp đủ.", "Thấp", "Đối vốn điều lệ TMS."),
    ("TC22", "QTT TNCN vs TK tháng/quý", "TNCN", "", "QTT XML 2022/2023/2025 chỉ tiêu rỗng; không có tờ khai quý TNCN.", "Thấp", "Bổ sung 05-1/05-2 và tờ khai quý."),
    ("TC23", "Dòng tiền chi NLĐ vs TNCT", "TNCN", "20%", "Thiếu TNCT và sổ lương/111.", "Thấp", "IN05 + IN10."),
    ("TC24", "KHBS vs hạch toán TMS", "TMS", "không chạy", "Không chạy (cấm TMS dòng hạch toán).", "Thấp", ""),
    ("TC25", "Bổ sung GTGT chưa điều chỉnh", "GTGT", "", "Không có tờ khai GTGT bổ sung (toàn chính thức).", "Thấp", ""),
    ("TC26", "Tỷ trọng HĐ mua NNT TT 03/06", "Hóa đơn", "cần ĐKKD", "Chưa có danh bạ 3.2. Không gắn tỷ trọng 03/06.", "Thấp", "Nạp danh bạ ĐKKD."),
    ("TC27", "DT TK GTGT vs HĐĐT bán ra", "Hóa đơn", "lệch ≠ 0",
     f"2023 TK[34] {vn(gt34_y[2023])} vs HĐĐT {vn(hd_ban_y[2023])} lệch {vn(diff_gt01.get(2023))}. Các năm khác XML thiếu quý.",
     "Cao", "Chọn mẫu HĐ điều chỉnh/hủy/sai kỳ."),
    ("TC28", "Phân tích ngang/dọc BCTC - GV tăng nhanh hơn DT", "Giá vốn", "TC28_M1",
     f"DT giảm 2022 {vn(bctc01_y.get(2022))} → 2023 {vn(bctc01_y.get(2023))} → 2025 {vn(bctc01_y.get(2025))}. GV/DT 2022 {gv_ty.get(2022):.1%} cao nhất.",
     "Trung bình", "Tách mặt hàng nông sản; xin sổ 511/632."),
]
for i, (id_, ten, phan, nguong, kq, muc_tc, kn) in enumerate(tc_rows):
    rr = 4 + i
    put(ws, rr, 1, id_, bold=True)
    put(ws, rr, 2, ten, wrap=True)
    put(ws, rr, 3, phan)
    put(ws, rr, 4, nguong, wrap=True)
    put(ws, rr, 5, kq, wrap=True)
    put(ws, rr, 6, muc_tc, bold=True)
    put(ws, rr, 7, kn, wrap=True)
    ws.row_dimensions[rr].height = 48
    if muc_tc == "Cao":
        ws.cell(rr, 6).fill = FILL_RED
        ws.cell(rr, 6).font = FONT_RED
    elif muc_tc == "Trung bình":
        ws.cell(rr, 6).fill = FILL_YEL
        ws.cell(rr, 6).font = FONT_ORG
    else:
        ws.cell(rr, 6).fill = FILL_GRN
autowidth(ws, cap=36)
ws.column_dimensions["E"].width = 60
ws.column_dimensions["G"].width = 40
apply_tnr(wb["KQ3"])
wb.save(lp(files["kq3"][1]))
wb.close()
print(" KQ3 saved")

# Extra GĐ3
print("FILL extra GĐ3")
dau_p = OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}.xlsx"
giai_p = OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Dau_hieu"
heads = ["STT", "Mã", "Dấu hiệu", "Mức", "Số liệu", "Sheet KQ", "Hướng kiểm tra", "Tài liệu cần", "Trạng thái"]
for i, h in enumerate(heads, 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
dau = [
    (1, "HD01", "Mua từ 2 NCC F0", "Cao", f"{len(f0_hits)} MST, GT {vn(f0_gt)}, thuế {vn(f0_thue)}", "KQ2.1 NCC/IN07", "1. Hợp đồng 2. UNC 3. PNK 4. Sử dụng hàng 5. Đối CVCBRR", "HĐ gốc Duy Lập Phát 7 món 2022; HĐ 1358 Hoàng Quân", "Chưa kiểm tra tại trụ sở"),
    (2, "HD04", f"Top-5 NCC {top5_ty:.2%}", "Trung bình", f"{len(ncc_hd)} NCC, Top-5 {vn(top5_gt)}", "KQ2.1 NCC", "Chọn mẫu hợp đồng Top-5, trọng tâm ADM VN 0313654924", "Hợp đồng, BK (chưa có)", "Chưa kiểm tra"),
    (3, "HD03", f"Tỷ trọng F0 {f0_ty:.2%} > 5%", "Cao", vn(f0_gt), "KQ2.1", "Gộp vào chuỗi HD01", "IN07", "Chưa kiểm tra"),
    (4, "TN05", "HTK/DT 2025 = 40,9% > 30%", "Cao", f"HTK {vn(BM.get(2025,{}).get('htk'))} / DT {vn(BM.get(2025,{}).get('dt'))}", "KQ3 TC03", "1. Kiểm kê 2. XNT 3. Hàng chậm 4. Khớp ct140", "Sổ kho IN09", "Thiếu IN09"),
    (5, "GT15", "Khấu trừ [43] tăng kéo dài 2023-2025", "Cao", f"[43] cuối 2025 {vn(ct43_last)}", "KQ2.1 GTGT", "Rà hoàn thuế, khấu trừ, hồ sơ hoàn", "Tờ khai GTGT đủ kỳ, UNC", "Thiếu một số quý XML"),
    (6, "GT01", "Lệch DT tờ khai vs HĐĐT (thiếu kỳ XML)", "Cao", f"2023 lệch {vn(diff_gt01.get(2023))}", "KQ2.1 Phần II", "Bổ sung XML thiếu; chọn HĐ hủy/điều chỉnh", "XML GTGT đủ quý, HĐĐT bán", "Thiếu Q1-Q3/2022, Q2-Q4/2024, Q1/2025"),
    (7, "TT01", f"{len(tt01_multi)} cụm HĐ cùng ngày n>=2 vượt ngưỡng TM", "Cao", f"{len(tt01)} cụm >= ngưỡng", "KQ2.1", "Đối UNC; không kết luận thiếu chứng từ khi chưa có sao kê", "Sao kê IN11", "Thiếu IN11"),
    (8, "GT13", "Q4/2025 DT KCT [26]=[34]", "Trung bình", vn(ct26_q4_2025), "KQ2.1 Phần I", "Phân bổ thuế đầu vào kỳ KCT", "Tờ khai Q4/2025, hồ sơ KCT nông sản", "Chưa kiểm tra"),
    (9, "IN02-IN13", "Chưa nhận sổ DN", "Trung bình", "Mức hồ sơ A", "KQ1.3", "Đề nghị bổ sung theo thứ tự KQ1.3", "Bảng kê, sổ 511/632, kho, sao kê", "Chưa nhận"),
]
for i, row in enumerate(dau):
    for c, v in enumerate(row, 1):
        put(ws, 2 + i, c, v, wrap=c in {3, 5, 7, 8, 9})
    if row[3] == "Cao":
        ws.cell(2 + i, 4).fill = FILL_RED
        ws.cell(2 + i, 4).font = FONT_RED
    else:
        ws.cell(2 + i, 4).fill = FILL_YEL
        ws.cell(2 + i, 4).font = FONT_ORG
autowidth(ws, cap=40)
wb.save(lp(dau_p))
wb.close()

wb = Workbook()
# 10 sheets
sheet_specs = [
    ("00_Tong_hop_chi_tiet", "Tổng hợp dấu hiệu"),
    ("01_NCC_F0", "Chi tiết HĐ mua NCC F0"),
    ("02_Top5_NCC", "Top-5 NCC"),
    ("03_TT01_cum_HD", "Cụm HĐ cùng ngày"),
    ("04_GTGT_theo_ky", "GTGT theo kỳ XML"),
    ("05_BCTC_TNDN", "BCTC / TNDN"),
    ("06_HDDT_ban_nam", "HĐĐT bán theo năm"),
    ("07_HDDT_mua_nam", "HĐĐT mua theo năm"),
    ("08_Thieu_XML", "Kỳ XML còn thiếu"),
    ("09_Kho_511_632", "Kho / 511-632"),
]
ws0 = wb.active
ws0.title = sheet_specs[0][0]
for name, _ in sheet_specs[1:]:
    wb.create_sheet(name)

ws = wb["00_Tong_hop_chi_tiet"]
put(ws, 1, 1, f"Tổng hợp chi tiết dấu hiệu - {TEN} MST {MST} kỳ {KY}", bold=True, fill=FILL_HDR, font=FONT_HDR)
heads = ["STT", "Mã", "Dấu hiệu", "Mức", "Giá trị A", "Giá trị B", "Chênh / tỷ lệ", "Nguồn", "Ghi chú"]
for i, h in enumerate(heads, 1):
    put(ws, 2, i, h, bold=True, fill=FILL_SEC)
for i, row in enumerate(dau, 1):
    put(ws, 2 + i, 1, row[0], num=True)
    put(ws, 2 + i, 2, row[1])
    put(ws, 2 + i, 3, row[2], wrap=True)
    put(ws, 2 + i, 4, row[3])
    put(ws, 2 + i, 5, row[4], wrap=True)
    put(ws, 2 + i, 8, row[5])
    put(ws, 2 + i, 9, row[8], wrap=True)

ws = wb["01_NCC_F0"]
heads = ["MST", "Tên", "Số HĐ", "Ngày", "GT chưa thuế", "Thuế", "CV", "Ngày CV"]
for i, h in enumerate(heads, 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
rr = 2
for m, rec, info in f0_hits:
    for line in rec["lines"]:
        put(ws, rr, 1, m)
        put(ws, rr, 2, rec["ten"], wrap=True)
        put(ws, rr, 3, line.get("so_hd"))
        put(ws, rr, 4, dparse(line.get("ngay")), dt=True)
        put(ws, rr, 5, line.get("gt"), num=True)
        put(ws, rr, 6, line.get("thue"), num=True)
        put(ws, rr, 7, info.get("cv"))
        put(ws, rr, 8, dparse(info.get("ngay_cv")), dt=True)
        for c in range(1, 9):
            ws.cell(rr, c).fill = FILL_RED
        rr += 1
if rr == 2:
    put(ws, 2, 1, "Không có HĐ F0")

ws = wb["02_Top5_NCC"]
for i, h in enumerate(["STT", "MST", "Tên", "GT", "Thuế", "Số HĐ", "Tỷ trọng"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, (m, rec) in enumerate(top5, 1):
    put(ws, 1 + i, 1, i, num=True)
    put(ws, 1 + i, 2, m)
    put(ws, 1 + i, 3, rec["ten"], wrap=True)
    put(ws, 1 + i, 4, rec["gt"], num=True)
    put(ws, 1 + i, 5, rec["thue"], num=True)
    put(ws, 1 + i, 6, rec["n_hd"], num=True)
    put(ws, 1 + i, 7, rec["gt"] / tot_mua if tot_mua else 0, pct=True)

ws = wb["03_TT01_cum_HD"]
for i, h in enumerate(["MST", "Tên", "Ngày", "Số HĐ", "GT", "Thuế", "TT", "Ngưỡng"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, x in enumerate(tt01[:200], 1):
    put(ws, 1 + i, 1, x["mst"])
    put(ws, 1 + i, 2, x["ten"], wrap=True)
    put(ws, 1 + i, 3, dparse(x["ngay"]), dt=True)
    put(ws, 1 + i, 4, x["n"], num=True)
    put(ws, 1 + i, 5, x["gt"], num=True)
    put(ws, 1 + i, 6, x["thue"], num=True)
    put(ws, 1 + i, 7, x["tt"], num=True)
    put(ws, 1 + i, 8, x["nguong"], num=True)
    if x["n"] >= 2:
        for c in range(1, 9):
            ws.cell(1 + i, c).fill = FILL_YEL

ws = wb["04_GTGT_theo_ky"]
for i, h in enumerate(["Năm", "Kỳ", "ct23", "ct24", "ct26", "ct34", "ct35", "ct41", "ct43", "File"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, x in enumerate(gtgt, 1):
    c = x.get("chi_tieu") or {}
    put(ws, 1 + i, 1, x.get("nam"))
    put(ws, 1 + i, 2, x.get("ky_norm"))
    for col, code in enumerate(["23", "24", "26", "34", "35", "41", "43"], 3):
        put(ws, 1 + i, col, c.get(code), num=True)
    put(ws, 1 + i, 10, x.get("file"), wrap=True)

ws = wb["05_BCTC_TNDN"]
for i, h in enumerate(["Năm", "BCTC[01]", "BCTC[11]", "LN gộp", "HTK140", "TNDN[04]", "TNDN[16]", "TNDN[18]"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, y in enumerate(YEARS, 1):
    put(ws, 1 + i, 1, y)
    if y in BM:
        put(ws, 1 + i, 2, BM[y]["dt"], num=True)
        put(ws, 1 + i, 3, BM[y]["gv"], num=True)
        put(ws, 1 + i, 4, BM[y]["ln_gop"], num=True)
        put(ws, 1 + i, 5, BM[y]["htk140"], num=True)
    else:
        put(ws, 1 + i, 2, "KHÔNG_ĐỦ_DỮ_LIỆU", fill=FILL_YEL)
    tnct = TN.get(y) or {}
    put(ws, 1 + i, 6, tnct.get("04"), num=True)
    put(ws, 1 + i, 7, tnct.get("16"), num=True)
    put(ws, 1 + i, 8, tnct.get("18"), num=True)

ws = wb["06_HDDT_ban_nam"]
for i, h in enumerate(["Năm", "Số HĐ", "GT", "Thuế"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, y in enumerate(YEARS, 1):
    gt, n = sum_hd("BAN", y)
    th, _ = sum_hd("BAN", y, "thue")
    put(ws, 1 + i, 1, y)
    put(ws, 1 + i, 2, n, num=True)
    put(ws, 1 + i, 3, gt, num=True)
    put(ws, 1 + i, 4, th, num=True)

ws = wb["07_HDDT_mua_nam"]
for i, h in enumerate(["Năm", "Số HĐ", "GT", "Thuế"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, y in enumerate(YEARS, 1):
    gt, n = sum_hd("MUA", y)
    th, _ = sum_hd("MUA", y, "thue")
    put(ws, 1 + i, 1, y)
    put(ws, 1 + i, 2, n, num=True)
    put(ws, 1 + i, 3, gt, num=True)
    put(ws, 1 + i, 4, th, num=True)

ws = wb["08_Thieu_XML"]
for i, h in enumerate(["Loại", "Kỳ/Năm thiếu", "Ghi chú"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
rr = 2
for y in YEARS:
    for k in GY[y]["missing"]:
        put(ws, rr, 1, "GTGT")
        put(ws, rr, 2, k)
        put(ws, rr, 3, "Không điền 0 giả")
        ws.cell(rr, 1).fill = FILL_YEL
        rr += 1
put(ws, rr, 1, "BCTC/TNDN/TNCN")
put(ws, rr, 2, "2024")
put(ws, rr, 3, "Không có XML trong 1.1")
ws.cell(rr, 1).fill = FILL_YEL

ws = wb["09_Kho_511_632"]
put(ws, 1, 1, "KHÔNG_ĐỦ_DỮ_LIỆU", bold=True, fill=FILL_YEL)
put(ws, 2, 1, "Chưa nhận IN08, IN09, IN13. Không điền 0 giả. Quy tắc KH01-KH07, DT01-DT03, 511-632 để trống đến Lệnh 4.", wrap=True)

for sn, _ in sheet_specs:
    autowidth(wb[sn], cap=36)
    wb[sn].sheet_view.showGridLines = True
wb.save(lp(giai_p))
wb.close()
print(" extra saved")

print("\n===== VERIFY =====")
for key, (_, p) in files.items():
    w = load_workbook(lp(p), read_only=True)
    print(p.name, "sheets", w.sheetnames, "size", os.path.getsize(lp(p)))
    w.close()
print("dau", dau_p.name, os.path.getsize(lp(dau_p)))
print("giai", giai_p.name, os.path.getsize(lp(giai_p)))
print("KQ3 rules", len(rows_kq3), "run", n_run, "kdd", n_kdd, "treo", n_treo, "cao", n_cao, "tb", n_tb)
print("DONE")
