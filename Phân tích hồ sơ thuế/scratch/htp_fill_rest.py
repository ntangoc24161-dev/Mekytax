# -*- coding: utf-8 -*-
"""Fill KQ1.3 + KQ2.1 + KQ2.2 + KQ3 + 2 extras for HUY THANH PHAT."""
from __future__ import annotations

import json
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

MST = "3603820015"
TEN = "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ VÀ SẢN XUẤT HUY THÀNH PHÁT"
KY = "2021-2025"
YEARS = [2021, 2022, 2023, 2024, 2025]
TODAY = date(2026, 11, 9)
VER = "v1"
MAU_VER = "1.4"
LEVEL = "E"
DN_STATUS = "Một phần"
P02, P13, P14, P05, P18, P23, P03, P07 = 1000, 0.05, 0.50, 0.30, 0.92, 0.10, 0.02, 0.20
P10, P11 = 20_000_000, 5_000_000

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
BO = MAU / "Bo_bieu_mau_chuan_phan_tich_ho_so_thue.xlsx"
ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Dang kiem tra\KE HOACH\HUY_THANH_PHAT_3603820015"
)
OUT = ROOT / "4. KET QUA PHAN TICH"
SAO = ROOT / "7. Sao luu tong hop"
VAULT_KQ = VAULT / "ket-qua"
CACHE = Path(r"C:\Users\NGOC\AppData\Local\Temp\htp_cache.json")

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FONT_HDR = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
FONT_RED = Font(name="Times New Roman", size=11, color="C00000")
FONT_ORG = Font(name="Times New Roman", size=11, color="B25900")
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
NUM, PCT, DATE_FMT = "#,##0", "0.00%", "dd/mm/yyyy"


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


def tnr(cell, bold=None):
    f = cell.font
    cell.font = Font(
        name="Times New Roman",
        size=f.size if f and f.size else 11,
        bold=f.bold if bold is None and f else (bool(bold) if bold is not None else False),
        italic=False,
        color=f.color if f else None,
    )


def master(ws, r, c):
    cell = ws.cell(r, c)
    if not isinstance(cell, MergedCell):
        return cell
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            return ws.cell(mr.min_row, mr.min_col)
    return cell


def put(ws, r, c, v, num=False, dt=False, pct=False, bold=None, fill=None, wrap=False):
    cell = master(ws, r, c)
    cell.value = v
    tnr(cell, bold=bold)
    cell.border = THIN
    cell.alignment = Alignment(vertical="center", wrap_text=wrap, horizontal="center" if dt else None)
    if fill:
        cell.fill = fill
    if num and isinstance(v, (int, float)):
        cell.number_format = NUM
        cell.alignment = Alignment(vertical="center", horizontal="right", wrap_text=wrap)
    if pct and isinstance(v, (int, float)):
        cell.number_format = PCT
        cell.alignment = Alignment(vertical="center", horizontal="right", wrap_text=wrap)
    if dt and v not in (None, ""):
        cell.number_format = DATE_FMT
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=wrap)
    return cell


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


def unmerge_all(ws):
    for mr in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mr))


def clear_from(ws, start_row):
    if ws.max_row and ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)


def banner(ws):
    for row in ws.iter_rows(max_row=5, max_col=12):
        for cell in row:
            if isinstance(cell, MergedCell) or not isinstance(cell.value, str):
                continue
            v = cell.value
            v = v.replace("[MST]", MST).replace("[KY]", KY).replace("[NGAY]", TODAY.strftime("%d/%m/%Y"))
            v = v.replace("[TÊN DOANH NGHIỆP]", TEN)
            v = v.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")
            if v != cell.value:
                cell.value = v
            tnr(cell)


def flag_gt(a, b):
    if a is None and b is None:
        return None
    if a is None or b is None:
        return "Chưa đủ nguồn"
    if abs((a or 0) - (b or 0)) <= P02:
        return "Khớp"
    return "Lệch"


print("LOAD CACHE")
C = json.loads(CACHE.read_text(encoding="utf-8"))
hddt, bk, soct = C["hddt"], C["bangke"], C["soct"]
gtgt, tndn, bctc = C["gtgt_eff"], C["tndn_eff"], C["bctc_eff"]
nld_qtt = {int(x["nam"]): x for x in C["nld_qtt"]}
kho = {k["nam"]: k for k in C["kho"]}


def valid_map(rows):
    out = {}
    for x in rows:
        m = mst_base(x.get("mst"))
        if len(m) in (10, 13):
            out[m] = x
    return out


f0, f1, ca, db = valid_map(C["f0"]), valid_map(C["f1"]), valid_map(C["ca"]), valid_map(C["danhba"])
nnt_f1, nnt_f0, nnt_ca = MST in f1, MST in f0, MST in ca
ps = [x for x in soct if x["kind"] == "PS"]
sodu = [x for x in soct if x["kind"] == "SODU"]


def tk_sum(code, field, nam=None):
    s = 0.0
    for x in ps:
        if x["tk"] != code:
            continue
        if nam is not None and x.get("nam") != nam:
            continue
        s += x.get(field) or 0
    return s


def tk_year(code, field):
    return {y: tk_sum(code, field, y) for y in YEARS}


dt_511, gv_632 = tk_year("511", "ps_co"), tk_year("632", "ps_no")
n621, n622, n627 = tk_year("621", "ps_no"), tk_year("622", "ps_no"), tk_year("627", "ps_no")
n131, c131 = tk_year("131", "ps_no"), tk_year("131", "ps_co")
n331, c331 = tk_year("331", "ps_no"), tk_year("331", "ps_co")
n133 = tk_year("1331", "ps_no")
if sum(n133.values()) == 0:
    n133 = tk_year("133", "ps_no")


def sodu_ck(code, nam, side="no"):
    rows = [x for x in sodu if x.get("tk") == code and x.get("nam") == nam]
    if not rows:
        return None
    key = "du_no" if side == "no" else "du_co"
    if rows[-1].get(key) is not None:
        return rows[-1].get(key)
    return (rows[-1].get("ps_no") if side == "no" else rows[-1].get("ps_co"))


BM, TN = {}, {}
for x in bctc:
    k, cd = x.get("kqkd") or {}, x.get("cdkt_cuoi") or {}
    BM[x["nam"]] = {
        "dt": k.get("01") or 0, "gv": k.get("11") or 0, "ln_gop": k.get("20") or 0,
        "dt_tc": k.get("21") or 0, "thue_tndn": k.get("22") or 0, "cp_lai": k.get("23") or 0,
        "lnst": k.get("60") or k.get("50") or 0, "ln_tt": k.get("50") or 0,
        "htk": cd.get("141") or cd.get("140") or 0, "htk140": cd.get("140") or 0,
        "htk141": cd.get("141") or 0, "pt": cd.get("131") or 0, "tr": cd.get("311") or cd.get("331") or 0,
        "tm": cd.get("111") or 0, "nh": cd.get("112") or 0, "von": cd.get("411") or 0,
        "thue_kt": cd.get("152") or 0, "tscd": cd.get("211") or 0, "file": x.get("file"),
    }
for x in tndn:
    TN[x["nam"]] = x.get("chi_tieu") or {}


def gy():
    d = {y: {"34": 0.0, "35": 0.0, "23": 0.0, "24": 0.0, "40": 0.0, "43": 0.0, "32": 0.0, "26": 0.0, "kys": []} for y in YEARS}
    for x in gtgt:
        y = x.get("nam")
        if y not in d:
            continue
        ct = x.get("chi_tieu") or {}
        for k in ("34", "35", "23", "24", "40", "32", "26"):
            d[y][k] += ct.get(k) or 0
        d[y]["43"] = ct.get("43") or ct.get("41") or d[y]["43"]
        d[y]["kys"].append(x)
    return d


GY = gy()


def sum_hd(rows, chieu, nam=None, field="gt"):
    s, n = 0.0, 0
    for r in rows:
        if r.get("chieu") != chieu:
            continue
        if nam is not None and r.get("nam") != nam:
            continue
        s += r.get(field) or 0
        n += 1
    return s, n


bk_u, seen = [], {}
for r in bk:
    k = (r.get("chieu"), so_key(r.get("so_key")), r.get("mst_dt"), r.get("ngay"))
    if k in seen:
        i = seen[k]
        bk_u[i]["gt"] += r.get("gt") or 0
        bk_u[i]["thue"] += r.get("thue") or 0
    else:
        seen[k] = len(bk_u)
        bk_u.append(dict(r))

ncc_hd = defaultdict(lambda: {"ten": "", "gt": 0.0, "thue": 0.0, "n_hd": 0, "n_bk": 0, "src": set(), "lines": []})
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
    rec["src"].add("HĐĐT")
    rec["lines"].append(r)
for r in bk_u:
    if r["chieu"] != "MUA":
        continue
    m = mst_base(r.get("mst_dt"))
    if not m:
        continue
    rec = ncc_hd[m]
    rec["ten"] = rec["ten"] or r.get("ten") or ""
    rec["n_bk"] += 1
    rec["src"].add("Bảng kê")
    if rec["n_hd"] == 0:
        rec["gt"] += r.get("gt") or 0
        rec["thue"] += r.get("thue") or 0

tot_mua_hddt = sum(r["gt"] for r in hddt if r["chieu"] == "MUA")
top5 = sorted(ncc_hd.items(), key=lambda x: -x[1]["gt"])[:5]
top5_gt = sum(v["gt"] for _, v in top5)
top5_ty = (top5_gt / tot_mua_hddt) if tot_mua_hddt else 0

f0_hits = []
for m, rec in ncc_hd.items():
    if m in f0:
        f0_hits.append((m, rec, f0[m], "F0", "CANH_BAO"))
    elif m in ca:
        f0_hits.append((m, rec, ca[m], "CA", "CA"))
f0_gt = sum(r["gt"] for _, r, _, _, _ in f0_hits)
f0_ty = (f0_gt / tot_mua_hddt) if tot_mua_hddt else 0


def ncc_flag(m):
    if m in f0:
        inf = f0[m]
        return "F0", "CANH_BAO", inf.get("cv") or "", inf.get("nd") or "", inf
    if m in ca:
        inf = ca[m]
        return "F0", "CA", inf.get("cv") or "", inf.get("nd") or "", inf
    if m in f1:
        inf = f1[m]
        return "F1", "CANH_BAO", inf.get("cv") or "", inf.get("nd") or "", inf
    return "Không", "", "", "", db.get(m)


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
    nguong = P11 if d and d >= date(2025, 7, 1) else P10
    if v["tt"] >= nguong or v["gt"] >= nguong:
        tt01.append({"mst": mst, "ngay": ngay, "n": v["n"], "gt": v["gt"], "thue": v["thue"], "tt": v["tt"], "ten": v["ten"], "nguong": nguong})
tt01.sort(key=lambda x: -x["tt"])

lines_511 = [x for x in ps if x["tk"] == "511"]
lines_632 = [x for x in ps if x["tk"] == "632"]
match_y = {y: {"khop": 0, "lech": 0, "no632": 0, "no511": 0, "duoi": 0, "duoi_dt": 0, "duoi_rows": []} for y in YEARS}
used_632 = set()
for a in lines_511:
    y = a.get("nam")
    if y not in match_y:
        continue
    key = (a.get("ngay_ct"), so_key(a.get("so_ct")))
    hit = None
    for i, b in enumerate(lines_632):
        if i in used_632 or b.get("nam") != y:
            continue
        if (b.get("ngay_ct"), so_key(b.get("so_ct"))) == key:
            hit = i
            break
    if hit is None:
        match_y[y]["no632"] += 1
        continue
    used_632.add(hit)
    b = lines_632[hit]
    lai = (a.get("ps_co") or 0) - (b.get("ps_no") or 0)
    if lai <= 0:
        match_y[y]["duoi"] += 1
        match_y[y]["duoi_dt"] += a.get("ps_co") or 0
        match_y[y]["lech"] += 1
        match_y[y]["duoi_rows"].append((a, b, lai))
    else:
        match_y[y]["khop"] += 1
for i, b in enumerate(lines_632):
    if i not in used_632 and b.get("nam") in match_y:
        match_y[b["nam"]]["no511"] += 1

chi_lines = [x for x in ps if x["tk"] in {"1111", "111"} and (x.get("ps_co") or 0) > 0]
seen_chi, chi_ok = set(), []
for x in chi_lines:
    k = (x.get("nam"), x.get("so_ct"), x.get("ngay_ct"), round(x.get("ps_co") or 0))
    if k in seen_chi:
        continue
    seen_chi.add(k)
    chi_ok.append(x)
cash_over = []
for x in chi_ok:
    d = dparse(x.get("ngay_ct"))
    nguong = P11 if d and d >= date(2025, 7, 1) else P10
    if (x.get("ps_co") or 0) >= nguong:
        cash_over.append({**x, "nguong": nguong})
cash_over.sort(key=lambda x: -(x.get("ps_co") or 0))

gt_tot = {y: {"34": GY[y]["34"], "hddt": sum_hd(hddt, "BAN", y)[0], "bk": sum_hd(bk_u, "BAN", y)[0],
              "23": GY[y]["23"], "hddt_m": sum_hd(hddt, "MUA", y)[0], "bk_m": sum_hd(bk_u, "MUA", y)[0],
              "hddt_thue_b": sum_hd(hddt, "BAN", y, "thue")[0], "bk_thue_b": sum_hd(bk_u, "BAN", y, "thue")[0],
              "hddt_thue_m": sum_hd(hddt, "MUA", y, "thue")[0], "bk_thue_m": sum_hd(bk_u, "MUA", y, "thue")[0],
              "35": GY[y]["35"], "24": GY[y]["24"]} for y in YEARS}

print("f0_hits", len(f0_hits), "f0_ty", round(f0_ty, 4), "top5", round(top5_ty, 4), "tt01", len(tt01), "cash", len(cash_over))
print("CA", [m for m, *_ in f0_hits])

files = {
    "kq13": OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx",
    "kq21": OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_{VER}.xlsx",
    "kq22": OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_{VER}.xlsx",
    "kq3": OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_{VER}.xlsx",
    "dau": OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}_{VER}.xlsx",
    "giai": OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}_{VER}.xlsx",
}

# ---------------------------------------------------------------------------
# KQ1.3
# ---------------------------------------------------------------------------
print("FILL KQ1.3")
wb = load_workbook(lp(files["kq13"]))
ws = wb["KQ1_Kiem_ke"]
put(ws, 2, 3, MST, bold=True)
put(ws, 2, 10, TEN, bold=True)
put(ws, 3, 3, KY)
put(ws, 3, 10, LEVEL)
put(ws, 4, 3, "Đủ 8 | Mỏng 4 | Thiếu 1 | Chưa nhận 1")
put(ws, 4, 10, TODAY, dt=True)
put(ws, 5, 3, VER)
put(ws, 5, 10, MAU_VER)
put(ws, 6, 3, TODAY, dt=True)
put(ws, 6, 10, DN_STATUS)
put(ws, 7, 3, "KQ2.1, KQ2.2, KQ3")
put(ws, 7, 10, "Kỳ suy ra K12: GTGT Q1/2022–Q3/2025, TNDN/BCTC 2021-2024, TNDN 2025, sổ DN 2021-2025. NNT không có trên F1/F0/CA.", wrap=True)

def set_in(ws, row, *vals):
    for i, v in enumerate(vals, 6):
        put(ws, row, i, v, num=isinstance(v, (int, float)))

set_in(ws, 10, "Excel", "2022-2025", "2021 HĐĐT mỏng", len(hddt), 0, "Không", "Đủ", "Portal HĐĐT mua 614 + bán 61", "Đã chuẩn hóa", "MST, khóa HĐ", 0, "CQT")
set_in(ws, 11, "Excel", "2021-2025", "Bán ra thiếu 2022, 2023 Q1-Q3, 2025 Q3", len(bk_u), sum(1 for x in bk_u if not x.get("mst_dt")), "Không", "Mỏng", "Có bảng kê gốc; bán ra 2024 lệch gấp ~2 lần tờ khai/HĐĐT", "Đã gộp, lọc Tổng cộng", "BT6 MST", 0, "DN")
set_in(ws, 12, "XML", "Q1/2022-Q3/2025", "2021 GTGT mỏng; Q4/2025 chưa có", len(gtgt), 0, "Có (Q2/2022)", "Đủ", "15 kỳ hiệu lực", "Đã chọn bản hiệu lực", "Bổ sung ưu tiên hơn chính thức", 0, "CQT")
set_in(ws, 13, "XML", "2021-2025", "BCTC 2025 chưa có XML", len(bctc) + len(tndn), 0, "Có TNDN 2025 L01", "Đủ", "BCTC 2021-2024; TNDN 2021-2025", "Đã trích chỉ tiêu", "Đọc mã XML", 0, "CQT")
set_in(ws, 14, "XML/Excel", "2021-2025", "Đa nguồn 3.3 mỏng", len(C.get("nld_ds") or []), 0, "Không", "Đủ", "QTT 5 năm; 113 dòng NLĐ", "Đã nạp 1.4", "QTT trước NLĐ", 0, "CQT")
set_in(ws, 15, "Excel sổ CVCBRR+CA", "cập nhật 09/11/2026", "", len(f0_hits), 0, "Không", "Đủ", "Nguồn chỉ định F0/F1/CA", "Đã khớp MST 10/13 số", "NNT không thuộc F1", 0, "CQT")
set_in(ws, 16, "Excel", "2021-2025", "2022 phát sinh rất mỏng", len(ps), 0, "Không", "Đủ", "Sổ chi tiết 8.464 dòng PS", "Đã gom theo TK/tháng", "Không cộng 111+1111", 0, "DN")
set_in(ws, 17, "Excel", "2021,2023-2025", "Thiếu XNT 2022", len(kho), 0, "Không", "Mỏng", "Có tồn kho + bảng giá thành + PDF định mức", "Đã cộng XNT", "SX thép/gỗ", 0, "DN")
set_in(ws, 18, "Excel (từ sổ 111)", "2021-2025", "Không có sổ 111 tách file", len(chi_ok), 0, "Không", "Đủ", "Chi TK 1111 lấy từ sổ chi tiết", "Đã lọc trùng 111/1111", "K5 20tr/5tr", 0, "DN")
set_in(ws, 19, "Excel Eximbank", "2023.Q4-2025.Q4", "Thiếu 2021–2023.Q3", 9, 0, "Không", "Mỏng", "9 file LichSuTaiKhoan", "Đã kê file", "SpreadsheetML chưa bóc dòng", 0, "DN")
set_in(ws, 20, "", "", "Chưa nhận PL GDLK", 0, 0, "Không", "Chưa nhận", "Không thấy phụ lục NĐ 132", "Để trống", "Không điền 0", 0, "DN")
set_in(ws, 21, "Excel", "2021-2025", "2022 không có DT 511", len(lines_511) + len(lines_632), 0, "Không", "Đủ", "511 và 632 từ sổ chi tiết", "Đã tách TK đúng", "Khớp BCTC [01]/[11] các năm có số", 0, "DN")
set_in(ws, 22, "Excel TMS PTDN", "2016-2025", "Không có bảng nợ thuế/chậm nộp", 2, 0, "Không", "Mỏng", "PTDN01 CĐKT, PTDN02 KQKD", "Nhận file", "LS01/LS02 chưa đủ", 0, "CQT")

ws = wb["KQ1_Kha_nang_phan_tich"]
put(ws, 2, 3, MST, bold=True)
put(ws, 2, 7, KY)
put(ws, 4, 3, VER)
put(ws, 4, 7, MAU_VER)
put(ws, 5, 3, TODAY, dt=True)
put(ws, 5, 7, "Mức E: trống IN11/IN09/IN12/IN14")
g3_part = {"GT16", "TT02", "TT04", "TT06"}
kho_part = {"KH01", "KH02", "KH03", "KH04", "KH05", "KH06", "KH07", "TN04"}
run_now = part = block = 0
for r in range(10, (ws.max_row or 10) + 1):
    ma = str(ws.cell(r, 1).value or "").strip()
    if not ma or len(ma) > 6:
        continue
    if ma == "HD07":
        put(ws, r, 8, "TREO"); put(ws, r, 9, "Thiếu tham chiếu ngành (K13)"); put(ws, r, 10, "K13"); put(ws, r, 11, "Chạy sơ bộ khi có số ngành"); block += 1
    elif ma in {"LS01", "LS02"}:
        put(ws, r, 8, "Chạy một phần"); put(ws, r, 9, "TMS PTDN có file nhưng chưa có nợ thuế"); put(ws, r, 10, "IN14"); put(ws, r, 11, "Xuất nghĩa vụ nộp thuế TMS"); part += 1
    elif ma in {"LK01", "LK02"}:
        put(ws, r, 8, "Chạy được ngay"); put(ws, r, 9, ""); put(ws, r, 10, "IN12"); put(ws, r, 11, "Không thấy PL GDLK"); run_now += 1
    elif ma in g3_part:
        put(ws, r, 8, "Chạy một phần"); put(ws, r, 9, "Sao kê Excel 2023.Q4-2025.Q4, thiếu trước đó"); put(ws, r, 10, "IN11"); put(ws, r, 11, "Xin sao kê 2021-2023.Q3"); part += 1
    elif ma in kho_part:
        put(ws, r, 8, "Chạy một phần"); put(ws, r, 9, "Thiếu XNT 2022"); put(ws, r, 10, "IN09"); put(ws, r, 11, "Xin sổ kho 2022"); part += 1
    else:
        put(ws, r, 8, "Chạy được ngay"); put(ws, r, 9, ""); put(ws, r, 10, ""); put(ws, r, 11, ""); run_now += 1
put(ws, 3, 3, run_now)
put(ws, 3, 7, f"{part} | {block}")

ws = wb["KQ1_De_nghi_bo_sung"]
put(ws, 2, 3, MST, bold=True)
put(ws, 2, 7, KY)
put(ws, 3, 3, VER)
put(ws, 3, 7, TODAY, dt=True)
sug = [
    (1, "Sổ XNT kho năm 2022 (TK 152/155/156)", "IN09", "2. HO SO NNT", "File Excel XNT 2022", 7, "KH01-KH07, TN04", "Thiếu đúng 1 năm giữa chuỗi 2021-2025"),
    (2, "Sao kê NH Excel/PDF 2021 đến 2023.Q3 và UNC các món >= 20 triệu", "IN11", "2. HO SO NNT", "Sao kê + UNC", 4, "GT16, TT02, TT04, TT06", "Hiện chỉ có Eximbank 2023.Q4-2025.Q4"),
    (3, "Phụ lục GDLK NĐ 132 hoặc xác nhận không phát sinh", "IN12", "2. HO SO NNT", "PL-01 hoặc CV không PS", 2, "LK01, LK02", "Chưa thấy PL trên hồ sơ"),
    (4, "Bảng kê bán ra các kỳ thiếu: 2022, 2023 Q1-Q3, 2025 Q3", "IN02", "2.5", "Bảng kê gốc từng quý", 2, "GT02, GT03", "Hiện dựa HĐĐT + tờ khai"),
    (5, "File TMS nghĩa vụ nộp thuế / nợ / chậm nộp (không chỉ PTDN)", "IN14", "1.3", "Excel nợ thuế TMS", 2, "LS01, LS02", "PTDN01/02 không thay bảng nợ thuế"),
    (6, "Giải trình lệch HĐĐT bán ra vs tờ khai GTGT 2024 (~977 triệu) và 2025 (~270 triệu)", "IN01/IN03", "1.1 + 1.2", "HĐ hủy/thay thế, cut-off", 2, "GT01, TC27", "Chênh lệch nguồn CQT"),
    (7, "Giải trình bảng kê bán ra 2024 gấp ~2 lần tờ khai/HĐĐT", "IN02", "2.5", "Đối chiếu dòng HĐ", 1, "GT02", "Tổng unique HĐ không phải trùng parse"),
    (8, "Giải trình HTK BCTC vs sổ kho 2023-2024 và thuế TNDN 2024 = 279.130 đ / LNTT 1,936 tỷ", "IN04/IN09", "1.1 + 2", "Bảng kê HTK, quyết toán", 2, "KH04, TN19", "Lệch tồn kho và thuế suất hiệu dụng"),
]
for i, rowv in enumerate(sug):
    for c, v in enumerate(rowv, 1):
        put(ws, 10 + i, c, v, bold=(c == 1))
autowidth(ws)
wb.save(lp(files["kq13"]))
wb.close()
print(" KQ1.3 saved")

# ---------------------------------------------------------------------------
# KQ2.1
# ---------------------------------------------------------------------------
print("FILL KQ2.1")
wb = load_workbook(lp(files["kq21"]))
for s in wb.worksheets:
    banner(s)

ws = wb["IN07_NCC_CanhBao"]
clear_from(ws, 3)
r = 3
put(ws, r, 1, MST); put(ws, r, 2, TEN); put(ws, r, 3, "NNT đang kiểm tra")
put(ws, r, 4, "Không"); put(ws, r, 5, ""); put(ws, r, 9, "Không có trên F0/F1/CA")
put(ws, r, 10, "Đang hoạt động"); put(ws, r, 16, "Sổ CVCBRR F0/F1 + sổ CA"); put(ws, r, 18, C.get("cvcbrr_mtime"))
r = 4
for m, rec, inf, nhom, nguon in f0_hits:
    put(ws, r, 1, m); put(ws, r, 2, rec["ten"] or inf.get("ten"))
    put(ws, r, 3, nhom); put(ws, r, 4, "F0"); put(ws, r, 5, nguon)
    put(ws, r, 6, inf.get("cv")); put(ws, r, 7, dparse(inf.get("ngay_cv")), dt=True)
    put(ws, r, 9, inf.get("nd") or inf.get("sheet"))
    put(ws, r, 10, inf.get("sheet") or "")
    put(ws, r, 16, "Sổ CA" if nguon == "CA" else "Sổ CVCBRR")
    put(ws, r, 18, inf.get("sheet"))
    for c in range(1, 11):
        master(ws, r, c).fill = FILL_RED
    r += 1

ws = wb["KQ2_NCC_mua_vao"]
unmerge_all(ws)
banner(ws)
if ws.max_row >= 8:
    ws.delete_rows(8, ws.max_row - 7)
put(ws, 5, 2, f0_ty, pct=True)
put(ws, 5, 4, "Vượt ngưỡng" if f0_ty > P13 else "Không vượt", fill=FILL_RED if f0_ty > P13 else FILL_GRN)
put(ws, 5, 8, "Vượt ngưỡng" if top5_ty > P14 else "Không vượt", fill=FILL_RED if top5_ty > P14 else FILL_GRN)
invoices = [r for r in hddt if r["chieu"] == "MUA"]
for i, it in enumerate(invoices, 8):
    m = mst_base(it.get("mst_dt"))
    nhom, nguon, cv, nd, inf = ncc_flag(m)
    src = "HĐĐT"
    k = (so_key(it.get("so_key")), m)
    if any((so_key(b.get("so_key")), mst_base(b.get("mst_dt"))) == k for b in bk_u if b["chieu"] == "MUA"):
        src = "Chung"
    fill = FILL_RED if nhom in {"F0", "F1"} else None
    put(ws, i, 1, i - 7, num=True)
    put(ws, i, 2, m)
    put(ws, i, 3, it.get("ten_nb") or (inf or {}).get("ten") or "")
    put(ws, i, 4, it.get("so_hd"))
    put(ws, i, 5, dparse(it.get("ngay")), dt=True)
    put(ws, i, 6, "")
    put(ws, i, 7, it.get("gt") or 0, num=True)
    put(ws, i, 8, it.get("thue") or 0, num=True)
    put(ws, i, 9, src)
    put(ws, i, 10, f"Có - {cv}" if cv else "Không")
    put(ws, i, 11, nhom, fill=fill)
    put(ws, i, 12, (inf or {}).get("sheet") or (db.get(m) or {}).get("ten") and "Có trong danh bạ" or "Không rõ")
last = 7 + len(invoices)
put(ws, last + 2, 1, "NHẬN XÉT MODULE (không phải kết luận KQ3)", bold=True, fill=FILL_YEL)
put(
    ws, last + 3, 1,
    f"Phần 1 - NNT {MST} trên F1: Không. NNT trên F0/CA: Không. "
    f"HD03 tỷ trọng F0/CA = {f0_ty:.2%} (P13=5%, {'vượt' if f0_ty>P13 else 'không vượt'}). "
    f"HD04 Top-5 = {top5_ty:.2%} (P14=50%, {'vượt' if top5_ty>P14 else 'không vượt'}). "
    f"1 NCC dính sổ CA: 0104567918. Tổng NCC union = {len(ncc_hd)}. "
    f"Nguồn CVCBRR {str(C.get('cvcbrr_mtime') or '')[:10]}, CA {str(C.get('ca_mtime') or '')[:10]} (trong 30 ngày).",
    wrap=True, fill=FILL_YEL,
)
last_inv = last
# fix Tong_hop_co HD01 to col K, ranges
th = wb["KQ2_Tong_hop_co"]
th["F12"] = f'=(SUMPRODUCT(--(LEN(KQ2_NCC_mua_vao!$K$8:$K${last_inv})>0)))-(COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"Chưa đủ nguồn"))'
th["G12"] = f'=COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"F0")+COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"F1")'
th["H12"] = f'=COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"Chưa đủ nguồn")'
th["F13"] = f'=(SUMPRODUCT(--(LEN(KQ2_NCC_mua_vao!$I$8:$I${last_inv})>0)))-(COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"Chưa đủ nguồn"))'
th["G13"] = f'=COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"Bảng kê")+COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"HĐĐT")'
th["H13"] = f'=COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"Chưa đủ nguồn")'

ws = wb["KQ2_Doi_chieu_GTGT"]
unmerge_all(ws)
banner(ws)
put(ws, 2, 1, f"Người nộp thuế: {TEN} - MST: {MST}")
put(ws, 3, 1, f"Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Kỳ: {KY} | Kê khai: theo quý")
year_rows = {}
for rr in range(1, 150):
    a = ws.cell(rr, 1).value
    if isinstance(a, str) and a.strip().startswith("Năm "):
        try:
            year_rows[int(a.strip().split()[1])] = rr
        except (IndexError, ValueError):
            pass
period_offset = {"1": 1, "2": 2, "3": 3, "4": 4, "Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
for x in gtgt:
    y, ky = x.get("nam"), str(x.get("ky") or "")
    base = year_rows.get(y)
    if not base:
        continue
    q = ky.split("/")[0]
    off = period_offset.get(q) or period_offset.get("Q" + q)
    if not off:
        continue
    rr = base + off
    ct = x.get("chi_tieu") or {}
    put(ws, rr, 5, ct.get("34") or 0, num=True)
    put(ws, rr, 9, ct.get("35") or 0, num=True)
    put(ws, rr, 11, ct.get("23") or 0, num=True)
    put(ws, rr, 12, ct.get("24") or 0, num=True)
    put(ws, rr, 17, ct.get("40") or 0, num=True)
    put(ws, rr, 18, ct.get("43") or ct.get("41") or 0, num=True)
# Phần II years 156-165 = 2016-2025
for y in YEARS:
    r = 156 + (y - 2016)
    g = gt_tot[y]
    put(ws, r, 1, y)
    put(ws, r, 2, g["hddt"] or None, num=True)
    put(ws, r, 3, g["hddt_thue_b"] or None, num=True)
    put(ws, r, 4, g["bk"] or None, num=True)
    put(ws, r, 5, g["bk_thue_b"] or None, num=True)
    put(ws, r, 6, g["34"] if y >= 2022 else None, num=True)
    put(ws, r, 7, g["35"] if y >= 2022 else None, num=True)
    put(ws, r, 8, (TN.get(y) or {}).get("04"), num=True)
    put(ws, r, 9, (BM.get(y) or {}).get("dt"), num=True)
    fl = flag_gt(g["34"] if y >= 2022 else None, g["hddt"] or None)
    if y >= 2022:
        put(ws, r, 10, fl, fill=FILL_RED if fl == "Lệch" else FILL_GRN if fl == "Khớp" else FILL_YEL)
    put(ws, r, 13, g["hddt_m"] or None, num=True)
    put(ws, r, 14, g["bk_m"] or None, num=True)
    put(ws, r, 15, g["23"] if y >= 2022 else None, num=True)
    put(ws, r, 19, g["hddt_thue_m"] or None, num=True)
    put(ws, r, 20, g["bk_thue_m"] or None, num=True)
    put(ws, r, 21, g["24"] if y >= 2022 else None, num=True)
# Q19 delete Txx then years 2016-2020
for r in range(149, 8, -1):
    a = str(ws.cell(r, 1).value or "").strip()
    if re.match(r"^T(0[5-9]|1[0-2])$", a):
        ws.delete_rows(r, 1)
year_rows = []
for r in range(1, min(ws.max_row or 1, 160)):
    a = ws.cell(r, 1).value
    if isinstance(a, str) and a.strip().startswith("Năm "):
        try:
            y = int(a.strip().split()[1])
            year_rows.append((r, y))
        except (IndexError, ValueError):
            pass
for r, y in sorted(year_rows, reverse=True):
    if 2021 <= y <= 2025:
        continue
    end = r
    for rr in range(r + 1, (ws.max_row or r) + 1):
        a = ws.cell(rr, 1).value
        if isinstance(a, str) and (a.startswith("Năm ") or a.startswith("PHẦN") or a.startswith("Hướng") or a.startswith("NGUỒN")):
            end = rr - 1
            break
        end = rr
    ws.delete_rows(r, max(1, end - r + 1))
# Phần II leftover 2016-2020 year numbers
for r in range(1, (ws.max_row or 1) + 1):
    a = ws.cell(r, 1).value
    if a in (2016, 2017, 2018, 2019, 2020):
        ws.delete_rows(r, 1)
p2_start = None
for r in range(1, (ws.max_row or 1) + 1):
    a = ws.cell(r, 1).value
    if a == 2021:
        p2_start = r
        break
if p2_start:
    p2_end = p2_start + 4
    th["F8"] = f'=(SUMPRODUCT(--(LEN(KQ2_Doi_chieu_GTGT!$J${p2_start}:$J${p2_end})>0)))-(COUNTIF(KQ2_Doi_chieu_GTGT!$J${p2_start}:$J${p2_end},"Chưa đủ nguồn"))'
    th["G8"] = f'=COUNTIF(KQ2_Doi_chieu_GTGT!$J${p2_start}:$J${p2_end},"Lệch")'
    th["H8"] = f'=COUNTIF(KQ2_Doi_chieu_GTGT!$J${p2_start}:$J${p2_end},"Chưa đủ nguồn")'
last_r = (ws.max_row or 1) + 2
put(ws, last_r, 1, "NHẬN XÉT / NGUYÊN NHÂN TRỌNG YẾU", bold=True, fill=FILL_YEL)
put(
    ws, last_r + 1, 1,
    "NNT kê khai GTGT theo quý. Đã xóa năm 2016-2020 ngoài kỳ và dòng T05-T12 (không kê khai tháng). "
    "2023 DT [34] khớp HĐĐT và BCTC [01] = 4.098.246.332. "
    "2024 [34]=22.700.193.753 vs HĐĐT bán 23.677.198.783 (lệch 977.005.030) vs BCTC/511 22.695.878.685. "
    "Bảng kê bán ra 2024 unique HĐ khoảng gấp 2 lần tờ khai (GT02). "
    "2025 Q1-Q3 [34]=6.495.442.200 vs HĐĐT 6.765.682.200 vs 511 6.634.562.200. Q4/2025 GTGT chưa có.",
    wrap=True, fill=FILL_YEL,
)
ws.column_dimensions["A"].width = 22

ws = wb["KQ2_TNCN_LD2"]
banner(ws)
for i, y in enumerate(YEARS):
    r = 8 + i
    q = nld_qtt.get(y)
    put(ws, r, 1, y)
    if q:
        put(ws, r, 2, q.get("so_ld"), num=True)
        put(ws, r, 4, q.get("tnct"), num=True)
        put(ws, r, 5, q.get("thue_kt") or q.get("thue_phai"), num=True)
put(ws, 21, 1, "NHẬN XÉT MODULE: QTT 2021-2025 đã nạp. 2024 LD=53 TNCT=3.720.760.500 thuế phải=3.669.317. "
    "2025 LD=30 TNCT=2.644.434.498 thuế KT=40.000.000. Đa nguồn 3.3 mỏng → LD04 sơ bộ. "
    "Không thấy 05-2 tách rõ trên XML QTT.", wrap=True, fill=FILL_YEL)
autowidth(ws, cap=28)
wb.save(lp(files["kq21"]))
wb.close()
print(" KQ2.1 saved")

# ---------------------------------------------------------------------------
# KQ2.2
# ---------------------------------------------------------------------------
print("FILL KQ2.2")
wb = load_workbook(lp(files["kq22"]))
for s in wb.worksheets:
    banner(s)
ws = wb["KQ2_511_632"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | {TODAY.strftime('%d/%m/%Y')} | {VER} | DN SX (gỗ 2021, thép/neo đá 2023-2024)")
core = [2021, 2022, 2023, 2024]
put(ws, 5, 2, "Tổng 2021-2025")
for i, y in enumerate(core):
    put(ws, 5, 3 + i, y, bold=True)
put(ws, 5, 7, "2025 + đánh giá")
tot511 = sum(dt_511.values())
tot632 = sum(gv_632.values())
put(ws, 6, 2, tot511, num=True)
put(ws, 7, 2, tot632, num=True)
put(ws, 8, 2, tot511 - tot632, num=True)
put(ws, 9, 2, (tot511 - tot632) / tot511 if tot511 else None, pct=True)
for i, y in enumerate(core):
    c = 3 + i
    put(ws, 6, c, dt_511[y], num=True)
    put(ws, 7, c, gv_632[y], num=True)
    lai = dt_511[y] - gv_632[y]
    put(ws, 8, c, lai, num=True)
    put(ws, 9, c, lai / dt_511[y] if dt_511[y] else None, pct=True)
    put(ws, 10, c, match_y[y]["khop"], num=True)
    put(ws, 13, c, match_y[y]["no632"], num=True)
    put(ws, 14, c, match_y[y]["no511"], num=True)
    put(ws, 15, c, match_y[y]["duoi"], num=True)
    ty = match_y[y]["duoi_dt"] / dt_511[y] if dt_511[y] else None
    put(ws, 16, c, ty, pct=True)
put(ws, 6, 7, dt_511[2025], num=True)
put(ws, 7, 7, gv_632[2025], num=True)
put(ws, 8, 7, dt_511[2025] - gv_632[2025], num=True)
put(ws, 9, 7, (dt_511[2025] - gv_632[2025]) / dt_511[2025] if dt_511[2025] else None, pct=True)
note = (
    "511 lấy đúng TK 511 (không cộng 5111/5112/5113). 2021-2024 khớp BCTC [01]. "
    f"2025 511={dt_511[2025]:,.0f} (chưa có BCTC XML). "
    f"Bán dưới giá: {sum(match_y[y]['duoi'] for y in YEARS)} dòng. "
    "2022 DT=0 khớp BCTC. Biên lãi 2021 1,4%; 2023 2,6% (sổ) / 6,5% (BCTC GV lệch); 2024 11,5%; 2025 3,8%."
).replace(",", ".")
put(ws, 6, 8, note, wrap=True, fill=FILL_YEL)

ws = wb["KQ2_Kho_Gia_thanh"]
banner(ws)
rr = 7
for y in YEARS:
    k = kho.get(y)
    if not k:
        continue
    put(ws, rr, 1, y)
    put(ws, rr, 2, k["dk_sl"], num=True)
    put(ws, rr, 3, k["dk_tt"], num=True)
    put(ws, rr, 4, k["nhap_sl"], num=True)
    put(ws, rr, 5, k["nhap_tt"], num=True)
    put(ws, rr, 6, k["xuat_sl"], num=True)
    put(ws, rr, 7, k["xuat_tt"], num=True)
    put(ws, rr, 8, k["ck_sl"], num=True)
    put(ws, rr, 9, k["ck_tt"], num=True)
    put(ws, rr, 12, "Tổng hợp tồn kho (không tách 152/155); thiếu 2022")
    rr += 1
# Phần 3 row 34 formulas follow A7 - fill comments at 45
put(ws, 45, 1,
    "NHẬN XÉT MODULE: 2022 không có XNT (KHÔNG_ĐỦ_DỮ_LIỆU). "
    "2023 CK kho 1.606.411.182 vs BCTC HTK [141]=821.334.499 (lệch). "
    "2024 CK kho 6.556.251.416 vs BCTC HTK 3.633.866.728 (lệch). "
    "Mẫu BCTC đọc XML; 141=140 cùng số → dùng [141] (K6). Xuất kho 2024 65,25 tỷ vs 632 20,08 tỷ (kho gồm NVL/TP, không chỉ giá vốn).",
    wrap=True, fill=FILL_YEL)

ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
banner(ws)
rr = 7
for y in [2021, 2023, 2024]:
    put(ws, rr, 1, y)
    put(ws, rr, 3, n621.get(y) or None, num=True)
    put(ws, rr, 5, n622.get(y) or None, num=True)
    put(ws, rr, 7, n627.get(y) or None, num=True)
    tot = (n621.get(y) or 0) + (n622.get(y) or 0) + (n627.get(y) or 0)
    put(ws, rr, 8, tot or None, num=True)
    put(ws, rr, 11, f"Bảng tính giá thành {y} có file; 2022/2025 thiếu bảng")
    rr += 1
rr = 21
for y in YEARS:
    if not dt_511.get(y) and y not in BM:
        continue
    put(ws, rr, 1, y)
    put(ws, rr, 2, gv_632.get(y) or None, num=True)
    put(ws, rr, 3, (BM.get(y) or {}).get("gv"), num=True)
    dt = (BM.get(y) or {}).get("dt") or dt_511.get(y)
    gv = (BM.get(y) or {}).get("gv") or gv_632.get(y)
    if dt:
        put(ws, rr, 5, (gv or 0) / dt, pct=True)
        put(ws, rr, 9, "Vượt ngưỡng" if dt and (gv or 0) / dt > P18 else "Không vượt")
    rr += 1
put(ws, 32, 1,
    "NHẬN XÉT: 2021 giá thành tổng 1.141.534.875 = 632. 2023 bảng GT 3.718.999.785 vs 632 3.992.714.785 vs BCTC [11] 3.832.364.999. "
    "2024 bảng GT 18.588.811.391 vs 632/BCTC 20.078.537.359. Có PDF định mức thép. TC28 sơ bộ, chưa so ngành (K13).",
    wrap=True, fill=FILL_YEL)

ws = wb["KQ2_Cong_no_Tien_mat"]
banner(ws)
rr = 8
for y in YEARS:
    put(ws, rr, 1, y)
    put(ws, rr, 3, n131.get(y) or None, num=True)
    put(ws, rr, 4, c131.get(y) or None, num=True)
    put(ws, rr, 7, c331.get(y) or None, num=True)
    put(ws, rr, 8, n331.get(y) or None, num=True)
    dt = (BM.get(y) or {}).get("dt")
    put(ws, rr, 11, dt, num=True)
    tr = (BM.get(y) or {}).get("tr") or 0
    if dt:
        ratio = tr / dt if dt else 0
        put(ws, rr, 10, "Vượt ngưỡng" if ratio > P23 else "Không vượt", fill=FILL_RED if ratio > P23 else FILL_GRN)
        put(ws, rr, 12, ratio, pct=True)
    rr += 1
# bank files
bank_files = [m for m in C["manifest"] if m.get("ma_in") == "IN11" and m.get("ext") == ".xlsx"]
rr = 21
for i, m in enumerate(bank_files[:8], 1):
    put(ws, rr, 1, i)
    put(ws, rr, 2, "Eximbank")
    put(ws, rr, 3, m.get("name"))
    put(ws, rr, 6, "Excel")
    put(ws, rr, 7, "Đã nạp danh mục - dòng chi tiết SpreadsheetML chưa kết luận UNC (K11)")
    rr += 1
rr = 34
for x in cash_over[:15]:
    put(ws, rr, 1, x.get("so_ct"))
    put(ws, rr, 2, dparse(x.get("ngay_ct")), dt=True)
    put(ws, rr, 3, x.get("dien_giai"))
    put(ws, rr, 4, x.get("ps_co"), num=True)
    put(ws, rr, 11, "Vượt ngưỡng", fill=FILL_YEL)
    rr += 1
rr = 52
for i, t in enumerate(tt01[:15], 1):
    put(ws, rr, 1, i)
    put(ws, rr, 2, dparse(t["ngay"]), dt=True)
    put(ws, rr, 3, t["mst"])
    put(ws, rr, 4, t["ten"])
    put(ws, rr, 5, t["n"], num=True)
    put(ws, rr, 6, t["gt"], num=True)
    put(ws, rr, 10, "Vượt ngưỡng", fill=FILL_YEL)
    rr += 1
put(ws, 66, 1,
    f"NHẬN XÉT: {len(cash_over)} phiếu chi TK 111 vượt ngưỡng K5; {len(tt01)} cụm HĐ cùng ngày cùng NCC vượt ngưỡng TT01. "
    "IN11 có Excel 2023.Q4-2025.Q4 nhưng chưa bóc dòng → TT02/TT06/GT16 = PDF-chưa kết luận, không kết luận thiếu UNC (K11). "
    "BCTC [311] phải trả 2023=940.322.767 / DT 4,10 tỷ = 22,9% > P23=10%.",
    wrap=True, fill=FILL_YEL)
autowidth(ws, cap=40)
wb.save(lp(files["kq22"]))
wb.close()
print(" KQ2.2 saved")

# ---------------------------------------------------------------------------
# KQ3
# ---------------------------------------------------------------------------
print("FILL KQ3")
wbbo = load_workbook(lp(BO), data_only=True)
w4 = wbbo["04_Quy_tac_rui_ro"]
rules = []
for r in range(6, 120):
    rid = w4.cell(r, 1).value
    if not rid or not re.match(r"^[A-Z]{2}\d{2}$", str(rid).strip()):
        continue
    rules.append({
        "id": str(rid).strip(), "nhom": w4.cell(r, 2).value, "ten": w4.cell(r, 3).value,
        "need": w4.cell(r, 7).value, "muc": w4.cell(r, 8).value, "goi_y": w4.cell(r, 9).value,
    })
w8 = wbbo["08_Master_28TC"]
tcs = []
for r in range(5, 33):
    tid = w8.cell(r, 1).value
    if not tid:
        continue
    tcs.append((str(tid), w8.cell(r, 2).value, w8.cell(r, 3).value, w8.cell(r, 4).value, str(w8.cell(r, 8).value or "")))
wbbo.close()

wb = load_workbook(lp(files["kq3"]))
ws = wb["KQ3"]
put(ws, 2, 2, MST, bold=True)
put(ws, 2, 7, TEN, bold=True)
put(ws, 3, 2, KY)
put(ws, 3, 7, LEVEL)
put(ws, 4, 2, VER)
put(ws, 4, 7, MAU_VER)
put(ws, 5, 2, TODAY, dt=True)
put(ws, 5, 7, f"KQ1.3 {VER}, KQ2.1 {VER}, KQ2.2 {VER}")

findings = []

def add_f(**kw):
    findings.append(kw)

if f0_hits:
    add_f(qt="HD01", nhom="Hóa đơn – NCC", ten="Mua từ NCC F0/CA", tc="TC26", obj="0104567918", ky=KY,
          a="Giá trị mua CA", va=f0_gt, b="Tổng mua HĐĐT", vb=tot_mua_hddt, do=f0_ty, nguong=P13,
          st="ĐÃ_CHẠY", kq="Có phát hiện", nhan="Chắc", mucr="Cao", hs=3, diem=3, sac="GTGT", tien=f0_gt,
          nguon="IN07 sổ CA", ct="KQ2.1!KQ2_NCC_mua_vao", kt="Đối chiếu HĐ, HĐL, UNC với CÔNG TY TNHH CÔNG NGHIỆP PHÚ THÁI - CN HƯNG YÊN",
          nhan_xet="NNT không thuộc F1. 1 NCC dính sổ Công an (K2 coi F0).")
add_f(qt="HD03", nhom="Hóa đơn – NCC", ten="Tỷ trọng mua NCC rủi ro", tc="TC26", obj="NCC", ky=KY,
      a="Tỷ trọng F0/CA", va=f0_ty, b="P13", vb=P13, do=f0_ty, nguong=P13, st="ĐÃ_CHẠY",
      kq="Không vượt" if f0_ty <= P13 else "Có phát hiện", nhan="Chắc",
      mucr="Thông tin" if f0_ty <= P13 else "Cao", hs=2, diem=0 if f0_ty <= P13 else 2, sac="GTGT", tien=f0_gt,
      nguon="KQ2_NCC", ct="KQ2.1!KQ2_NCC_mua_vao")
add_f(qt="HD04", nhom="Hóa đơn – NCC", ten="Tập trung Top-5 NCC", tc="", obj="Top-5", ky=KY,
      a="Tỷ trọng Top-5", va=top5_ty, b="P14", vb=P14, do=top5_ty, nguong=P14, st="ĐÃ_CHẠY",
      kq="Có phát hiện" if top5_ty > P14 else "Không vượt", nhan="Chắc",
      mucr="Trung bình" if top5_ty > P14 else "Thông tin", hs=1, diem=1 if top5_ty > P14 else 0,
      sac="GTGT", tien=top5_gt, nguon="pivot MST HĐĐT", ct="KQ2.1")

for y in YEARS:
    g = gt_tot[y]
    if y >= 2022 and (g["34"] or g["hddt"]):
        fl = flag_gt(g["34"], g["hddt"])
        add_f(qt="GT01", nhom="GTGT", ten="DT [34] vs HĐĐT bán ra", tc="TC27", obj=TEN, ky=str(y),
              a="TK [34]", va=g["34"], b="HĐĐT bán", vb=g["hddt"], do=(g["34"] or 0) - (g["hddt"] or 0),
              nguong=P02, st="ĐÃ_CHẠY", kq=fl, nhan="Chắc",
              mucr="Trung bình" if fl == "Lệch" else "Thông tin", hs=2, diem=2 if fl == "Lệch" else 0,
              sac="GTGT", tien=abs((g["34"] or 0) - (g["hddt"] or 0)), nguon="IN03, IN01",
              ct=f"KQ2.1!KQ2_Doi_chieu_GTGT {y}")
    if g["34"] or g["bk"]:
        fl2 = flag_gt(g["34"] if y >= 2022 else None, g["bk"] or None)
        if y >= 2023:
            add_f(qt="GT02", nhom="GTGT", ten="DT [34] vs Bảng kê bán ra", tc="TC27", obj=TEN, ky=str(y),
                  a="TK [34]", va=g["34"], b="BK bán", vb=g["bk"], do=(g["34"] or 0) - (g["bk"] or 0),
                  nguong=P02, st="ĐÃ_CHẠY", kq=fl2 or "Chưa đủ nguồn", nhan="Chắc",
                  mucr="Cao" if y == 2024 and fl2 == "Lệch" else ("Trung bình" if fl2 == "Lệch" else "Thông tin"),
                  hs=2, diem=3 if y == 2024 and fl2 == "Lệch" else (2 if fl2 == "Lệch" else 0),
                  sac="GTGT", tien=abs((g["34"] or 0) - (g["bk"] or 0)), nguon="IN03, IN02", ct=f"KQ2.1 năm {y}")
    dt_bctc = (BM.get(y) or {}).get("dt")
    if dt_bctc is not None and (g["34"] or y <= 2021):
        fl3 = flag_gt(dt_bctc, g["34"] if y >= 2022 else dt_511.get(y))
        add_f(qt="TN01", nhom="TNDN – BCTC", ten="DT BCTC [01] vs GTGT [34]/511", tc="TC1", obj=TEN, ky=str(y),
              a="BCTC [01]", va=dt_bctc, b="GTGT [34]" if y >= 2022 else "511", vb=g["34"] if y >= 2022 else dt_511.get(y),
              do=(dt_bctc or 0) - (g["34"] if y >= 2022 else dt_511.get(y) or 0), nguong=P02, st="ĐÃ_CHẠY",
              kq=fl3, nhan="Sơ bộ", mucr="Trung bình" if fl3 == "Lệch" else "Thông tin", hs=2,
              diem=2 if fl3 == "Lệch" else 0, sac="TNDN",
              tien=abs((dt_bctc or 0) - (g["34"] if y >= 2022 else dt_511.get(y) or 0)), nguon="IN04, IN03", ct=f"IN04 KQKD {y}")
    if y in BM:
        dt, gv = BM[y]["dt"], BM[y]["gv"]
        lai = (dt or 0) - (gv or 0)
        add_f(qt="TN03", nhom="TNDN – BCTC", ten="Lợi nhuận gộp âm / biên mỏng", tc="TC5", obj=TEN, ky=str(y),
              a="DT thuần", va=dt, b="GV", vb=gv, do=lai / dt if dt else None, nguong=0, st="ĐÃ_CHẠY",
              kq="Có phát hiện" if lai < 0 else "Không vượt", nhan="Sơ bộ",
              mucr="Cao" if lai < 0 else "Thông tin", hs=2, diem=2 if lai < 0 else 0, sac="TNDN",
              tien=lai if lai < 0 else 0, nguon="IN04", ct=f"KQKD {y}")
        htk, dt2 = BM[y]["htk"], BM[y]["dt"]
        if dt2:
            ratio = htk / dt2
            add_f(qt="TN05", nhom="TNDN – BCTC", ten="HTK / DT [141]", tc="TC3", obj=TEN, ky=str(y),
                  a="HTK [141]", va=htk, b="DT", vb=dt2, do=ratio, nguong=P05, st="ĐÃ_CHẠY",
                  kq="Vượt ngưỡng" if ratio > P05 else "Không vượt", nhan="Sơ bộ",
                  mucr="Trung bình" if ratio > P05 else "Thông tin", hs=2, diem=1 if ratio > P05 else 0,
                  sac="TNDN", tien=htk, nguon="IN04", ct=f"CDKT {y} [141]")
        thue, lntt = BM[y]["thue_tndn"], BM[y]["ln_tt"]
        if lntt and lntt > 0:
            rate = (thue or 0) / lntt
            add_f(qt="TN19", nhom="TNDN – BCTC", ten="Thuế TNDN hiện hành / LNTT", tc="", obj=TEN, ky=str(y),
                  a="Thuế [22]", va=thue, b="LNTT [50]", vb=lntt, do=rate, nguong=0.20, st="ĐÃ_CHẠY",
                  kq="Có phát hiện" if rate < 0.10 and lntt > 1e8 else "Không vượt", nhan="Sơ bộ",
                  mucr="Cao" if y == 2024 and rate < 0.05 else "Thông tin", hs=2,
                  diem=3 if y == 2024 and rate < 0.05 else 0, sac="TNDN", tien=thue,
                  nguon="IN04 KQKD [22]/[50]", ct=f"BCTC {y}",
                  nhan_xet="2024 thuế 279.130 đ trên LNTT 1.931.870.396 (TNDN [22] TN chịu thuế). Sơ bộ - chưa so ngành (K13).")
        tr, dt3 = BM[y]["tr"], BM[y]["dt"]
        if dt3:
            ratio = (tr or 0) / dt3
            add_f(qt="TN17", nhom="TNDN – BCTC", ten="Phải trả / DT thuần", tc="TC20", obj=TEN, ky=str(y),
                  a="Phải trả [311]", va=tr, b="DT", vb=dt3, do=ratio, nguong=P23, st="ĐÃ_CHẠY",
                  kq="Vượt ngưỡng" if ratio > P23 else "Không vượt", nhan="Sơ bộ",
                  mucr="Trung bình" if ratio > P23 else "Thông tin", hs=1, diem=1 if ratio > P23 else 0,
                  sac="TNDN", tien=tr, nguon="IN04", ct=f"CDKT {y} [311]")
    k = kho.get(y)
    if k and y in BM:
        ck, htk = k["ck_tt"], BM[y]["htk"]
        add_f(qt="KH04", nhom="Kho – giá thành", ten="Tồn CK sổ vs HTK BCTC", tc="TC3", obj=TEN, ky=str(y),
              a="CK kho", va=ck, b="BCTC [141]", vb=htk, do=(ck or 0) - (htk or 0), nguong=P02, st="ĐÃ_CHẠY",
              kq="Lệch" if abs((ck or 0) - (htk or 0)) > P02 else "Khớp", nhan="Sơ bộ",
              mucr="Trung bình" if abs((ck or 0) - (htk or 0)) > P02 else "Thông tin", hs=2,
              diem=2 if abs((ck or 0) - (htk or 0)) > 1e8 else 0, sac="TNDN",
              tien=abs((ck or 0) - (htk or 0)), nguon="IN09, IN04", ct=f"KQ2.2!KQ2_Kho_Gia_thanh {y}")
    elif y == 2022:
        add_f(qt="KH04", nhom="Kho – giá thành", ten="Tồn CK sổ vs HTK BCTC", tc="TC3", obj=TEN, ky="2022",
              a="", va=None, b="", vb=None, do=None, nguong=None, st="KHÔNG_ĐỦ_DỮ_LIỆU", kq="Chưa kết luận",
              nhan="Sơ bộ", mucr="", hs=0, diem=0, sac="", tien=None, nguon="IN09", ct="", thieu="IN09 năm 2022")
    add_f(qt="DT03", nhom="Doanh thu – giá vốn", ten="511/632 năm vs BCTC", tc="TC28", obj="511-632", ky=str(y),
          a="511 PS Có", va=dt_511.get(y), b="632 PS Nợ", vb=gv_632.get(y),
          do=(dt_511.get(y) or 0) - (gv_632.get(y) or 0), nguong=None, st="ĐÃ_CHẠY",
          kq="Đã chạy tổng năm", nhan="Sơ bộ", mucr="Thông tin", hs=1, diem=0, sac="TNDN",
          tien=(dt_511.get(y) or 0) - (gv_632.get(y) or 0), nguon="IN13", ct="KQ2.2!KQ2_511_632")
    if match_y[y]["duoi"]:
        add_f(qt="DT01", nhom="Doanh thu – giá vốn", ten="Dòng lãi gộp ≤ 0", tc="", obj="511-632 dòng", ky=str(y),
              a="Số dòng P<=0", va=match_y[y]["duoi"], b="DT dưới giá", vb=match_y[y]["duoi_dt"],
              do=match_y[y]["duoi"], nguong=0, st="ĐÃ_CHẠY", kq="Có phát hiện", nhan="Sơ bộ",
              mucr="Trung bình", hs=2, diem=2, sac="TNDN", tien=match_y[y]["duoi_dt"], nguon="IN13",
              ct="KQ2.2!KQ2_511_632")

add_f(qt="TT01", nhom="Tiền mặt – thanh toán", ten="Cụm HĐ cùng ngày cùng MST vượt ngưỡng", tc="", obj="Cụm HĐ", ky=KY,
      a="Số cụm vượt", va=len(tt01), b="Ngưỡng K5", vb=P10, do=len(tt01), nguong=1, st="ĐÃ_CHẠY",
      kq="Có phát hiện" if tt01 else "Không vượt", nhan="Sơ bộ",
      mucr="Trung bình" if tt01 else "Thông tin", hs=2, diem=2 if tt01 else 0, sac="GTGT",
      tien=sum(x["tt"] for x in tt01), nguon="IN01", ct="KQ2.2!KQ2_Cong_no_Tien_mat 3b")
add_f(qt="TT03", nhom="Tiền mặt – thanh toán", ten="Chi tiền mặt TK 111 vượt ngưỡng", tc="", obj="TK 111", ky=KY,
      a="Số phiếu vượt", va=len(cash_over), b="Ngưỡng K5", vb=P10, do=len(cash_over), nguong=1, st="ĐÃ_CHẠY",
      kq="Có phát hiện" if cash_over else "Không vượt", nhan="Chắc",
      mucr="Trung bình" if cash_over else "Thông tin", hs=2, diem=2 if cash_over else 0, sac="TNDN",
      tien=sum(x.get("ps_co") or 0 for x in cash_over), nguon="IN10", ct="KQ2.2 Phần 3a")
add_f(qt="TT06", nhom="Tiền mặt – thanh toán", ten="Sao kê NH chưa bóc dòng", tc="", obj="IN11", ky=KY,
      a="File Excel", va=9, b="", vb=None, do=9, nguong=None, st="KHÔNG_ĐỦ_DỮ_LIỆU", kq="Chưa kết luận",
      nhan="Sơ bộ", mucr="", hs=0, diem=0, sac="", tien=None, nguon="IN11", ct="KQ1.3 IN11",
      thieu="Dòng sao kê SpreadsheetML + UNC", kt="Không kết luận thiếu UNC (K11)")
add_f(qt="HD07", nhom="Hóa đơn – NCC", ten="Đơn giá/mặt hàng lệch ngành", tc="", obj="ngành", ky=KY,
      a="", va=None, b="", vb=None, do=None, nguong=None, st="TREO", kq="Chưa kết luận", nhan="", mucr="",
      hs=0, diem=0, sac="", tien=None, nguon="K13", ct="", thieu="Tham chiếu ngành")
add_f(qt="LS01", nhom="Lịch sử tuân thủ", ten="Nợ thuế / KHBS TMS", tc="TC24", obj="TMS", ky=KY,
      a="", va=None, b="", vb=None, do=None, nguong=None, st="KHÔNG_ĐỦ_DỮ_LIỆU", kq="Chưa kết luận",
      nhan="", mucr="", hs=0, diem=0, sac="", tien=None, nguon="IN14", ct="1.3 chỉ có PTDN", thieu="IN14 dòng nợ thuế")
add_f(qt="LS02", nhom="Lịch sử tuân thủ", ten="Chậm nộp / nợ thuế kỳ trước", tc="TC24", obj="TMS", ky=KY,
      a="", va=None, b="", vb=None, do=None, nguong=None, st="KHÔNG_ĐỦ_DỮ_LIỆU", kq="Chưa kết luận",
      nhan="", mucr="", hs=0, diem=0, sac="", tien=None, nguon="IN14", ct="", thieu="IN14")
add_f(qt="LK01", nhom="Giao dịch liên kết", ten="GDLK không nộp PL", tc="TC18", obj="GDLK", ky=KY,
      a="PL GDLK", va=0, b="", vb=None, do=0, nguong=None, st="ĐÃ_CHẠY", kq="Không thấy phụ lục",
      nhan="Sơ bộ", mucr="Thông tin", hs=1, diem=0, sac="TNDN", tien=None, nguon="IN12 chưa nhận", ct="KQ1.3",
      kt="Xác nhận có/không quan hệ liên kết")

have = {f["qt"] for f in findings}
for rule in rules:
    rid = rule["id"]
    if rid in have:
        continue
    st = "ĐÃ_CHẠY"
    kq = "Không vượt / không có cờ máy"
    thieu = ""
    if rid in {"GT16", "TT02", "TT04"}:
        st, kq, thieu = "KHÔNG_ĐỦ_DỮ_LIỆU", "Chưa kết luận", "IN11 dòng sao kê"
    elif rid in {"KH01", "KH02", "KH03", "KH05", "KH06", "KH07"} and 2022 not in kho:
        st, kq, thieu = "ĐÃ_CHẠY", "Chạy một phần - thiếu 2022", "IN09 2022"
    add_f(qt=rid, nhom=rule["nhom"], ten=rule["ten"], tc="", obj=TEN, ky=KY, a="", va=None, b="", vb=None,
          do=None, nguong=None, st=st, kq=kq, nhan="Sơ bộ", mucr="", hs=0, diem=0, sac="", tien=None,
          nguon=rule["need"], ct="", thieu=thieu, kt=rule["goi_y"])

clear_from(ws, 10)
for i, f in enumerate(findings, 10):
    mph = f"{f['qt']}|{f.get('obj') or TEN}|{f.get('ky') or KY}"
    put(ws, i, 1, mph)
    put(ws, i, 2, "KQ2.1" if f["qt"][:2] in {"HD", "GT", "LD"} else ("KQ2.2" if f["qt"][:2] in {"DT", "KH", "TT"} else "Trực tiếp"))
    put(ws, i, 3, f.get("nhom")); put(ws, i, 4, f.get("qt")); put(ws, i, 5, f.get("ten"))
    put(ws, i, 6, f.get("tc")); put(ws, i, 7, f.get("obj")); put(ws, i, 8, f.get("ky"))
    put(ws, i, 9, f.get("a"))
    put(ws, i, 10, f.get("va"), num=isinstance(f.get("va"), (int, float)), pct=isinstance(f.get("va"), float) and f.get("va") is not None and abs(f.get("va")) <= 2 and f["qt"] in {"HD03", "HD04", "TN05", "TN17", "TN19"})
    put(ws, i, 11, f.get("b"))
    put(ws, i, 12, f.get("vb"), num=isinstance(f.get("vb"), (int, float)))
    put(ws, i, 13, f.get("do"), num=isinstance(f.get("do"), (int, float)))
    put(ws, i, 14, f.get("nguong"), num=isinstance(f.get("nguong"), (int, float)))
    put(ws, i, 15, f.get("st")); put(ws, i, 16, f.get("kq")); put(ws, i, 17, f.get("nhan"))
    fill = FILL_RED if f.get("mucr") == "Cao" else FILL_YEL if f.get("mucr") == "Trung bình" else None
    put(ws, i, 18, f.get("mucr"), fill=fill)
    put(ws, i, 19, f.get("hs"), num=True); put(ws, i, 20, f.get("diem"), num=True)
    put(ws, i, 21, f.get("sac"))
    put(ws, i, 22, f.get("tien"), num=isinstance(f.get("tien"), (int, float)))
    w = f.get("tien") if f.get("diem") and f.get("tien") else 0
    put(ws, i, 23, w if f.get("diem") else 0, num=True)
    put(ws, i, 25, f.get("nguon")); put(ws, i, 26, f.get("ct")); put(ws, i, 27, f.get("thieu"))
    put(ws, i, 28, f.get("kt")); put(ws, i, 32, f.get("nhan_xet"))
autowidth(ws, cap=36)

ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
clear_from(ws, 4)
put(ws, 2, 1, f"Áp dụng cho {TEN} (MST {MST}) - Kỳ {KY} | {VER} | {TODAY.strftime('%d/%m/%Y')}")

def tc_result(tid, ten, qt):
    rel = [f for f in findings if f.get("qt") == qt or f.get("tc") == tid or str(f.get("tc") or "") == tid.replace("TC", "TC")]
    rel = [f for f in findings if f.get("qt") == qt or str(f.get("tc") or "") in {tid, tid.replace("TC", "TC0") if len(tid) == 3 else tid}]
    if not rel:
        return "Chưa đánh giá / thiếu nguồn – không ghi 0 giả", "Chưa đánh giá", "Bổ sung nguồn hoặc xác minh tại trụ sở"
    top = max(rel, key=lambda x: x.get("diem") or 0)
    kq = f"{top.get('kq')} | {top.get('a')}={top.get('va')} vs {top.get('b')}={top.get('vb')} (kỳ {top.get('ky')})"
    return kq, top.get("mucr") or "Chưa đánh giá", top.get("kt") or ten

for i, (tid, ten, loai, nguong, qt) in enumerate(tcs, 4):
    kq, mucr, kt = tc_result(tid, ten, qt)
    fill = FILL_RED if mucr == "Cao" else FILL_YEL if mucr == "Trung bình" else FILL_GRN if mucr == "Thông tin" else None
    put(ws, i, 1, tid, bold=True)
    put(ws, i, 2, ten, wrap=True)
    put(ws, i, 3, loai)
    put(ws, i, 4, nguong)
    put(ws, i, 5, kq, wrap=True)
    put(ws, i, 6, mucr, fill=fill)
    put(ws, i, 7, kt, wrap=True)
    ws.row_dimensions[i].height = 42
autowidth(ws, cap=50)
ws.column_dimensions["E"].width = 60
wb.save(lp(files["kq3"]))
wb.close()
print(" KQ3 saved", "findings", len(findings), "rules", len(rules))

# ---------------------------------------------------------------------------
# Extra GĐ3
# ---------------------------------------------------------------------------
print("EXTRAS")
wb = Workbook()
ws = wb.active
ws.title = "Danh_gia_Dau_hieu_Bat_thuong"
ws.sheet_view.showGridLines = True
put(ws, 1, 1, f"KQ3_Danh_gia_Dau_hieu_Bat_thuong – {TEN} – MST {MST} – {KY}", bold=True, fill=FILL_HDR)
ws.merge_cells("A1:H1")
master(ws, 1, 1).font = FONT_HDR
hdr = ["STT", "Dấu hiệu", "Kỳ", "Số liệu then chốt", "Mức", "Hướng kiểm tra", "Chứng từ cần xem", "Mã quy tắc"]
for c, h in enumerate(hdr, 1):
    put(ws, 3, c, h, bold=True, fill=FILL_HDR)
    master(ws, 3, c).font = FONT_HDR
dau = [
    (1, "HĐĐT bán ra lệch tờ khai GTGT [34]", "2024, 2025",
     "2024 lệch 977.005.030; 2025 lệch 270.240.000", "Trung bình",
     "Đối chiếu HĐ hủy/thay thế, cut-off kỳ, hóa đơn điều chỉnh", "XML GTGT, portal HĐĐT, biên bản điều chỉnh", "GT01/TC27"),
    (2, "Bảng kê bán ra 2024 gấp ~2 lần tờ khai/HĐĐT", "2024",
     "BK unique HĐ vs [34]=22.700.193.753", "Cao",
     "Rà từng HĐ bảng kê vs portal; loại HĐ hủy, kỳ chồng", "Bảng kê gốc 2024 các quý", "GT02"),
    (3, "Thuế TNDN 2024 rất thấp so với LNTT", "2024",
     "Thuế BCTC [22]=279.130; LNTT [50]=1.931.870.396; TNDN [18]=1.936.132.137", "Cao",
     "Rà ưu đãi, chuyển lỗ, điều chỉnh tăng/giảm [B1]-[B15], phụ lục QTT", "QTT TNDN 2024 XML + PL", "TN19"),
    (4, "HTK BCTC lệch sổ kho", "2023-2024",
     "2023 kho CK 1.606.411.182 vs BCTC 821.334.499; 2024 6.556.251.416 vs 3.633.866.728", "Trung bình",
     "Đối chiếu XNT 152/155/156, phiếu nhập xuất, kiểm kê", "Sổ kho, biên bản KK, BCTC thuyết minh", "KH04/TC3"),
    (5, "Tập trung Top-5 NCC 72,8% > P14=50%", "2021-2025",
     f"Top-5 = {top5_ty:.2%} trên tổng mua HĐĐT {tot_mua_hddt:,.0f}".replace(",", "."),
     "Trung bình", "Chọn mẫu HĐ Top-5, hợp đồng, UNC, giá so thị trường", "HĐĐT mua, HĐKT, UNC", "HD04"),
    (6, "1 NCC trên sổ Công an", "kỳ kiểm tra",
     "MST 0104567918 – CÔNG TY TNHH CÔNG NGHIỆP PHÚ THÁI - CN HƯNG YÊN, GT 8.094.000, 2 HĐ",
     "Cao", "Xác minh hàng hóa thật, vận chuyển, thanh toán", "HĐ, phiếu xuất NCC, UNC, sổ CA", "HD01/TC26"),
    (7, "Chi tiền mặt và cụm HĐ vượt ngưỡng K5", "2021-2025",
     f"{len(cash_over)} phiếu chi 111; {len(tt01)} cụm HĐ TT01", "Trung bình",
     "Đối chiếu UNC/sao kê; không kết luận thiếu UNC khi chưa bóc dòng NH (K11)", "Sổ 111, sao kê Eximbank", "TT01/TT03"),
    (8, "Phải trả [311] 2023 / DT = 22,9% > P23=10%", "2023",
     "Phải trả 940.322.767 / DT 4.098.246.332", "Trung bình",
     "Tuổi nợ 331, đối chiếu công nợ NCC", "Sổ 331, xác nhận công nợ", "TN17/TC20"),
    (9, "632 năm 2023 lệch BCTC [11]", "2023",
     "632 sổ 3.992.714.785 vs BCTC GV 3.832.364.999 (lệch ~160 triệu)", "Trung bình",
     "Rà kết chuyển 632, xuất kho, giá thành", "Sổ 632, bảng GT 2023", "DT03/KH07"),
]
for i, row in enumerate(dau):
    for c, v in enumerate(row, 1):
        fill = FILL_RED if row[4] == "Cao" else FILL_YEL if row[4] == "Trung bình" else None
        put(ws, 4 + i, c, v, wrap=True, fill=fill if c == 5 else None)
autowidth(ws, cap=50)
wb.save(lp(files["dau"]))
wb.close()

wb = Workbook()
sheets = [
    "00_Tong_hop_chi_tiet", "1_NCC_F0_CanhBao", "2_NCC_DKKD_bat_thuong", "3_Ty_le_gia_von",
    "4_HD_ban_duoi_gia", "5_Khong_vay_no_EBITDA", "6_Top5_NCC", "7_Lech_HDDT_BangKe",
    "8_Cong_no_TK131_TK331", "9_TNCN_NLD",
]
wb.active.title = sheets[0]
for s in sheets[1:]:
    wb.create_sheet(s)

def hdr_row(ws, headers):
    for c, h in enumerate(headers, 1):
        put(ws, 1, c, h, bold=True, fill=FILL_HDR)
        master(ws, 1, c).font = FONT_HDR

ws = wb["00_Tong_hop_chi_tiet"]
put(ws, 1, 1, f"{TEN} - {MST} - {KY} - Tổng hợp chi tiết dấu hiệu", bold=True, fill=FILL_HDR)
master(ws, 1, 1).font = FONT_HDR
ws.merge_cells("A1:D1")
hdr_row(ws, ["Phần", "Chỉ tiêu", "Giá trị", "Nguồn"])
summary = [
    ("I", "Số HĐĐT mua", sum(1 for x in hddt if x["chieu"] == "MUA"), "IN01"),
    ("I", "Số HĐĐT bán", sum(1 for x in hddt if x["chieu"] == "BAN"), "IN01"),
    ("I", "Số dòng bảng kê unique", len(bk_u), "IN02"),
    ("I", "Số NCC mua", len(ncc_hd), "IN01+IN02"),
    ("II", "NCC F0/CA", len(f0_hits), "IN07"),
    ("II", "Doanh số F0/CA", f0_gt, "IN07"),
    ("II", "Top-5 %", top5_ty, "HD04"),
    ("III", "Tổng DT 511", tot511, "IN13"),
    ("III", "Tổng GV 632", tot632, "IN13"),
    ("III", "Lãi gộp sổ", tot511 - tot632, "IN13"),
    ("III", "Dòng P<=0", sum(match_y[y]["duoi"] for y in YEARS), "DT01"),
    ("IV", "Cụm HĐ TT01", len(tt01), "TT01"),
    ("IV", "Phiếu chi 111 vượt", len(cash_over), "TT03"),
    ("IV", "Sao kê NH", "9 file Excel, chưa bóc dòng", "IN11"),
    ("V", "NNT trên F1", "Không", "HD01"),
]
for i, row in enumerate(summary):
    for c, v in enumerate(row, 1):
        put(ws, 3 + i, c, v, num=isinstance(v, (int, float)) and c == 3, pct=isinstance(v, float) and v <= 2 and c == 3)

ws = wb["1_NCC_F0_CanhBao"]
hdr_row(ws, ["MST", "Tên", "DS", "Thuế", "Nhóm", "Nguồn", "CV"])
rr = 2
for m, rec, inf, nhom, nguon in f0_hits:
    put(ws, rr, 1, m); put(ws, rr, 2, rec["ten"] or inf.get("ten"))
    put(ws, rr, 3, rec["gt"], num=True); put(ws, rr, 4, rec["thue"], num=True)
    put(ws, rr, 5, "F0"); put(ws, rr, 6, nguon); put(ws, rr, 7, inf.get("cv"))
    for c in range(1, 8):
        master(ws, rr, c).fill = FILL_RED
        master(ws, rr, c).font = FONT_RED
    rr += 1
if rr == 2:
    put(ws, 2, 1, "Không có NCC F0/CA")

ws = wb["2_NCC_DKKD_bat_thuong"]
hdr_row(ws, ["MST", "Tên", "DS", "Trạng thái"])
put(ws, 2, 1, "Chưa gắn được ĐKKD bất thường từ danh bạ (danh bạ mỏng / không có ngày hiệu lực).")

ws = wb["3_Ty_le_gia_von"]
hdr_row(ws, ["Năm", "DT_511", "GV_632", "Lãi gộp", "Tỷ lệ", "BCTC[01]", "BCTC[11]"])
for j, y in enumerate(YEARS):
    put(ws, 2 + j, 1, y)
    put(ws, 2 + j, 2, dt_511[y], num=True)
    put(ws, 2 + j, 3, gv_632[y], num=True)
    put(ws, 2 + j, 4, dt_511[y] - gv_632[y], num=True)
    put(ws, 2 + j, 5, (dt_511[y] - gv_632[y]) / dt_511[y] if dt_511[y] else None, pct=True)
    put(ws, 2 + j, 6, (BM.get(y) or {}).get("dt"), num=True)
    put(ws, 2 + j, 7, (BM.get(y) or {}).get("gv"), num=True)

ws = wb["4_HD_ban_duoi_gia"]
hdr_row(ws, ["Năm", "Ngày", "Số CT", "DT 511", "GV 632", "Lãi", "Diễn giải"])
rr = 2
for y in YEARS:
    for a, b, lai in match_y[y]["duoi_rows"][:50]:
        put(ws, rr, 1, y)
        put(ws, rr, 2, dparse(a.get("ngay_ct")), dt=True)
        put(ws, rr, 3, a.get("so_ct"))
        put(ws, rr, 4, a.get("ps_co"), num=True)
        put(ws, rr, 5, b.get("ps_no"), num=True)
        put(ws, rr, 6, lai, num=True)
        put(ws, rr, 7, a.get("dien_giai"))
        for c in range(1, 8):
            master(ws, rr, c).fill = FILL_RED
        rr += 1
if rr == 2:
    put(ws, 2, 1, "Không có dòng P<=0")

ws = wb["5_Khong_vay_no_EBITDA"]
put(ws, 1, 1, "EBITDA sơ bộ từ BCTC - lãi vay [23]=0 mọi năm có BCTC", bold=True, fill=FILL_HDR)
master(ws, 1, 1).font = FONT_HDR
for j, y in enumerate([yy for yy in YEARS if yy in BM]):
    put(ws, 3, 2 + j, y, bold=True)
    put(ws, 4, 1, "LN trước thuế [50]")
    put(ws, 4, 2 + j, BM[y]["ln_tt"], num=True)
    put(ws, 5, 1, "Lãi vay [23]")
    put(ws, 5, 2 + j, BM[y]["cp_lai"], num=True)
    put(ws, 6, 1, "Thuế TNDN [22]")
    put(ws, 6, 2 + j, BM[y]["thue_tndn"], num=True)
put(ws, 8, 1, "Lãi vay = 0 trên BCTC 2021-2024: số liệu kê khai không có chi phí lãi. Vẫn cần sổ 635 khi kiểm tra tại trụ sở. Mẫu KQ2.1 hiện tại không có sheet EBITDA riêng.", wrap=True, fill=FILL_GRN)

ws = wb["6_Top5_NCC"]
hdr_row(ws, ["Hạng", "MST", "Tên", "DS", "Tỷ trọng"])
for i, (k, a) in enumerate(top5, 1):
    put(ws, 1 + i, 1, i)
    put(ws, 1 + i, 2, k)
    put(ws, 1 + i, 3, a["ten"])
    put(ws, 1 + i, 4, a["gt"], num=True)
    put(ws, 1 + i, 5, a["gt"] / tot_mua_hddt if tot_mua_hddt else None, pct=True)
put(ws, 8, 1, f"Tổng Top-5 = {top5_ty:.2%} (P14=50%)")

ws = wb["7_Lech_HDDT_BangKe"]
hdr_row(ws, ["Năm", "TK[34]", "HĐĐT bán", "BK bán", "Lệch TK-HĐ", "Lệch TK-BK", "TK[23]", "HĐĐT mua", "BK mua"])
for j, y in enumerate(YEARS):
    g = gt_tot[y]
    put(ws, 2 + j, 1, y)
    put(ws, 2 + j, 2, g["34"] if y >= 2022 else None, num=True)
    put(ws, 2 + j, 3, g["hddt"] or None, num=True)
    put(ws, 2 + j, 4, g["bk"] or None, num=True)
    put(ws, 2 + j, 5, None if y < 2022 else g["34"] - g["hddt"], num=True)
    put(ws, 2 + j, 6, None if y < 2022 else g["34"] - g["bk"], num=True)
    put(ws, 2 + j, 7, g["23"] if y >= 2022 else None, num=True)
    put(ws, 2 + j, 8, g["hddt_m"] or None, num=True)
    put(ws, 2 + j, 9, g["bk_m"] or None, num=True)

ws = wb["8_Cong_no_TK131_TK331"]
put(ws, 1, 1, "IN11 Excel 2023.Q4-2025.Q4 - chưa bóc dòng, không kết luận thiếu UNC (K11).", wrap=True, fill=FILL_YEL)
put(ws, 2, 1, f"IN10 số dòng chi 111 (unique): {len(chi_ok)}")
put(ws, 3, 1, f"TT01 cụm HĐ vượt: {len(tt01)}; TT03 phiếu chi vượt: {len(cash_over)}")
hdr_row = None
put(ws, 5, 1, "Năm"); put(ws, 5, 2, "PS Nợ 131"); put(ws, 5, 3, "PS Có 131"); put(ws, 5, 4, "PS Có 331"); put(ws, 5, 5, "BCTC [311]"); put(ws, 5, 6, "[311]/DT")
for j, y in enumerate(YEARS):
    put(ws, 6 + j, 1, y)
    put(ws, 6 + j, 2, n131.get(y), num=True)
    put(ws, 6 + j, 3, c131.get(y), num=True)
    put(ws, 6 + j, 4, c331.get(y), num=True)
    put(ws, 6 + j, 5, (BM.get(y) or {}).get("tr"), num=True)
    dt = (BM.get(y) or {}).get("dt")
    tr = (BM.get(y) or {}).get("tr")
    put(ws, 6 + j, 6, (tr / dt) if dt else None, pct=True)

ws = wb["9_TNCN_NLD"]
put(ws, 1, 1, "Năm"); put(ws, 1, 2, "Số LĐ QTT"); put(ws, 1, 3, "TNCT"); put(ws, 1, 4, "Thuế KT"); put(ws, 1, 5, "Thuế phải")
for j, y in enumerate(YEARS):
    q = nld_qtt.get(y)
    put(ws, 2 + j, 1, y)
    if q:
        put(ws, 2 + j, 2, q.get("so_ld"), num=True)
        put(ws, 2 + j, 3, q.get("tnct"), num=True)
        put(ws, 2 + j, 4, q.get("thue_kt"), num=True)
        put(ws, 2 + j, 5, q.get("thue_phai"), num=True)
put(ws, 8, 1, f"Danh sách NLĐ nạp {len(C.get('nld_ds') or [])} dòng (IN05). Đa nguồn 3.3 mỏng.")

for s in wb.sheetnames:
    w = wb[s]
    w.sheet_view.showGridLines = True
    for row in w.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            tnr(cell)
    autowidth(w, cap=40)
wb.save(lp(files["giai"]))
wb.close()
print(" extras saved")

# copy
SAO.mkdir(parents=True, exist_ok=True)
VAULT_KQ.mkdir(parents=True, exist_ok=True)
copied = []
for p in sorted(OUT.glob(f"KQ*_{MST}_{KY}_{VER}.xlsx")):
    shutil.copy2(lp(p), lp(SAO / p.name))
    shutil.copy2(lp(p), lp(VAULT_KQ / p.name))
    copied.append((p.name, p.stat().st_size))
    print("copied", p.name, p.stat().st_size)

print("DONE", json.dumps({
    "files": copied,
    "level": LEVEL,
    "f0_hits": len(f0_hits),
    "top5": round(top5_ty, 4),
    "tt01": len(tt01),
    "cash": len(cash_over),
    "findings": len(findings),
    "dt_511": dt_511,
}, ensure_ascii=False, default=str))
Path(r"C:\Users\NGOC\AppData\Local\Temp\htp_fill_rest_ok.txt").write_text("ok", encoding="utf-8")
