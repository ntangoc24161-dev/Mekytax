# -*- coding: utf-8 -*-
"""Fill v1.4 KQ1.1-KQ3 for HUY THANH PHAT 3603820015 from htp_cache.json."""
from __future__ import annotations

import json
import re
import shutil
import sys
from collections import defaultdict
from copy import copy
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

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
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


def clear_from(ws, start_row, max_col):
    if ws.max_row and ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)


def write_row(ws, r, values, nums=(), dates=(), fill=None):
    for c, v in enumerate(values, 1):
        cell = ws.cell(r, c, v)
        tnr(cell)
        cell.border = THIN
        cell.alignment = Alignment(vertical="center", wrap_text=c in {8, 9, 10} or False)
        if fill:
            cell.fill = fill
        if c in nums and isinstance(v, (int, float)):
            cell.number_format = NUM
        if c in dates and v not in (None, ""):
            cell.number_format = DATE_FMT


def header_banner(ws, r, text):
    if isinstance(ws.cell(r, 1).value, str) and "[" in str(ws.cell(r, 1).value):
        v = ws.cell(r, 1).value
        v = v.replace("[MST]", MST).replace("[KY]", KY).replace("[NGAY]", TODAY.strftime("%d/%m/%Y"))
        v = v.replace("[TÊN DOANH NGHIỆP]", TEN).replace("[DN]", TEN)
        v = v.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")
        v = v.replace("[Đủ | Một phần | Chưa có]", DN_STATUS)
        ws.cell(r, 1).value = v
        tnr(ws.cell(r, 1), bold=True)


def sub_all(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str):
                v = cell.value
                v = v.replace("[MST]", MST).replace("[KY]", KY).replace("[NGAY]", TODAY.strftime("%d/%m/%Y"))
                v = v.replace("[TÊN DOANH NGHIỆP]", TEN)
                v = v.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")
                v = v.replace("[Đủ | Một phần | Chưa có]", DN_STATUS)
                if v != cell.value:
                    cell.value = v
            tnr(cell)


print("LOAD CACHE")
C = json.loads(CACHE.read_text(encoding="utf-8"))
hddt = C["hddt"]
bk = C["bangke"]
soct = C["soct"]
gtgt = C["gtgt_eff"]
tndn = C["tndn_eff"]
bctc = C["bctc_eff"]
f0 = {mst_base(x["mst"]): x for x in C["f0"]}
f1 = {mst_base(x["mst"]): x for x in C["f1"]}
ca = {mst_base(x["mst"]): x for x in C["ca"]}
db = {mst_base(x["mst"]): x for x in C["danhba"]}
nnt_f1 = MST in f1
nnt_f0 = MST in f0
nnt_ca = MST in ca

ps = [x for x in soct if x["kind"] == "PS"]


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


dt_511 = tk_year("511", "ps_co")
gv_632 = tk_year("632", "ps_no")
chi_111 = tk_year("1111", "ps_co")
if sum(chi_111.values()) == 0:
    chi_111 = tk_year("111", "ps_co")
n133 = tk_year("1331", "ps_no")
if sum(n133.values()) == 0:
    n133 = tk_year("133", "ps_no")
n131 = tk_year("131", "ps_no")
c331 = tk_year("331", "ps_co")
n621 = tk_year("621", "ps_no")
n622 = tk_year("622", "ps_no")
n627 = tk_year("627", "ps_no")


def bctc_map():
    out = {}
    for x in bctc:
        k = x.get("kqkd") or {}
        cd = x.get("cdkt_cuoi") or {}
        out[x["nam"]] = {
            "dt": k.get("01") or 0,
            "gv": k.get("11") or 0,
            "ln_gop": k.get("20") or 0,
            "dt_tc": k.get("21") or 0,
            "thue_tndn": k.get("22") or 0,
            "lnst": k.get("60") or k.get("50") or 0,
            "htk": cd.get("141") or cd.get("140") or 0,
            "htk140": cd.get("140") or 0,
            "htk141": cd.get("141") or 0,
            "pt": cd.get("131") or 0,
            "tr": cd.get("311") or cd.get("331") or 0,
            "tm": cd.get("111") or 0,
            "nh": cd.get("112") or 0,
            "von": cd.get("411") or 0,
            "thue_kt": cd.get("152") or 0,
            "tscd": cd.get("211") or 0,
            "ngmua_tt": cd.get("312") or cd.get("313") or 0,
            "cp_tt": cd.get("242") or 0,
            "file": x.get("file"),
            "cdkt": cd,
            "kqkd": k,
        }
    return out


BM = bctc_map()
TN = {x["nam"]: x.get("chi_tieu") or {} for x in tndn}


def gtgt_year():
    d = {y: {"34": 0.0, "35": 0.0, "23": 0.0, "24": 0.0, "40": 0.0, "43": 0.0, "32": 0.0, "22": 0.0, "kys": []} for y in YEARS}
    for x in gtgt:
        y = x.get("nam")
        if y not in d:
            continue
        ct = x.get("chi_tieu") or {}
        d[y]["34"] += ct.get("34") or 0
        d[y]["35"] += ct.get("35") or 0
        d[y]["23"] += ct.get("23") or 0
        d[y]["24"] += ct.get("24") or 0
        d[y]["40"] += ct.get("40") or 0
        d[y]["43"] = ct.get("43") or ct.get("41") or d[y]["43"]
        d[y]["32"] += ct.get("32") or 0
        d[y]["22"] += ct.get("22") or 0
        d[y]["kys"].append(x)
    return d


GY = gtgt_year()


def sum_hd(rows, chieu, nam=None, field="gt"):
    s = 0.0
    n = 0
    for r in rows:
        if r.get("chieu") != chieu:
            continue
        if nam is not None and r.get("nam") != nam:
            continue
        s += r.get(field) or 0
        n += 1
    return s, n


# unique bang ke by so+mst+ngay (sum line items)
bk_u = []
seen = {}
for r in bk:
    k = (r.get("chieu"), so_key(r.get("so_key")), r.get("mst_dt"), r.get("ngay"))
    if k in seen:
        i = seen[k]
        bk_u[i]["gt"] += r.get("gt") or 0
        bk_u[i]["thue"] += r.get("thue") or 0
        if r.get("hang") and r["hang"] not in (bk_u[i].get("hang") or ""):
            bk_u[i]["hang"] = (bk_u[i].get("hang") or "") + "; " + r["hang"]
    else:
        seen[k] = len(bk_u)
        bk_u.append(dict(r))

# NCC from HDDT mua (CQT) union BK mua
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
    info = f0.get(m) or ca.get(m)
    if info:
        f0_hits.append((m, rec, info, "F0" if m in f0 else "CA"))
f0_gt = sum(r["gt"] for _, r, _, _ in f0_hits)
f0_ty = (f0_gt / tot_mua_hddt) if tot_mua_hddt else 0

# same-day clusters TT01
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
        tt01.append({"mst": mst, "ngay": ngay, "n": v["n"], "gt": v["gt"], "thue": v["thue"], "tt": v["tt"], "ten": v["ten"], "nguong": nguong})
tt01.sort(key=lambda x: -x["tt"])

# 511-632 line match
lines_511 = [x for x in ps if x["tk"] == "511"]
lines_632 = [x for x in ps if x["tk"] == "632"]


def hd_from_dg(s):
    s = str(s or "")
    m = re.search(r"(?:HĐ|HD|hđ|số\s*)\s*([0-9]{1,8})", s, re.I)
    if m:
        return so_key(m.group(1))
    m = re.search(r"\b(\d{1,6})\b", s)
    return so_key(m.group(1)) if m else ""


match_y = {y: {"khop": 0, "lech": 0, "no632": 0, "no511": 0, "duoi": 0, "duoi_dt": 0} for y in YEARS}
used_632 = set()
for a in lines_511:
    y = a.get("nam")
    if y not in match_y:
        continue
    key = (a.get("ngay_ct"), so_key(a.get("so_ct")))
    hit = None
    for i, b in enumerate(lines_632):
        if i in used_632:
            continue
        if b.get("nam") != y:
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
    else:
        match_y[y]["khop"] += 1
for i, b in enumerate(lines_632):
    if i not in used_632 and b.get("nam") in match_y:
        match_y[b["nam"]]["no511"] += 1

kho = {k["nam"]: k for k in C["kho"]}
nld_qtt = {int(x["nam"]): x for x in C["nld_qtt"]}

# diffs
diff_gt01 = {y: GY[y]["34"] - sum_hd(hddt, "BAN", y)[0] for y in YEARS}
diff_gt02 = {y: GY[y]["34"] - sum_hd(bk_u, "BAN", y)[0] for y in YEARS}
diff_tn01 = {y: GY[y]["34"] - (BM.get(y, {}).get("dt") or 0) for y in YEARS}

print("511", dt_511)
print("632", gv_632)
print("GTGT34", {y: GY[y]["34"] for y in YEARS})
print("HDDT ban", {y: sum_hd(hddt, "BAN", y)[0] for y in YEARS})
print("top5 ty", round(top5_ty, 4), "f0 hits", len(f0_hits), "tt01", len(tt01))
print("NNT F1/F0/CA", nnt_f1, nnt_f0, nnt_ca)

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

# ---------------------------------------------------------------------------
# KQ1.1
# ---------------------------------------------------------------------------
print("FILL KQ1.1")
wb = load_workbook(lp(files["kq11"][1]))

# IN00
ws = wb["IN00_Manifest"]
cqt_files = [m for m in C["manifest"] if m["nguon"] == "CQT"]
clear_from(ws, 3, 15)
for i, m in enumerate(cqt_files, 1):
    write_row(
        ws,
        2 + i,
        [
            f"N{i:03d}",
            m.get("ma_in"),
            "CQT",
            m.get("ext", "").lstrip(".").upper() or "KHAC",
            dparse(m.get("mtime")),
            2021,
            2025,
            m.get("sha256"),
            "Có" if "bo sung" in m["name"].lower() or "bổ sung" in m["name"].lower() else "Không",
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

# IN01
ws = wb["IN01_HDDT"]
clear_from(ws, 3, 26)
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
            "",
            "",
            None,
            None,
            r.get("gt"),
            None,
            r.get("thue"),
            r.get("tt"),
            "Hợp lệ",
            "",
            r.get("tinh_vao_tong"),
            "",
            f"{r.get('ky_hieu')}|{r.get('so_key')}|{dt}",
            r.get("file"),
        ],
        nums={17, 19, 20},
        dates={7},
    )

# IN03
ws = wb["IN03_ToKhai_GTGT"]
clear_from(ws, 3, 27)
for i, x in enumerate(gtgt, 1):
    ct = x.get("chi_tieu") or {}
    ky = x.get("ky") or ""
    ky_s = ky
    if re.match(r"^[1-4]/\d{4}$", ky):
        q, y = ky.split("/")
        ky_s = f"Q{q}/{y}"
    write_row(
        ws,
        2 + i,
        [
            x.get("nam"),
            ky_s,
            "Bổ sung" if str(x.get("loai_tkhai")).upper() in {"B", "2"} else "Chính thức",
            x.get("so_lan") or 0,
            "Có",
            None,
            ct.get("22"),
            ct.get("23"),
            ct.get("24"),
            ct.get("25"),
            ct.get("26"),
            ct.get("29"),
            ct.get("30"),
            ct.get("31"),
            ct.get("32"),
            ct.get("33"),
            ct.get("34"),
            ct.get("35"),
            ct.get("36"),
            ct.get("37"),
            ct.get("38"),
            ct.get("39"),
            ct.get("40"),
            ct.get("41"),
            ct.get("42"),
            ct.get("43"),
            x.get("file"),
        ],
        nums=set(range(7, 27)),
    )

# IN04
ws = wb["IN04_BCTC_TNDN"]
clear_from(ws, 3, 8)
r = 3
for y, b in BM.items():
    pairs = [
        ("KQKD", "01", "Doanh thu bán hàng và cung cấp dịch vụ", b["dt"]),
        ("KQKD", "11", "Giá vốn hàng bán", b["gv"]),
        ("KQKD", "20", "Lợi nhuận gộp", b["ln_gop"]),
        ("KQKD", "21", "Doanh thu hoạt động tài chính", b["dt_tc"]),
        ("KQKD", "22", "Chi phí thuế TNDN hiện hành", b["thue_tndn"]),
        ("CDKT", "141", "Hàng tồn kho", b["htk141"]),
        ("CDKT", "140", "Hàng tồn kho (ct140)", b["htk140"]),
        ("CDKT", "131", "Phải thu khách hàng", b["pt"]),
        ("CDKT", "111", "Tiền mặt", b["tm"]),
        ("CDKT", "112", "Tiền gửi ngân hàng", b["nh"]),
        ("CDKT", "411", "Vốn góp chủ sở hữu", b["von"]),
        ("CDKT", "152", "Thuế GTGT được khấu trừ", b["thue_kt"]),
        ("CDKT", "211", "TSCĐ hữu hình", b["tscd"]),
    ]
    tnct = TN.get(y) or {}
    if tnct.get("04"):
        pairs.append(("TNDN", "04", "Doanh thu QTT TNDN [04]", tnct.get("04")))
    if tnct.get("13"):
        pairs.append(("TNDN", "13", "Giá vốn QTT TNDN [13]", tnct.get("13")))
    if tnct.get("18") is not None:
        pairs.append(("TNDN", "18", "Lợi nhuận trước thuế [18]", tnct.get("18")))
    if tnct.get("22") is not None:
        pairs.append(("TNDN", "22", "Thu nhập chịu thuế [22]", tnct.get("22")))
    for bao, ma, tenct, so in pairs:
        write_row(ws, r, [y, bao, ma, tenct, so, None, "TT200/TT133 - đọc từ XML BCTC", b.get("file")], nums={5})
        r += 1

# IN05
ws = wb["IN05_NLD_TNCN"]
clear_from(ws, 3, 16)
r = 3
for y in YEARS:
    q = nld_qtt.get(y)
    if not q:
        continue
    write_row(
        ws,
        r,
        ["QTT", y, f"Quyết toán TNCN {y}", "", MST, "05/QTT-TNCN", "Tiền lương", q.get("tnct"), q.get("thue_kt"), "Không", None, "", "", f"Số LĐ {int(q.get('so_ld') or 0)}", "Không", q.get("mau")],
        nums={8, 9},
    )
    r += 1
for i, row in enumerate(C.get("nld_ds") or []):
    raw = row.get("raw") or []
    if not raw or not any(raw):
        continue
    write_row(ws, r, ["NLD", row.get("nam"), raw[1] if len(raw) > 1 else "", raw[2] if len(raw) > 2 else "", raw[3] if len(raw) > 3 else "", "05-1", "", None, None, "", None, "", "", "", "", ""], nums={8, 9})
    r += 1
    if r > 400:
        break

# IN14
ws = wb["IN14_TMS_Nghia_vu"]
clear_from(ws, 3, 16)
tms_files = [m for m in C["manifest"] if m.get("ma_in") == "IN14"]
for i, m in enumerate(tms_files, 1):
    write_row(
        ws,
        2 + i,
        [MST, "PHAN_TICH_BCTC_TMS", "BCTC", "2016-2025", None, None, None, None, None, None, None, "", None, "File PTDN TMS (CĐKT/KQKD 2016-2025) - SpreadsheetML, chưa bóc hết dòng nợ thuế. LS01/LS02: chưa đủ lịch sử nợ/chậm nộp.", "TMS_PTDN", m.get("rel")],
        dates={6},
    )

apply_tnr(wb["IN00_Manifest"])
wb.save(lp(files["kq11"][1]))
wb.close()
print(" KQ1.1 saved")

# ---------------------------------------------------------------------------
# KQ1.2
# ---------------------------------------------------------------------------
print("FILL KQ1.2")
wb = load_workbook(lp(files["kq12"][1]))
ws = wb["IN00_Manifest"]
dn_files = [m for m in C["manifest"] if m["nguon"] in {"DN", "TRA_CUU", "BB"}]
clear_from(ws, 3, 15)
for i, m in enumerate(dn_files, 1):
    write_row(ws, 2 + i, [f"D{i:03d}", m.get("ma_in"), m.get("nguon"), (m.get("ext") or "").lstrip(".").upper(), dparse(m.get("mtime")), 2021, 2025, m.get("sha256"), "Không", "Có", None, 0, "Giữ nguyên", "", m.get("rel")], dates={5})
ws.cell(1, 2).value = f"{TEN} - MST {MST} - kỳ {KY} - {TODAY.strftime('%d/%m/%Y')} - {VER}"

ws = wb["IN02_BangKe"]
clear_from(ws, 3, 24)
hddt_keys = {(so_key(r.get("so_key")), mst_base(r.get("mst_dt"))) for r in hddt}
for i, r in enumerate(bk_u, 1):
    k = (so_key(r.get("so_key")), mst_base(r.get("mst_dt")))
    st = "KHOP" if k in hddt_keys else "CHI_CO_BK"
    write_row(
        ws,
        2 + i,
        [
            r.get("chieu"),
            r.get("nam"),
            r.get("quy"),
            r.get("thang"),
            "",
            r.get("so_hd"),
            dparse(r.get("ngay")),
            r.get("mst_goc"),
            r.get("mst_dt"),
            r.get("ten"),
            r.get("hang"),
            "",
            None,
            r.get("gt"),
            r.get("thue_suat"),
            r.get("thue"),
            (r.get("gt") or 0) + (r.get("thue") or 0),
            "",
            "Có",
            "Có",
            st,
            "",
            f"|{r.get('so_key')}|{r.get('mst_dt')}",
            r.get("file"),
        ],
        nums={14, 16, 17},
        dates={7},
    )

# IN08 monthly TK
ws = wb["IN08_So_KeToan"]
clear_from(ws, 3, 10)
agg = defaultdict(lambda: {"no": 0.0, "co": 0.0})
keep_tk = {"111", "1111", "112", "1121", "131", "133", "1331", "152", "155", "156", "331", "511", "5111", "5112", "632", "621", "622", "627", "154"}
for x in ps:
    tk = x.get("tk")
    if tk not in keep_tk:
        continue
    d = dparse(x.get("ngay_ct") or x.get("ngay_ht"))
    th = d.month if d else 0
    agg[(x.get("nam"), th, tk)]["no"] += x.get("ps_no") or 0
    agg[(x.get("nam"), th, tk)]["co"] += x.get("ps_co") or 0
r = 3
for (nam, th, tk) in sorted(agg):
    a = agg[(nam, th, tk)]
    write_row(ws, r, [nam, th, tk, a["no"], a["co"], None, None, None, "SO_CHI_TIET", f"SO_CHI_TIET_CAC_TAI_KHOAN {nam}.xlsx"], nums={4, 5})
    r += 1

# IN09
ws = wb["IN09_Kho_GiaThanh"]
clear_from(ws, 3, 16)
r = 3
for y in YEARS:
    k = kho.get(y)
    if not k:
        continue
    write_row(
        ws,
        r,
        [y, "Tổng hợp tồn kho", k["dk_sl"], k["dk_tt"], k["nhap_sl"], k["nhap_tt"], k["xuat_sl"], k["xuat_tt"], k["ck_sl"], k["ck_tt"], None, n621.get(y), n622.get(y), n627.get(y), None, k["file"]],
        nums=set(range(3, 15)),
    )
    r += 1

# IN10 chi 1111
ws = wb["IN10_Chi_tien_mat"]
clear_from(ws, 3, 12)
chi_lines = [x for x in ps if x["tk"] in {"1111", "111"} and (x.get("ps_co") or 0) > 0]
# de-dup 111 vs 1111
seen_chi = set()
chi_ok = []
for x in chi_lines:
    k = (x.get("nam"), x.get("so_ct"), x.get("ngay_ct"), round(x.get("ps_co") or 0))
    if k in seen_chi:
        continue
    seen_chi.add(k)
    chi_ok.append(x)
chi_ok.sort(key=lambda x: -(x.get("ps_co") or 0))
r = 3
for x in chi_ok[:800]:
    write_row(
        ws,
        r,
        [dparse(x.get("ngay_ct")), x.get("nam"), x.get("so_ct"), x.get("tk_du"), "", "", x.get("dien_giai"), x.get("ps_co"), "Chi TK111", "Chưa rõ", "", x.get("file")],
        nums={8},
        dates={1},
    )
    r += 1

# IN11
ws = wb["IN11_Sao_ke_NH"]
clear_from(ws, 3, 14)
bank_files = [m for m in C["manifest"] if m.get("ma_in") == "IN11" and m["ext"] == ".xlsx"]
r = 3
for m in bank_files:
    write_row(ws, r, [dparse(m.get("mtime")), None, "Eximbank", "", m.get("name"), None, None, "", "", "", "", "Excel", "Đã nạp danh mục file - dòng chi tiết SpreadsheetML cần đọc bổ sung", m.get("rel")], dates={1})
    r += 1

# IN12 empty - leave example deleted
ws = wb["IN12_GDLK"]
clear_from(ws, 3, 9)
ws.cell(3, 1).value = None

# IN13
ws = wb["IN13_So_511_632"]
clear_from(ws, 3, 20)
r = 3
for x in lines_511:
    write_row(
        ws,
        r,
        [x.get("nam"), dparse(x.get("ngay_ct")).month if dparse(x.get("ngay_ct")) else None, dparse(x.get("ngay_ct")), x.get("so_ct"), x.get("tk"), x.get("tk_du"), "", x.get("dien_giai"), "", None, x.get("ps_co"), 0, "", hd_from_dg(x.get("dien_giai")), x.get("so_ct"), None, None, None, "", x.get("file")],
        nums={11, 12},
        dates={3},
    )
    r += 1
for x in lines_632:
    write_row(
        ws,
        r,
        [x.get("nam"), dparse(x.get("ngay_ct")).month if dparse(x.get("ngay_ct")) else None, dparse(x.get("ngay_ct")), x.get("so_ct"), x.get("tk"), x.get("tk_du"), "", x.get("dien_giai"), "", None, 0, x.get("ps_no"), "", hd_from_dg(x.get("dien_giai")), x.get("so_ct"), None, None, None, "", x.get("file")],
        nums={11, 12},
        dates={3},
    )
    r += 1

wb.save(lp(files["kq12"][1]))
wb.close()
print(" KQ1.2 saved")

# ---------------------------------------------------------------------------
# KQ1.3
# ---------------------------------------------------------------------------
print("FILL KQ1.3")
wb = load_workbook(lp(files["kq13"][1]))
ws = wb["KQ1_Kiem_ke"]
ws["B2"] = MST
ws["H2"] = TEN
ws["B3"] = KY
ws["H3"] = LEVEL
ws["B5"] = VER
ws["H5"] = MAU_VER
ws["B6"] = TODAY
ws["B6"].number_format = DATE_FMT
ws["H6"] = DN_STATUS
ws["B7"] = "KQ2.1, KQ2.2, KQ3"
ws["H7"] = "Kỳ suy ra theo K12 từ tờ khai GTGT Q1/2022–Q3/2025, TNDN/BCTC 2021-2024, TNDN 2025, sổ DN 2021-2025. NNT không có trên F1/F0/CA."

def set_in_row(ws, row, dinh_dang, ky_phu, ky_thieu, so_dong, thieu_mst, bs, muc, lydo, tt, viec, ghichu):
    ws.cell(row, 6).value = dinh_dang
    ws.cell(row, 7).value = ky_phu
    ws.cell(row, 8).value = ky_thieu
    ws.cell(row, 9).value = so_dong
    ws.cell(row, 10).value = thieu_mst
    ws.cell(row, 11).value = bs
    ws.cell(row, 12).value = muc
    ws.cell(row, 13).value = lydo
    ws.cell(row, 14).value = tt
    ws.cell(row, 15).value = viec
    ws.cell(row, 17).value = ghichu
    for c in range(6, 18):
        tnr(ws.cell(row, c))

set_in_row(ws, 10, "Excel", "2022-2025", "2021 HĐĐT mỏng", len(hddt), 0, "Không", "Đủ", "Portal HĐĐT mua 614 + bán 61", "Đã chuẩn hóa", "MST, khóa HĐ, Tinh_vao_tong", "CQT")
set_in_row(ws, 11, "Excel", "2021-2025", "Bán ra thiếu 2022 và 2023 Q1-Q3, 2025 Q3", len(bk_u), sum(1 for x in bk_u if not x.get("mst_dt")), "Không", "Mỏng", "Có bảng kê gốc; tổng bán ra 2024 lệch gấp ~2 lần tờ khai/HĐĐT", "Đã gộp, lọc Tổng cộng", "BT6 MST, BT5 đối chiếu HĐĐT", "DN")
set_in_row(ws, 12, "XML", "Q1/2022-Q3/2025", "2021 GTGT mỏng; Q4/2025 chưa có", len(gtgt), 0, "Có (Q2/2022)", "Đủ", "15 kỳ hiệu lực", "Đã chọn bản hiệu lực", "Bổ sung ưu tiên hơn chính thức", "CQT")
set_in_row(ws, 13, "XML", "2021-2025", "BCTC 2025 chưa có XML", len(bctc) + len(tndn), 0, "Có TNDN 2025 L01", "Đủ", "BCTC 2021-2024; TNDN 2021-2025", "Đã trích chỉ tiêu", "Đọc mã XML, không suy thông tư", "CQT")
set_in_row(ws, 14, "XML/Excel", "2021-2025", "Đa nguồn 3.3 mỏng", len(C.get("nld_ds") or []), 0, "Không", "Đủ", "QTT 5 năm; 113 dòng NLĐ", "Đã nạp 1.4", "QTT trước NLĐ theo năm", "CQT")
set_in_row(ws, 15, "Excel sổ CVCBRR+CA", "cập nhật 09/11/2026", "", len(f0_hits), 0, "Không", "Đủ", "Nguồn chỉ định F0/F1/CA", "Đã khớp MST", "NNT không thuộc F1", "CQT")
set_in_row(ws, 16, "Excel", "2021-2025", "2022 phát sinh rất mỏng", len(ps), 0, "Không", "Đủ", "Sổ chi tiết 8.464 dòng PS", "Đã gom theo TK/tháng", "Không cộng 111+1111", "DN")
set_in_row(ws, 17, "Excel", "2021,2023-2025", "Thiếu XNT 2022", len(kho), 0, "Không", "Mỏng", "Có tồn kho + bảng giá thành + PDF định mức", "Đã cộng XNT", "DN sản xuất vì thép", "DN")
set_in_row(ws, 18, "Excel (từ sổ 111)", "2021-2025", "Không có sổ 111 tách file", len(chi_ok), 0, "Không", "Đủ", "Chi TK 1111 lấy từ sổ chi tiết", "Đã lọc trùng 111/1111", "DN")
set_in_row(ws, 19, "Excel Eximbank", "2023.Q4-2025.Q4", "Thiếu 2021–2023.Q3", 9, 0, "Không", "Mỏng", "9 file LichSuTaiKhoan", "Đã kê file", "SpreadsheetML, dòng chi tiết cần đọc bổ sung", "DN")
set_in_row(ws, 20, "", "", "Chưa nhận PL GDLK", 0, 0, "Không", "Chưa nhận", "Không thấy phụ lục NĐ 132", "Để trống", "Không điền 0", "DN")
set_in_row(ws, 21, "Excel", "2021-2025", "2022 không có DT 511", len(lines_511) + len(lines_632), 0, "Không", "Đủ", "511 và 632 từ sổ chi tiết", "Đã tách TK đúng", "Khớp BCTC [01]/[11] các năm có số", "DN")
set_in_row(ws, 22, "Excel TMS PTDN", "2016-2025", "Không có bảng nợ thuế/chậm nộp", 2, 0, "Không", "Mỏng", "PTDN01 CĐKT, PTDN02 KQKD", "Nhận file", "LS01/LS02 chưa đủ", "CQT")

# counts
ws["B4"] = 8
ws["C4"] = 4
ws["D4"] = 1
ws["E4"] = 1

ws = wb["KQ1_Kha_nang_phan_tich"]
ws["B2"] = MST
ws["F2"] = KY
ws["B4"] = VER
ws["F4"] = MAU_VER
ws["B5"] = TODAY
ws["B5"].number_format = DATE_FMT
ws["F5"] = "Mức E có khoảng trống IN11/IN09/IN12/IN14"

# rule capability from col A
g3 = {"DQ01", "GT06", "GT16", "TN04", "TN17", "TN18", "TT02", "TT03", "TT04", "TT05", "TT06", "KH01", "KH02", "KH03", "KH04", "KH05", "KH06", "KH07", "DT01", "DT02", "DT03", "LD02", "LD05"}
run_now = 0
part = 0
block = 0
for r in range(10, (ws.max_row or 10) + 1):
    ma = ws.cell(r, 1).value
    if not ma:
        continue
    ma = str(ma).strip()
    need = str(ws.cell(r, 4).value or "")
    if ma == "HD07":
        ws.cell(r, 8).value = "TREO"
        ws.cell(r, 9).value = "Thiếu tham chiếu ngành (K13)"
        ws.cell(r, 10).value = "K13"
        ws.cell(r, 11).value = "Chạy sơ bộ khi có số ngành"
        block += 1
    elif ma in {"LS01", "LS02"}:
        ws.cell(r, 8).value = "Chạy một phần"
        ws.cell(r, 9).value = "TMS PTDN có file nhưng chưa có nợ thuế"
        ws.cell(r, 10).value = "IN14"
        ws.cell(r, 11).value = "Xuất nghĩa vụ nộp thuế TMS"
        part += 1
    elif ma in {"LK01", "LK02"}:
        ws.cell(r, 8).value = "Chạy được ngay"
        ws.cell(r, 9).value = ""
        ws.cell(r, 10).value = "IN12"
        ws.cell(r, 11).value = "Không thấy PL GDLK trên hồ sơ hiện có"
        run_now += 1
    elif ma in g3 and ("IN11" in need and "2023.Q4" ):
        if ma in {"GT16", "TT02", "TT04", "TT06"}:
            ws.cell(r, 8).value = "Chạy một phần"
            ws.cell(r, 9).value = "Sao kê Excel 2023.Q4-2025.Q4, thiếu trước đó"
            ws.cell(r, 10).value = "IN11"
            ws.cell(r, 11).value = "Xin sao kê 2021-2023.Q3"
            part += 1
        elif ma.startswith("KH") or ma == "TN04":
            ws.cell(r, 8).value = "Chạy một phần"
            ws.cell(r, 9).value = "Thiếu XNT 2022"
            ws.cell(r, 10).value = "IN09"
            ws.cell(r, 11).value = "Xin sổ kho 2022"
            part += 1
        else:
            ws.cell(r, 8).value = "Chạy được ngay"
            ws.cell(r, 9).value = ""
            ws.cell(r, 10).value = ""
            ws.cell(r, 11).value = ""
            run_now += 1
    else:
        ws.cell(r, 8).value = "Chạy được ngay"
        ws.cell(r, 9).value = ""
        ws.cell(r, 10).value = ""
        ws.cell(r, 11).value = ""
        run_now += 1
    tnr(ws.cell(r, 8))
    tnr(ws.cell(r, 9))
    tnr(ws.cell(r, 10))
    tnr(ws.cell(r, 11))
ws["B3"] = run_now
ws["E3"] = part
ws["F3"] = block

ws = wb["KQ1_De_nghi_bo_sung"]
ws["B2"] = MST
ws["F2"] = KY
ws["B3"] = VER
ws["F3"] = TODAY
ws["F3"].number_format = DATE_FMT
# fill suggestion rows starting 9
sug = [
    (1, "Sổ XNT kho năm 2022 (TK 152/155/156)", "IN09", "Mở KH01-KH07, TN04 năm 2022", "Ưu tiên 1", "Thiếu đúng 1 năm giữa chuỗi 2021-2025"),
    (2, "Sao kê ngân hàng Excel/PDF 2021 đến 2023.Q3 và UNC các món >= 20 triệu", "IN11", "Mở GT16, TT02, TT04, TT06", "Ưu tiên 2", "Hiện chỉ có Eximbank 2023.Q4-2025.Q4"),
    (3, "Phụ lục GDLK NĐ 132 (nếu có quan hệ liên kết) hoặc xác nhận không phát sinh", "IN12", "Chốt LK01/LK02", "Ưu tiên 3", "Chưa thấy PL trên hồ sơ"),
    (4, "Bảng kê bán ra các kỳ thiếu: 2022, 2023 Q1-Q3, 2025 Q3", "IN02", "GT02/GT03 đủ kỳ", "Ưu tiên 4", "Hiện dựa HĐĐT + tờ khai"),
    (5, "File TMS nghĩa vụ nộp thuế / nợ / chậm nộp (không chỉ PTDN CĐKT-KQKD)", "IN14", "LS01, LS02", "Ưu tiên 5", "PTDN01/02 không thay bảng nợ thuế"),
    (6, "Giải trình lệch HĐĐT bán ra vs tờ khai GTGT 2024 (~977 triệu) và 2025 (~270 triệu)", "IN01/IN03", "GT01", "Ưu tiên 1b", "Chắc chắn có chênh lệch nguồn CQT"),
]
# find start
start = 9
for r in range(8, 20):
    if ws.cell(r, 1).value in {1, "1", "STT"}:
        start = r + 1 if ws.cell(r, 1).value == "STT" else r
        break
for i, rowv in enumerate(sug):
    for c, v in enumerate(rowv, 1):
        ws.cell(start + i, c).value = v
        tnr(ws.cell(start + i, c), bold=(c == 5))

wb.save(lp(files["kq13"][1]))
wb.close()
print(" KQ1.3 saved")
print("NEXT KQ2")
Path(r"C:\Users\NGOC\AppData\Local\Temp\htp_fill_part1_ok.txt").write_text("ok", encoding="utf-8")
