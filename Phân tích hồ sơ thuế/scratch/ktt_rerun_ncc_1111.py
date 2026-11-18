# -*- coding: utf-8 -*-
"""Chay lai NCC/IN07/HD01-HD05 Kiến Thành Tân theo nguyên tắc 11/11 (2026-11-18)."""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603216935"
TEN = "CÔNG TY TNHH MỘT THÀNH VIÊN KIẾN THÀNH TÂN"
KY = "2015-2025"
TODAY = date(2026, 11, 18)
VER = "v2"
P13, P14 = 0.05, 0.50

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\1. TD\14. KIEN THANH TAN_3603216935"
)
OUT = ROOT / "4. KET QUA PHAN TICH"
VAULT_OUT = VAULT / "ket-qua" / "KIEN_THANH_TAN_3603216935"
KQ11 = VAULT_OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_v2.xlsx"
KQ12 = VAULT_OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_v1.xlsx"
KQ13 = VAULT_OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_v2.xlsx"
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
DANHBA = ROOT / "3. DU LIEU BO SUNG & TRA CUU QLT" / "3.2. Danh ba trang thai NNT" / "Danh ba NNT.xls"

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
FILL_NONE = PatternFill(fill_type=None)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
NUM = "#,##0"
PCT = "0.00%"
DATE_FMT = "dd/mm/yyyy"
SS_NS = "urn:schemas-microsoft-com:office:spreadsheet"


def lp(p) -> str:
    s = os.path.abspath(str(p))
    if s.startswith("\\\\?\\"):
        return s
    if s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + s[2:]
    return "\\\\?\\" + s


def nnum(v):
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
    s = str(v).strip().replace("00:00:00", "").strip()
    if "T" in s:
        s = s.split("T", 1)[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    try:
        n = float(str(v).replace(",", "."))
        if 20000 < n < 80000:
            return date(1899, 12, 30) + timedelta(days=int(n))
    except ValueError:
        pass
    return None


def mst_base(s) -> str:
    s = str(s or "").split("-")[0]
    return re.sub(r"\D", "", s)


def so_key(v) -> str:
    s = str(v or "").strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.lstrip("0") or "0"


def vn(n):
    if n is None:
        return ""
    return f"{n:,.0f}".replace(",", ".")


def tnr(cell, bold=None):
    f = cell.font
    cell.font = Font(
        name="Times New Roman",
        size=f.size if f and f.size else 11,
        bold=f.bold if bold is None and f else (bool(bold) if bold is not None else False),
        italic=False,
        color=f.color if f else None,
    )


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


def autowidth(ws, cap=60, min_w=10, max_row=None):
    widths = {}
    mr = min(ws.max_row or 1, max_row or 4000)
    for row in ws.iter_rows(max_row=mr, max_col=ws.max_column or 1):
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            widths[cell.column] = max(widths.get(cell.column, min_w), min(len(s), cap) + 2)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


def copy_file(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(lp(src), lp(dst))
    print(" COPY", src.name, "->", dst)


# ---------------------------------------------------------------------------
# LOAD IN01 / IN02
# ---------------------------------------------------------------------------
print("LOAD IN01")
wb = load_workbook(lp(KQ11), read_only=True, data_only=True)
ws = wb["IN01_HDDT"]
hddt_raw = []
seen_sig = set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    ch = str(row[0]).upper()
    nam = int(row[1]) if row[1] else None
    if nam == 2026:
        continue
    rec = {
        "chieu": "MUA" if ch.startswith("M") else "BAN",
        "nam": nam,
        "kyhieu": str(row[4] or "").strip(),
        "so": so_key(row[5]),
        "so_goc": row[5],
        "ngay": ndate(row[6]),
        "mst_nb": mst_base(row[7]),
        "ten_nb": str(row[8] or "").strip(),
        "mst_dt": mst_base(row[11]),
        "mh": str(row[12] or "").strip(),
        "gt": nnum(row[16]),
        "thue": nnum(row[18]),
    }
    sig = (
        rec["chieu"], rec["nam"], rec["kyhieu"], rec["so"], rec["mst_dt"],
        rec["mh"], round(rec["gt"], 2), round(rec["thue"], 2),
    )
    if sig in seen_sig:
        continue
    seen_sig.add(sig)
    hddt_raw.append(rec)
wb.close()
mua_ok = [x for x in hddt_raw if x["chieu"] == "MUA"]
print(" HDDT unique mua", len(mua_ok), "all", len(hddt_raw))

print("LOAD IN02")
wb = load_workbook(lp(KQ12), read_only=True, data_only=True)
ws = wb["IN02_BangKe"]
bk_raw = []
seen_bk = set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    ch = str(row[0]).upper()
    nam = int(row[1]) if row[1] else None
    if nam == 2026:
        continue
    rec = {
        "chieu": "MUA" if ch.startswith("M") else "BAN",
        "nam": nam,
        "kyhieu": str(row[4] or "").strip(),
        "so": so_key(row[5]),
        "ngay": ndate(row[6]),
        "mst": mst_base(row[8] or row[7]),
        "ten": str(row[9] or "").strip(),
        "mh": str(row[10] or "").strip(),
        "gt": nnum(row[13]),
        "thue": nnum(row[15]),
    }
    sig = (
        rec["chieu"], rec["nam"], rec["kyhieu"], rec["so"], rec["mst"],
        rec["mh"], round(rec["gt"], 2), round(rec["thue"], 2),
    )
    if sig in seen_bk:
        continue
    seen_bk.add(sig)
    bk_raw.append(rec)
wb.close()
bk_mua = [x for x in bk_raw if x["chieu"] == "MUA"]
print(" BK unique mua", len(bk_mua), "all", len(bk_raw))

# ---------------------------------------------------------------------------
# F0 / F1 / CA
# ---------------------------------------------------------------------------
print("LOAD CVCBRR")
f0_map, f1_map, ca_map = {}, {}, {}
f1_buyers_nnt = []


def load_f0_f1():
    if not os.path.exists(lp(CVCBRR)):
        print(" CVCBRR missing")
        return
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    if "F0" in wb.sheetnames:
        ws = wb["F0"]
        for row in ws.iter_rows(min_row=8, values_only=True):
            m = mst_base(row[1] if row and len(row) > 1 else "")
            if not m or not m[0].isdigit():
                continue
            if m not in f0_map:
                f0_map[m] = {
                    "mst": m,
                    "ten": str(row[2] or "").strip() if len(row) > 2 else "",
                    "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                    "ngay_cv": ndate(row[4] if len(row) > 4 else None),
                    "dv": str(row[5] or "").strip() if len(row) > 5 else "",
                    "nd": str(row[6] or "").strip()[:240] if len(row) > 6 else "",
                }
    if "F1" in wb.sheetnames:
        ws = wb["F1"]
        for row in ws.iter_rows(min_row=8, values_only=True):
            m_ban = mst_base(row[1] if row and len(row) > 1 else "")
            m_mua = mst_base(row[8] if row and len(row) > 8 else "")
            if not m_mua:
                continue
            rec = {
                "mst_ban": m_ban,
                "ten_ban": str(row[2] or "").strip() if len(row) > 2 else "",
                "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                "ngay_cv": ndate(row[4] if len(row) > 4 else None),
                "ten_mua": str(row[7] or "").strip() if len(row) > 7 else "",
                "mst_mua": m_mua,
            }
            f1_map.setdefault(m_mua, []).append(rec)
            if m_mua == MST:
                f1_buyers_nnt.append(rec)
    wb.close()


def load_ca():
    if not os.path.exists(lp(CA)):
        print(" CA missing")
        return
    wb = load_workbook(lp(CA), read_only=True, data_only=True)
    for name in wb.sheetnames:
        if str(name).upper() in {"DASHBOARD", "DANH_MUC"}:
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, row in enumerate(rows[:20]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "mst" in joined or "mã số thuế" in joined:
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [str(c or "").strip().lower() for c in rows[hdr_i]]
        i_mst = next((i for i, h in enumerate(hdr) if "mst" in h or "mã số thuế" in h), None)
        i_ten = next((i for i, h in enumerate(hdr) if "tên" in h or "ten" in h), None)
        i_cv = next((i for i, h in enumerate(hdr) if "công văn" in h or "so cv" in h or "số cv" in h), None)
        if i_mst is None:
            continue
        for row in rows[hdr_i + 1:]:
            m = mst_base(row[i_mst] if i_mst < len(row) else "")
            if not m:
                continue
            if m not in ca_map:
                ca_map[m] = {
                    "mst": m,
                    "ten": str(row[i_ten] or "").strip() if i_ten is not None and i_ten < len(row) else "",
                    "cv": str(row[i_cv] or "").strip() if i_cv is not None and i_cv < len(row) else "",
                    "ngay_cv": None,
                    "nd": f"Sổ Công an sheet {name}",
                }
    wb.close()


load_f0_f1()
load_ca()
nnt_on_f0 = MST in f0_map
nnt_on_f1 = MST in f1_map
nnt_on_ca = MST in ca_map
cvcbrr_mtime = datetime.fromtimestamp(os.path.getmtime(lp(CVCBRR))).date() if os.path.exists(lp(CVCBRR)) else None
print(
    " F0", len(f0_map), "F1 buyers", len(f1_map), "CA", len(ca_map),
    "NNT F0/F1/CA", nnt_on_f0, nnt_on_f1, nnt_on_ca,
    "NNT F1 hits", len(f1_buyers_nnt), "CVCBRR", cvcbrr_mtime,
)

# ---------------------------------------------------------------------------
# Danh ba SpreadsheetML
# ---------------------------------------------------------------------------
print("LOAD DANHBA", DANHBA.exists(), DANHBA)
dkkd = {}


def dkkd_bat_thuong(tt: str) -> bool:
    s = (tt or "").strip().lower()
    if not s:
        return False
    if "đang hoạt động" in s or "dang hoat dong" in s:
        return False
    return True


def ss_index(cell, default):
    for key, val in cell.attrib.items():
        if key == "Index" or key.endswith("}Index"):
            return int(val)
    return default


def ss_text(cell):
    data = None
    for child in list(cell):
        tag = child.tag
        if tag == "Data" or tag.endswith("}Data"):
            data = child
            break
    if data is None:
        return "", ""
    typ = ""
    for key, val in data.attrib.items():
        if key == "Type" or key.endswith("}Type"):
            typ = val
    return (data.text or "").strip(), typ


def load_danhba():
    if not DANHBA.exists():
        print(" Danh ba missing")
        return
    tree = ET.parse(lp(DANHBA))
    root = tree.getroot()
    table = None
    for el in root.iter():
        if el.tag.endswith("}Table") or el.tag == "Table":
            table = el
            break
    if table is None:
        print(" Danh ba no Table")
        return
    rows = []
    for row in list(table):
        if not (row.tag.endswith("}Row") or row.tag == "Row"):
            continue
        cells = {}
        idx = 1
        for cell in list(row):
            if not (cell.tag.endswith("}Cell") or cell.tag == "Cell"):
                continue
            idx = ss_index(cell, idx)
            txt, typ = ss_text(cell)
            cells[idx] = (txt, typ)
            idx += 1
        if cells:
            rows.append(cells)
    print(" Danh ba rows", len(rows))
    if len(rows) < 3:
        return
    # Row 1 = title, row 2 = header. MST = col 4, ten = 5, nganh = 7,
    # dia chi = 8, giam doc = 19, ngay cap GCN = 33, tt to chuc = 45,
    # ly do dong to chuc = 53, ngay dong to chuc = 54, ngay BD KD = 65.
    for cells in rows[2:]:
        mst_raw = cells.get(4, ("", ""))[0]
        m = mst_base(mst_raw)
        if not m:
            continue
        tt = cells.get(45, ("", ""))[0]
        rec = {
            "mst_raw": mst_raw,
            "ten": cells.get(5, ("", ""))[0],
            "nganh": cells.get(7, ("", ""))[0],
            "dia_chi": cells.get(8, ("", ""))[0],
            "dd": cells.get(19, ("", ""))[0],
            "ngay_cap": ndate(cells.get(33, ("", ""))[0]),
            "tt": tt,
            "ly_do_dong": cells.get(53, ("", ""))[0],
            "ngay_dong": ndate(cells.get(54, ("", ""))[0]),
            "ngay_bd": ndate(cells.get(65, ("", ""))[0]),
        }
        old = dkkd.get(m)
        if old is None:
            dkkd[m] = rec
            continue
        old_ab = dkkd_bat_thuong(old.get("tt"))
        new_ab = dkkd_bat_thuong(tt)
        old_hyphen = "-" in (old.get("mst_raw") or "")
        new_hyphen = "-" in (mst_raw or "")
        if old_ab and not new_ab:
            dkkd[m] = rec
        elif old_hyphen and not new_hyphen and not new_ab:
            dkkd[m] = rec


load_danhba()
print(" DKKD map", len(dkkd), "NNT in map", MST in dkkd, "NNT tt", (dkkd.get(MST) or {}).get("tt"))


def ncc_flag(base: str):
    """CA -> F0 CVCBRR -> F1 (NCC là bên mua trên sheet F1). Cả F0 và F1 thì ghi F0."""
    if not base:
        return "", "", "", "", {}
    if base in ca_map:
        rec = ca_map[base]
        return "F0", "CA", rec.get("cv", ""), rec.get("nd", ""), rec
    if base in f0_map:
        rec = f0_map[base]
        return "F0", "CANH_BAO", rec.get("cv", ""), rec.get("nd", ""), rec
    if base in f1_map:
        hits = f1_map[base]
        rec0 = hits[0]
        bits = []
        for h in hits[:3]:
            if h.get("ten_ban") or h.get("mst_ban"):
                bits.append(f"{h.get('ten_ban','')} MST {h.get('mst_ban','')}".strip())
            elif h.get("cv"):
                bits.append(f"CV {h.get('cv')}")
            else:
                bits.append("dòng F1 nguồn CA, MST bán trống")
        nd = "F1 - NCC là bên mua dùng HĐ F0 (" + "; ".join(bits) + ")"
        return "F1", "CANH_BAO", rec0.get("cv", ""), nd, rec0
    return "", "", "", "", {}


# ---------------------------------------------------------------------------
# Invoice union: Bảng kê | HĐĐT | Chung ; Chung lấy tiền HĐĐT
# ---------------------------------------------------------------------------
hddt_keys = set()
bk_keys = set()
for inv in mua_ok:
    m = inv["mst_nb"] or inv["mst_dt"]
    hddt_keys.add((m, inv["so"], inv["ngay"]))
for inv in bk_mua:
    bk_keys.add((inv["mst"], inv["so"], inv["ngay"]))

mua_inv = []
for inv in mua_ok:
    m = inv["mst_nb"] or inv["mst_dt"]
    k = (m, inv["so"], inv["ngay"])
    src = "Chung" if k in bk_keys else "HĐĐT"
    mua_inv.append({
        "src": src, "mst": m, "ten": inv["ten_nb"],
        "so": inv["so_goc"] if inv["so_goc"] is not None else inv["so"],
        "ngay": inv["ngay"], "mh": inv["mh"], "gt": inv["gt"], "thue": inv["thue"],
        "nam": inv["nam"],
    })
for inv in bk_mua:
    k = (inv["mst"], inv["so"], inv["ngay"])
    if k in hddt_keys:
        continue
    mua_inv.append({
        "src": "Bảng kê", "mst": inv["mst"], "ten": inv["ten"],
        "so": inv["so"], "ngay": inv["ngay"], "mh": inv["mh"],
        "gt": inv["gt"], "thue": inv["thue"], "nam": inv["nam"],
    })
mua_inv.sort(key=lambda x: (x["ngay"] or date(1900, 1, 1), str(x["so"]), x["mst"] or ""))

ncc_agg = defaultdict(lambda: {
    "ten": "", "gt": 0.0, "thue": 0.0, "n": 0, "src": set(),
    "first": None, "n_hd": set(),
})
src_count = Counter()
for inv in mua_inv:
    src_count[inv["src"]] += 1
    k = inv["mst"]
    if not k:
        continue
    a = ncc_agg[k]
    a["ten"] = a["ten"] or inv["ten"]
    a["gt"] += inv["gt"]
    a["thue"] += inv["thue"]
    a["n"] += 1
    a["src"].add(inv["src"])
    a["n_hd"].add((inv["so"], inv["ngay"]))
    if inv["ngay"] and (a["first"] is None or inv["ngay"] < a["first"]):
        a["first"] = inv["ngay"]

tot_mua_gt = sum(a["gt"] for a in ncc_agg.values())
ranked = sorted(ncc_agg.items(), key=lambda x: x[1]["gt"], reverse=True)
top5 = ranked[:5]
top5_gt = sum(a["gt"] for _, a in top5)
top5_pct = (top5_gt / tot_mua_gt) if tot_mua_gt else 0.0
f0_ncc = [(k, a) for k, a in ranked if ncc_flag(k)[0] == "F0"]
f1_ncc = [(k, a) for k, a in ranked if ncc_flag(k)[0] == "F1"]
f0_gt = sum(a["gt"] for _, a in f0_ncc)
f0_thue = sum(a["thue"] for _, a in f0_ncc)
f1_gt = sum(a["gt"] for _, a in f1_ncc)
hd03_pct = ((f0_gt + f1_gt) / tot_mua_gt) if tot_mua_gt else 0.0

hd02_hits = []
hd05_hits = []
ncc_no_dkkd = 0
for k, a in ranked:
    rec = dkkd.get(k)
    if not rec:
        ncc_no_dkkd += 1
        continue
    if dkkd_bat_thuong(rec.get("tt")):
        hd02_hits.append((k, a, rec))
    ngay_tl = rec.get("ngay_cap") or rec.get("ngay_bd")
    first = a.get("first")
    if ngay_tl and first and (first - ngay_tl).days < 365 and (first - ngay_tl).days >= 0:
        hd05_hits.append((k, a, rec, (first - ngay_tl).days))

hd01_hit = bool(nnt_on_f1 or f0_ncc)
print(
    "NCC", len(ncc_agg), "inv", len(mua_inv), dict(src_count),
    "F0", len(f0_ncc), vn(f0_gt), "F1", len(f1_ncc), vn(f1_gt),
    "top5", f"{top5_pct:.2%}", "hd03", f"{hd03_pct:.2%}",
    "HD01", hd01_hit, "HD02", len(hd02_hits), "HD05", len(hd05_hits),
    "NCC not in danh ba", ncc_no_dkkd,
)

# ---------------------------------------------------------------------------
# Copy v1 -> v2
# ---------------------------------------------------------------------------
files = {
    "kq21": (
        VAULT_OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_v1.xlsx",
        VAULT_OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_{VER}.xlsx",
    ),
    "kq3": (
        VAULT_OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_v1.xlsx",
        VAULT_OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_{VER}.xlsx",
    ),
    "dau": (
        VAULT_OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}_v1.xlsx",
        VAULT_OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}_{VER}.xlsx",
    ),
    "giai": (
        VAULT_OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}_v1.xlsx",
        VAULT_OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}_{VER}.xlsx",
    ),
}
for src, dst in files.values():
    copy_file(src, dst)

hdr2 = (
    f"Người nộp thuế: {TEN} - MST: {MST} - Kỳ: {KY} | Phiên bản KQ: {VER} | "
    f"{TODAY.strftime('%d/%m/%Y')} | nguyên tắc F0/F1 11/11"
)

nnt_f0_sellers = f1_map.get(MST) or []
if nnt_on_f1:
    nnt_phan1 = (
        f"NNT {TEN} (MST {MST}) nằm F1 - đã mua HĐ từ F0 "
        + "; ".join(f"{h['ten_ban']} MST {h['mst_ban']} CV {h['cv']}" for h in nnt_f0_sellers[:5])
    )
    nnt_kl = "có cảnh báo - HD01 mức Cao"
elif nnt_on_f0:
    rec = f0_map[MST]
    nnt_phan1 = f"NNT {TEN} (MST {MST}) bị cảnh báo rủi ro thuế - Công văn {rec.get('cv')}"
    nnt_kl = "mục tiêu kiểm tra lớn - NNT F0"
elif nnt_on_ca:
    rec = ca_map[MST]
    nnt_phan1 = f"NNT {TEN} (MST {MST}) có trên sổ Công an - {rec.get('cv')}"
    nnt_kl = "có cảnh báo CA"
else:
    nnt_phan1 = (
        f"NNT {TEN} (MST {MST}) không có trên F0; không trên F1 (cột MST người mua); "
        "không trên sổ Công an. Không kích HD01 từ nhánh NNT."
    )
    nnt_kl = "an toàn (nhánh NNT)"

nnt_dkkd = dkkd.get(MST) or {}
nnt_tt = nnt_dkkd.get("tt") or "Không có trong danh bạ"


def src_label(srcset):
    s = set(srcset)
    if "Chung" in s or ({"Bảng kê", "HĐĐT"} <= s):
        return "Chung"
    if s == {"HĐĐT"}:
        return "HĐĐT"
    if s == {"Bảng kê"}:
        return "Bảng kê"
    return "/".join(sorted(s)) or ""


# ---------------------------------------------------------------------------
# FILL KQ2.1 IN07 + NCC
# ---------------------------------------------------------------------------
print("FILL KQ2.1")
wb = load_workbook(lp(files["kq21"][1]))
for sn in wb.sheetnames:
    for row in wb[sn].iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and "Phiên bản KQ: v1" in cell.value:
                cell.value = cell.value.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")
                tnr(cell)

ws = wb["IN07_NCC_CanhBao"]
if ws.max_row and ws.max_row >= 3:
    ws.delete_rows(3, ws.max_row - 2)
r = 3
put(ws, r, 1, "PHẦN 1. ĐÁNH GIÁ RỦI RO NNT (bắt buộc, kể cả kết quả âm)", bold=True, fill=FILL_SEC)
for c in range(2, 19):
    put(ws, r, c, "", fill=FILL_SEC)
r += 1
put(ws, r, 1, MST)
put(ws, r, 2, TEN)
put(ws, r, 3, "Phần 1 - NNT đang kiểm tra")
put(ws, r, 4, "F0" if nnt_on_f0 or nnt_on_ca else ("F1" if nnt_on_f1 else "Không"))
put(ws, r, 5, "CA" if nnt_on_ca else ("CANH_BAO" if (nnt_on_f0 or nnt_on_f1) else ""))
if nnt_on_ca:
    rec = ca_map[MST]
    put(ws, r, 6, rec.get("cv", ""))
    put(ws, r, 8, "Công an")
    put(ws, r, 9, rec.get("nd", ""))
elif nnt_on_f0:
    rec = f0_map[MST]
    put(ws, r, 6, rec.get("cv", ""))
    put(ws, r, 7, rec.get("ngay_cv"), dt=True)
    put(ws, r, 8, rec.get("dv") or "Chi cục Thuế")
    put(ws, r, 9, rec.get("nd", ""))
elif nnt_on_f1:
    rec = nnt_f0_sellers[0]
    put(ws, r, 6, rec.get("cv", ""))
    put(ws, r, 7, rec.get("ngay_cv"), dt=True)
    put(ws, r, 8, "Chi cục Thuế")
    put(ws, r, 9, nnt_phan1)
put(ws, r, 10, nnt_tt)
put(ws, r, 11, nnt_dkkd.get("ngay_dong"), dt=True)
put(ws, r, 12, nnt_dkkd.get("ngay_cap"), dt=True)
put(ws, r, 13, nnt_dkkd.get("nganh", ""))
put(ws, r, 14, nnt_dkkd.get("dia_chi", ""), wrap=True)
put(ws, r, 15, nnt_dkkd.get("dd", ""))
put(ws, r, 16, "Sổ CVCBRR F0/F1 + CA + Danh ba NNT.xls")
put(ws, r, 17, nnt_phan1 + f" Kết luận nhánh NNT: {nnt_kl}.", wrap=True)
put(ws, r, 18, str(CVCBRR.name) + " | Danh ba NNT.xls")
if nnt_on_f0 or nnt_on_ca:
    for c in range(1, 19):
        ws.cell(r, c).fill = FILL_RED
        ws.cell(r, c).font = FONT_RED
elif nnt_on_f1:
    for c in range(1, 19):
        ws.cell(r, c).fill = FILL_YEL
        ws.cell(r, c).font = FONT_ORG
else:
    for c in range(1, 19):
        ws.cell(r, c).fill = FILL_GRN
r += 1
put(ws, r, 1, "PHẦN 2. NCC F0 / F1 / CA / ĐKKD bất thường", bold=True, fill=FILL_SEC)
for c in range(2, 19):
    put(ws, r, c, "", fill=FILL_SEC)
r += 1
in07_ncc = 0
seen = {MST}
# F0/F1/CA first, then ĐKKD bất thường
ordered = []
for k, a in ranked:
    nhom, nguon, cv, nd, rec = ncc_flag(k)
    if nhom:
        ordered.append((k, a, nhom, nguon, cv, nd, rec, "F"))
for k, a, rec in hd02_hits:
    if k in seen or ncc_flag(k)[0]:
        continue
    ordered.append((k, a, "", "DKKD", "", rec.get("tt", ""), rec, "D"))

for k, a, nhom, nguon, cv, nd, rec, kind in ordered:
    if k in seen:
        continue
    seen.add(k)
    db = dkkd.get(k) or {}
    put(ws, r, 1, k)
    put(ws, r, 2, a["ten"] or rec.get("ten") or db.get("ten", ""))
    put(ws, r, 3, "NCC mua vào")
    put(ws, r, 4, nhom or ("ĐKKD" if kind == "D" else ""))
    put(ws, r, 5, nguon)
    put(ws, r, 6, cv or rec.get("cv", ""))
    put(ws, r, 7, rec.get("ngay_cv"), dt=True)
    put(ws, r, 8, "Công an" if nguon == "CA" else ("Chi cục Thuế" if nhom else "Danh bạ"))
    put(ws, r, 9, nd or rec.get("nd") or db.get("tt", ""), wrap=True)
    put(ws, r, 10, db.get("tt") or "Không có trong danh bạ")
    put(ws, r, 11, db.get("ngay_dong"), dt=True)
    put(ws, r, 12, db.get("ngay_cap") or db.get("ngay_bd"), dt=True)
    put(ws, r, 13, db.get("nganh", ""))
    put(ws, r, 14, db.get("dia_chi", ""), wrap=True)
    put(ws, r, 15, db.get("dd", ""))
    put(ws, r, 16, "HĐĐT/BK + CVCBRR/CA + Danh ba NNT.xls")
    put(ws, r, 17, f"GT mua {vn(a['gt'])}; nguồn {src_label(a['src'])}", wrap=True)
    put(ws, r, 18, "IN01+IN02+CVCBRR+CA+Danh ba")
    if nhom == "F0":
        fill, font = FILL_RED, FONT_RED
    elif nhom == "F1" or kind == "D":
        fill, font = FILL_YEL, FONT_ORG
    else:
        fill, font = None, FONT
    if fill:
        for c in range(1, 19):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = font
    r += 1
    in07_ncc += 1
if in07_ncc == 0:
    put(ws, r, 1, "Không có NCC trên F0, F1, sổ Công an hoặc ĐKKD bất thường.")
    put(ws, r, 17, "Đã tra đủ sổ chỉ định + danh bạ 3.2.")
    r += 1
r += 1
put(ws, r, 1, "PHẦN 3. NHẬN XÉT", bold=True, fill=FILL_SEC)
r += 1
top5_txt = "; ".join(f"{k} {a['ten'][:40]} {vn(a['gt'])}" for k, a in top5)
phan3 = (
    f"(1) Rủi ro NNT: {nnt_kl}. {nnt_phan1} "
    f"(2) Rủi ro NCC: F0={len(f0_ncc)} GT {vn(f0_gt)}; F1 tầng hai={len(f1_ncc)} GT {vn(f1_gt)}; "
    f"ĐKKD bất thường={len(hd02_hits)}; NCC mới <12 tháng={len(hd05_hits)}. "
    f"HD01={'Có' if hd01_hit else 'Không'} (chỉ NNT F1 hoặc mua NCC F0, không kích vì NCC F1). "
    f"HD03 F0+F1={hd03_pct:.2%} (P13=5%). Top-5={top5_pct:.2%}. "
    f"(3) Trọng tâm chứng từ: chuỗi 5 mắt xích với NCC F0 (nếu có); NCC F1 giữ mức Cao cho rủi ro NCC/HD03; "
    f"NCC ĐKKD bất thường; Top-5 {top5_txt}."
)
put(ws, r, 1, phan3, wrap=True)
try:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=18)
except Exception:
    pass
ws.row_dimensions[r].height = 72
autowidth(ws, cap=36, max_row=40)
print("IN07 NCC rows", in07_ncc)

ws = wb["KQ2_NCC_mua_vao"]
put(ws, 2, 1, hdr2, bold=True)
if ws.max_row > 8:
    ws.delete_rows(8, ws.max_row - 7)
r = 8
for i, inv in enumerate(mua_inv, 1):
    nhom, nguon, cv, nd, rec = ncc_flag(inv["mst"])
    db = dkkd.get(inv["mst"]) or {}
    tt = db.get("tt") or "Không có trong danh bạ"
    put(ws, r, 1, i)
    put(ws, r, 2, inv["mst"])
    put(ws, r, 3, inv["ten"])
    put(ws, r, 4, inv["so"])
    put(ws, r, 5, inv["ngay"], dt=True)
    put(ws, r, 6, inv["mh"])
    put(ws, r, 7, inv["gt"], num=True)
    put(ws, r, 8, inv["thue"], num=True)
    put(ws, r, 9, inv["src"])
    put(ws, r, 10, f"Có - {cv}" if cv else ("Có" if nhom else "Không"))
    put(ws, r, 11, nhom or "Không")
    put(ws, r, 12, tt)
    if nhom == "F0":
        for c in range(1, 13):
            ws.cell(r, c).fill = FILL_RED
            ws.cell(r, c).font = FONT_RED
    elif nhom == "F1":
        for c in range(1, 13):
            ws.cell(r, c).fill = FILL_YEL
            ws.cell(r, c).font = FONT_ORG
    elif dkkd_bat_thuong(tt):
        for c in range(1, 13):
            ws.cell(r, c).fill = FILL_YEL
    r += 1
last_inv = r - 1
put(ws, 5, 2, hd03_pct, pct=True)
put(ws, 5, 5, top5_pct, pct=True)
rr = last_inv + 2
put(ws, rr, 1, "Nhận xét module (3.4) - không phải kết luận hồ sơ", bold=True, fill=FILL_SEC)
note = (
    f"Nguyên tắc 11/11: F0/F1 NNT trước, rồi rà NCC. NNT F0={'Có' if nnt_on_f0 else 'Không'}; "
    f"F1={'Có' if nnt_on_f1 else 'Không'} ({len(f1_buyers_nnt)} dòng F1 MST mua); "
    f"CA={'Có' if nnt_on_ca else 'Không'}. "
    f"HD01={'Có phát hiện' if hd01_hit else 'Không'} (NNT F1 hoặc NCC F0; không kích vì NCC F1). "
    f"NCC={len(ncc_agg)}. HĐ union {len(mua_inv)} dòng (HĐĐT {src_count['HĐĐT']}, "
    f"Bảng kê {src_count['Bảng kê']}, Chung {src_count['Chung']}). "
    f"F0={len(f0_ncc)} GT {vn(f0_gt)} thuế {vn(f0_thue)}. F1 tầng hai={len(f1_ncc)} GT {vn(f1_gt)}. "
    f"HD03 F0+F1={hd03_pct:.2%} (P13=5%). HD04 Top-5={top5_pct:.2%} (P14=50%). "
    f"HD02 ĐKKD bất thường={len(hd02_hits)} MST. HD05 NCC mới <12 tháng={len(hd05_hits)}. "
    f"Danh bạ SpreadsheetML {len(dkkd)} MST; {ncc_no_dkkd} NCC không có trong danh bạ. "
    f"Top-5: {top5_txt}."
)
put(ws, rr + 1, 1, note, wrap=True)
try:
    ws.merge_cells(start_row=rr + 1, start_column=1, end_row=rr + 1, end_column=12)
except Exception:
    pass
ws.row_dimensions[rr + 1].height = 90
ws.auto_filter.ref = f"A7:L{max(7, last_inv)}"
autowidth(ws, cap=36, max_row=30)
print("NCC invoice rows", last_inv - 7)
wb.save(lp(files["kq21"][1]))
wb.close()
print(" KQ2.1 saved")

# ---------------------------------------------------------------------------
# PATCH KQ3
# ---------------------------------------------------------------------------
print("PATCH KQ3")
wb = load_workbook(lp(files["kq3"][1]))
ws = wb["KQ3"] if "KQ3" in wb.sheetnames else wb[wb.sheetnames[0]]
for row in ws.iter_rows():
    for cell in row:
        if isinstance(cell, MergedCell):
            continue
        if isinstance(cell.value, str) and "Phiên bản KQ: v1" in cell.value:
            cell.value = cell.value.replace("Phiên bản KQ: v1", f"Phiên bản KQ: {VER}")

hm = {str(ws.cell(9, c).value or "").strip(): c for c in range(1, 40) if ws.cell(9, c).value}
col_ma = hm.get("Mã quy tắc")
print(" KQ3 header", list(hm)[:12], "col_ma", col_ma)


def find_rule_row(ma):
    for rr in range(10, (ws.max_row or 10) + 1):
        if str(ws.cell(rr, col_ma).value or "").strip() == ma:
            return rr
    return None


def set_rule(ma, **kw):
    rr = find_rule_row(ma)
    if not rr:
        print("  missing rule", ma)
        return
    key_map = {
        "status": "Trạng thái chạy",
        "ket": "Kết quả",
        "nhan": "Nhãn",
        "muc": "Mức rủi ro",
        "hs": "Hệ số trọng yếu",
        "diem": "Điểm",
        "ca": "Chỉ tiêu A",
        "va": "Giá trị A",
        "cb": "Chỉ tiêu B",
        "vb": "Giá trị B",
        "nguong": "Ngưỡng",
        "thieu": "Còn thiếu để kết luận",
        "ghi": "Ghi chú",
        "nd": "Nội dung kiểm tra tại trụ sở",
        "ct_ref": "Tham chiếu chứng cứ",
        "sac": "Sắc thuế",
        "v": "Số tiền liên quan (đồng)",
        "w": "Số tiền tính vào tổng (đồng)",
        "nguon": "Nguồn",
    }
    for kwk, col in key_map.items():
        if kwk not in kw or col not in hm:
            continue
        v = kw[kwk]
        c = hm[col]
        num = isinstance(v, (int, float)) and not isinstance(v, bool) and col not in {"Hệ số trọng yếu", "Điểm"}
        put(ws, rr, c, v, num=num, wrap=col in {"Ghi chú", "Còn thiếu để kết luận", "Nội dung kiểm tra tại trụ sở"})
        if col == "Mức rủi ro":
            cell = ws.cell(rr, c)
            if v == "Cao":
                cell.fill = FILL_RED
                cell.font = FONT_RED
            elif v == "Trung bình":
                cell.fill = FILL_YEL
                cell.font = FONT_ORG
            else:
                cell.fill = FILL_GRN
                cell.font = FONT
        if col == "Trạng thái chạy":
            cell = ws.cell(rr, c)
            if v == "KHÔNG_ĐỦ_DỮ_LIỆU":
                cell.fill = FILL_YEL
            elif v == "TREO":
                cell.fill = FILL_RED
            else:
                cell.fill = FILL_GRN if v == "ĐÃ_CHẠY" else FILL_NONE
                cell.font = FONT


# HD01
if hd01_hit:
    set_rule(
        "HD01",
        status="ĐÃ_CHẠY", ket="Có phát hiện", nhan="Chắc", muc="Cao", hs=3, diem=9,
        ca="NNT F1 / NCC F0", va=len(f0_ncc), cb="GT F0", vb=f0_gt,
        thieu="", sac="GTGT", v=f0_gt if f0_ncc else None, w=f0_gt if f0_ncc else 0,
        ct_ref="IN07 Phần 1 + NCC_mua_vao",
        nd="Chuỗi 5 mắt xích NCC F0 hoặc HĐ F0 mà NNT đã mua",
        ghi=(
            f"Nguyên tắc 11/11: HD01 = NNT là F1 hoặc mua từ NCC F0. "
            f"NNT F0={nnt_on_f0} F1={nnt_on_f1} CA={nnt_on_ca}. F0 NCC={len(f0_ncc)}. "
            f"Không kích HD01 vì NCC F1 (tầng hai)."
        ),
        nguon="IN07, CVCBRR, CA",
    )
else:
    set_rule(
        "HD01",
        status="ĐÃ_CHẠY", ket="Không", nhan="Chắc", muc="", hs=0, diem=0,
        ca="NNT F1 / NCC F0", va=0, cb="GT F0", vb=0,
        thieu="", sac="", v=None, w=0,
        ct_ref="IN07 Phần 1 + NCC_mua_vao",
        nd="",
        ghi=(
            f"Nguyên tắc 11/11: HD01 = NNT là F1 hoặc mua từ NCC F0. "
            f"NNT F0={nnt_on_f0} F1={nnt_on_f1} CA={nnt_on_ca}. F0 NCC={len(f0_ncc)}. "
            f"Có {len(f1_ncc)} NCC F1 tầng hai - xếp Cao cho rủi ro NCC/HD03, không kích HD01."
        ),
        nguon="IN07, CVCBRR, CA",
    )

# HD02
if hd02_hits:
    set_rule(
        "HD02",
        status="ĐÃ_CHẠY", ket="Có phát hiện", nhan="Chắc", muc="Trung bình", hs=2, diem=6,
        ca="NCC ĐKKD bất thường", va=len(hd02_hits),
        thieu="",
        ghi="; ".join(f"{k} {(dkkd.get(k) or {}).get('tt')}" for k, _, _ in hd02_hits[:8]),
        nguon="Danh ba NNT.xls",
    )
else:
    set_rule(
        "HD02",
        status="ĐÃ_CHẠY", ket="Không", nhan="Chắc", muc="", hs=0, diem=0,
        ca="NCC ĐKKD bất thường", va=0,
        thieu="",
        ghi=f"Đã tra Danh ba NNT.xls ({len(dkkd)} MST). {ncc_no_dkkd} NCC không có trong danh bạ - không gán bất thường.",
        nguon="Danh ba NNT.xls",
    )

# HD03
if hd03_pct > P13:
    set_rule(
        "HD03",
        status="ĐÃ_CHẠY", ket="Có phát hiện", nhan="Chắc", muc="Cao", hs=3, diem=8,
        ca="Tỷ trọng F0+F1", va=hd03_pct, nguong="P13=5%",
        thieu="",
        ghi=f"F0={len(f0_ncc)} GT {vn(f0_gt)}; F1={len(f1_ncc)} GT {vn(f1_gt)}; NCC {len(ncc_agg)}",
    )
else:
    set_rule(
        "HD03",
        status="ĐÃ_CHẠY", ket="Không", nhan="Chắc", muc="", hs=0, diem=0,
        ca="Tỷ trọng F0+F1", va=hd03_pct, nguong="P13=5%",
        thieu="",
        ghi=f"F0={len(f0_ncc)} GT {vn(f0_gt)}; F1={len(f1_ncc)} GT {vn(f1_gt)}; {hd03_pct:.2%} <= 5%",
    )

# HD05
if hd05_hits:
    set_rule(
        "HD05",
        status="ĐÃ_CHẠY", ket="Có phát hiện", nhan="Chắc", muc="Trung bình", hs=2, diem=5,
        ca="NCC thành lập <12 tháng trước HĐ đầu", va=len(hd05_hits),
        thieu="",
        ghi="; ".join(
            f"{k} {days} ngày / GT {vn(a['gt'])}" for k, a, rec, days in hd05_hits[:8]
        ),
        nguon="Danh ba NNT.xls col 32/64",
    )
else:
    set_rule(
        "HD05",
        status="ĐÃ_CHẠY", ket="Không", nhan="Chắc", muc="", hs=0, diem=0,
        ca="NCC thành lập <12 tháng trước HĐ đầu", va=0,
        thieu="",
        ghi="Đã tra ngày cấp GCN ĐKKD / ngày bắt đầu KD trên Danh ba NNT.xls. Không có NCC mới <12 tháng trước HĐ đầu.",
        nguon="Danh ba NNT.xls col 32/64",
    )

# HD08 = NNT trên F1 bên mua
if nnt_on_f1:
    set_rule(
        "HD08",
        status="ĐÃ_CHẠY", ket="Có phát hiện", nhan="Chắc", muc="Cao", hs=3, diem=8,
        ca="NNT trên F1 bên mua", va=len(f1_buyers_nnt),
        thieu="",
        ghi="Đối chiếu sheet F1 cột MST mua (index 8), min_row=8.",
    )
else:
    set_rule(
        "HD08",
        status="ĐÃ_CHẠY", ket="Không", nhan="Chắc", muc="", hs=0, diem=0,
        ca="NNT trên F1 bên mua", va=0,
        thieu="",
        ghi="NNT không có trên sheet F1 cột MST người mua.",
    )

# recount C6 J6
n_run = n_kdd = n_treo = n_kad = n_cao = n_tb = 0
c_status = hm.get("Trạng thái chạy")
c_ket = hm.get("Kết quả")
c_muc = hm.get("Mức rủi ro")
for rr in range(10, (ws.max_row or 10) + 1):
    st = str(ws.cell(rr, c_status).value or "").strip()
    ket = str(ws.cell(rr, c_ket).value or "").strip()
    muc = str(ws.cell(rr, c_muc).value or "").strip()
    if not st and not str(ws.cell(rr, col_ma).value or "").strip():
        continue
    if st == "ĐÃ_CHẠY":
        n_run += 1
    elif st == "KHÔNG_ĐỦ_DỮ_LIỆU":
        n_kdd += 1
    elif st == "TREO":
        n_treo += 1
    elif st == "KHÔNG_ÁP_DỤNG":
        n_kad += 1
    if ket == "Có phát hiện" and muc == "Cao":
        n_cao += 1
    if ket == "Có phát hiện" and muc == "Trung bình":
        n_tb += 1
ws["C6"] = f"ĐÃ_CHẠY {n_run} | KHÔNG_ĐỦ {n_kdd} | TREO {n_treo} | KHÔNG_ÁP_DỤNG {n_kad}"
ws["J6"] = f"Cao {n_cao} | Trung bình {n_tb}"
print("KQ3 recount run", n_run, "kdd", n_kdd, "treo", n_treo, "kad", n_kad, "cao", n_cao, "tb", n_tb)

# TC26
if "Danh_gia_28_tieu_chi_rui_ro" in wb.sheetnames:
    tw = wb["Danh_gia_28_tieu_chi_rui_ro"]
    for rr in range(3, 40):
        if str(tw.cell(rr, 1).value or "").strip() == "TC26":
            kq = (
                f"HD03 F0+F1={hd03_pct:.2%} trên {len(ncc_agg)} NCC. "
                f"Đã tra Danh ba NNT.xls: ĐKKD bất thường {len(hd02_hits)} MST; "
                f"NCC mới <12 tháng {len(hd05_hits)}. HD01={'Có' if hd01_hit else 'Không'} "
                f"(không kích vì NCC F1)."
            )
            muc_tc = "Cao" if hd03_pct > P13 else "Thấp"
            put(tw, rr, 5, kq, wrap=True)
            put(tw, rr, 6, muc_tc, bold=True)
            put(tw, rr, 7, "Chuỗi 5 mắt xích NCC F0; rà NCC F1/ĐKKD khi kiểm tra trụ sở.", wrap=True)
            if muc_tc == "Cao":
                tw.cell(rr, 6).fill = FILL_RED
                tw.cell(rr, 6).font = FONT_RED
            else:
                tw.cell(rr, 6).fill = FILL_GRN
                tw.cell(rr, 6).font = FONT
            print(" patched TC26 row", rr, muc_tc)
            break

wb.save(lp(files["kq3"][1]))
wb.close()
print(" KQ3 saved")

# ---------------------------------------------------------------------------
# PATCH dau hieu + giai trinh
# ---------------------------------------------------------------------------
print("PATCH dau hieu")
wb = load_workbook(lp(files["dau"][1]))
ws = wb.active
# find HD01/HD03 row
for rr in range(2, 20):
    ma = str(ws.cell(rr, 2).value or "")
    if "HD01" in ma:
        muc = "Cao" if (hd01_hit or hd03_pct > P13) else "Thấp"
        put(ws, rr, 3, f"NCC F0={len(f0_ncc)}; F1 tầng hai={len(f1_ncc)}; tỷ trọng {hd03_pct:.2%}; HD01={'Có' if hd01_hit else 'Không'}", wrap=True)
        put(ws, rr, 4, muc)
        put(ws, rr, 5, f"F0={len(f0_ncc)} GT {vn(f0_gt)}; F1={len(f1_ncc)} GT {vn(f1_gt)}; NCC {len(ncc_agg)}", wrap=True)
        if muc == "Cao":
            ws.cell(rr, 4).fill = FILL_RED
            ws.cell(rr, 4).font = FONT_RED
        else:
            ws.cell(rr, 4).fill = FILL_YEL
            ws.cell(rr, 4).font = FONT_ORG
        print(" patched dau", rr, muc)
        break
wb.save(lp(files["dau"][1]))
wb.close()

print("PATCH giai trinh")
wb = load_workbook(lp(files["giai"][1]))
if "00_Tong_hop_chi_tiet" in wb.sheetnames:
    ws = wb["00_Tong_hop_chi_tiet"]
    for rr in range(3, 20):
        ma = str(ws.cell(rr, 2).value or "")
        if "HD01" in ma:
            muc = "Cao" if (hd01_hit or hd03_pct > P13) else "Thấp"
            put(ws, rr, 3, f"NCC F0={len(f0_ncc)}; F1 tầng hai={len(f1_ncc)}; HD01={'Có' if hd01_hit else 'Không'}", wrap=True)
            put(ws, rr, 4, muc)
            put(ws, rr, 5, f"F0 GT {vn(f0_gt)}; F1 GT {vn(f1_gt)}", wrap=True)
            put(ws, rr, 9, "Đã chạy lại 11/11; không kích HD01 vì NCC F1", wrap=True)
            break
if "01_NCC_F0" in wb.sheetnames:
    ws = wb["01_NCC_F0"]
    if ws.max_row and ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)
    f0_set = {k for k, _ in f0_ncc}
    if not f0_set:
        put(ws, 2, 1, "Không có HĐ mua NCC F0/CA trên sổ chỉ định (nguyên tắc 11/11, F0 = bên bán).")
    else:
        rr = 2
        for inv in mua_inv:
            if inv["mst"] not in f0_set:
                continue
            nhom, nguon, cv, nd, rec = ncc_flag(inv["mst"])
            put(ws, rr, 1, inv["mst"])
            put(ws, rr, 2, inv["ten"], wrap=True)
            put(ws, rr, 3, inv["so"])
            put(ws, rr, 4, inv["ngay"], dt=True)
            put(ws, rr, 5, inv["gt"], num=True)
            put(ws, rr, 6, inv["thue"], num=True)
            put(ws, rr, 7, cv)
            put(ws, rr, 8, nguon)
            for c in range(1, 9):
                ws.cell(rr, c).fill = FILL_RED
            rr += 1
wb.save(lp(files["giai"][1]))
wb.close()
print(" giai saved")

# ---------------------------------------------------------------------------
# PATCH KQ1.3 IN07 status (không đổi kết luận GĐ1 NNT F0/F1)
# ---------------------------------------------------------------------------
print("PATCH KQ1.3 IN07 status")
wb = load_workbook(lp(KQ13))
patched = False
for sn in wb.sheetnames:
    ws = wb[sn]
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value == "Chưa lập bảng IN07":
                cell.value = "Đã lập IN07 v2 theo nguyên tắc 11/11 (Phần 1 NNT + ĐKKD danh bạ)"
                tnr(cell)
                patched = True
wb.save(lp(KQ13))
wb.close()
print(" KQ1.3 patched", patched)

# ---------------------------------------------------------------------------
# Copy to company folder
# ---------------------------------------------------------------------------
print("COPY company")
company_files = [
    files["kq21"][1], files["kq3"][1], files["dau"][1], files["giai"][1], KQ13,
]
for p in company_files:
    copy_file(p, OUT / p.name)

stats = {
    "today": TODAY.isoformat(),
    "ver": VER,
    "nnt_on_f0": nnt_on_f0,
    "nnt_on_f1": nnt_on_f1,
    "nnt_on_ca": nnt_on_ca,
    "nnt_f1_hits": len(f1_buyers_nnt),
    "f0_size": len(f0_map),
    "f1_buyers": len(f1_map),
    "ca_size": len(ca_map),
    "dkkd_size": len(dkkd),
    "ncc": len(ncc_agg),
    "inv": len(mua_inv),
    "src": dict(src_count),
    "f0_ncc": [(k, a["ten"], a["gt"]) for k, a in f0_ncc],
    "f1_ncc": [(k, a["ten"], a["gt"], ncc_flag(k)[3]) for k, a in f1_ncc],
    "f0_gt": f0_gt,
    "f1_gt": f1_gt,
    "tot_mua_gt": tot_mua_gt,
    "top5_pct": top5_pct,
    "top5": [(k, a["ten"], a["gt"]) for k, a in top5],
    "hd01": hd01_hit,
    "hd02": [(k, (dkkd.get(k) or {}).get("tt"), a["gt"]) for k, a, _ in hd02_hits],
    "hd03_pct": hd03_pct,
    "hd05": [(k, days, a["gt"]) for k, a, rec, days in hd05_hits],
    "ncc_no_dkkd": ncc_no_dkkd,
    "kq3": {"run": n_run, "kdd": n_kdd, "treo": n_treo, "kad": n_kad, "cao": n_cao, "tb": n_tb},
    "cvcbrr_mtime": str(cvcbrr_mtime),
    "nnt_tt": nnt_tt,
    "files": [str(p.name) for p in company_files],
}
stat_path = VAULT / "scratch" / "ktt_ncc_1111_stats.json"
stat_path.write_text(json.dumps(stats, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
print("STATS", json.dumps({k: stats[k] for k in ("hd01", "hd03_pct", "kq3", "ncc", "inv", "f0_ncc", "f1_ncc", "hd02", "hd05", "nnt_on_f0", "nnt_on_f1", "nnt_on_ca")}, ensure_ascii=False, default=str))
print("DONE")
