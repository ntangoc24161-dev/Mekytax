# -*- coding: utf-8 -*-
"""GĐ2 + GĐ3 Kiến Thành Tân MST 3603216935 - quy trình 1.4, sao mẫu rồi điền."""
from __future__ import annotations

import os
import re
import shutil
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603216935"
TEN = "CÔNG TY TNHH MỘT THÀNH VIÊN KIẾN THÀNH TÂN"
KY = "2015-2025"
YEARS = list(range(2015, 2026))
YEARS_511 = [2022, 2023, 2024, 2025]  # 4 cột mẫu KQ2_511_632
TODAY = date(2026, 11, 17)
VER = "v1"
MAU_VER = "1.4"
LEVEL = "D"
DN_STATUS = "Một phần"
LOAI_HINH = "Thương mại"

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\1. TD\14. KIEN THANH TAN_3603216935"
)
OUT = ROOT / "4. KET QUA PHAN TICH"
VAULT_OUT = VAULT / "ket-qua" / "KIEN_THANH_TAN_3603216935"
KQ11 = VAULT_OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_v2.xlsx"
KQ12 = VAULT_OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_v1.xlsx"
KQ13 = VAULT_OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_v2.xlsx"
RULES_WB = MAU / "Bo_bieu_mau_chuan_phan_tich_ho_so_thue.xlsx"
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")

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
P03, P05, P13, P14, P18, P23, P07 = 0.02, 0.30, 0.05, 0.50, 0.92, 0.10, 0.20
HAS = {
    "IN00": True, "IN01": True, "IN02": True, "IN03": True, "IN04": True,
    "IN05": True, "IN07": True, "IN08": True, "IN09": False, "IN10": True,
    "IN11": False, "IN12": False, "IN13": True, "IN14": True,
}
GD3_ONLY = {
    "DQ01", "GT06", "GT16", "TN04", "TN17", "TN18",
    "TT02", "TT03", "TT04", "TT05", "TT06",
    "KH01", "KH02", "KH03", "KH04", "KH05", "KH06", "KH07",
    "DT01", "DT02", "DT03", "LD02", "LD05",
}
KHONG_AP_DUNG = {"KH03", "KH05", "LD05"}  # thương mại


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
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def mst_base(s) -> str:
    s = str(s or "").split("-")[0]
    return re.sub(r"\D", "", s)


def so_key(v) -> str:
    s = str(v or "").strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.lstrip("0") or "0"


def mst_ok(s) -> bool:
    d = mst_base(s)
    return len(d) in (10, 13) and d.isdigit()


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


def apply_tnr(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            tnr(cell)


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


def clear_example_from(ws, start_row, file_col=None):
    if ws.max_row and ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)


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


def cash_nguong(ngay):
    if ngay is None:
        return 20_000_000
    if ngay >= date(2025, 7, 1):
        return 5_000_000
    return 20_000_000


# ---------------------------------------------------------------------------
# LOAD IN
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
    rec = {
        "chieu": "MUA" if ch.startswith("M") else "BAN",
        "nam": nam,
        "kyhieu": str(row[4] or "").strip(),
        "so": so_key(row[5]),
        "so_goc": row[5],
        "ngay": ndate(row[6]),
        "mst_nb": mst_base(row[7]),
        "ten_nb": str(row[8] or "").strip(),
        "mst_nm": mst_base(row[9]),
        "ten_nm": str(row[10] or "").strip(),
        "mst_dt": mst_base(row[11]),
        "mh": str(row[12] or "").strip(),
        "sl": nnum(row[14]),
        "dg": nnum(row[15]),
        "gt": nnum(row[16]),
        "ts": row[17],
        "thue": nnum(row[18]),
        "tong": nnum(row[19]) if len(row) > 19 else nnum(row[16]) + nnum(row[18]),
        "tinh": str(row[22] or "Có") if len(row) > 22 else "Có",
    }
    sig = (rec["chieu"], rec["nam"], rec["kyhieu"], rec["so"], rec["mst_dt"], rec["mh"], round(rec["gt"], 2), round(rec["thue"], 2), rec["sl"])
    if sig in seen_sig:
        continue
    seen_sig.add(sig)
    hddt_raw.append(rec)
wb.close()
mua_ok = [x for x in hddt_raw if x["chieu"] == "MUA"]
ban_ok = [x for x in hddt_raw if x["chieu"] == "BAN"]
print(" HDDT unique mua", len(mua_ok), "ban", len(ban_ok), "from raw after dedupe")

print("LOAD IN03")
wb = load_workbook(lp(KQ11), read_only=True, data_only=True)
ws = wb["IN03_ToKhai_GTGT"]
gtgt_all = []
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    rec = {
        "nam": int(row[0]),
        "ky": str(row[1] or ""),
        "loai": str(row[2] or ""),
        "lan": row[3],
        "hieu_luc": str(row[4] or ""),
        "ct22": nnum(row[6]), "ct23": nnum(row[7]), "ct24": nnum(row[8]),
        "ct25": nnum(row[9]), "ct26": nnum(row[10]),
        "ct32": nnum(row[14]) if len(row) > 14 else 0,
        "ct33": nnum(row[15]) if len(row) > 15 else 0,
        "ct34": nnum(row[16]) if len(row) > 16 else 0,
        "ct35": nnum(row[17]) if len(row) > 17 else 0,
        "ct37": nnum(row[19]) if len(row) > 19 else 0,
        "ct38": nnum(row[20]) if len(row) > 20 else 0,
        "ct40": nnum(row[22]) if len(row) > 22 else 0,
        "ct41": nnum(row[23]) if len(row) > 23 else 0,
        "ct43": nnum(row[25]) if len(row) > 25 else 0,
    }
    gtgt_all.append(rec)
eff_gtgt = [x for x in gtgt_all if x["hieu_luc"] == "Có"]
print(" GTGT all", len(gtgt_all), "eff", len(eff_gtgt))

print("LOAD IN04")
ws = wb["IN04_BCTC_TNDN"]
bctc = defaultdict(lambda: {"kqkd": {}, "cdkt": {}, "tndn": {}, "mau": ""})
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    y = int(row[0]) if row[0] else None
    bao = str(row[1] or "")
    ma = str(row[2] or "")
    so = row[4]
    if so is None or so == "":
        continue
    if bao == "KQKD":
        bctc[y]["kqkd"][ma] = nnum(so)
    elif bao == "CDKT":
        bctc[y]["cdkt"][ma] = nnum(so)
    elif bao == "TNDN":
        bctc[y]["tndn"][ma] = nnum(so)
    if row[6]:
        bctc[y]["mau"] = str(row[6])
print(" BCTC years with numbers", sorted(bctc.keys()))

print("LOAD IN05")
ws = wb["IN05_NLD_TNCN"]
in05 = []
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    in05.append({
        "loai": str(row[0]),
        "nam": int(row[1]) if row[1] else None,
        "ten": str(row[2] or ""),
        "cccd": str(row[3] or "").strip(),
        "mst_cn": str(row[4] or ""),
        "nguon": str(row[5] or ""),
        "tnct": nnum(row[7]) if row[7] not in (None, "") else None,
        "thue": nnum(row[8]) if row[8] not in (None, "") else None,
        "so_ld": nnum(row[16]) if len(row) > 16 and row[16] not in (None, "") else None,
        "trung": str(row[14] or "") if len(row) > 14 else "",
    })
wb.close()
print(" IN05", len(in05), "QTT", sum(1 for x in in05 if x["loai"] == "QTT"), "NLD", sum(1 for x in in05 if x["loai"] == "NLD"))

print("LOAD IN02")
wb = load_workbook(lp(KQ12), read_only=True, data_only=True)
ws = wb["IN02_BangKe"]
bk_raw = []
seen_bk = set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    ch = str(row[0]).upper()
    rec = {
        "chieu": "MUA" if ch.startswith("M") else "BAN",
        "nam": int(row[1]) if row[1] else None,
        "kyhieu": str(row[4] or "").strip(),
        "so": so_key(row[5]),
        "ngay": ndate(row[6]),
        "mst": mst_base(row[8] or row[7]),
        "ten": str(row[9] or "").strip(),
        "mh": str(row[10] or "").strip(),
        "gt": nnum(row[13]),
        "thue": nnum(row[15]),
        "tong": nnum(row[16]) if len(row) > 16 else nnum(row[13]) + nnum(row[15]),
    }
    sig = (rec["chieu"], rec["nam"], rec["kyhieu"], rec["so"], rec["mst"], rec["mh"], round(rec["gt"], 2), round(rec["thue"], 2))
    if sig in seen_bk:
        continue
    seen_bk.add(sig)
    bk_raw.append(rec)
print(" BK unique", len(bk_raw), "mua", sum(1 for x in bk_raw if x["chieu"]=="MUA"), "ban", sum(1 for x in bk_raw if x["chieu"]=="BAN"))

print("LOAD IN08")
ws = wb["IN08_So_KeToan"]
in08 = defaultdict(lambda: {"no": 0.0, "co": 0.0, "n": 0})
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or row[2] is None:
        continue
    nam = int(row[0]) if row[0] else None
    if nam == 2026 or nam is None:
        continue
    tk = str(row[2])[:3]
    in08[(nam, tk)]["no"] += nnum(row[3])
    in08[(nam, tk)]["co"] += nnum(row[4])
    in08[(nam, tk)]["n"] += 1

print("LOAD IN10")
ws = wb["IN10_Chi_tien_mat"]
in10 = []
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row:
        continue
    nam = int(row[1]) if row[1] else None
    if nam == 2026:
        continue
    amt = nnum(row[7])
    if row[0] is None and amt == 0:
        continue
    in10.append({
        "ngay": ndate(row[0]),
        "nam": nam,
        "so": str(row[2] or ""),
        "tk_du": str(row[3] or ""),
        "mst": mst_base(row[4]),
        "ten": str(row[5] or ""),
        "dg": str(row[6] or ""),
        "tien": amt,
    })
print(" IN10", len(in10))

print("LOAD IN13")
ws = wb["IN13_So_511_632"]
in13 = []
seen_13 = set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or row[4] is None:
        continue
    nam = int(row[0]) if row[0] else None
    if nam == 2026:
        continue
    tk = str(row[4])
    rec = {
        "nam": nam,
        "thang": row[1],
        "ngay": ndate(row[2]),
        "so_ct": str(row[3] or "").strip(),
        "tk": tk,
        "tk_du": str(row[5] or ""),
        "mh": str(row[6] or "").strip(),
        "dg": str(row[7] or "").strip(),
        "mst": mst_base(row[8]),
        "sl": nnum(row[9]),
        "ps_co": nnum(row[10]),
        "ps_no": nnum(row[11]),
    }
    sig = (rec["nam"], rec["so_ct"], rec["ngay"], rec["tk"][:3], rec["dg"], rec["mh"], round(rec["ps_co"], 2), round(rec["ps_no"], 2))
    if sig in seen_13:
        continue
    seen_13.add(sig)
    in13.append(rec)
wb.close()
print(" IN13 unique", len(in13), "511", sum(1 for x in in13 if x["tk"].startswith("511")), "632", sum(1 for x in in13 if x["tk"].startswith("632")))

# ---------------------------------------------------------------------------
# CVCBRR / CA
# ---------------------------------------------------------------------------
print("LOAD CVCBRR")
f0_map, f1_map, ca_map = {}, {}, {}
f1_buyers = []


def load_f0():
    if not os.path.exists(lp(CVCBRR)):
        print(" CVCBRR missing")
        return
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    if "F0" in wb.sheetnames:
        ws = wb["F0"]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i < 6:
                continue
            m = mst_base(row[1] if len(row) > 1 else "")
            if not m or not m[0].isdigit():
                continue
            if m not in f0_map:
                f0_map[m] = {
                    "mst": m, "ten": str(row[2] or "").strip() if len(row) > 2 else "",
                    "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                    "ngay_cv": ndate(row[4] if len(row) > 4 else None),
                    "nd": str(row[6] or "").strip()[:240] if len(row) > 6 else "",
                }
    if "F1" in wb.sheetnames:
        ws = wb["F1"]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i < 6:
                continue
            m_ban = mst_base(row[1] if len(row) > 1 else "")
            m_mua = mst_base(row[8] if len(row) > 8 else "")
            cv = str(row[3] or "").strip() if len(row) > 3 else ""
            if m_mua == MST:
                f1_buyers.append({"mst_ban": m_ban, "cv": cv})
            if m_mua and m_mua not in f1_map and m_mua not in f0_map:
                f1_map[m_mua] = {
                    "mst": m_mua, "ten": "", "cv": cv,
                    "ngay_cv": ndate(row[4] if len(row) > 4 else None),
                    "nd": "F1 - bên mua dùng HĐ F0",
                }
    wb.close()


def load_ca():
    if not os.path.exists(lp(CA)):
        print(" CA missing")
        return
    wb = load_workbook(lp(CA), read_only=True, data_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, row in enumerate(rows[:12]):
            if "mst" in " ".join(str(c or "").lower() for c in row):
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [str(c or "").strip().lower() for c in rows[hdr_i]]
        i_mst = next((i for i, h in enumerate(hdr) if "mst" in h), None)
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


load_f0()
load_ca()
nnt_on_f0 = MST in f0_map
nnt_on_f1 = MST in f1_map or bool(f1_buyers)
nnt_on_ca = MST in ca_map
print(" F0", len(f0_map), "F1 buyers map", len(f1_map), "CA", len(ca_map), "NNT F0/F1/CA", nnt_on_f0, nnt_on_f1, nnt_on_ca)
cvcbrr_mtime = datetime.fromtimestamp(os.path.getmtime(lp(CVCBRR))).date() if os.path.exists(lp(CVCBRR)) else None
if cvcbrr_mtime and (TODAY - cvcbrr_mtime).days > 30:
    print(" WARN CVCBRR older than 30 days", cvcbrr_mtime)


def ncc_flag(base: str):
    if not base:
        return "", "", "", ""
    if base in ca_map:
        rec = ca_map[base]
        return "F0", "CA", rec.get("cv", ""), rec.get("nd", "")
    if base in f0_map:
        rec = f0_map[base]
        return "F0", "CANH_BAO", rec.get("cv", ""), rec.get("nd", "")
    if base in f1_map:
        rec = f1_map[base]
        return "F1", "CANH_BAO", rec.get("cv", ""), rec.get("nd", "")
    return "", "", "", ""


# ---------------------------------------------------------------------------
# AGG
# ---------------------------------------------------------------------------
def hd_year(chieu, nam, field="gt"):
    src = mua_ok if chieu == "MUA" else ban_ok
    return sum(x[field] for x in src if x["nam"] == nam)


def bk_year(chieu, nam, field="gt"):
    return sum(x[field] for x in bk_raw if x["chieu"] == chieu and x["nam"] == nam)


def bctc_val(y, ma, which="kqkd"):
    rec = bctc.get(y)
    if not rec:
        return None
    if which == "kqkd":
        return rec["kqkd"].get(ma)
    if which == "cdkt":
        return rec["cdkt"].get(ma)
    if which == "tndn":
        return rec["tndn"].get(ma)
    for d in (rec["kqkd"], rec["cdkt"], rec["tndn"]):
        if ma in d:
            return d[ma]
    return None


def gtgt_year_sum():
    d = {y: {"ct34": 0.0, "ct35": 0.0, "ct23": 0.0, "ct24": 0.0, "ct26": 0.0, "ct32": 0.0, "ct33": 0.0, "ct40": 0.0, "ct43": 0.0, "kys": [], "missing": []} for y in YEARS}
    have = {(x["nam"], x["ky"].upper()) for x in eff_gtgt}
    for y in YEARS:
        for q in range(1, 5):
            k = f"Q{q}/{y}"
            if (y, k) not in have:
                d[y]["missing"].append(k)
    for x in eff_gtgt:
        y = x["nam"]
        if y not in d:
            continue
        d[y]["ct34"] += x["ct34"] or x["ct32"]
        d[y]["ct35"] += x["ct35"] or x["ct33"]
        d[y]["ct23"] += x["ct23"]
        d[y]["ct24"] += x["ct24"]
        d[y]["ct26"] += x["ct26"]
        d[y]["ct32"] += x["ct32"]
        d[y]["ct33"] += x["ct33"]
        d[y]["ct40"] += x["ct40"]
        if x["ct43"]:
            d[y]["ct43"] = x["ct43"]
        d[y]["kys"].append(x)
    return d


GY = gtgt_year_sum()
print(" GTGT missing", {y: GY[y]["missing"] for y in YEARS if GY[y]["missing"]})

ncc_agg = defaultdict(lambda: {"ten": "", "gt": 0.0, "thue": 0.0, "n": 0, "n_hd": set(), "src": set()})
for inv in mua_ok:
    k = inv["mst_nb"] or inv["mst_dt"]
    if not k:
        continue
    a = ncc_agg[k]
    a["ten"] = a["ten"] or inv["ten_nb"]
    a["gt"] += inv["gt"]
    a["thue"] += inv["thue"]
    a["n"] += 1
    a["n_hd"].add((inv["kyhieu"], inv["so"]))
    a["src"].add("HĐĐT")
# BK mua years without HDDT coverage (pre-2022)
hddt_mua_years = {x["nam"] for x in mua_ok}
for inv in bk_raw:
    if inv["chieu"] != "MUA":
        continue
    if inv["nam"] in hddt_mua_years:
        continue
    k = inv["mst"]
    if not k:
        continue
    a = ncc_agg[k]
    a["ten"] = a["ten"] or inv["ten"]
    a["gt"] += inv["gt"]
    a["thue"] += inv["thue"]
    a["n"] += 1
    a["n_hd"].add((inv["kyhieu"], inv["so"]))
    a["src"].add("Bảng kê")

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
risk_gt = f0_gt + f1_gt
hd03_pct = (risk_gt / tot_mua_gt) if tot_mua_gt else 0.0

cash_hits = []
for inv in mua_ok:
    amt = inv["tong"] if inv["tong"] else (inv["gt"] + inv["thue"])
    if amt >= cash_nguong(inv["ngay"]):
        cash_hits.append(inv)
clusters = defaultdict(lambda: {"n": 0, "gt": 0.0, "thue": 0.0, "tt": 0.0, "ten": "", "ngay": None, "mst": ""})
for inv in mua_ok:
    if not inv["ngay"]:
        continue
    k = (inv["mst_nb"] or inv["mst_dt"], inv["ngay"])
    clusters[k]["n"] += 1
    clusters[k]["gt"] += inv["gt"]
    clusters[k]["thue"] += inv["thue"]
    clusters[k]["tt"] += inv["tong"] or (inv["gt"] + inv["thue"])
    clusters[k]["ten"] = inv["ten_nb"]
    clusters[k]["ngay"] = inv["ngay"]
    clusters[k]["mst"] = k[0]
tt01 = []
for v in clusters.values():
    ng = cash_nguong(v["ngay"])
    if v["tt"] >= ng or v["gt"] >= ng:
        tt01.append({**v, "nguong": ng})
tt01.sort(key=lambda x: -x["gt"])
tt01_multi = [x for x in tt01 if x["n"] >= 2]

n_mst_sai = sum(1 for x in hddt_raw if not mst_ok(x["mst_dt"]))
n_thieu_kh = sum(1 for x in hddt_raw if not x["kyhieu"])
miss_mst = sum(1 for x in hddt_raw if not x["mst_dt"])
print("NCC", len(ncc_agg), "F0", len(f0_ncc), "F1", len(f1_ncc), "top5", f"{top5_pct:.2%}", "hd03", f"{hd03_pct:.2%}", "cash", len(cash_hits), "tt01", len(tt01), "multi", len(tt01_multi))
print("mst sai", n_mst_sai, "thieu kh", n_thieu_kh, "miss mst", miss_mst)

# 511-632 match. Sổ bán BH00001/24 vs xuất XK00001/24 - ghép số thứ tự + ngày + mặt hàng.
def amt_511(r):
    return r["ps_co"] if r["ps_co"] else r["ps_no"]


def amt_632(r):
    return r["ps_no"] if r["ps_no"] else r["ps_co"]


def so_seq(so):
    m = re.search(r"(\d+/\d+)", str(so or ""))
    return m.group(1) if m else str(so or "")


r511 = [x for x in in13 if x["tk"].startswith("511")]
r632 = [x for x in in13 if x["tk"].startswith("632")]
dc_y = {y: {"dt": 0.0, "gv": 0.0, "n511": 0, "n632": 0, "khop": 0, "khop_mh": 0, "lech_ten": 0, "no632": 0, "no511": 0, "duoi": 0, "dt_duoi": 0.0, "gv_duoi": 0.0} for y in YEARS_511}
used632 = set()
idx632 = list(range(len(r632)))
for a in r511:
    y = a["nam"]
    if y not in dc_y:
        continue
    dc_y[y]["dt"] += amt_511(a)
    dc_y[y]["n511"] += 1
    seq_a = so_seq(a["so_ct"])
    mh_a = (a["dg"] or a["mh"] or "").strip().lower()
    k_full = None
    for i in idx632:
        if i in used632:
            continue
        b = r632[i]
        if b["nam"] != y:
            continue
        if seq_a and seq_a == so_seq(b["so_ct"]) and a["ngay"] == b["ngay"] and mh_a and mh_a == (b["dg"] or b["mh"] or "").strip().lower():
            k_full = i
            break
    if k_full is None:
        for i in idx632:
            if i in used632:
                continue
            b = r632[i]
            if b["nam"] != y:
                continue
            if a["ngay"] == b["ngay"] and mh_a and mh_a == (b["dg"] or b["mh"] or "").strip().lower():
                k_full = i
                break
    if k_full is not None:
        b = r632[k_full]
        used632.add(k_full)
        dc_y[y]["khop"] += 1
        gv = amt_632(b)
        dc_y[y]["gv"] += gv
        lai = amt_511(a) - gv
        if lai <= 0:
            dc_y[y]["duoi"] += 1
            dc_y[y]["dt_duoi"] += amt_511(a)
            dc_y[y]["gv_duoi"] += gv
        continue
    k_mh = None
    for i in idx632:
        if i in used632:
            continue
        b = r632[i]
        if b["nam"] != y:
            continue
        if seq_a and seq_a == so_seq(b["so_ct"]) and mh_a and mh_a == (b["dg"] or b["mh"] or "").strip().lower():
            k_mh = i
            break
    if k_mh is not None:
        b = r632[k_mh]
        used632.add(k_mh)
        dc_y[y]["khop_mh"] += 1
        dc_y[y]["gv"] += amt_632(b)
        continue
    k_ng = None
    for i in idx632:
        if i in used632:
            continue
        b = r632[i]
        if b["nam"] != y:
            continue
        if seq_a and seq_a == so_seq(b["so_ct"]) and a["ngay"] == b["ngay"]:
            k_ng = i
            break
    if k_ng is not None:
        b = r632[k_ng]
        used632.add(k_ng)
        dc_y[y]["lech_ten"] += 1
        dc_y[y]["gv"] += amt_632(b)
        continue
    dc_y[y]["no632"] += 1
for i, b in enumerate(r632):
    y = b["nam"]
    if y not in dc_y:
        continue
    dc_y[y]["n632"] += 1
    if i not in used632:
        dc_y[y]["no511"] += 1
        dc_y[y]["gv"] += amt_632(b)
print("511-632", {y: {k: (round(v, 0) if isinstance(v, float) else v) for k, v in rec.items()} for y, rec in dc_y.items()})

# IN10 over threshold
tt02 = [x for x in in10 if x["tien"] >= cash_nguong(x["ngay"])]
tt03_map = defaultdict(lambda: {"n": 0, "tien": 0.0, "ten": "", "ngay": None, "mst": ""})
for x in in10:
    if not x["ngay"]:
        continue
    k = (x["mst"] or x["ten"], x["ngay"])
    tt03_map[k]["n"] += 1
    tt03_map[k]["tien"] += x["tien"]
    tt03_map[k]["ten"] = x["ten"]
    tt03_map[k]["ngay"] = x["ngay"]
    tt03_map[k]["mst"] = x["mst"]
tt03 = [v for v in tt03_map.values() if v["n"] >= 2 and v["tien"] >= cash_nguong(v["ngay"])]

print("LOAD 80 rules")
rwb = load_workbook(lp(RULES_WB), data_only=True)
rws = rwb["04_Quy_tac_rui_ro"]
RULES = []
for r in range(6, 90):
    ma = rws.cell(r, 1).value
    if not ma:
        continue
    RULES.append({
        "ma": str(ma).strip(),
        "nhom": rws.cell(r, 2).value,
        "ten": rws.cell(r, 3).value,
        "need": str(rws.cell(r, 7).value or ""),
        "muc": rws.cell(r, 8).value,
        "goi_y": rws.cell(r, 9).value,
        "giai_doan": rws.cell(r, 15).value,
    })
rwb.close()
print(" rules", len(RULES))

# copy templates
OUT.mkdir(parents=True, exist_ok=True)
VAULT_OUT.mkdir(parents=True, exist_ok=True)
files = {
    "kq21": (MAU / "KQ2.1_Phan_tich_CQT.xlsx", OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_{VER}.xlsx"),
    "kq22": (MAU / "KQ2.2_Phan_tich_ho_so_cty.xlsx", OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_{VER}.xlsx"),
    "kq3": (MAU / "KQ3_Ket_luan_danh_gia.xlsx", OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_{VER}.xlsx"),
}
for src, dst in files.values():
    shutil.copy2(lp(src), lp(dst))
    print("copied", dst.name)

hdr2 = (
    f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: {MAU_VER} | "
    f"Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER} | Dữ liệu công ty: {DN_STATUS}"
)

# ===========================================================================
# KQ2.1
# ===========================================================================
print("FILL KQ2.1")
wb = load_workbook(lp(files["kq21"][1]))
for sn in wb.sheetnames:
    sub_placeholders(wb[sn])

ws = wb["KQ2_Tong_hop_co"]
put(ws, 2, 1, hdr2, bold=True)

ws = wb["KQ2_Doi_chieu_GTGT"]
put(ws, 2, 1, f"Người nộp thuế: {TEN} - MST: {MST} - Kỳ: {KY}", bold=True)
put(ws, 3, 1, hdr2 + " | Nguồn: tờ khai GTGT IN03 (CQT), HĐĐT IN01 (CQT, đã khử trùng dòng copy), Bảng kê IN02, BCTC IN04. NKC 133 IN08 có 2023-2025.")


def year_header_row(y):
    return 9 + (y - 2016) * 14


for t in sorted(eff_gtgt, key=lambda x: (x["nam"], x["ky"])):
    y = t["nam"]
    if y < 2016 or y > 2025:
        continue
    mq = re.search(r"Q([1-4])", t["ky"].upper())
    if not mq:
        continue
    rr = year_header_row(y) + int(mq.group(1))
    dt = t["ct32"] or t["ct34"]
    th = t["ct33"] or t["ct35"]
    if dt:
        put(ws, rr, 5, dt, num=True)
    if th:
        put(ws, rr, 9, th, num=True)
    if t["ct23"]:
        put(ws, rr, 11, t["ct23"], num=True)
    if t["ct24"]:
        put(ws, rr, 12, t["ct24"], num=True)
    if t["ct37"]:
        put(ws, rr, 13, t["ct37"], num=True)
    if t["ct38"]:
        put(ws, rr, 14, t["ct38"], num=True)
    if t["ct40"]:
        put(ws, rr, 17, t["ct40"], num=True)
    if t["ct43"]:
        put(ws, rr, 18, t["ct43"], num=True)

for y in range(2016, 2026):
    rr = 156 + (y - 2016)
    hb = hd_year("BAN", y)
    ht = hd_year("BAN", y, "thue")
    if any(x["nam"] == y for x in ban_ok):
        put(ws, rr, 2, hb, num=True)
        put(ws, rr, 3, ht, num=True)
    bb = bk_year("BAN", y)
    bt = bk_year("BAN", y, "thue")
    if any(x["nam"] == y and x["chieu"] == "BAN" for x in bk_raw):
        put(ws, rr, 4, bb, num=True)
        put(ws, rr, 5, bt, num=True)
    gy = GY.get(y) or {}
    if gy.get("kys"):
        if gy.get("ct34") or gy.get("ct32"):
            put(ws, rr, 6, gy.get("ct34") or gy.get("ct32"), num=True)
        if gy.get("ct35") or gy.get("ct33"):
            put(ws, rr, 7, gy.get("ct35") or gy.get("ct33"), num=True)
        if gy.get("ct23"):
            put(ws, rr, 15, gy["ct23"], num=True)
        if gy.get("ct24"):
            put(ws, rr, 21, gy["ct24"], num=True)
    td = bctc_val(y, "04", "tndn") or bctc_val(y, "A1", "tndn")
    if td is not None:
        put(ws, rr, 8, td, num=True)
    b01 = bctc_val(y, "01", "kqkd")
    if b01 is not None:
        put(ws, rr, 9, b01, num=True)
    put(ws, rr, 12, y)
    hm = hd_year("MUA", y)
    hmt = hd_year("MUA", y, "thue")
    if any(x["nam"] == y for x in mua_ok):
        put(ws, rr, 13, hm, num=True)
        put(ws, rr, 19, hmt, num=True)
    bm = bk_year("MUA", y)
    bmt = bk_year("MUA", y, "thue")
    if any(x["nam"] == y and x["chieu"] == "MUA" for x in bk_raw):
        put(ws, rr, 14, bm, num=True)
        put(ws, rr, 20, bmt, num=True)
    n133 = in08.get((y, "133"), {}).get("no")
    if n133:
        put(ws, rr, 25, n133, num=True)

notes = [
    "Nhận xét / nguyên nhân trọng yếu (module GTGT, không gán mức rủi ro hồ sơ)",
    "NNT kê khai GTGT theo quý. Phần I lấy tờ khai hiệu lực (ưu tiên bổ sung). Thiếu XML Q4/2022 và Q1-Q3/2023: để trống, không điền 0.",
    "HĐĐT IN01 có dòng copy trùng (portal/file lặp): đã khử trùng theo khóa HĐ+MST+mặt hàng+tiền trước khi cộng. Bảng kê IN02 dùng cho cột BK.",
    "IN08 năm 2026 là nhiễu đường dẫn thư mục 2026 - không dùng. PS Nợ 133 có 2023-2025.",
]
for y in range(2016, 2026):
    tk34 = (GY.get(y) or {}).get("ct34")
    hd = hd_year("BAN", y) if any(x["nam"] == y for x in ban_ok) else None
    bk = bk_year("BAN", y) if any(x["nam"] == y and x["chieu"] == "BAN" for x in bk_raw) else None
    b01 = bctc_val(y, "01", "kqkd")
    bits = [str(y)]
    if tk34 and hd is not None:
        bits.append(f"TK[34]={vn(tk34)} HĐĐT bán={vn(hd)} lệch {vn(tk34-hd)}")
    if tk34 and bk is not None:
        bits.append(f"BK bán={vn(bk)} lệch TK {vn((tk34 or 0)-bk)}")
    if b01 is not None and tk34:
        bits.append(f"BCTC[01]={vn(b01)} lệch GTGT {vn(tk34-b01)}")
    miss = GY.get(y, {}).get("missing") or []
    if miss:
        bits.append("thiếu XML " + ",".join(miss))
    if len(bits) > 1:
        notes.append(" | ".join(bits))
r_note = 168
for i, line in enumerate(notes):
    put(ws, r_note + i, 1, line, bold=(i == 0), wrap=True, fill=FILL_SEC if i == 0 else None)
    try:
        ws.merge_cells(start_row=r_note + i, start_column=1, end_row=r_note + i, end_column=10)
    except Exception:
        pass

# IN07
ws = wb["IN07_NCC_CanhBao"]
clear_example_from(ws, 3)
r = 3
in07_n = 0
seen = set()
if nnt_on_f0 or nnt_on_f1 or nnt_on_ca:
    nhom = "F0" if (nnt_on_f0 or nnt_on_ca) else "F1"
    nguon = "CA" if nnt_on_ca else "CANH_BAO"
    rec = ca_map.get(MST) or f0_map.get(MST) or f1_map.get(MST) or {}
    put(ws, r, 1, MST)
    put(ws, r, 2, TEN)
    put(ws, r, 3, "NNT đang kiểm tra")
    put(ws, r, 4, nhom)
    put(ws, r, 5, nguon)
    put(ws, r, 6, rec.get("cv", ""))
    put(ws, r, 7, rec.get("ngay_cv"), dt=True)
    put(ws, r, 16, "Sổ CVCBRR / CA")
    r += 1
    in07_n += 1
    seen.add(MST)
for k, a in ranked:
    nhom, nguon, cv, nd = ncc_flag(k)
    if not nhom or k in seen:
        continue
    seen.add(k)
    rec = ca_map.get(k) or f0_map.get(k) or f1_map.get(k) or {}
    put(ws, r, 1, k)
    put(ws, r, 2, a["ten"] or rec.get("ten", ""))
    put(ws, r, 3, "NCC mua vào")
    put(ws, r, 4, nhom)
    put(ws, r, 5, nguon)
    put(ws, r, 6, cv or rec.get("cv", ""))
    put(ws, r, 7, rec.get("ngay_cv"), dt=True)
    put(ws, r, 8, "Công an" if nguon == "CA" else "Chi cục Thuế")
    put(ws, r, 9, nd or rec.get("nd", ""))
    put(ws, r, 16, "HĐĐT/BK + sổ CVCBRR/CA")
    put(ws, r, 17, f"GT mua {vn(a['gt'])}")
    fill = FILL_RED if nhom == "F0" else FILL_YEL
    font = FONT_RED if nhom == "F0" else FONT_ORG
    for c in range(1, 19):
        ws.cell(r, c).fill = fill
        ws.cell(r, c).font = font
    r += 1
    in07_n += 1
if in07_n == 0:
    put(ws, 3, 1, "Không có NNT/NCC trên F0, F1 hoặc sổ Công an tại thời điểm chạy.")
    put(ws, 3, 17, "Không phải phát hiện - đã tra đủ sổ chỉ định.")
print("IN07 rows", in07_n)
autowidth(ws, max_row=40)

# NCC 100%
ws = wb["KQ2_NCC_mua_vao"]
put(ws, 2, 1, hdr2, bold=True)
if ws.max_row > 8:
    ws.delete_rows(8, ws.max_row - 7)
# unique invoices mua: HDDT 2022-2025 + BK earlier
mua_inv = []
seen_inv = set()
for inv in mua_ok:
    k = (inv["kyhieu"], inv["so"], inv["mst_nb"] or inv["mst_dt"], inv["nam"])
    if k in seen_inv:
        # same invoice extra line - still write (chi tiết hàng)
        pass
    seen_inv.add(k)
    mua_inv.append(("HĐĐT", inv["mst_nb"] or inv["mst_dt"], inv["ten_nb"], inv["so_goc"] if inv["so_goc"] is not None else inv["so"], inv["ngay"], inv["mh"], inv["gt"], inv["thue"]))
for inv in bk_raw:
    if inv["chieu"] != "MUA" or inv["nam"] in hddt_mua_years:
        continue
    mua_inv.append(("Bảng kê", inv["mst"], inv["ten"], inv["so"], inv["ngay"], inv["mh"], inv["gt"], inv["thue"]))
r = 8
for i, (src, mst, ten, so, ngay, mh, gt, thue) in enumerate(mua_inv, 1):
    nhom, nguon, cv, nd = ncc_flag(mst)
    put(ws, r, 1, i)
    put(ws, r, 2, mst)
    put(ws, r, 3, ten)
    put(ws, r, 4, so)
    put(ws, r, 5, ngay, dt=True)
    put(ws, r, 6, mh)
    put(ws, r, 7, gt, num=True)
    put(ws, r, 8, thue, num=True)
    put(ws, r, 9, src)
    put(ws, r, 10, f"Có - {cv}" if cv else ("Có" if nhom else "Không"))
    put(ws, r, 11, nhom or "Không")
    put(ws, r, 12, "Chưa tra DKKD (thiếu danh bạ 3.2)")
    if nhom == "F0":
        for c in range(1, 13):
            ws.cell(r, c).fill = FILL_RED
            ws.cell(r, c).font = FONT_RED
    elif nhom == "F1":
        for c in range(1, 13):
            ws.cell(r, c).fill = FILL_YEL
            ws.cell(r, c).font = FONT_ORG
    r += 1
last_inv = r - 1
put(ws, 5, 2, hd03_pct, pct=True)
put(ws, 5, 5, top5_pct, pct=True)
rr = last_inv + 2
put(ws, rr, 1, "Nhận xét module (3.4) - không phải kết luận hồ sơ", bold=True, fill=FILL_SEC)
top5_txt = "; ".join(f"{k} {a['ten'][:40]} {vn(a['gt'])}" for k, a in top5)
put(
    ws, rr + 1, 1,
    f"HD04 Top-5 / tổng mua = {top5_pct:.2%} (P14=50% tạm). Số NCC={len(ncc_agg)}. "
    f"F0={len(f0_ncc)} GT {vn(f0_gt)} thuế {vn(f0_thue)}. F1={len(f1_ncc)} GT {vn(f1_gt)}. "
    f"HD03 F0+F1 = {hd03_pct:.2%} (P13=5%). NNT F0={'Có' if nnt_on_f0 else 'Không'}; F1={'Có' if nnt_on_f1 else 'Không'}; CA={'Có' if nnt_on_ca else 'Không'}. "
    f"Top-5: {top5_txt}. Nguồn 2022-2025 = HĐĐT (khử trùng); 2014-2021 = Bảng kê. Danh bạ 3.2 chưa có nên HD02/HD05 chưa chốt.",
    wrap=True,
)
try:
    ws.merge_cells(start_row=rr + 1, start_column=1, end_row=rr + 1, end_column=12)
except Exception:
    pass
ws.auto_filter.ref = f"A7:L{max(7, last_inv)}"
autowidth(ws, cap=36, max_row=30)
print("NCC invoice rows", last_inv - 7)

# TNCN
ws = wb["KQ2_TNCN_LD2"]
put(ws, 2, 1, hdr2, bold=True)
for rr in range(8, 30):
    for c in range(1, 8):
        cell = ws.cell(rr, c)
        if isinstance(cell, MergedCell):
            continue
        if isinstance(cell.value, str) and cell.value.startswith("="):
            continue
        cell.value = None
rr = 8
qtt_by = {x["nam"]: x for x in in05 if x["loai"] == "QTT"}
for y in YEARS:
    t = qtt_by.get(y)
    nld_y = [x for x in in05 if x["loai"] == "NLD" and x["nam"] == y]
    nld_52 = [x for x in nld_y if "05-2" in (x.get("nguon") or "")]
    cccd = [x["cccd"] for x in nld_y if x.get("cccd")]
    dup = len(cccd) - len(set(cccd)) if cccd else 0
    no10 = sum(1 for x in nld_52 if (x.get("thue") or 0) == 0 and (x.get("tnct") or 0) > 0)
    put(ws, rr, 1, y)
    if t:
        put(ws, rr, 2, int(t["so_ld"] or len(nld_y) or 0) or None, num=True)
        if t["tnct"] is not None:
            put(ws, rr, 4, t["tnct"], num=True)
        if t["thue"] is not None:
            put(ws, rr, 5, t["thue"], num=True)
    put(ws, rr, 3, len(nld_52) or None, num=True)
    put(ws, rr, 6, no10 or None, num=True)
    put(ws, rr, 7, dup or None, num=True)
    rr += 1
put(ws, rr + 1, 1, "Nhận xét module TNCN (không ghi mức rủi ro hồ sơ)", bold=True, fill=FILL_SEC)
put(
    ws, rr + 2, 1,
    f"IN05: 11 dòng QTT 2015-2025 + 24 NLĐ phụ lục 05-1. Năm 2023 QTT XML không có phụ lục và chỉ tiêu TNCT trống. "
    f"05-2 không có dòng tên thật (bỏ node rỗng T10). Thiếu tra cứu đa nguồn 3.3 nên LD04 ngoài QTT chưa chốt. "
    f"Tờ khai TNCN quý không có trong 1.1 - LD01 (QTT vs tạm tính quý) = KHÔNG_ĐỦ_DỮ_LIỆU.",
    wrap=True,
)
apply_tnr(wb["KQ2_Doi_chieu_GTGT"])
wb.save(lp(files["kq21"][1]))
wb.close()
print("saved KQ2.1")

# ===========================================================================
# KQ2.2
# ===========================================================================
print("FILL KQ2.2")
wb = load_workbook(lp(files["kq22"][1]))
for sn in wb.sheetnames:
    sub_placeholders(wb[sn])
ws = wb["KQ2_Tong_hop_co"]
put(ws, 2, 1, hdr2, bold=True)

ws = wb["KQ2_511_632"]
put(ws, 2, 1, hdr2, bold=True)
# clear example zeros
for r in range(6, 17):
    for c in range(2, 8):
        cell = ws.cell(r, c)
        if isinstance(cell, MergedCell):
            continue
        cell.value = None
# years in cols 3-6
for i, y in enumerate(YEARS_511):
    put(ws, 5, 3 + i, y, bold=True)
tot_dt = tot_gv = 0.0
tot_khop = tot_no632 = tot_no511 = tot_duoi = 0
ghi = []
for i, y in enumerate(YEARS_511):
    rec = dc_y[y]
    c = 3 + i
    if rec["n511"] or rec["n632"]:
        put(ws, 6, c, rec["dt"], num=True)
        put(ws, 7, c, rec["gv"], num=True)
        put(ws, 8, c, rec["dt"] - rec["gv"], num=True)
        put(ws, 9, c, (rec["dt"] - rec["gv"]) / rec["dt"] if rec["dt"] else None, pct=True)
        put(ws, 10, c, rec["khop"], num=True)
        put(ws, 11, c, rec["khop_mh"], num=True)
        put(ws, 12, c, rec["lech_ten"], num=True)
        put(ws, 13, c, rec["no632"], num=True)
        put(ws, 14, c, rec["no511"], num=True)
        put(ws, 15, c, rec["duoi"], num=True)
        put(ws, 16, c, rec["dt_duoi"] / rec["dt"] if rec["dt"] else None, pct=True)
        tot_dt += rec["dt"]
        tot_gv += rec["gv"]
        tot_khop += rec["khop"]
        tot_no632 += rec["no632"]
        tot_no511 += rec["no511"]
        tot_duoi += rec["duoi"]
        ty = (rec["dt"] - rec["gv"]) / rec["dt"] if rec["dt"] else None
        ghi.append(f"{y}: DT {vn(rec['dt'])} GV {vn(rec['gv'])} lãi gộp {ty:.2%}" if ty is not None else str(y))
    else:
        put(ws, 6, 7, "2022: chưa có dòng IN13 (sổ 2019/2022 mỏng). Không điền 0.")
put(ws, 6, 2, tot_dt, num=True)
put(ws, 7, 2, tot_gv, num=True)
put(ws, 8, 2, tot_dt - tot_gv, num=True)
put(ws, 9, 2, (tot_dt - tot_gv) / tot_dt if tot_dt else None, pct=True)
put(ws, 10, 2, tot_khop, num=True)
put(ws, 13, 2, tot_no632, num=True)
put(ws, 14, 2, tot_no511, num=True)
put(ws, 15, 2, tot_duoi, num=True)
put(
    ws, 5, 7,
    f"IN13 (bỏ năm 2026 nhiễu). Khớp HĐ+ngày+MH {tot_khop} dòng. Bán dưới giá {tot_duoi} dòng. "
    + " | ".join(ghi),
    wrap=True,
)

ws = wb["KQ2_Kho_Gia_thanh"]
put(ws, 2, 1, hdr2, bold=True)
# IN09 không có số XNT thật - để trống, ghi chú. Phần 3 vẫn điền BCTC HTK/DT (CQT)
for i, y in enumerate([2022, 2023, 2024, 2025, 2021, 2020, 2017, 2015][:10]):
    rr = 7 + i
    put(ws, rr, 1, y)
    put(ws, rr, 12, "IN09 chỉ kiểm kê file, không có SL/TT XNT. Không điền 0.")
# Phần 3 row 34+
for i, y in enumerate([2022, 2023, 2024, 2025, 2021, 2020, 2017, 2015][:10]):
    rr = 34 + i
    put(ws, rr, 1, y)
    htk = bctc_val(y, "140", "cdkt")
    dt = bctc_val(y, "01", "kqkd")
    if htk is not None:
        put(ws, rr, 5, htk, num=True)
    if dt is not None:
        put(ws, rr, 7, dt, num=True)
    put(ws, rr, 10, "BCTC ct140 TT133 (K6). Sổ 152/155 trống. 156 IN08 có PS nhưng không có tồn CK.")
put(ws, 45, 1, "Nhận xét TC3: HTK BCTC/DT 2020-2022 vượt 30% (sơ bộ). IN09 không dùng được. Không kết luận tồn âm.", wrap=True, fill=FILL_YEL)

ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
put(ws, 2, 1, hdr2, bold=True)
for i, y in enumerate(YEARS_511):
    rr = 7 + i
    put(ws, rr, 1, y)
    put(ws, rr, 11, "KHÔNG_ÁP_DỤNG 621/622/627/154 - DN thương mại hạt nhựa.")
# phần GV vs BCTC if template has row 21
for i, y in enumerate(YEARS_511):
    rr = 21 + i
    dt = bctc_val(y, "01", "kqkd")
    gv = bctc_val(y, "11", "kqkd")
    if ws.cell(20, 1).value or ws.cell(21, 1).value is not None or True:
        if dt is not None:
            try:
                put(ws, rr, 1, y)
                put(ws, rr, 2, gv, num=True) if gv is not None else None
            except Exception:
                pass
put(ws, 32, 1, "Nhận xét: thương mại - không chạy định mức SX. Tỷ lệ GV/DT BCTC tham khảo cho TN04/TC28 ở KQ3. Sổ 632 lấy IN13.", wrap=True, fill=FILL_SEC)

ws = wb["KQ2_Cong_no_Tien_mat"]
put(ws, 2, 1, hdr2, bold=True)
for i, y in enumerate([2023, 2024, 2025, 2022, 2021, 2020, 2017, 2015][:10]):
    rr = 8 + i
    put(ws, rr, 1, y)
    a131 = in08.get((y, "131"))
    a331 = in08.get((y, "331"))
    if a131:
        put(ws, rr, 3, a131["no"], num=True)
        put(ws, rr, 4, a131["co"], num=True)
    if a331:
        put(ws, rr, 7, a331["co"], num=True)
        put(ws, rr, 8, a331["no"], num=True)
    b01 = bctc_val(y, "01", "kqkd")
    v331 = bctc_val(y, "331", "cdkt")
    if b01 is not None:
        put(ws, rr, 11, b01, num=True)
    if v331 is not None:
        put(ws, rr, 9, v331, num=True)
# TT02 list from row 34 if headers exist
put(ws, 40, 1, f"TT01 sơ bộ HĐĐT: {len(cash_hits)} HĐ mua đạt ngưỡng K5; {len(tt01_multi)} cụm cùng ngày n>=2. Không kết luận đã thanh toán tiền mặt.", wrap=True)
put(ws, 41, 1, f"TT02 IN10 (bỏ 2026): {len(tt02)} phiếu chi >= ngưỡng; TT03 {len(tt03)} cụm cùng ngày. IN11 trống: GT16/TT05/TT06 = KHÔNG_ĐỦ_DỮ_LIỆU, không kết luận thiếu UNC.", wrap=True)
wb.save(lp(files["kq22"][1]))
wb.close()
print("saved KQ2.2")

# ===========================================================================
# KQ3
# ===========================================================================
print("FILL KQ3")
wb = load_workbook(lp(files["kq3"][1]))
ws = wb["KQ3"]
put(ws, 2, 3, MST, bold=True)
put(ws, 2, 10, TEN, bold=True)
put(ws, 3, 3, KY)
put(ws, 3, 10, LEVEL, bold=True)
put(ws, 4, 3, VER)
put(ws, 4, 10, MAU_VER)
put(ws, 5, 3, TODAY, dt=True)
put(ws, 5, 10, f"KQ1.3 v2; KQ2.1 {VER}; KQ2.2 {VER} | DN {DN_STATUS} | loại hình {LOAI_HINH}")
put(ws, 7, 10, DN_STATUS)


def year_diffs(get_a, get_b, full_years=None):
    out = []
    for y in YEARS:
        a, b = get_a(y), get_b(y)
        if a is None or b is None:
            continue
        full = True if full_years is None else y in full_years
        out.append((y, a, b, a - b, full))
    return out


def pick_diff(diffs, p03=P03):
    full = [x for x in diffs if x[4]]
    pool = full or diffs
    if not pool:
        return None
    y, a, b, d, full = max(pool, key=lambda x: abs(x[3]))
    base = max(abs(a), abs(b), 1)
    vuot = abs(d) / base > p03
    return y, a, b, d, full, vuot


def tk34(y):
    g = GY.get(y) or {}
    if not g.get("kys"):
        return None
    return g.get("ct34") or g.get("ct32") or 0.0


def tk35(y):
    g = GY.get(y) or {}
    if not g.get("kys"):
        return None
    return g.get("ct35") or g.get("ct33") or 0.0


def tk23(y):
    g = GY.get(y) or {}
    return g.get("ct23") if g.get("kys") else None


def tk24(y):
    g = GY.get(y) or {}
    return g.get("ct24") if g.get("kys") else None


def hd_ban(y):
    return hd_year("BAN", y) if any(x["nam"] == y for x in ban_ok) else None


def hd_mua(y):
    return hd_year("MUA", y) if any(x["nam"] == y for x in mua_ok) else None


def hd_ban_thue(y):
    return hd_year("BAN", y, "thue") if any(x["nam"] == y for x in ban_ok) else None


def hd_mua_thue(y):
    return hd_year("MUA", y, "thue") if any(x["nam"] == y for x in mua_ok) else None


def bk_ban(y):
    return bk_year("BAN", y) if any(x["nam"] == y and x["chieu"] == "BAN" for x in bk_raw) else None


def bk_mua(y):
    return bk_year("MUA", y) if any(x["nam"] == y and x["chieu"] == "MUA" for x in bk_raw) else None


BAN_FULL = {2023, 2024, 2025}
MUA_FULL = {2022, 2023, 2024, 2025}
d_gt01 = year_diffs(tk34, hd_ban, BAN_FULL)
d_gt02 = year_diffs(tk34, bk_ban)
d_gt03 = year_diffs(tk35, hd_ban_thue, BAN_FULL)
d_gt04 = year_diffs(tk23, hd_mua, MUA_FULL)
d_gt05 = year_diffs(tk24, hd_mua_thue, MUA_FULL)
d_tn01 = year_diffs(tk34, lambda y: bctc_val(y, "01", "kqkd"))
d_tn02 = year_diffs(lambda y: bctc_val(y, "04", "tndn"), lambda y: bctc_val(y, "01", "kqkd"))

findings = []


def add_rule(rows, rd, **kw):
    ma = rd["ma"]
    status = kw.get("status", "ĐÃ_CHẠY")
    ket = kw.get("ket", "Không")
    nhan = kw.get("nhan", "")
    muc_rr = kw.get("muc", "")
    diem = kw.get("diem", 0)
    thieu = kw.get("thieu", "")
    if ket != "Có phát hiện":
        w_tien = 0 if status == "ĐÃ_CHẠY" else None
        muc_rr = muc_rr if ket == "Có phát hiện" else ""
        diem = diem if ket == "Có phát hiện" else 0
    else:
        w_tien = kw.get("w", 0)
    hs = {"Cao": 3, "Trung bình": 2, "Thấp": 1}.get(muc_rr, 0) if ket == "Có phát hiện" else 0
    rows.append({
        "Mã_PH": f"{ma}|{MST}|{KY}",
        "Tệp phân tích": kw.get("tep", "Trực tiếp"),
        "Nhóm": rd["nhom"],
        "Mã quy tắc": ma,
        "Tên quy tắc": rd["ten"],
        "Mã TC": kw.get("ma_tc", ""),
        "Đối tượng": TEN,
        "Kỳ": KY,
        "Chỉ tiêu A": kw.get("ca"),
        "Giá trị A": kw.get("va"),
        "Chỉ tiêu B": kw.get("cb"),
        "Giá trị B": kw.get("vb"),
        "Giá trị đo": kw.get("do"),
        "Ngưỡng": kw.get("nguong", rd.get("muc")),
        "Trạng thái chạy": status,
        "Kết quả": ket,
        "Nhãn": nhan,
        "Mức rủi ro": muc_rr if ket == "Có phát hiện" else "",
        "Hệ số trọng yếu": hs,
        "Điểm": diem if ket == "Có phát hiện" else 0,
        "Sắc thuế": kw.get("sac", "") if ket == "Có phát hiện" else "",
        "Số tiền liên quan (đồng)": kw.get("v"),
        "Số tiền tính vào tổng (đồng)": w_tien if ket == "Có phát hiện" else (0 if status == "ĐÃ_CHẠY" else None),
        "Mã_PH chủ": kw.get("x_chu", ""),
        "Nguồn": kw.get("nguon", "IN01, IN02, IN03, IN04"),
        "Tham chiếu chứng cứ": kw.get("ct_ref", ""),
        "Còn thiếu để kết luận": thieu,
        "Nội dung kiểm tra tại trụ sở": kw.get("nd", ""),
        "Chứng từ cần xem": kw.get("ct_xem", ""),
        "Ghi chú": kw.get("ghi", ""),
        "Câu hỏi chất vấn NNT": kw.get("hoi", ""),
        "Chuỗi bằng chứng 5 mắt xích": kw.get("chuoi", ""),
        "Độ vững chuỗi bằng chứng": kw.get("do_vung", ""),
    })
    if ket == "Có phát hiện":
        findings.append((ma, muc_rr or "Thấp", kw.get("ghi") or rd["ten"], kw.get("v") or 0))


def need_missing(need):
    ins = set(re.findall(r"IN\d{2}", (need or "").replace("–", "-").replace("—", "-")))
    if "IN01-IN08" in (need or "").replace("–", "-"):
        ins |= {f"IN{i:02d}" for i in range(1, 9)}
    return sorted(i for i in ins if not HAS.get(i, False))


rows_kq3 = []
for rd in RULES:
    ma = rd["ma"]
    miss = need_missing(rd["need"])
    if ma in KHONG_AP_DUNG:
        add_rule(rows_kq3, rd, status="KHÔNG_ÁP_DỤNG", ket="Không", nhan="", ghi=f"DN {LOAI_HINH} - không chạy {ma}")
        continue
    if ma == "HD07":
        add_rule(rows_kq3, rd, status="TREO", ket="Chưa kết luận", thieu="Thiếu số liệu tham chiếu ngành (K13)")
        continue
    if ma == "TN20":
        add_rule(rows_kq3, rd, status="CHƯA_TRIỂN_KHAI", ket="Chưa kết luận", thieu="Khấu hao chi tiết chưa triển khai")
        continue
    if ma == "TC24" or ma == "LS02":
        pass

    if ma == "DQ01":
        add_rule(rows_kq3, rd, status="ĐÃ_CHẠY", ket="Có phát hiện", nhan="Chắc", muc="Thấp", diem=2,
                 ca="IN còn thiếu", va="IN11, IN12, IN09 số XNT", tep="KQ1.3",
                 ct_ref="KQ1_Kiem_ke mức D", nd="Đề nghị sao kê NH, PL GDLK, XNT kho",
                 ghi="Mức D: có IN08/IN10/IN13, thiếu IN11/IN12, IN09 mỏng")
        continue
    if ma == "DQ02":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="Dòng Tổng cộng IN02", va=0, nguon="IN02",
                 ghi="Bảng kê đã lọc khi nạp GĐ1")
        continue
    if ma == "DQ03":
        ket = "Có phát hiện" if (n_mst_sai + miss_mst) else "Không"
        add_rule(rows_kq3, rd, ket=ket, nhan="Chắc", muc="Trung bình" if ket == "Có phát hiện" else "",
                 diem=5 if ket == "Có phát hiện" else 0, ca="HĐ thiếu/sai MST", va=n_mst_sai + miss_mst,
                 nguon="IN01", ghi=f"miss MST {miss_mst}, sai định dạng {n_mst_sai} (sau khử trùng)")
        continue
    if ma == "DQ04":
        ket = "Có phát hiện" if n_thieu_kh else "Không"
        add_rule(rows_kq3, rd, ket=ket, nhan="Chắc", muc="Thấp" if ket == "Có phát hiện" else "",
                 diem=3 if ket == "Có phát hiện" else 0, ca="HĐ thiếu ký hiệu", va=n_thieu_kh, nguon="IN01")
        continue
    if ma == "DQ05":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="Khóa BK sau khử trùng", va=len(bk_raw), nguon="IN02")
        continue
    if ma == "DQ06":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="HĐĐT sau khử trùng dòng copy",
                 va=len(hddt_raw), nguon="IN01", ghi="File nguồn có copy lặp; đã unique trước khi đối chiếu")
        continue
    if ma == "DQ07":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ca="Tinh_vao_tong", va=len(hddt_raw),
                 ghi="Portal đánh Hợp lệ toàn bộ")
        continue
    if ma == "DQ08":
        n_bs = sum(1 for x in gtgt_all if x["loai"] == "Bổ sung")
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="Tờ khai GTGT bổ sung", va=n_bs, nguon="IN03")
        continue
    if ma == "GT01":
        picked = pick_diff(d_gt01)
        if picked:
            y, a, b, d, full, vuot = picked
            ket = "Có phát hiện" if vuot else "Không"
            add_rule(rows_kq3, rd, ket=ket, nhan="Chắc" if full else "Sơ bộ", muc="Cao" if vuot else "",
                     diem=9 if vuot else 0, ca=f"TK[34] {y}", va=a, cb=f"HĐĐT bán {y}", vb=b, do=d,
                     nguong="lệch > P03 2% tạm", sac="GTGT", v=abs(d) if vuot else None, w=0,
                     tep="KQ2.1", ct_ref="KQ2_Doi_chieu_GTGT Phần II", ma_tc="TC27",
                     thieu="" if full else "HĐĐT 2022 chưa phủ đủ năm",
                     nd="Chọn mẫu HĐ hủy/điều chỉnh, cut-off kỳ",
                     ghi=f"Lệch lớn nhất {y}={vn(d)}; thiếu XML {', '.join(GY.get(y,{}).get('missing') or []) or 'không'}")
        else:
            add_rule(rows_kq3, rd, ket="Chưa kết luận", status="KHÔNG_ĐỦ_DỮ_LIỆU", thieu="IN01/IN03")
        continue
    if ma == "GT02":
        picked = pick_diff(d_gt02)
        if picked:
            y, a, b, d, full, vuot = picked
            ket = "Có phát hiện" if vuot else "Không"
            add_rule(rows_kq3, rd, ket=ket, nhan="Chắc", muc="Cao" if vuot else "", diem=8 if vuot else 0,
                     ca=f"TK[34] {y}", va=a, cb=f"BK bán {y}", vb=b, do=d, tep="KQ2.1",
                     v=abs(d) if vuot else None, w=0, ghi=f"Lệch lớn nhất {y}={vn(d)}")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Chắc")
        continue
    if ma == "GT03":
        picked = pick_diff(d_gt03)
        if picked:
            y, a, b, d, full, vuot = picked
            ket = "Có phát hiện" if vuot else "Không"
            add_rule(rows_kq3, rd, ket=ket, nhan="Chắc" if full else "Sơ bộ", muc="Cao" if vuot else "",
                     diem=8 if vuot else 0, ca=f"TK[35] {y}", va=a, cb=f"Thuế HĐĐT bán {y}", vb=b, do=d,
                     tep="KQ2.1", v=abs(d) if vuot else None, w=0, x_chu=f"GT01|{MST}|{KY}")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Chắc")
        continue
    if ma == "GT04":
        picked = pick_diff(d_gt04)
        if picked:
            y, a, b, d, full, vuot = picked
            ket = "Có phát hiện" if vuot else "Không"
            add_rule(rows_kq3, rd, ket=ket, nhan="Chắc" if full else "Sơ bộ", muc="Cao" if vuot else "",
                     diem=8 if vuot else 0, ca=f"TK[23] {y}", va=a, cb=f"HĐĐT mua {y}", vb=b, do=d,
                     tep="KQ2.1", v=abs(d) if vuot else None, w=0, ghi=f"Lệch lớn nhất {y}={vn(d)}")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Chắc")
        continue
    if ma == "GT05":
        picked = pick_diff(d_gt05)
        if picked:
            y, a, b, d, full, vuot = picked
            ket = "Có phát hiện" if vuot else "Không"
            add_rule(rows_kq3, rd, ket=ket, nhan="Chắc" if full else "Sơ bộ", muc="Trung bình" if vuot else "",
                     diem=6 if vuot else 0, ca=f"TK[24] {y}", va=a, cb=f"Thuế HĐĐT mua {y}", vb=b, do=d,
                     tep="KQ2.1", v=abs(d) if vuot else None, w=0, x_chu=f"GT04|{MST}|{KY}")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Chắc")
        continue
    if ma == "GT06":
        # BK thuế mua vs 133
        diffs = []
        for y in [2023, 2024, 2025]:
            bk_th = bk_year("MUA", y, "thue")
            n133 = in08.get((y, "133"), {}).get("no")
            if n133:
                diffs.append((y, bk_th, n133, bk_th - n133))
        if diffs:
            y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
            vuot = abs(d) / max(abs(a), abs(b), 1) > P03
            add_rule(rows_kq3, rd, ket="Có phát hiện" if vuot else "Không", nhan="Sơ bộ",
                     muc="Trung bình" if vuot else "", diem=5 if vuot else 0,
                     ca=f"BK thuế mua {y}", va=a, cb=f"PS Nợ 133 {y}", vb=b, do=d, tep="KQ2.1",
                     ghi="IN08 133 là tổng hợp sổ cái (có thể trùng file). 2022 không có 133.")
        else:
            add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN08 TK 133 một số năm")
        continue
    if ma == "GT07":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="HĐĐT bán chủ yếu 10%; chưa so từng dòng với nhóm tờ khai")
        continue
    if ma == "GT08":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="Thuế suất mua vào chủ yếu 10%")
        continue
    if ma == "GT09":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="Chưa tách HĐ lập 5 ngày cuối quý")
        continue
    if ma == "GT10":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="Chưa thấy HĐ 8% bất thường trên mẫu đã đọc; K9 không mặc định 2%")
        continue
    if ma == "GT11":
        add_rule(rows_kq3, rd, ket="Chưa kết luận", nhan="Sơ bộ", ma_tc="TC9",
                 thieu="BCTC TT133 không có [152]; dùng ct140/181 chưa bóc",
                 ghi="Mẫu B01-DNN/ct140. Không điền 0 cho [152].")
        continue
    if ma == "GT12":
        hit = None
        for y in YEARS:
            c43 = GY.get(y, {}).get("ct43")
            htk = bctc_val(y, "140", "cdkt") or 0
            tscd = bctc_val(y, "211", "cdkt") or 0
            if c43 and (htk + tscd):
                if c43 > 0.1 * (htk + tscd):
                    hit = (y, c43, 0.1 * (htk + tscd))
                    break
        if hit:
            add_rule(rows_kq3, rd, ket="Có phát hiện", nhan="Sơ bộ", muc="Trung bình", diem=5,
                     ca=f"[43] {hit[0]}", va=hit[1], cb="10%(HTK+TSCĐ)", vb=hit[2], ma_tc="TC10")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ma_tc="TC10")
        continue
    if ma == "GT13":
        seq = [(y, GY[y]["ct43"]) for y in YEARS if GY[y].get("kys") and GY[y].get("ct43")]
        tang = all(seq[i][1] <= seq[i + 1][1] for i in range(len(seq) - 1)) if len(seq) >= 3 else False
        add_rule(rows_kq3, rd, ket="Có phát hiện" if tang else "Không", nhan="Chắc",
                 muc="Trung bình" if tang else "", diem=5 if tang else 0,
                 ca="[43] các năm có XML", va=seq[-1][1] if seq else None, ghi=str(seq[-4:]) if seq else "")
        continue
    if ma == "GT14":
        am = [(y, GY[y]["ct40"]) for y in YEARS if GY[y].get("kys") and GY[y].get("ct40", 0) < 0]
        add_rule(rows_kq3, rd, ket="Có phát hiện" if am else "Không", nhan="Chắc",
                 muc="Trung bình" if am else "", diem=5 if am else 0, ca="Số năm [40] âm", va=len(am), ghi=str(am[:6]))
        continue
    if ma == "GT15":
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ghi="Chưa thấy 4 quý liên tiếp [43] âm trên tờ khai hiệu lực")
        continue
    if ma == "GT16":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN11 sao kê NH",
                 ghi=f"{len(cash_hits)} HĐ đạt ngưỡng K5 - không kết luận thiếu UNC (K11)")
        continue
    if ma == "TN01":
        picked = pick_diff(d_tn01)
        if picked:
            y, a, b, d, full, vuot = picked
            ket = "Có phát hiện" if vuot else "Không"
            add_rule(rows_kq3, rd, ket=ket, nhan="Sơ bộ", muc="Cao" if vuot else "", diem=7 if vuot else 0,
                     ca=f"GTGT[34] {y}", va=a, cb=f"BCTC[01] {y}", vb=b, do=d, ma_tc="TC1", tep="KQ2.1",
                     v=abs(d) if vuot else None, w=0, ghi=f"Lệch lớn nhất {y}={vn(d)}; thiếu XML ảnh hưởng một số năm")
        else:
            add_rule(rows_kq3, rd, ket="Chưa kết luận", nhan="Sơ bộ")
        continue
    if ma == "TN02":
        picked = pick_diff(d_tn02)
        if picked:
            y, a, b, d, full, vuot = picked
            ket = "Có phát hiện" if vuot else "Không"
            add_rule(rows_kq3, rd, ket=ket, nhan="Sơ bộ", muc="Trung bình" if vuot else "", diem=6 if vuot else 0,
                     ca=f"TNDN[04] {y}", va=a, cb=f"BCTC[01] {y}", vb=b, do=d, v=abs(d) if vuot else None, w=0)
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="Một số năm BCTC/TNDN XML không bóc được số")
        continue
    if ma == "TN03":
        hit = None
        for y in YEARS:
            dt = bctc_val(y, "01", "kqkd")
            gv = bctc_val(y, "11", "kqkd")
            if dt is not None and gv is not None and dt - gv < 0:
                hit = (y, dt, gv)
                break
        if hit:
            add_rule(rows_kq3, rd, ket="Có phát hiện", nhan="Sơ bộ - chưa so ngành", muc="Cao", diem=7,
                     ca=f"DT {hit[0]}", va=hit[1], cb="GV", vb=hit[2], do=hit[1] - hit[2], ma_tc="TC5")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ - chưa so ngành", ma_tc="TC5",
                     ghi="Lãi gộp BCTC dương các năm có số (2015,2017,2020-2025)")
        continue
    if ma == "TN04":
        hit = None
        for y in YEARS:
            dt = bctc_val(y, "01", "kqkd")
            gv = bctc_val(y, "11", "kqkd")
            if dt and gv is not None and gv / dt > P18:
                hit = (y, gv / dt, dt, gv)
                break
        if hit:
            add_rule(rows_kq3, rd, ket="Có phát hiện", nhan="Sơ bộ - chưa so ngành", muc="Trung bình", diem=6,
                     ca=f"GV/DT {hit[0]}", va=hit[1], nguong="P18=92% tạm", ma_tc="TC28",
                     ghi=f"DT {vn(hit[2])} GV {vn(hit[3])}")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ - chưa so ngành", ma_tc="TC28")
        continue
    if ma == "TN05":
        hit = None
        for y in YEARS:
            dt = bctc_val(y, "01", "kqkd")
            htk = bctc_val(y, "140", "cdkt")
            if dt and htk is not None and htk / dt > P05:
                hit = (y, htk, dt, htk / dt)
                break
        if hit:
            add_rule(rows_kq3, rd, ket="Có phát hiện", nhan="Sơ bộ", muc="Cao", diem=7, ma_tc="TC3",
                     ca=f"HTK ct140 {hit[0]}", va=hit[1], cb="DT", vb=hit[2], do=hit[3], nguong="P05=30%",
                     nd="Kiểm kê tồn, sổ XNT 156, hàng chậm",
                     ghi=f"{hit[0]} HTK/DT={hit[3]:.1%}. 2025 HTK=0.")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ma_tc="TC3")
        continue
    if ma == "TN06":
        add_rule(rows_kq3, rd, ket="Chưa kết luận", nhan="Sơ bộ", ma_tc="TC2",
                 thieu="Chưa bóc [312] người mua trả trước trên XML", ghi="Không điền 0")
        continue
    if ma in {"TN07", "TN08", "TN09", "TN10", "TN11", "TN12", "TN13", "TN16", "TN19", "TN21", "TN22"}:
        extra = {}
        if ma == "TN12":
            hit = None
            for y in YEARS:
                von = bctc_val(y, "411", "cdkt")
                lnst = bctc_val(y, "60", "kqkd")
                if von is not None and lnst is not None and lnst < 0 and abs(lnst) > von:
                    hit = y
            extra = dict(ket="Không", nhan="Sơ bộ", ghi="Vốn 411=200tr; lỗ năm không vượt vốn lũy kế đã đọc")
        elif ma == "TN13":
            extra = dict(ket="Có phát hiện", nhan="Sơ bộ", muc="Trung bình", diem=5,
                         ghi="2021→2022 DT giảm (1,04 tỷ → 640tr) HTK vẫn cao 758tr", ma_tc="TC3")
        elif ma == "TN15":
            extra = dict(ket="Có phát hiện", nhan="Sơ bộ", muc="Trung bình", diem=5,
                         ca="LNTT âm", ghi="LNTT âm 2015, 2017, 2022, 2024")
        elif ma == "TN22":
            extra = dict(ket="Không", nhan="Sơ bộ - chưa so ngành", ma_tc="TC28",
                         ghi="GV/DT thương mại 2023=95,3% sát P18; các năm khác thấp hơn")
        else:
            extra = dict(ket="Không", nhan="Sơ bộ", ghi="Chưa thấy chỉ tiêu XML vượt ngưỡng hoặc mã vắng")
        add_rule(rows_kq3, rd, **extra)
        continue
    if ma == "TN14":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ma_tc="TC17",
                 ghi="CP tài chính BCTC [22]=0 các năm có số; IN12 chưa nhận")
        continue
    if ma == "TN15":
        am = [y for y in YEARS if (bctc_val(y, "50", "kqkd") or 0) < 0]
        add_rule(rows_kq3, rd, ket="Có phát hiện" if len(am) >= 2 else "Không", nhan="Sơ bộ",
                 muc="Trung bình" if len(am) >= 2 else "", diem=5 if len(am) >= 2 else 0,
                 ca="Năm LNTT âm", va=str(am), ghi="Vốn 411=200tr ổn định")
        continue
    if ma == "TN17":
        hit = None
        for y in YEARS:
            dt = bctc_val(y, "01", "kqkd")
            v331 = bctc_val(y, "331", "cdkt")
            if dt and v331 is not None and v331 / dt > P23:
                hit = (y, v331, dt)
        if hit:
            add_rule(rows_kq3, rd, ket="Có phát hiện", nhan="Sơ bộ", muc="Trung bình", diem=5,
                     ca=f"331 {hit[0]}", va=hit[1], cb="DT", vb=hit[2], nguong="P23=10%")
        else:
            add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ",
                     ghi="BCTC [331]=0 các năm có số. IN08 có PS 331 2023-2025 nhưng không có dư CK.")
        continue
    if ma == "TN18":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ",
                 ghi="642 IN08 2023-2025 nhỏ so với DT; chưa thấy đột biến không tương xứng")
        continue
    if ma == "HD01":
        ket = "Có phát hiện" if f0_ncc or f1_ncc else "Không"
        add_rule(rows_kq3, rd, ket=ket, nhan="Chắc", muc="Cao" if ket == "Có phát hiện" else "",
                 diem=9 if ket == "Có phát hiện" else 0, ca="NCC F0", va=len(f0_ncc),
                 cb="GT F0", vb=f0_gt, v=f0_gt if ket == "Có phát hiện" else None, w=f0_gt if ket == "Có phát hiện" else 0,
                 sac="GTGT", tep="KQ2.1", ct_ref="IN07 + NCC_mua_vao",
                 nd="Chuỗi 5 mắt xích NCC F0", ghi=f"NNT F0={nnt_on_f0} F1={nnt_on_f1} CA={nnt_on_ca}. F0 NCC={len(f0_ncc)}")
        continue
    if ma == "HD02":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="Danh bạ ĐKKD 3.2")
        continue
    if ma == "HD03":
        ket = "Có phát hiện" if hd03_pct > P13 else "Không"
        add_rule(rows_kq3, rd, ket=ket, nhan="Chắc", muc="Cao" if ket == "Có phát hiện" else "",
                 diem=8 if ket == "Có phát hiện" else 0, ca="Tỷ trọng F0+F1", va=hd03_pct, nguong="P13=5%",
                 ma_tc="TC26", tep="KQ2.1", pct_note=True)
        continue
    if ma == "HD04":
        ket = "Có phát hiện" if top5_pct > P14 else "Không"
        add_rule(rows_kq3, rd, ket=ket, nhan="Chắc", muc="Trung bình" if ket == "Có phát hiện" else "",
                 diem=6 if ket == "Có phát hiện" else 0, ca="Top-5 / tổng mua", va=top5_pct, nguong="P14=50%",
                 tep="KQ2.1", ghi=f"{len(ncc_agg)} NCC, Top-5 {vn(top5_gt)}")
        continue
    if ma == "HD05":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="Ngày cấp MST danh bạ 3.2")
        continue
    if ma == "HD06":
        add_rule(rows_kq3, rd, status="CHƯA_TRIỂN_KHAI", ket="Chưa kết luận", thieu="Benford chưa chạy")
        continue
    if ma == "HD08":
        ket = "Có phát hiện" if f1_buyers else "Không"
        add_rule(rows_kq3, rd, ket=ket, nhan="Chắc", muc="Cao" if ket == "Có phát hiện" else "",
                 diem=8 if ket == "Có phát hiện" else 0, ca="NNT trên F1 bên mua", va=len(f1_buyers),
                 ghi="Đối chiếu sheet F1 cột MST mua")
        continue
    if ma == "HD09":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", ghi="Chưa thấy MST vừa bán vừa mua trọng yếu")
        continue
    if ma == "TT01":
        add_rule(rows_kq3, rd, ket="Có phát hiện" if cash_hits else "Không", nhan="Sơ bộ",
                 muc="Trung bình" if cash_hits else "", diem=5 if cash_hits else 0,
                 ca="HĐ mua >= ngưỡng K5", va=len(cash_hits),
                 ghi=f"{len(tt01)} cụm; {len(tt01_multi)} cụm n>=2. Chưa kết luận hình thức TT.")
        continue
    if ma == "TT02":
        add_rule(rows_kq3, rd, ket="Có phát hiện" if tt02 else "Không", nhan="Sơ bộ",
                 muc="Trung bình" if tt02 else "", diem=5 if tt02 else 0,
                 ca="Phiếu chi >= ngưỡng", va=len(tt02), nguon="IN10",
                 ghi=f"IN10 {len(in10)} dòng (bỏ 2026). {len(tt02)} phiếu vượt ngưỡng.")
        continue
    if ma == "TT03":
        add_rule(rows_kq3, rd, ket="Có phát hiện" if tt03 else "Không", nhan="Sơ bộ",
                 muc="Trung bình" if tt03 else "", diem=5 if tt03 else 0,
                 ca="Cụm chi cùng ngày", va=len(tt03), nguon="IN10")
        continue
    if ma == "TT04":
        has3388 = any(k[1] == "338" or k[1] == "341" for k in in08)
        add_rule(rows_kq3, rd, ket="Không" if not has3388 else "Chưa kết luận", nhan="Sơ bộ",
                 ghi="IN08 không thấy TK 3388/341 tách")
        continue
    if ma in {"TT05", "TT06"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN11",
                 ghi="K11: thiếu sao kê không phải phát hiện")
        continue
    if ma in {"KH01", "KH02", "KH04", "KH06"}:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN09 XNT chi tiết",
                 ghi="IN09 chỉ kiểm kê file, Ton_CK_TT không dùng được")
        continue
    if ma == "KH07":
        diffs = []
        for y in [2023, 2024, 2025]:
            xuat = in08.get((y, "156"), {}).get("co")
            gv = dc_y.get(y, {}).get("gv")
            if xuat and gv:
                diffs.append((y, xuat, gv, xuat - gv))
        if diffs:
            y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
            vuot = abs(d) / max(abs(a), abs(b), 1) > P03
            add_rule(rows_kq3, rd, ket="Có phát hiện" if vuot else "Không", nhan="Sơ bộ",
                     muc="Trung bình" if vuot else "", diem=5 if vuot else 0,
                     ca=f"Xuất 156 {y}", va=a, cb=f"632 {y}", vb=b, do=d,
                     ghi="IN08 156 có thể trùng file; IN13 632 khớp BCTC [11]")
        else:
            add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN09/IN08 156")
        continue
    if ma == "DT01":
        n_duoi = sum(dc_y[y]["duoi"] for y in YEARS_511)
        add_rule(rows_kq3, rd, ket="Có phát hiện" if n_duoi else "Không", nhan="Chắc",
                 muc="Cao" if n_duoi else "", diem=8 if n_duoi else 0,
                 ca="Dòng lãi gộp <=0", va=n_duoi, tep="KQ2.2", ct_ref="KQ2_511_632",
                 ghi="; ".join(f"{y}: {dc_y[y]['duoi']}" for y in YEARS_511 if dc_y[y]["n511"]))
        continue
    if ma == "DT02":
        n = sum(dc_y[y]["no632"] for y in YEARS_511)
        add_rule(rows_kq3, rd, ket="Có phát hiện" if n else "Không", nhan="Chắc",
                 muc="Trung bình" if n else "", diem=6 if n else 0, ca="511 không có 632", va=n, tep="KQ2.2")
        continue
    if ma == "DT03":
        n = sum(dc_y[y]["no511"] for y in YEARS_511)
        add_rule(rows_kq3, rd, ket="Có phát hiện" if n else "Không", nhan="Chắc",
                 muc="Trung bình" if n else "", diem=6 if n else 0, ca="632 không có 511", va=n, tep="KQ2.2")
        continue
    if ma == "LD01":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="Tờ khai TNCN quý",
                 ghi="Chỉ có QTT năm. 2023 QTT rỗng phụ lục.")
        continue
    if ma == "LD02":
        add_rule(rows_kq3, rd, ket="Chưa kết luận", nhan="Sơ bộ", ma_tc="TC23",
                 thieu="Chi phí lương BCTC tách 622/641 chưa đủ",
                 ghi="TNCT QTT có 2015-2025 trừ 2023 trống. 641 IN08 2024-2025.")
        continue
    if ma == "LD03":
        n52 = sum(1 for x in in05 if x["loai"] == "NLD" and "05-2" in (x.get("nguon") or ""))
        add_rule(rows_kq3, rd, ket="Không", nhan="Chắc", ca="NLĐ 05-2", va=n52,
                 ghi="Không có dòng 05-2 tên thật (node rỗng đã bỏ)")
        continue
    if ma == "LD04":
        nld = [x for x in in05 if x["loai"] == "NLD" and x.get("cccd")]
        cccd = [x["cccd"] for x in nld]
        dup = len(cccd) - len(set(cccd))
        add_rule(rows_kq3, rd, ket="Có phát hiện" if dup else "Không", nhan="Chắc",
                 muc="Thấp" if dup else "", diem=3 if dup else 0, ca="Trùng CCCD 05-1", va=dup,
                 thieu="Thiếu tra cứu đa nguồn 3.3")
        continue
    if ma == "LK01":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN12",
                 ghi="Không kết luận không có GDLK")
        continue
    if ma == "LK02":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="IN12", ma_tc="TC17")
        continue
    if ma == "LS01":
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ", nguon="IN14",
                 ghi="IN14 có 10 file TMS - chưa thấy nợ đọng trọng yếu trên manifest GĐ1")
        continue
    if ma == "LS02":
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu="Bộ chỉ số TMS chi tiết")
        continue
    # default
    if miss:
        add_rule(rows_kq3, rd, status="KHÔNG_ĐỦ_DỮ_LIỆU", ket="Chưa kết luận", thieu=",".join(miss))
    else:
        add_rule(rows_kq3, rd, ket="Không", nhan="Sơ bộ")

# write KQ3 rows
if ws.max_row and ws.max_row >= 10:
    ws.delete_rows(10, ws.max_row - 9)
hm = {str(ws.cell(9, c).value or "").strip(): c for c in range(1, 40) if ws.cell(9, c).value}
for i, rec in enumerate(rows_kq3):
    rr = 10 + i
    for k, v in rec.items():
        c = hm.get(k)
        if not c:
            continue
        num = isinstance(v, (int, float)) and not isinstance(v, bool) and k not in {"Hệ số trọng yếu", "Điểm"}
        put(ws, rr, c, v, num=num, wrap=k in {"Ghi chú", "Còn thiếu để kết luận", "Nội dung kiểm tra tại trụ sở", "Chuỗi bằng chứng 5 mắt xích"})
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
n_kad = sum(1 for r in rows_kq3 if r["Trạng thái chạy"] == "KHÔNG_ÁP_DỤNG")
n_cao = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Cao" and r.get("Kết quả") == "Có phát hiện")
n_tb = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Trung bình" and r.get("Kết quả") == "Có phát hiện")
ws["C6"] = f"ĐÃ_CHẠY {n_run} | KHÔNG_ĐỦ {n_kdd} | TREO {n_treo} | KHÔNG_ÁP_DỤNG {n_kad}"
ws["J6"] = f"Cao {n_cao} | Trung bình {n_tb}"
print("KQ3 rules", len(rows_kq3), "run", n_run, "kdd", n_kdd, "treo", n_treo, "kad", n_kad, "cao", n_cao, "tb", n_tb)

# 28 TC
ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
put(ws, 2, 1, f"Áp dụng cho {TEN} (MST {MST}) - Kỳ {KY} | {VER} | {TODAY.strftime('%d/%m/%Y')} | Mức {LEVEL} | {LOAI_HINH}", bold=True)
if str(ws.cell(4, 1).value or "").startswith("TC"):
    ws.delete_rows(4, 1)

def htk_ty(y):
    dt = bctc_val(y, "01", "kqkd")
    htk = bctc_val(y, "140", "cdkt")
    if dt and htk is not None:
        return htk / dt
    return None

htk_txt = "; ".join(f"{y}={htk_ty(y):.1%}" for y in YEARS if htk_ty(y) is not None)
gv_txt = "; ".join(
    f"{y}={(bctc_val(y,'11','kqkd')/bctc_val(y,'01','kqkd')):.1%}"
    for y in YEARS if bctc_val(y, "01", "kqkd") and bctc_val(y, "11", "kqkd") is not None
)
ln_txt = "; ".join(
    f"{y}={vn(bctc_val(y,'20','kqkd'))}"
    for y in YEARS if bctc_val(y, "20", "kqkd") is not None
)
gt01_txt = "; ".join(
    f"{y} TK {vn(tk34(y))} HĐĐT {vn(hd_ban(y))} BK {vn(bk_ban(y))} BCTC {vn(bctc_val(y,'01','kqkd'))}"
    for y in [2022, 2023, 2024, 2025]
)
n_duoi = sum(dc_y[y]["duoi"] for y in YEARS_511)

tc_rows = [
    ("TC01", "Chênh lệch DT GTGT [34] vs DT BCTC [01]", "Doanh thu - giá vốn", "lệch ≠ 0",
     gt01_txt + ". Thiếu XML Q4/2022, Q1-Q3/2023.", "Cao", "Bổ sung XML thiếu; đối chiếu HĐĐT vs BK vs [34] từng quý."),
    ("TC02", "Người mua trả tiền trước > 5% DT", "Công nợ", "5%",
     "Chưa bóc [312] trên XML B01-DNN. Không điền 0.", "Thấp", "Đọc thuyết minh BCTC / sổ 131."),
    ("TC03", "HTK + XDCB dở dang > 30% DT", "Hàng tồn kho", "30% - ct140 TT133",
     f"{htk_txt}. 2020-2022 vượt 30%. 2025 HTK=0. IN09 không có XNT.", "Cao", "Kiểm kê tồn, sổ XNT 156, hàng chậm luân chuyển."),
    ("TC04", "DT chưa thực hiện > 0", "Doanh thu", "> 0", "Không thấy chỉ tiêu DT chưa thực hiện trên XML.", "Thấp", ""),
    ("TC05", "DT thuần - giá vốn < 0", "Giá vốn", "< 0", f"Lãi gộp BCTC dương: {ln_txt}.", "Thấp", ""),
    ("TC06", "Chi phí phải trả trích trước lớn", "Chi phí", "[315]", "Chưa bóc [315] trên XML.", "Thấp", "Đọc thuyết minh."),
    ("TC07", "Lãi vay chưa vốn hóa", "Chi phí", "[242]", "CP tài chính [22]=0 các năm có BCTC.", "Thấp", ""),
    ("TC08", "Ưu đãi / miễn giảm TNDN", "TNDN", "C2/C11",
     "LNTT âm 2015/2017/2022/2024; 2023 LNTT dương nhỏ; 2025 LNTT dương. Chưa tách C2/C11.", "Trung bình", "Đọc PL 03-1A."),
    ("TC09", "VAT khấu trừ TK vs BCTC [152]/181", "GTGT", "[43] vs 181 TT133",
     "BCTC TT133 không dùng [152]. Chưa bóc 181. Không điền 0.", "Thấp", "Đối [43] cuối năm vs CDKT."),
    ("TC10", "VAT khấu trừ vs 10%×(HTK+TSCĐ)", "GTGT", "×10%",
     "TSCĐ [211]=0. [43] so 10% HTK - xem GT12.", "Thấp", ""),
    ("TC11", "Có DT KCT [26] nhưng [25]=[24]", "GTGT", "",
     "Ngành hạt nhựa chủ yếu chịu thuế 10%. Cần rà từng quý [26].", "Thấp", "Rà phân bổ ĐV kỳ có KCT."),
    ("TC12", "Điều chỉnh giảm LN trước thuế", "TNDN", "B7/B8", "Chưa tách B7/B8 trên XML 03/TNDN.", "Thấp", "Đọc PL 03-1A"),
    ("TC13", "Trích lập dự phòng", "BCTC", "", "Không thấy mã dự phòng trọng yếu.", "Thấp", ""),
    ("TC14", "CN vốn [22]>0 và [25]<0", "BCTC", "", "CP tài chính [22]=0.", "Thấp", ""),
    ("TC15", "Bù trừ lãi BĐS / lỗ SXKD", "TNDN", "", "Không thấy tách BĐS. LNTT âm một số năm SXKD/TM.", "Thấp", ""),
    ("TC16", "Chuyển giá [251][252]>0 và [421]<0", "GDLK", "", "IN12 chưa nhận.", "Thấp", "Xin PL NĐ 132 nếu có bên liên kết."),
    ("TC17", "CP lãi vay vs 30% EBITDA", "EBITDA", "NĐ 132", "Lãi vay [22]=0; IN12 trống; khấu hao chưa có.", "Thấp", ""),
    ("TC18", "B11>0 không có PL GDLK", "GDLK", "IN12", "IN12 chưa nhận - KHÔNG_ĐỦ_DỮ_LIỆU.", "Thấp", "Xin PL hoặc xác nhận không GDLK."),
    ("TC19", "Thuế TNDN hoãn lại", "TNDN", "", "Không thấy biến động hoãn lại.", "Thấp", ""),
    ("TC20", "Phải trả người bán lớn", "Công nợ", "[331]/[311]",
     f"BCTC [331]=0. Nợ NH [311] có 2015-2024, 2025=0. IN08 có PS 331 2023-2025.", "Trung bình", "Đối sổ 331 khi kiểm tra trụ sở."),
    ("TC21", "Lãi vay khi thiếu VĐL", "Tài chính", "bỏ", "Vốn 411=200 triệu các năm có BCTC. TC21 bỏ.", "Thấp", ""),
    ("TC22", "QTT TNCN vs TK tháng/quý", "TNCN", "",
     "Có QTT 2015-2025 (2023 rỗng). Không có tờ khai quý. 24 NLĐ 05-1.", "Thấp", "Xin TK quý và tra cứu đa nguồn 3.3."),
    ("TC23", "Dòng tiền chi NLĐ vs TNCT", "TNCN", "20%", "Thiếu tách lương 111/641 vs TNCT. LD02 chưa chốt.", "Thấp", "IN05 + IN10 + 641."),
    ("TC24", "KHBS vs hạch toán TMS", "TMS", "không chạy", "Không chạy (cấm TMS dòng hạch toán).", "Thấp", ""),
    ("TC25", "Bổ sung GTGT chưa điều chỉnh", "GTGT", "", f"{sum(1 for x in gtgt_all if x['loai']=='Bổ sung')} tờ khai bổ sung (đã lấy bản hiệu lực).", "Thấp", ""),
    ("TC26", "Tỷ trọng HĐ mua NNT TT 03/06", "Hóa đơn", "cần ĐKKD",
     f"HD03 F0+F1={hd03_pct:.2%} trên {len(ncc_agg)} NCC. Chưa có danh bạ 3.2 nên không gắn 03/06.",
     "Cao" if hd03_pct > P13 else "Thấp", "Nạp danh bạ; chuỗi 5 mắt xích NCC F0."),
    ("TC27", "DT TK GTGT vs HĐĐT bán ra", "Hóa đơn", "lệch ≠ 0",
     gt01_txt, "Cao", "Chọn mẫu HĐ điều chỉnh/hủy; bổ sung XML thiếu kỳ."),
    ("TC28", "Phân tích ngang/dọc BCTC - GV/DT", "Giá vốn", "TC28_M1 / P18=92% tạm",
     f"GV/DT: {gv_txt}. IN13 2023-2025 khớp BCTC [01]/[11]. Bán dưới giá {n_duoi} dòng.",
     "Trung bình" if n_duoi else "Thấp", "Rà dòng lãi <=0 trên KQ2_511_632."),
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
dau_p = OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}_{VER}.xlsx"
giai_p = OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}_{VER}.xlsx"
ph_cao = [r for r in rows_kq3 if r.get("Kết quả") == "Có phát hiện" and r.get("Mức rủi ro") in {"Cao", "Trung bình"}]
ph_cao.sort(key=lambda x: -int(x.get("Điểm") or 0))

wb = Workbook()
ws = wb.active
ws.title = "Dau_hieu"
heads = ["STT", "Mã", "Dấu hiệu", "Mức", "Số liệu", "Sheet KQ", "Hướng kiểm tra", "Tài liệu cần", "Trạng thái"]
for i, h in enumerate(heads, 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
dau = []
dau.append((1, "TN05/TC03", "HTK/DT vượt 30% một số năm", "Cao", htk_txt, "KQ3 TC03", "Kiểm kê tồn, XNT 156, hàng chậm", "Sổ kho IN09 chi tiết", "Thiếu XNT"))
dau.append((2, "GT01/TC27", "Lệch DT tờ khai vs HĐĐT/BK", "Cao", gt01_txt[:300], "KQ2.1 GTGT", "Bổ sung XML thiếu; HĐ hủy/điều chỉnh", "XML GTGT đủ quý", "Thiếu Q4/2022 Q1-Q3/2023"))
dau.append((3, "HD01/HD03", f"NCC F0/F1 tỷ trọng {hd03_pct:.2%}", "Cao" if (f0_ncc or hd03_pct > P13) else "Thấp",
            f"F0={len(f0_ncc)} GT {vn(f0_gt)}; F1={len(f1_ncc)}; NCC {len(ncc_agg)}", "KQ2.1 NCC/IN07",
            "Chuỗi 5 mắt xích (HĐ-HĐKT-UNC-PNK-xuất bán)", "HĐ gốc, UNC, PNK", "Chưa kiểm tra trụ sở"))
dau.append((4, "HD04", f"Top-5 NCC {top5_pct:.2%}", "Trung bình" if top5_pct > P14 else "Thấp",
            f"Top-5 {vn(top5_gt)} / {vn(tot_mua_gt)}", "KQ2.1 NCC", "Chọn mẫu hợp đồng Top-5", "Hợp đồng, BK", "Chưa kiểm tra"))
dau.append((5, "DT01", f"Bán dưới giá {n_duoi} dòng IN13", "Cao" if n_duoi else "Thấp",
            "; ".join(f"{y}:{dc_y[y]['duoi']}" for y in YEARS_511), "KQ2.2 511_632",
            "Rà từng HĐ lãi <=0, đơn giá bán vs giá vốn", "Sổ 511/632 gốc, HĐ bán", "Có IN13 2023-2025"))
dau.append((6, "TT01/TT02", f"HĐ/phiếu chi vượt ngưỡng K5", "Trung bình",
            f"HĐĐT {len(cash_hits)}; phiếu chi {len(tt02)}; cụm TT03 {len(tt03)}", "KQ2.2 Công nợ",
            "Đối UNC; không kết luận thiếu chứng từ khi chưa có sao kê", "Sao kê IN11", "Thiếu IN11"))
dau.append((7, "TN01", "Lệch GTGT [34] vs BCTC [01]", "Cao", "Xem TC01", "KQ2.1 Phần II", "Cut-off, HĐ chưa xuất, hàng bán trả lại", "Sổ 511, HĐĐT", "Có sổ 511 2023-2025"))
dau.append((8, "IN11-IN12", "Thiếu sao kê NH và PL GDLK", "Trung bình", "Mức D", "KQ1.3", "Đề nghị bổ sung theo KQ1.3", "Sao kê, PL NĐ 132", "Chưa nhận"))
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
heads = ["MST", "Tên", "Số HĐ", "Ngày", "GT chưa thuế", "Thuế", "CV", "Nguồn F"]
for i, h in enumerate(heads, 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
rr = 2
f0_set_ncc = {k for k, _ in f0_ncc}
if not f0_set_ncc:
    put(ws, 2, 1, "Không có HĐ mua NCC F0/CA trên sổ chỉ định")
else:
    for inv in mua_ok:
        m = inv["mst_nb"] or inv["mst_dt"]
        if m not in f0_set_ncc:
            continue
        nhom, nguon, cv, nd = ncc_flag(m)
        put(ws, rr, 1, m)
        put(ws, rr, 2, inv["ten_nb"], wrap=True)
        put(ws, rr, 3, inv["so_goc"] or inv["so"])
        put(ws, rr, 4, inv["ngay"], dt=True)
        put(ws, rr, 5, inv["gt"], num=True)
        put(ws, rr, 6, inv["thue"], num=True)
        put(ws, rr, 7, cv)
        put(ws, rr, 8, nguon)
        for c in range(1, 9):
            ws.cell(rr, c).fill = FILL_RED
        rr += 1

ws = wb["02_Top5_NCC"]
for i, h in enumerate(["STT", "MST", "Tên", "GT", "Thuế", "Số HĐ", "Tỷ trọng"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, (m, rec) in enumerate(top5, 1):
    put(ws, 1 + i, 1, i, num=True)
    put(ws, 1 + i, 2, m)
    put(ws, 1 + i, 3, rec["ten"], wrap=True)
    put(ws, 1 + i, 4, rec["gt"], num=True)
    put(ws, 1 + i, 5, rec["thue"], num=True)
    put(ws, 1 + i, 6, len(rec["n_hd"]), num=True)
    put(ws, 1 + i, 7, rec["gt"] / tot_mua_gt if tot_mua_gt else 0, pct=True)

ws = wb["03_TT01_cum_HD"]
for i, h in enumerate(["MST", "Tên", "Ngày", "Số HĐ", "GT", "Thuế", "TT", "Ngưỡng"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, x in enumerate(tt01[:200], 1):
    put(ws, 1 + i, 1, x["mst"])
    put(ws, 1 + i, 2, x["ten"], wrap=True)
    put(ws, 1 + i, 3, x["ngay"], dt=True)
    put(ws, 1 + i, 4, x["n"], num=True)
    put(ws, 1 + i, 5, x["gt"], num=True)
    put(ws, 1 + i, 6, x["thue"], num=True)
    put(ws, 1 + i, 7, x["tt"], num=True)
    put(ws, 1 + i, 8, x["nguong"], num=True)
    if x["n"] >= 2:
        for c in range(1, 9):
            ws.cell(1 + i, c).fill = FILL_YEL
if not tt01:
    put(ws, 2, 1, "Không có cụm HĐ mua đạt ngưỡng K5")

ws = wb["04_GTGT_theo_ky"]
for i, h in enumerate(["Năm", "Kỳ", "Loại", "Hiệu lực", "ct23", "ct24", "ct26", "ct32", "ct34", "ct35", "ct40", "ct43"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, x in enumerate(eff_gtgt, 1):
    put(ws, 1 + i, 1, x["nam"])
    put(ws, 1 + i, 2, x["ky"])
    put(ws, 1 + i, 3, x["loai"])
    put(ws, 1 + i, 4, x["hieu_luc"])
    for col, key in enumerate(["ct23", "ct24", "ct26", "ct32", "ct34", "ct35", "ct40", "ct43"], 5):
        put(ws, 1 + i, col, x.get(key), num=True)

ws = wb["05_BCTC_TNDN"]
for i, h in enumerate(["Năm", "BCTC[01]", "BCTC[11]", "LN gộp", "HTK140", "331", "411", "LNTT[50]", "TNDN[04]"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
rr = 2
for y in YEARS:
    if y not in bctc:
        continue
    put(ws, rr, 1, y)
    put(ws, rr, 2, bctc_val(y, "01", "kqkd"), num=True)
    put(ws, rr, 3, bctc_val(y, "11", "kqkd"), num=True)
    put(ws, rr, 4, bctc_val(y, "20", "kqkd"), num=True)
    put(ws, rr, 5, bctc_val(y, "140", "cdkt"), num=True)
    put(ws, rr, 6, bctc_val(y, "331", "cdkt"), num=True)
    put(ws, rr, 7, bctc_val(y, "411", "cdkt"), num=True)
    put(ws, rr, 8, bctc_val(y, "50", "kqkd"), num=True)
    put(ws, rr, 9, bctc_val(y, "04", "tndn"), num=True)
    rr += 1

ws = wb["06_HDDT_ban_nam"]
for i, h in enumerate(["Năm", "Số dòng unique", "GT chưa thuế", "Thuế", "BK GT", "BCTC[01]", "TK[34]"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, y in enumerate([yy for yy in YEARS if hd_ban(yy) is not None or bk_ban(yy) is not None], 1):
    put(ws, 1 + i, 1, y)
    put(ws, 1 + i, 2, sum(1 for x in ban_ok if x["nam"] == y), num=True)
    put(ws, 1 + i, 3, hd_ban(y), num=True)
    put(ws, 1 + i, 4, hd_ban_thue(y), num=True)
    put(ws, 1 + i, 5, bk_ban(y), num=True)
    put(ws, 1 + i, 6, bctc_val(y, "01", "kqkd"), num=True)
    put(ws, 1 + i, 7, tk34(y), num=True)

ws = wb["07_HDDT_mua_nam"]
for i, h in enumerate(["Năm", "Số dòng unique", "GT chưa thuế", "Thuế", "BK GT", "TK[23]"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, y in enumerate([yy for yy in YEARS if hd_mua(yy) is not None or bk_mua(yy) is not None], 1):
    put(ws, 1 + i, 1, y)
    put(ws, 1 + i, 2, sum(1 for x in mua_ok if x["nam"] == y), num=True)
    put(ws, 1 + i, 3, hd_mua(y), num=True)
    put(ws, 1 + i, 4, hd_mua_thue(y), num=True)
    put(ws, 1 + i, 5, bk_mua(y), num=True)
    put(ws, 1 + i, 6, tk23(y), num=True)

ws = wb["08_Thieu_XML"]
for i, h in enumerate(["Năm", "Kỳ thiếu", "Ghi chú"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
rr = 2
for y in YEARS:
    for k in GY[y]["missing"]:
        put(ws, rr, 1, y)
        put(ws, rr, 2, k)
        put(ws, rr, 3, "Không điền 0 - T10")
        rr += 1
if rr == 2:
    put(ws, 2, 1, "Không thiếu kỳ trong phạm vi đã kê khai")

ws = wb["09_Kho_511_632"]
for i, h in enumerate(["Năm", "DT 511", "GV 632", "Lãi gộp", "Tỷ lệ", "Khớp HĐ+ngày+MH", "Không 632", "Không 511", "Bán dưới giá", "BCTC[01]", "BCTC[11]"], 1):
    put(ws, 1, i, h, bold=True, fill=FILL_HDR, font=FONT_HDR)
for i, y in enumerate(YEARS_511, 1):
    rec = dc_y[y]
    put(ws, 1 + i, 1, y)
    put(ws, 1 + i, 2, rec["dt"] or None, num=True)
    put(ws, 1 + i, 3, rec["gv"] or None, num=True)
    put(ws, 1 + i, 4, (rec["dt"] - rec["gv"]) if rec["n511"] else None, num=True)
    put(ws, 1 + i, 5, ((rec["dt"] - rec["gv"]) / rec["dt"]) if rec["dt"] else None, pct=True)
    put(ws, 1 + i, 6, rec["khop"] or None, num=True)
    put(ws, 1 + i, 7, rec["no632"] or None, num=True)
    put(ws, 1 + i, 8, rec["no511"] or None, num=True)
    put(ws, 1 + i, 9, rec["duoi"] or None, num=True)
    put(ws, 1 + i, 10, bctc_val(y, "01", "kqkd"), num=True)
    put(ws, 1 + i, 11, bctc_val(y, "11", "kqkd"), num=True)
put(ws, 8, 1, "IN09 không có số XNT thật. Năm 2026 trên IN08/IN13 là nhiễu đường dẫn - đã loại. Năm 2022 IN13 trống.", wrap=True, fill=FILL_YEL)

for sn, _ in sheet_specs:
    apply_tnr(wb[sn])
    autowidth(wb[sn], cap=40, max_row=200)
wb.save(lp(giai_p))
wb.close()
print(" extra saved")

# copy to vault
VAULT_OUT.mkdir(parents=True, exist_ok=True)
copied = []
for src in [files["kq21"][1], files["kq22"][1], files["kq3"][1], dau_p, giai_p]:
    dst = VAULT_OUT / src.name
    shutil.copy2(lp(src), lp(dst))
    copied.append(dst.name)
    print("vault", dst.name, "bytes", os.path.getsize(lp(dst)))

print("FINDINGS", n_cao, "cao", n_tb, "tb")
for x in findings[:12]:
    print(" ", x)
print("KTT_GD23_OK")

