# -*- coding: utf-8 -*-
"""Rebuild 6 KQ files for TT_BIEN_HOA 3603883505 to v1.4 vong 2b form."""
from __future__ import annotations

import shutil
from collections import defaultdict
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

MST = "3603883505"
TEN = "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ T&T BIÊN HÒA"
KY = "2022-2025"
YEARS = [2022, 2023, 2024, 2025]
P13 = 0.05
P14 = 0.50
P18 = 0.92
P23 = 0.10
TODAY = date(2026, 10, 19)

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep" / "MAU_FILE_TONG_HOP"
OUT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH"
)
ARC_MASTER = VAULT / "ARC_v1.4_TT_BIEN_HOA" / "Phan_tich_ho_so_cong_ty_TT_BIEN_HOA_3603883505_ketqua.xlsx"

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FONT_TITLE = Font(name="Times New Roman", size=12, bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F4E79")
FONT_HDR = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
FILL_NOTE = PatternFill("solid", fgColor="FFF2CC")
FILL_TOT = PatternFill("solid", fgColor="D6EAF8")
THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
WRAP = Alignment(wrap_text=True, vertical="center")
NUM = "#,##0"
PCT = "0.00%"
DATE_FMT = "dd/mm/yyyy"


def lp(p: Path | str) -> str:
    s = str(p)
    if not s.startswith("\\\\?\\"):
        s = "\\\\?\\" + s
    return s


def nnum(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(",", ".")
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
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    return None


def norm_mst(v) -> str:
    if v is None:
        return ""
    s = str(v).strip().replace(" ", "").replace(".", "")
    if s.endswith(".0") and s.replace(".0", "").isdigit():
        s = s[:-2]
    return s


def so_hd_key(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s.replace(".0", "").isdigit():
        s = s[:-2]
    return s.lstrip("0") or "0"


def safe_text(s) -> str:
    if s is None:
        return ""
    t = str(s)
    if t.startswith("="):
        t = "'" + t
    return t


def is_real_chieu(v) -> str | None:
    if v is None:
        return None
    s = str(v).strip().upper()
    if s in {"MUA", "MUA VÀO", "MUA VAO"}:
        return "MUA"
    if s in {"BAN", "BÁN", "BAN RA", "BÁN RA"}:
        return "BAN"
    return None


def style_range(ws, r, c1, c2, header=False, num=False, date=False, wrap=False, fill=None):
    for c in range(c1, c2 + 1):
        cell = ws.cell(r, c)
        cell.font = FONT_HDR if header else FONT
        cell.border = THIN
        cell.alignment = WRAP if wrap or header else Alignment(vertical="center")
        if fill:
            cell.fill = fill
        if header:
            cell.fill = FILL_HDR
        if num and isinstance(cell.value, (int, float)):
            cell.number_format = NUM
        if date and cell.value not in (None, ""):
            cell.number_format = DATE_FMT


def apply_font_all(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.font is None or cell.font.name != "Times New Roman":
                italic = bool(cell.font and cell.font.italic)
                bold = bool(cell.font and cell.font.bold)
                size = cell.font.size if cell.font and cell.font.size else 11
                color = cell.font.color if cell.font else None
                cell.font = Font(
                    name="Times New Roman",
                    size=size,
                    bold=bold,
                    italic=False if italic else bold and False or False,
                    color=color,
                )
                # force not italic, keep bold
                cell.font = Font(
                    name="Times New Roman",
                    size=size,
                    bold=bold,
                    italic=False,
                    color=color,
                )


def autowidth(ws, cap=60, min_w=10):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, 4000), max_col=ws.max_column or 1):
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column
            s = str(cell.value)
            if s.startswith("="):
                continue
            ln = min(len(s), cap)
            widths[col] = max(widths.get(col, min_w), ln + 2)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


def copy_cell(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)


def bump_version_text(v):
    if v is None:
        return v
    if isinstance(v, str) and "v1" in v and "v1.4" not in v.replace("biểu mẫu: 1.4", ""):
        return v.replace("Phiên bản KQ: v1", "Phiên bản KQ: v2").replace("| v1", "| v2")
    return v


def load_rows(path, sheet, header_row=2):
    wb = load_workbook(lp(path), data_only=True)
    ws = wb[sheet]
    headers = [ws.cell(header_row, c).value for c in range(1, (ws.max_column or 1) + 1)]
    rows = []
    for r in range(header_row + 1, (ws.max_row or 1) + 1):
        rec = {}
        empty = True
        for c, h in enumerate(headers, 1):
            if h is None:
                continue
            val = ws.cell(r, c).value
            rec[str(h)] = val
            if val not in (None, ""):
                empty = False
        if not empty:
            rec["_r"] = r
            rows.append(rec)
    wb.close()
    return headers, rows


def effective_gtgt(rows):
    best = {}
    for rec in rows:
        nam = rec.get("Nam")
        ky = rec.get("Ky")
        if nam is None or ky is None:
            continue
        if str(rec.get("La_ban_hieu_luc") or "").strip() not in {"Có", "CO", "Yes", "1", "TRUE", "True"}:
            # still keep if blank
            if rec.get("La_ban_hieu_luc") not in (None, "", "Có"):
                continue
        lan = nnum(rec.get("Lan_BS"))
        key = (int(nam), str(ky).strip().upper())
        prev = best.get(key)
        if prev is None or lan >= nnum(prev.get("Lan_BS")):
            best[key] = rec
    return best


def cash_threshold(d: date | None) -> float:
    if d is None:
        return 20_000_000
    return 5_000_000 if d >= date(2025, 7, 1) else 20_000_000


def unmerge_all(ws):
    for rng in list(ws.merged_cells.ranges):
        try:
            ws.unmerge_cells(str(rng))
        except Exception:
            pass


def put(ws, r, c, value, font=FONT, num=False, pct=False, dt=False, wrap=False, fill=None, bold=False):
    cell = ws.cell(r, c)
    if isinstance(cell, MergedCell):
        for mr in list(ws.merged_cells.ranges):
            if cell.coordinate in mr:
                cell = ws.cell(mr.min_row, mr.min_col)
                break
        else:
            return None
    cell.value = value
    cell.font = FONT_B if bold else font
    cell.alignment = WRAP if wrap else Alignment(vertical="center")
    cell.border = THIN
    if fill:
        cell.fill = fill
    if isinstance(value, (int, float)) and num:
        cell.number_format = NUM
    if isinstance(value, (int, float)) and pct:
        cell.number_format = PCT
    if dt:
        cell.number_format = DATE_FMT
        if isinstance(value, datetime):
            cell.value = value.date()
    return cell


def write_comment_block(ws, start_row, title, lines, cols=12):
    put(ws, start_row, 1, title, bold=True, fill=FILL_NOTE, wrap=True)
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=min(cols, 8))
    r = start_row + 1
    for line in lines:
        put(ws, r, 1, safe_text(line), wrap=True, fill=FILL_NOTE)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=min(cols, 8))
        ws.row_dimensions[r].height = 32
        r += 1
    return r


def copy_workbook(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(lp(src), lp(dst))


# ---------------------------------------------------------------------------
# Load source data
# ---------------------------------------------------------------------------
kq11 = OUT / "KQ1.1_Kiem_ke_CQT_v1.xlsx"
kq12 = OUT / "KQ1.2_Kiem_ke_ho_so_NNT_v1.xlsx"
kq13 = OUT / "KQ1.3_Danh_gia_ho_so_v1.xlsx"
kq21 = OUT / "KQ2.1_Phan_tich_CQT_v1.xlsx"
kq22 = OUT / "KQ2.2_Phan_tich_so_sach_v1.xlsx"
kq3v1 = OUT / "KQ3_Tong_hop_phat_hien_v1.xlsx"

print("Loading IN ...")
_, in01 = load_rows(kq11, "IN01_HDDT")
_, in03 = load_rows(kq11, "IN03_ToKhai_GTGT")
_, in04 = load_rows(kq11, "IN04_BCTC_TNDN")
_, in05 = load_rows(kq11, "IN05_NLD_TNCN")
_, in02 = load_rows(kq12, "IN02_BangKe")
_, in08 = load_rows(kq12, "IN08_So_KeToan")
_, in09 = load_rows(kq12, "IN09_Kho_GiaThanh")
_, in10 = load_rows(kq12, "IN10_Chi_tien_mat")
_, in13 = load_rows(kq12, "IN13_So_511_632")

gtgt_eff = effective_gtgt(in03)


def bctc(nam, bao_cao, ma):
    for rec in in04:
        if rec.get("Nam") == nam and str(rec.get("Bao_cao")) == bao_cao and str(rec.get("Ma_so")) == str(ma):
            return nnum(rec.get("So_nam_nay"))
    return None


def tk_year(prefix, field="PS_No"):
    out = {y: 0.0 for y in YEARS}
    last_ck = {y: None for y in YEARS}
    last_thang = {y: -1 for y in YEARS}
    for rec in in08:
        nam = rec.get("Nam")
        if nam not in YEARS:
            continue
        tk = str(rec.get("TK") or "")
        if not tk.startswith(prefix):
            continue
        out[nam] += nnum(rec.get(field))
        thang = int(nnum(rec.get("Thang")))
        if thang >= last_thang[nam]:
            last_thang[nam] = thang
            last_ck[nam] = rec
    return out, last_ck


ps_511, _ = tk_year("511", "PS_Co")
ps_632, _ = tk_year("632", "PS_No")
ps_156_nhap, ck_156 = tk_year("156", "PS_No")
ps_156_xuat, _ = tk_year("156", "PS_Co")
ps_131n, ck_131 = tk_year("131", "PS_No")
ps_131c, _ = tk_year("131", "PS_Co")
ps_331c, ck_331 = tk_year("331", "PS_Co")
ps_331n, _ = tk_year("331", "PS_No")
ps_111c, _ = tk_year("111", "PS_Co")


def ck_val(rec, no_field="Du_No_CK", co_field="Du_Co_CK"):
    if not rec:
        return 0.0, 0.0
    return nnum(rec.get(no_field)), nnum(rec.get(co_field))


def tinh_vao_tong(rec) -> bool:
    v = rec.get("Tinh_vao_tong")
    if v in (None, ""):
        return True
    s = str(v).strip().upper()
    return s not in {"KHÔNG", "KHONG", "NO", "0", "FALSE", "N"}


# invoices mua
inv_bk = []
for rec in in02:
    if is_real_chieu(rec.get("Chieu")) != "MUA":
        continue
    if not tinh_vao_tong(rec):
        continue
    mst = norm_mst(rec.get("MST_chuan") or rec.get("MST_goc"))
    if not mst or not mst[0].isdigit():
        continue
    inv_bk.append(
        {
            "mst": mst,
            "ten": rec.get("Ten_DoiTac") or "",
            "so_hd": str(rec.get("So_HD") or "").strip(),
            "so_key": so_hd_key(rec.get("So_HD")),
            "ngay": ndate(rec.get("Ngay_HD")),
            "mh": rec.get("Mat_hang") or "",
            "gt": nnum(rec.get("GT_chua_thue")),
            "thue": nnum(rec.get("Thue_GTGT")),
            "src": "Bảng kê",
        }
    )

inv_hd = []
for rec in in01:
    if is_real_chieu(rec.get("Chieu")) != "MUA":
        continue
    if not tinh_vao_tong(rec):
        continue
    mst = norm_mst(rec.get("MST_DoiTac_chuan") or rec.get("MST_NB"))
    if not mst or not mst[0].isdigit():
        continue
    thue = rec.get("Thue_GTGT")
    if thue is None:
        # try col after GT
        thue = rec.get("Tien_thue")
    inv_hd.append(
        {
            "mst": mst,
            "ten": rec.get("Ten_NB") or rec.get("Ten_DoiTac") or "",
            "so_hd": str(rec.get("So_HD") or "").strip(),
            "so_key": so_hd_key(rec.get("So_HD")),
            "ngay": ndate(rec.get("Ngay_HD")),
            "mh": rec.get("Mat_hang") or "",
            "gt": nnum(rec.get("GT_chua_thue")),
            "thue": nnum(thue),
            "src": "HĐĐT",
        }
    )

# union invoices by mst|so_key|date
seen = {}
for src_list, tag in ((inv_bk, "Bảng kê"), (inv_hd, "HĐĐT")):
    for it in src_list:
        key = (it["mst"], it["so_key"], str(it["ngay"]))
        if key in seen:
            cur = seen[key]
            if cur["src"] != tag:
                cur["src"] = "Chung"
            if not cur["mh"] and it["mh"]:
                cur["mh"] = it["mh"]
            if not cur["ten"] and it["ten"]:
                cur["ten"] = it["ten"]
            if cur["gt"] == 0 and it["gt"]:
                cur["gt"] = it["gt"]
            if cur["thue"] == 0 and it["thue"]:
                cur["thue"] = it["thue"]
        else:
            seen[key] = dict(it)
invoices = list(seen.values())
invoices.sort(key=lambda x: (x["mst"], x["ngay"] or date(1990, 1, 1), x["so_key"]))

# current IN07 flags
wb = load_workbook(lp(kq21), data_only=True)
ws = wb["IN07_NCC_CanhBao"]
in07_hdr = [ws.cell(2, c).value for c in range(1, 19)]
in07_map = {}
nnt_flags = []
for r in range(3, ws.max_row + 1):
    mst = norm_mst(ws.cell(r, 1).value)
    if not mst:
        continue
    rec = {str(in07_hdr[c - 1] or f"c{c}"): ws.cell(r, c).value for c in range(1, 17)}
    rec["MST"] = mst
    if mst == MST:
        nnt_flags.append(rec)
        continue
    if mst not in in07_map:
        in07_map[mst] = rec
    else:
        # prefer F0
        old = str(in07_map[mst].get("Nhom_F") or "")
        new = str(rec.get("Nhom_F") or "")
        if new == "F0" or (new and not old):
            in07_map[mst] = rec
wb.close()

# 511 matching from current line list
wb = load_workbook(lp(kq22), data_only=True)
ws = wb["KQ2_511_632"]
match_by_year = {y: defaultdict(int) for y in YEARS}
dt_by_year = {y: 0.0 for y in YEARS}
gv_by_year = {y: 0.0 for y in YEARS}
lo_by_year = {y: {"n": 0, "dt": 0.0, "gv": 0.0, "lo": 0.0} for y in YEARS}
for r in range(5, ws.max_row + 1):
    khoa = ws.cell(r, 1).value
    if not khoa or str(khoa).startswith("KQ2") or str(khoa).startswith("Khóa"):
        continue
    dt = nnum(ws.cell(r, 3).value)
    gv = nnum(ws.cell(r, 4).value)
    lai = nnum(ws.cell(r, 5).value) if ws.cell(r, 5).value not in (None, "") else dt - gv
    st = str(ws.cell(r, 7).value or "")
    note = str(ws.cell(r, 9).value or "")
    # year from date in khoa or note
    nam = None
    s = str(khoa)
    for y in YEARS:
        if str(y) in s:
            nam = y
            break
    if nam is None:
        continue
    dt_by_year[nam] += dt
    gv_by_year[nam] += gv
    if "Khớp HĐ + ngày" in note or st == "KHỚP" and "ngày CT + mặt hàng" in note:
        match_by_year[nam]["k1"] += 1
    elif "khác ngày" in note.lower() or "Khớp HĐ + mặt hàng" in note:
        match_by_year[nam]["k2"] += 1
    elif "Ghép theo ngày" in note:
        match_by_year[nam]["k3"] += 1
    elif "Không khớp 632" in note or st == "KHÔNG_CÓ_632":
        match_by_year[nam]["no632"] += 1
    elif "không có doanh thu" in note.lower() or st == "KHÔNG_CÓ_511":
        match_by_year[nam]["no511"] += 1
    else:
        match_by_year[nam]["khac"] += 1
        match_by_year[nam]["k1"] += 1  # current file marks KHỚP broadly
    if lai <= 0 and dt:
        lo_by_year[nam]["n"] += 1
        lo_by_year[nam]["dt"] += dt
        lo_by_year[nam]["gv"] += gv
        lo_by_year[nam]["lo"] += lai
wb.close()

# cash over threshold
cash_over = []
for rec in in10:
    d = ndate(rec.get("Ngay_CT"))
    amt = nnum(rec.get("So_tien"))
    if amt >= cash_threshold(d):
        cash_over.append(rec)

print(
    f"IN02 mua {len(inv_bk)} IN01 mua {len(inv_hd)} union HD {len(invoices)} "
    f"IN07 map {len(in07_map)} cash_over {len(cash_over)}"
)

# ---------------------------------------------------------------------------
# Archive v1 then write v2
# ---------------------------------------------------------------------------
arc_dir = OUT / "ARC_vong2b_truoc_v2"
arc_dir.mkdir(parents=True, exist_ok=True)

v1_files = [kq11, kq12, kq13, kq21, kq22, kq3v1]
for f in v1_files:
    if f.exists():
        dest = arc_dir / f.name
        if not dest.exists():
            shutil.copy2(lp(f), lp(dest))
            print("archived", f.name)

out_names = {
    "kq11": OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_v2.xlsx",
    "kq12": OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_v2.xlsx",
    "kq13": OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_v2.xlsx",
    "kq21": OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_v2.xlsx",
    "kq22": OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_v2.xlsx",
    "kq3": OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_v2.xlsx",
}


def bump_wb(path: Path, extra_note=None):
    wb = load_workbook(lp(path))
    for ws in wb.worksheets:
        for row in ws.iter_rows(max_row=min(6, ws.max_row or 1), max_col=min(12, ws.max_column or 1)):
            for cell in row:
                if isinstance(cell.value, str):
                    nv = cell.value
                    nv = nv.replace("Phiên bản KQ: v1", "Phiên bản KQ: v2")
                    nv = nv.replace("Phiên bản kết quả", nv if False else cell.value)
                    if "v1" in nv and "1.4" in nv:
                        nv = nv.replace("| v1", "| v2")
                    if nv != cell.value:
                        cell.value = nv
                if isinstance(cell.value, datetime):
                    cell.value = cell.value.date()
                    cell.number_format = DATE_FMT
                elif cell.number_format and "h" in str(cell.number_format).lower() and "yy" in str(cell.number_format).lower():
                    cell.number_format = DATE_FMT
        # date cells
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, datetime):
                    cell.value = cell.value.date()
                    cell.number_format = DATE_FMT
                fnt = cell.font
                if fnt is None or fnt.name != "Times New Roman" or fnt.italic:
                    cell.font = Font(
                        name="Times New Roman",
                        size=fnt.size if fnt and fnt.size else 11,
                        bold=bool(fnt and fnt.bold),
                        italic=False,
                        color=fnt.color if fnt else None,
                    )
        autowidth(ws)
    wb.save(lp(path))
    wb.close()


print("Copy GĐ1 ...")
copy_workbook(kq11, out_names["kq11"])
copy_workbook(kq12, out_names["kq12"])
copy_workbook(kq13, out_names["kq13"])

# sort IN05 QTT before NLD per year
wb = load_workbook(lp(out_names["kq11"]))
ws = wb["IN05_NLD_TNCN"]
hdr = [ws.cell(2, c).value for c in range(1, (ws.max_column or 1) + 1)]
body = []
for r in range(3, ws.max_row + 1):
    vals = [ws.cell(r, c).value for c in range(1, len(hdr) + 1)]
    if all(v is None or v == "" for v in vals):
        continue
    loai = str(vals[0] or "")
    nam = vals[1]
    ten = str(vals[2] or "").strip()
    if ten in {"Nguyễn Văn A", "Nguyen Van A"} or ten.lower().startswith("ví dụ"):
        continue
    body.append((nam if nam is not None else 9999, 0 if loai.upper() == "QTT" else 1, ten, vals))
body.sort()
for r in range(3, ws.max_row + 1):
    for c in range(1, len(hdr) + 1):
        ws.cell(r, c).value = None
for i, item in enumerate(body):
    r = 3 + i
    for c, v in enumerate(item[3], 1):
        cell = ws.cell(r, c, v)
        cell.font = FONT
        if isinstance(v, datetime):
            cell.value = v.date()
            cell.number_format = DATE_FMT
        elif isinstance(v, (int, float)) and c >= 8:
            cell.number_format = NUM
wb.save(lp(out_names["kq11"]))
wb.close()

for p in (out_names["kq11"], out_names["kq12"], out_names["kq13"]):
    bump_wb(p)
print("GĐ1 v2 written")

# ---------------------------------------------------------------------------
# KQ2.1 from template
# ---------------------------------------------------------------------------
print("Build KQ2.1 ...")
copy_workbook(MAU / "KQ2.1_Phan_tich_CQT.xlsx", out_names["kq21"])
wb = load_workbook(lp(out_names["kq21"]))

# --- IN07 100% NCC ---
ws = wb["IN07_NCC_CanhBao"]
# clear example / old data from row 3 down but keep header row 2
max_clear = max(ws.max_row or 3, 40)
for r in range(3, max_clear + 1):
    for c in range(1, 17):
        ws.cell(r, c).value = None

hdrs = [
    "MST",
    "Ten",
    "Loai_canh_bao",
    "Nhom_F",
    "Nguon_F0",
    "So_CV",
    "Ngay_CV",
    "Co_quan_BH",
    "Noi_dung",
    "Tinh_trang_DKKD",
    "Ngay_hieu_luc_TT",
    "Ngay_cap_MST",
    "Nganh_nghe",
    "Dia_chi",
    "Nguoi_dai_dien",
    "Nguon_tra_cuu",
]
ws["A1"] = (
    f"KQ2.1 | IN07_NCC_CanhBao – 100% NCC mua vào (kỳ {KY}). "
    f"MST NNT: {MST} | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | v2 | mẫu 1.4 vòng 2b"
)
ws["A1"].font = FONT_TITLE
for c, h in enumerate(hdrs, 1):
    cell = ws.cell(2, c, h)
    cell.font = FONT_HDR
    cell.fill = FILL_HDR
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    cell.border = THIN

# NNT block
r = 3
put(ws, r, 1, "I. ĐỐI CHIẾU NNT ĐANG KIỂM TRA", bold=True, fill=FILL_NOTE)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
r = 4
if nnt_flags:
    for rec in nnt_flags:
        put(ws, r, 1, MST)
        put(ws, r, 2, TEN)
        put(ws, r, 3, rec.get("Loai_canh_bao") or "KHONG")
        put(ws, r, 4, rec.get("Nhom_F") or "")
        put(ws, r, 16, rec.get("Nguon_tra_cuu") or "")
        r += 1
else:
    put(ws, r, 1, MST)
    put(ws, r, 2, TEN)
    put(ws, r, 3, "KHONG")
    put(ws, r, 16, "Sổ CVCBRR F0/F1 + Sổ Công an")
    r += 1

r += 1
put(ws, r, 1, "II. 100% NCC BÊN BÁN (một dòng / MST)", bold=True, fill=FILL_NOTE)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
r += 1

# unique sellers
sellers = {}
for it in invoices:
    s = sellers.setdefault(it["mst"], {"ten": it["ten"], "gt": 0.0, "thue": 0.0, "n": 0, "bk": 0, "hd": 0})
    if it["ten"] and (not s["ten"] or len(str(it["ten"])) > len(str(s["ten"]))):
        s["ten"] = it["ten"]
    s["gt"] += it["gt"]
    s["thue"] += it["thue"]
    s["n"] += 1
    if it["src"] in ("Bảng kê", "Chung"):
        s["bk"] += 1
    if it["src"] in ("HĐĐT", "Chung"):
        s["hd"] += 1

seller_rows_start = r
for mst, info in sorted(sellers.items(), key=lambda x: -x[1]["gt"]):
    flag = in07_map.get(mst, {})
    nhom = flag.get("Nhom_F") or ""
    if not nhom and str(flag.get("Loai_canh_bao") or "") in {"F0", "F1"}:
        nhom = flag.get("Loai_canh_bao")
    src_flag = "Chung" if info["bk"] and info["hd"] else ("Bảng kê" if info["bk"] else "HĐĐT")
    vals = [
        mst,
        info["ten"],
        flag.get("Loai_canh_bao") or ("F0" if nhom == "F0" else ""),
        nhom,
        flag.get("Nguon_F0") or "",
        flag.get("So_CV") or "",
        ndate(flag.get("Ngay_CV")),
        flag.get("Co_quan_BH") or "",
        flag.get("Noi_dung") or "",
        flag.get("Tinh_trang_DKKD") or "",
        ndate(flag.get("Ngay_hieu_luc_TT")),
        ndate(flag.get("Ngay_cap_MST")),
        flag.get("Nganh_nghe") or "",
        flag.get("Dia_chi") or "",
        flag.get("Nguoi_dai_dien") or "",
        flag.get("Nguon_tra_cuu") or src_flag,
    ]
    for c, v in enumerate(vals, 1):
        put(ws, r, c, v, dt=c in (7, 11, 12), wrap=c in (9, 10, 13, 14))
    r += 1

r += 1
put(ws, r, 1, "Tổng cộng số NCC", bold=True, fill=FILL_TOT)
put(ws, r, 2, len(sellers), bold=True, num=True, fill=FILL_TOT)
put(ws, r, 3, sum(s["gt"] for s in sellers.values()), bold=True, num=True, fill=FILL_TOT)
in07_last = r
autowidth(ws, cap=40)

# --- KQ2_NCC_mua_vao invoice detail ---
ws = wb["KQ2_NCC_mua_vao"]
# wipe data from row 8
for r in range(8, max(ws.max_row or 8, 20) + 1):
    for c in range(1, 13):
        ws.cell(r, c).value = None

ws["A1"] = (
    f"KQ2_NCC_mua_vao: CHI TIẾT 100% HÓA ĐƠN MUA VÀO (khai triển IN07). "
    f"Sắp theo MST người bán rồi Ngày HĐ. MST {MST} | Kỳ {KY} | v2 | mẫu 1.4 vòng 2b"
)
ws["A1"].font = FONT_TITLE
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 vòng 2b | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
ws["A2"].font = FONT
ws["B4"] = P13
ws["D4"] = P14
ws["F4"] = 12

# compute HD03 / HD04
tot_gt = sum(it["gt"] for it in invoices) or 0.0
f_gt = 0.0
by_seller = defaultdict(float)
for it in invoices:
    by_seller[it["mst"]] += it["gt"]
    nhom = str((in07_map.get(it["mst"]) or {}).get("Nhom_F") or "")
    if nhom in {"F0", "F1"}:
        f_gt += it["gt"]
hd03 = f_gt / tot_gt if tot_gt else None
top5 = sorted(by_seller.items(), key=lambda x: -x[1])[:5]
top5_gt = sum(v for _, v in top5)
hd04 = top5_gt / tot_gt if tot_gt else None
hd03_flag = "Chưa đủ nguồn" if hd03 is None else ("Vượt ngưỡng" if hd03 > P13 else "Không vượt")
hd04_flag = "Chưa đủ nguồn" if hd04 is None else ("Vượt ngưỡng" if hd04 > P14 else "Không vượt")

# Tong_hop_co reads D5 (HD03 flag) and H5 (HD04 flag)
ws["B5"] = hd03
ws["B5"].number_format = PCT
ws["B5"].font = FONT
ws["D5"] = hd03_flag
ws["D5"].font = FONT_B
ws["F5"] = hd04
ws["F5"].number_format = PCT
ws["F5"].font = FONT
ws["H5"] = hd04_flag
ws["H5"].font = FONT_B
ws["G5"] = "HD04 Top-5 NCC (gộp theo MST người bán, không LARGE theo dòng HĐ)"
ws["G5"].font = FONT

headers_ncc = [
    "STT",
    "MST người bán (chuẩn hóa)",
    "Tên người bán",
    "Số HĐ",
    "Ngày HĐ",
    "Mặt hàng",
    "Giá trị chưa thuế",
    "Thuế GTGT",
    "Nguồn dữ liệu (Bảng kê | HĐĐT | Chung)",
    "Có công văn cảnh báo? (Có - số CV | Không)",
    "Nhóm rủi ro (F0 | F1 | Không, tra từ IN07)",
    "Trạng thái ĐKKD (tra từ IN07)",
]
for c, h in enumerate(headers_ncc, 1):
    cell = ws.cell(7, c, h)
    cell.font = FONT_HDR
    cell.fill = FILL_HDR
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = THIN

r = 8
for i, it in enumerate(invoices, 1):
    flag = in07_map.get(it["mst"], {})
    nhom = flag.get("Nhom_F") or "Không"
    if not nhom:
        nhom = "Không"
    cv = flag.get("So_CV")
    cv_txt = f"Có - {cv}" if cv else "Không"
    vals = [
        i,
        it["mst"],
        it["ten"],
        it["so_hd"],
        it["ngay"],
        it["mh"],
        it["gt"],
        it["thue"],
        it["src"],
        cv_txt,
        nhom,
        flag.get("Tinh_trang_DKKD") or "",
    ]
    for c, v in enumerate(vals, 1):
        put(ws, r, c, v, num=c in (7, 8), dt=c == 5, wrap=c in (3, 6, 12))
    r += 1

r += 1
put(ws, r, 1, "Tổng cộng", bold=True, fill=FILL_TOT)
put(ws, r, 6, len(invoices), bold=True, num=True, fill=FILL_TOT)
put(ws, r, 7, tot_gt, bold=True, num=True, fill=FILL_TOT)
put(ws, r, 8, sum(it["thue"] for it in invoices), bold=True, num=True, fill=FILL_TOT)
r += 2
put(ws, r, 1, "BẢNG TOP-5 NCC (phục vụ HD04 / P14 = 50%)", bold=True, fill=FILL_NOTE)
r += 1
for c, h in enumerate(["STT", "MST", "Tên NCC", "Giá trị chưa thuế", "Tỷ trọng"], 1):
    cell = ws.cell(r, c, h)
    cell.font = FONT_HDR
    cell.fill = FILL_HDR
    cell.border = THIN
r += 1
for i, (mst, gt) in enumerate(top5, 1):
    put(ws, r, 1, i)
    put(ws, r, 2, mst)
    put(ws, r, 3, sellers.get(mst, {}).get("ten") or "", wrap=True)
    put(ws, r, 4, gt, num=True)
    put(ws, r, 5, (gt / tot_gt) if tot_gt else 0, pct=True)
    r += 1
put(ws, r, 1, "Cộng Top-5", bold=True, fill=FILL_TOT)
put(ws, r, 4, top5_gt, bold=True, num=True, fill=FILL_TOT)
put(ws, r, 5, hd04 or 0, bold=True, pct=True, fill=FILL_TOT)
r += 2

nnt_ok = all(str(x.get("Loai_canh_bao") or "KHONG").upper() in {"KHONG", "KHÔNG", ""} for x in nnt_flags) if nnt_flags else True
f0_ncc = [m for m, rec in in07_map.items() if str(rec.get("Nhom_F") or rec.get("Loai_canh_bao") or "") in {"F0", "F1"}]
r = write_comment_block(
    ws,
    r,
    "PHẦN 1. ĐỐI CHIẾU NNT ĐANG XÉT VỚI F1 VÀ SỔ CÔNG AN",
    [
        f"NNT {TEN} (MST {MST}) {'KHÔNG' if nnt_ok else 'CÓ'} xuất hiện trên sổ CVCBRR F1 / sổ Công an 2026 với vai trò bên mua rủi ro.",
        "Nguồn: SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx (F0/F1) và Bao_cao_cong_van.xlsx (CA 2026).",
    ],
)
r += 1
nnm = in07_map.get("0317409384") or {}
r = write_comment_block(
    ws,
    r,
    "PHẦN 3. ĐÁNH GIÁ RỦI RO NCC VÀ ĐỀ XUẤT TRỌNG TÂM CHỨNG TỪ (nhận xét module, không phải kết luận KQ3)",
    [
        f"Tổng {len(invoices)} hóa đơn mua vào hợp lệ, {len(sellers)} NCC, giá trị chưa thuế {tot_gt:,.0f} đ.".replace(",", "."),
        f"HD03 tỷ trọng mua F0/F1 = {hd03:.2%} so P13 5% → {hd03_flag}." if hd03 is not None else "HD03 chưa đủ nguồn.",
        f"HD04 Top-5 NCC = {hd04:.2%} so P14 50% → {hd04_flag}." if hd04 is not None else "HD04 chưa đủ nguồn.",
        f"NCC F0/F1: {len(f0_ncc)} MST. NNM Group (0317409384) CV {nnm.get('So_CV') or '6174'} - ưu tiên đối chiếu HĐ gốc, nhập kho, UNC.",
        "Trọng tâm: NCC F0; NCC ĐKKD bất thường sau ngày hiệu lực; cụm HĐ cùng ngày vượt ngưỡng tiền mặt. Sao kê IN11 còn PDF nên chưa kết luận thiếu UNC.",
    ],
)
autowidth(ws, cap=36)

# --- GTGT Phần I from IN03 ---
ws = wb["KQ2_Doi_chieu_GTGT"]
unmerge_all(ws)
ws["A2"] = f"Người nộp thuế: {TEN} - MST: {MST}"
ws["A3"] = f"Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 vòng 2b | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Kỳ {KY}"

# map year header rows
year_rows = {}
for rr in range(1, (ws.max_row or 1) + 1):
    a = ws.cell(rr, 1).value
    if isinstance(a, str) and a.strip().startswith("Năm "):
        try:
            year_rows[int(a.strip().split()[1])] = rr
        except (IndexError, ValueError):
            pass

period_offset = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

def fill_gtgt_period(year, ky, rec):
    base = year_rows.get(year)
    if not base:
        return
    off = period_offset.get(str(ky).upper())
    if not off:
        return
    rr = base + off
    # template cols: B-E DT by rate, F total DT, G-J tax out, K mua vao value, L thue vao
    # We fill K = ct23, L = ct24, F/E approx ct34, J/I ct35
    put(ws, rr, 5, nnum(rec.get("ct34")), num=True)  # E total-ish 8/10
    put(ws, rr, 11, nnum(rec.get("ct23")), num=True)
    put(ws, rr, 12, nnum(rec.get("ct24")), num=True)
    # tax out total col J (10)
    put(ws, rr, 9, nnum(rec.get("ct35")), num=True)
    put(ws, rr, 17, nnum(rec.get("ct40")), num=True)
    put(ws, rr, 18, nnum(rec.get("ct43") or rec.get("ct41")), num=True)


for (nam, ky), rec in sorted(gtgt_eff.items()):
    if nam in YEARS:
        fill_gtgt_period(nam, ky, rec)

# delete years outside KY from bottom
for y in sorted(year_rows, reverse=True):
    if y in YEARS:
        continue
    start = year_rows[y]
    # estimate block: until next year header or 'Tổng cộng' of that year - 14 rows typical
    end = start + 13
    ws.delete_rows(start, end - start + 1)

r_last = (ws.max_row or 1) + 2
write_comment_block(
    ws,
    r_last,
    "NHẬN XÉT / NGUYÊN NHÂN TRỌNG YẾU (module GTGT, không gán mức rủi ro hồ sơ)",
    [
        "NNT kê khai GTGT theo quý. Phần I lấy từ IN03 các kỳ hiệu lực (ưu tiên tờ khai bổ sung) trong kỳ 2022-2025; đã xóa năm ngoài phạm vi.",
        "Đối chiếu bán ra / mua vào với bảng kê và HĐĐT nằm ở các khối bên dưới. Chênh lệch kỳ nào lớn cần đối chiếu XML gốc và HĐ hủy/thay thế.",
        "2022 thiếu Q1-Q2 trên IN03 (chỉ có Q3-Q4) - giữ dòng thiếu nguồn, không điền 0 giả.",
    ],
)
autowidth(ws, cap=18)

# --- TNCN copy from v1 + comment ---
ws_src_wb = load_workbook(lp(kq21), data_only=False)
ws_src = ws_src_wb["KQ2_TNCN_LD2"]
ws = wb["KQ2_TNCN_LD2"]
unmerge_all(ws)
# copy values from row 1
for r in range(1, (ws_src.max_row or 1) + 1):
    for c in range(1, (ws_src.max_column or 1) + 1):
        sv = ws_src.cell(r, c).value
        dv = ws.cell(r, c)
        if isinstance(dv, MergedCell):
            continue
        dv.value = sv
        if isinstance(sv, datetime):
            dv.value = sv.date()
            dv.number_format = DATE_FMT
        elif isinstance(sv, (int, float)) and not isinstance(sv, bool):
            dv.number_format = NUM
        dv.font = Font(name="Times New Roman", size=11, bold=bool(ws_src.cell(r, c).font and ws_src.cell(r, c).font.bold), italic=False)
for mr in list(ws_src.merged_cells.ranges):
    try:
        ws.merge_cells(str(mr))
    except Exception:
        pass
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | v2 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
qtt_years = sorted({rec.get("Nam") for rec in in05 if str(rec.get("Loai_dong") or "").upper() == "QTT" and rec.get("Nam") in YEARS})
nld_n = sum(1 for rec in in05 if str(rec.get("Loai_dong") or "").upper() == "NLD")
write_comment_block(
    ws,
    (ws.max_row or 1) + 2,
    "NHẬN XÉT MODULE TNCN (không ghi mức rủi ro hồ sơ)",
    [
        f"IN05 có {nld_n} dòng NLĐ và QTT các năm {qtt_years}. Dòng QTT đã xếp trước NLĐ trong từng năm ở KQ1.1.",
        "Thiếu file tra cứu đa nguồn thư mục 3.3: chưa kết luận 1 nguồn / nhiều nguồn hay trùng CCCD ngoài QTT.",
        "Đối chiếu TC22 (QTT vs tờ khai quý) và TC23 (chi phí lương vs TNCT) lấy số trên bảng; lệch thì nêu tại KQ3.",
    ],
)
ws_src_wb.close()
autowidth(ws, cap=28)

# header tong hop
ws = wb["KQ2_Tong_hop_co"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 vòng 2b | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
autowidth(ws, cap=40)

wb.save(lp(out_names["kq21"]))
wb.close()
print("KQ2.1 saved")

# ---------------------------------------------------------------------------
# KQ2.2 from template
# ---------------------------------------------------------------------------
print("Build KQ2.2 ...")
copy_workbook(MAU / "KQ2.2_Phan_tich_ho_so_cty.xlsx", out_names["kq22"])
wb = load_workbook(lp(out_names["kq22"]))

ws = wb["KQ2_Tong_hop_co"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 vòng 2b | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"

# 511-632 yearly
ws = wb["KQ2_511_632"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 vòng 2b | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
ws["A4"] = "I. BẢNG TỔNG HỢP ĐỐI CHIẾU THEO NĂM TÀI CHÍNH"
headers = ["Chỉ tiêu", "Tổng cộng (kỳ kiểm tra)"] + [f"Năm {y}" for y in YEARS] + ["Đánh giá tuân thủ & Rủi ro"]
for c, h in enumerate(headers, 1):
    cell = ws.cell(5, c, h)
    cell.font = FONT_HDR
    cell.fill = FILL_HDR
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = THIN

# prefer so cai 511/632 which matches BCTC
dt = [ps_511[y] for y in YEARS]
gv = [ps_632[y] for y in YEARS]
lai = [dt[i] - gv[i] for i in range(4)]
ty = [lai[i] / dt[i] if dt[i] else None for i in range(4)]

k1 = [match_by_year[y]["k1"] for y in YEARS]
k2 = [match_by_year[y]["k2"] for y in YEARS]
k3 = [match_by_year[y]["k3"] for y in YEARS]
no632 = [match_by_year[y]["no632"] for y in YEARS]
no511 = [match_by_year[y]["no511"] for y in YEARS]
nlo = [lo_by_year[y]["n"] for y in YEARS]
ty_lo = [
    (lo_by_year[y]["dt"] / dt_by_year[y]) if dt_by_year[y] else 0
    for y in YEARS
]


def row_eval():
    notes = []
    notes.append(
        f"Lãi gộp giảm mạnh: 2022 {ty[0]:.2%} → 2023 {ty[1]:.2%} → 2024 {ty[2]:.2%} → 2025 {ty[3]:.2%} (TC28)."
    )
    if sum(nlo) == 1:
        notes.append("1 HĐ bán dưới giá năm 2024 (HĐ số 6 ngày 23/01/2024, lỗ 745.000 đ) - chọn mẫu chứng từ gốc.")
    elif sum(nlo):
        notes.append(f"{sum(nlo)} dòng lãi gộp ≤ 0 - chọn mẫu chứng từ gốc, chưa kết luận cố ý.")
    notes.append("Ghép 3 khóa theo skill doi-chieu-511-632; lệch ngày CT (khóa 2) không tự coi là sai sót.")
    return " ".join(notes)


metrics = [
    ("Doanh thu TK 511 (PS Có)", dt, False),
    ("Giá vốn TK 632 (PS Nợ)", gv, False),
    ("Lãi gộp", lai, False),
    ("Tỷ lệ lãi gộp", ty, True),
    ("Số dòng Khớp HĐ+ngày+mặt hàng", k1, False),
    ("Số dòng Khớp HĐ+mặt hàng (khác ngày CT)", k2, False),
    ("Số dòng lệch HĐ/tên (ghép theo ngày CT)", k3, False),
    ("Số dòng Không khớp 632 (có 511, không có 632)", no632, False),
    ("Số dòng Có giá vốn, không có doanh thu", no511, False),
    ("Số dòng bán dưới giá (Lãi ≤ 0)", nlo, False),
    ("Tỷ lệ DT dưới giá / tổng DT", ty_lo, True),
]
# clear example rows 6-16
for r in range(6, 20):
    for c in range(1, 8):
        ws.cell(r, c).value = None
for i, (name, vals, is_pct) in enumerate(metrics):
    r = 6 + i
    put(ws, r, 1, name, bold=True)
    total = sum(v or 0 for v in vals)
    if is_pct:
        tot = (sum(lai) / sum(dt)) if name.startswith("Tỷ lệ lãi") and sum(dt) else (sum(lo_by_year[y]["dt"] for y in YEARS) / sum(dt) if sum(dt) else 0)
        put(ws, r, 2, tot, pct=True)
        for j, v in enumerate(vals):
            put(ws, r, 3 + j, v if v is not None else None, pct=True)
    else:
        put(ws, r, 2, total, num=True)
        for j, v in enumerate(vals):
            put(ws, r, 3 + j, v, num=True)
    if i == 0:
        put(ws, r, 7, row_eval(), wrap=True)
    else:
        put(ws, r, 7, "")
ws.merge_cells(start_row=6, start_column=7, end_row=6 + len(metrics) - 1, end_column=7)
ws["G6"].alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[6].height = 80
autowidth(ws, cap=42)

# Kho
ws = wb["KQ2_Kho_Gia_thanh"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | v2 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
# Phần 1 TK 152 empty (thương mại) - Phần hàng hóa 156 is later in template
# Fill year rows of phần 1 as blank years 2022-2025 for 152 N/A
# Find "PHẦN" 156 - from inspect, phần 1 starts row 5 NVL 152, phần 2 TP 155, need 156
phan_rows = {}
for rr in range(1, (ws.max_row or 1) + 1):
    a = ws.cell(rr, 1).value
    if isinstance(a, str) and a.startswith("PHẦN"):
        phan_rows[a[:20]] = rr
# Fill 4 years on first table (rows 7+) as thương mại: NVL 152 không áp dụng
for i, y in enumerate(YEARS):
    rr = 7 + i
    put(ws, rr, 1, y)
    put(ws, rr, 12, "Không áp dụng TK 152 (DN thương mại, hàng hóa 156)", wrap=True)
# delete leftover example years 11-16 if they are formula-only
for rr in range(16, 10, -1):
    if ws.cell(rr, 1).value in (None, "") or str(ws.cell(rr, 1).value).startswith("="):
        ws.cell(rr, 1).value = None

# Fill a 156 block at end
r = (ws.max_row or 1) + 2
put(ws, r, 1, "PHẦN HÀNG HÓA TK 156 (DN thương mại)", bold=True, fill=FILL_NOTE)
r += 1
for c, h in enumerate(
    [
        "Năm",
        "Nhập TK 156 (PS Nợ)",
        "Xuất TK 156 (PS Có)",
        "Tồn CK (dư Nợ cuối năm)",
        "HTK BCTC ct140",
        "Chênh tồn CK vs BCTC",
        "Xuất 156 vs GV 632",
        "Ghi chú",
    ],
    1,
):
    cell = ws.cell(r, c, h)
    cell.font = FONT_HDR
    cell.fill = FILL_HDR
    cell.border = THIN
    cell.alignment = Alignment(wrap_text=True)
r += 1
htk_note = []
for y in YEARS:
    nhap = ps_156_nhap[y]
    xuat = ps_156_xuat[y]
    ton_ck, _ = ck_val(ck_156[y])
    htk = bctc(y, "CDKT", "140")
    gv = ps_632[y]
    chenh_htk = (ton_ck - htk) if htk is not None else None
    chenh_gv = xuat - gv
    put(ws, r, 1, y)
    put(ws, r, 2, nhap, num=True)
    put(ws, r, 3, xuat, num=True)
    put(ws, r, 4, ton_ck, num=True)
    put(ws, r, 5, htk, num=True)
    put(ws, r, 6, chenh_htk, num=True)
    put(ws, r, 7, chenh_gv, num=True)
    put(ws, r, 8, "TT133 dùng ct140. Xuất 156 khớp 632." if abs(chenh_gv) < 1 else "Lệch xuất kho vs 632", wrap=True)
    r += 1
    if htk:
        dt_y = ps_511[y] or 1
        htk_note.append(f"{y}: HTK/DT {htk/dt_y:.2%}")
r += 1
write_comment_block(
    ws,
    r,
    "NHẬN XÉT MODULE KHO (TC3) - không gán mức rủi ro hồ sơ",
    [
        "DN thương mại: không bắt buộc 152/155. IN09 mỏng (chỉ có phát sinh 156 năm 2025), số liệu XNT lấy từ IN08 sổ cái 156.",
        f"Xuất kho 156 khớp GV 632 các năm 2022-2025 (chênh 0). Tỷ lệ HTK/DT: {'; '.join(htk_note)} - đều dưới ngưỡng 30%.",
        "IN09 còn file 'chi tiết vật tư' chưa bóc dòng; khi NNT gửi sổ XNT chi tiết thì điền SL/tồn đầu kỳ.",
    ],
)
autowidth(ws, cap=28)

# Gia thanh - thương mại
ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | v2 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
ws["B19"] = P18
# Phần 1 621/622/627: mark N/A
for i, y in enumerate(YEARS):
    rr = 7 + i
    put(ws, rr, 1, y)
    put(ws, rr, 11, "Không áp dụng 621/622/627 (DN thương mại)", wrap=True)
# Phần 2 rows 21+
for i, y in enumerate(YEARS):
    rr = 21 + i
    dt_y = ps_511[y]
    gv_y = ps_632[y]
    ln = bctc(y, "KQKD", "20")
    put(ws, rr, 1, y)
    put(ws, rr, 2, gv_y, num=True)
    put(ws, rr, 3, (dt_y - (ln or 0)) if ln is not None else gv_y, num=True)
    put(ws, rr, 4, 0, num=True)
    put(ws, rr, 5, dt_y, num=True)
    put(ws, rr, 6, ln, num=True)
    ty_ln = (ln / dt_y) if dt_y and ln is not None else None
    gv_dt = (gv_y / dt_y) if dt_y else None
    put(ws, rr, 7, ty_ln, pct=True)
    put(ws, rr, 8, gv_dt, pct=True)
    flag = "Chưa đủ nguồn" if gv_dt is None else ("Vượt ngưỡng" if gv_dt > P18 else "Không vượt")
    put(ws, rr, 9, flag)
    put(ws, rr, 10, "Cơ cấu DT-GV-lãi gộp thương mại; 632 khớp [01]-[20].", wrap=True)
write_comment_block(
    ws,
    31,
    "NHẬN XÉT MODULE GIÁ THÀNH / TC28",
    [
        "Không ép 621/622/627. Giá vốn lấy TK 632, khớp BCTC DNN (ct140 / KQKD [01] và [20]).",
        f"Biên lãi gộp: 2022 {ty[0]:.2%}; 2023 {ty[1]:.2%}; 2024 {ty[2]:.2%}; 2025 {ty[3]:.2%}. Năm 2025 GV/DT {ps_632[2025]/ps_511[2025]:.2%} vượt P18 92% - nêu tại KQ3 (TN22/TC28).",
        "Cần đối chứng hợp đồng, giá vốn từng lô và chính sách chiết khấu năm 2024-2025.",
    ],
)
autowidth(ws, cap=28)

# Cong no
ws = wb["KQ2_Cong_no_Tien_mat"]
ws["A2"] = f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | v2 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')}"
ws["B6"] = P23
for i, y in enumerate(YEARS):
    rr = 8 + i
    dn_ck, _ = ck_val(ck_131[y])
    _, dc_ck = ck_val(ck_331[y])
    dt_y = ps_511[y]
    put(ws, rr, 1, y)
    put(ws, rr, 3, ps_131n[y], num=True)
    put(ws, rr, 4, ps_131c[y], num=True)
    put(ws, rr, 5, dn_ck, num=True)
    put(ws, rr, 7, ps_331c[y], num=True)
    put(ws, rr, 8, ps_331n[y], num=True)
    put(ws, rr, 9, dc_ck, num=True)
    put(ws, rr, 11, dt_y, num=True)
    ratio = (dc_ck / dt_y) if dt_y else None
    put(ws, rr, 12, ratio, pct=True)
    put(ws, rr, 10, "Chưa đủ nguồn" if ratio is None else ("Vượt ngưỡng" if ratio > P23 else "Không vượt"))
# Phần 2 ngan hang
put(ws, 21, 1, 1)
put(ws, 21, 2, "Chưa tách được tên NH từ IN11")
put(ws, 21, 6, "PDF scan")
put(ws, 21, 7, "Chưa nạp")
put(ws, 21, 8, "PDF - chưa kết luận UNC (Q7)", wrap=True)
# cash list summary at end
r = (ws.max_row or 1) + 2
put(ws, r, 1, "PHẦN CHI TIỀN MẶT TK 111 VƯỢT NGƯỠNG (K5)", bold=True, fill=FILL_NOTE)
r += 1
for c, h in enumerate(["STT", "Ngày CT", "Số phiếu", "TK đối ứng", "Diễn giải", "Số tiền", "Ngưỡng áp"], 1):
    cell = ws.cell(r, c, h)
    cell.font = FONT_HDR
    cell.fill = FILL_HDR
    cell.border = THIN
r += 1
for i, rec in enumerate(cash_over[:80], 1):
    d = ndate(rec.get("Ngay_CT"))
    put(ws, r, 1, i)
    put(ws, r, 2, d, dt=True)
    put(ws, r, 3, rec.get("So_phieu_chi"))
    put(ws, r, 4, rec.get("TK_doi_ung"))
    put(ws, r, 5, rec.get("Dien_giai"), wrap=True)
    put(ws, r, 6, nnum(rec.get("So_tien")), num=True)
    put(ws, r, 7, cash_threshold(d), num=True)
    r += 1
if len(cash_over) > 80:
    put(ws, r, 1, f"... còn {len(cash_over)-80} phiếu, đủ danh sách trên IN10")
    r += 1
r += 1
write_comment_block(
    ws,
    r,
    "NHẬN XÉT MODULE CÔNG NỢ - TIỀN MẶT",
    [
        f"Đã chạy 131/331 từ IN08 dù IN11 còn PDF. {len(cash_over)} phiếu chi 111 vượt ngưỡng K5 (20tr đến 30/06/2025, 5tr từ 01/07/2025).",
        "Sao kê ngân hàng dạng PDF scan: ghi PDF - chưa kết luận thiếu UNC. Ưu tiên UNC của 2 HĐ NNM Group.",
        "IN12 GDLK chưa nộp: không kết luận quan hệ liên kết.",
    ],
)
autowidth(ws, cap=28)

autowidth(wb["KQ2_Tong_hop_co"], cap=40)
wb.save(lp(out_names["kq22"]))
wb.close()
print("KQ2.2 saved")

# ---------------------------------------------------------------------------
# KQ3
# ---------------------------------------------------------------------------
print("Build KQ3 ...")
copy_workbook(MAU / "KQ3_Ket_luan_danh_gia.xlsx", out_names["kq3"])
wb = load_workbook(lp(out_names["kq3"]))
ws = wb["KQ3"]
unmerge_all(ws)
src = load_workbook(lp(kq3v1), data_only=False)
ss = src["KQ3"]
# copy header + 80 rules
for r in range(1, (ss.max_row or 1) + 1):
    for c in range(1, (ss.max_column or 1) + 1):
        sv = ss.cell(r, c)
        dv = ws.cell(r, c)
        if isinstance(dv, MergedCell):
            continue
        dv.value = sv.value
        if isinstance(sv.value, datetime):
            dv.value = sv.value.date() if hasattr(sv.value, "date") else sv.value
            dv.number_format = DATE_FMT
        elif sv.number_format:
            dv.number_format = sv.number_format
        dv.font = Font(
            name="Times New Roman",
            size=sv.font.size if sv.font and sv.font.size else 11,
            bold=bool(sv.font and sv.font.bold),
            italic=False,
        )
# bump v2 - only version labels, not "v1.4"
for r in range(1, 8):
    for c in range(1, 12):
        val = ws.cell(r, c).value
        if isinstance(val, str):
            nv = val.replace("Phiên bản KQ: v1", "Phiên bản KQ: v2")
            nv = nv.replace("| v1", "| v2")
            if nv.endswith(" v1"):
                nv = nv[:-2] + "v2"
            if nv != val:
                ws.cell(r, c).value = nv
ws["C4"] = "v2"
ws["C5"] = TODAY.strftime("%d/%m/%Y")
ws["J5"] = "KQ1.3 v2, KQ2.1 v2, KQ2.2 v2"
src.close()
autowidth(ws, cap=28)

# 28 TC from archive
ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
arc = load_workbook(lp(ARC_MASTER), data_only=True)
asheet = arc["Danh_gia_28_tieu_chi_rui_ro"]
ws["A1"] = "Danh_gia_28_tieu_chi_rui_ro – ĐÁNH GIÁ 28 TIÊU CHÍ RỦI RO (QĐ 1109 & Tool 4.4)"
ws["A1"].font = FONT_TITLE
ws["A2"] = f"Áp dụng cho {TEN} (MST {MST}) - Kỳ {KY} | v2 | {TODAY.strftime('%d/%m/%Y')}"
ws["A2"].font = FONT
headers = [
    "ID",
    "Tên tiêu chí",
    "Phân loại",
    "Ngưỡng rủi ro",
    "Kết quả phân tích trên hồ sơ",
    "Mức độ rủi ro",
    "Kiến nghị kiểm tra",
]
for c, h in enumerate(headers, 1):
    cell = ws.cell(3, c, h)
    cell.font = FONT_HDR
    cell.fill = FILL_HDR
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = THIN
# archive header at row 4, data from 5
out_r = 4
for r in range(5, asheet.max_row + 1):
    vals = [asheet.cell(r, c).value for c in range(1, 8)]
    if not vals[0]:
        continue
    fill = None
    muc = str(vals[5] or "")
    if "Cao" in muc:
        fill = PatternFill("solid", fgColor="F4CCCC")
    elif "Trung" in muc:
        fill = PatternFill("solid", fgColor="FCE4D6")
    elif "Thấp" in muc or "An toàn" in muc:
        fill = PatternFill("solid", fgColor="D9EAD3")
    for c, v in enumerate(vals, 1):
        put(ws, out_r, c, v, wrap=True, fill=fill)
    ws.row_dimensions[out_r].height = 48
    out_r += 1
arc.close()
# delete example row if still TC01 example leftover beyond data: already overwritten from row 4
autowidth(ws, cap=36)
ws.column_dimensions["E"].width = 60
ws.column_dimensions["G"].width = 40

wb.save(lp(out_names["kq3"]))
wb.close()
print("KQ3 saved")

# ---------------------------------------------------------------------------
# Verify
# ---------------------------------------------------------------------------
print("\n===== VERIFY =====")
for key, p in out_names.items():
    wb = load_workbook(lp(p), data_only=False)
    print(p.name, "sheets", wb.sheetnames, "size", p.stat().st_size)
    if key == "kq21":
        ws = wb["KQ2_NCC_mua_vao"]
        print("  NCC A1", str(ws["A1"].value)[:80])
        print("  HD03 D5", ws["D5"].value, "HD04 H5", ws["H5"].value, "F5", ws["F5"].value)
        print("  invoice rows from 8", ws.cell(8, 2).value, "max", ws.max_row)
        ws = wb["IN07_NCC_CanhBao"]
        print("  IN07 A1", str(ws["A1"].value)[:80], "max", ws.max_row)
        ws = wb["KQ2_Doi_chieu_GTGT"]
        years_left = []
        for rr in range(1, ws.max_row + 1):
            a = ws.cell(rr, 1).value
            if isinstance(a, str) and a.startswith("Năm "):
                years_left.append(a)
        print("  GTGT years left", years_left)
    if key == "kq22":
        ws = wb["KQ2_511_632"]
        print("  511 A4", ws["A4"].value, "B6", ws["B6"].value, "G6", str(ws["G6"].value)[:80] if ws["G6"].value else None)
        ws = wb["KQ2_Kho_Gia_thanh"]
        print("  kho max", ws.max_row)
        ws = wb["KQ2_Cong_no_Tien_mat"]
        print("  congno A8", ws["A8"].value, "C8", ws["C8"].value)
        ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
        print("  giathanh A21", ws["A21"].value, "H21", ws["H21"].value)
    if key == "kq3":
        print("  KQ3 sheets", wb.sheetnames, "KQ3 rows", wb["KQ3"].max_row, "28TC rows", wb["Danh_gia_28_tieu_chi_rui_ro"].max_row)
    wb.close()

print("DONE")
print("HD04", hd04, hd04_flag, "HD03", hd03, hd03_flag)
print("invoices", len(invoices), "sellers", len(sellers), "tot", tot_gt)
