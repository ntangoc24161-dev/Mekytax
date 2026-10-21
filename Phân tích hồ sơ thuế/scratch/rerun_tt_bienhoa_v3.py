# -*- coding: utf-8 -*-
"""Patch TT_BIEN_HOA 6 KQ files v2 -> v3 (Q9 bump). Fix flags, Q19, IN07, dates, Unicode, 511-632 note."""
from __future__ import annotations

import shutil
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MST = "3603883505"
TEN = "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ T&T BIÊN HÒA"
KY = "2022-2025"
TODAY = date(2026, 10, 20)
VER = "v3"

OUT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH"
)

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FILL_NOTE = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
DATE_FMT = "dd/mm/yyyy"
NUM = "#,##0"
PCT = "0.00%"

# TCVN3 (ABC) subset seen in IN10
TCVN3 = str.maketrans(
    {
        "µ": "à",
        "¸": "á",
        "¶": "ả",
        "·": "ã",
        "¹": "ạ",
        "¨": "ă",
        "»": "ắ",
        "¾": "ặ",
        "©": "ê",
        "Ê": "Ê",
        "Ç": "ề",
        "È": "ế",
        "É": "ể",
        "Ë": "ễ",
        "Ì": "ệ",
        "Ò": "ề",
        "Õ": "ế",
        "Ô": "ố",
        "Ö": "ồ",
        "Ø": "ổ",
        "ª": "ô",
        "ß": "ơ",
        "ú": "ớ",
        "û": "ờ",
        "ü": "ở",
        "ý": "ỡ",
        "þ": "ợ",
        "­": "",
        "¬": "ư",
        "Ù": "ừ",
        "Ụ": "ự",
        "¡": "ă",
        "£": "ô",
        "¤": "ơ",
        "¥": "ư",
        "®": "ỷ",
        "¯": "ỹ",
        "°": "ỵ",
        "±": "đ",
        "²": "Đ",
        "Ë": "ễ",
        "í": "í",
        "ì": "ì",
        "î": "ỉ",
        "ï": "ĩ",
        "ó": "ó",
        "ò": "ò",
        "ô": "ô",
        "õ": "õ",
        "ö": "ọ",
        "Ñ": "đ",
        "ñ": "đ",
        "Æ": "ấ",
        "Å": "ầ",
        "Ä": "ẩ",
        "Ã": "ẫ",
        "Ä": "ẩ",
        "á": "á",
        "à": "à",
        "ã": "ã",
        "ä": "ä",
        "å": "å",
        "æ": "æ",
        "ç": "ç",
        "è": "è",
        "é": "é",
        "ê": "ê",
        "ë": "ë",
        "¨": "ă",
    }
)


def lp(p: Path | str) -> str:
    s = str(p)
    if not s.startswith("\\\\?\\"):
        s = "\\\\?\\" + s
    return s


def looks_tcvn(s: str) -> bool:
    return any(ch in s for ch in "Òµ¬¨ËÕ¶·¹»¾¸")


def tcvn_fix(s):
    if not isinstance(s, str) or not s:
        return s
    if not looks_tcvn(s):
        return s
    t = s.replace("\xad", "")
    t = t.translate(TCVN3)
    repl = {
        "tiÒn": "tiền",
        "hµng": "hàng",
        "l¬ng": "lương",
        "tiÕp": "tiếp",
        "vËt t": "vật t",
        "x¨ng": "xăng",
        "Céng": "Cộng",
        "kh¸ch": "khách",
    }
    for a, b in repl.items():
        t = t.replace(a, b)
    return t


def unmerge_all(ws):
    for rng in list(ws.merged_cells.ranges):
        try:
            ws.unmerge_cells(str(rng))
        except Exception:
            pass


def bump_text(v):
    if not isinstance(v, str):
        return v
    nv = v
    nv = nv.replace("Ngày chạy: 19/10/2026", f"Ngày chạy: {TODAY.strftime('%d/%m/%Y')}")
    nv = nv.replace("| v2", f"| {VER}")
    nv = nv.replace("Phiên bản KQ: v2", f"Phiên bản KQ: {VER}")
    nv = nv.replace(" | v2 | ", f" | {VER} | ")
    if nv.endswith(" v2"):
        nv = nv[:-2] + VER
    return nv


def fix_dates_font(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str):
                nv = bump_text(cell.value)
                nv2 = tcvn_fix(nv)
                if nv2 != cell.value:
                    cell.value = nv2
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


def autowidth(ws, cap=60, min_w=10):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, 4000), max_col=ws.max_column or 1):
        for cell in row:
            if cell.value is None or isinstance(cell, MergedCell):
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            ln = min(len(s), cap)
            widths[cell.column] = max(widths.get(cell.column, min_w), ln + 2)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


v2 = {
    "kq11": OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_v2.xlsx",
    "kq12": OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_v2.xlsx",
    "kq13": OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_v2.xlsx",
    "kq21": OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_v2.xlsx",
    "kq22": OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_v2.xlsx",
    "kq3": OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_v2.xlsx",
}
v3 = {
    "kq11": OUT / f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_v3.xlsx",
    "kq12": OUT / f"KQ1.2_Kiem_ke_ho_so_NNT_{MST}_{KY}_v3.xlsx",
    "kq13": OUT / f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_v3.xlsx",
    "kq21": OUT / f"KQ2.1_Phan_tich_CQT_{MST}_{KY}_v3.xlsx",
    "kq22": OUT / f"KQ2.2_Phan_tich_ho_so_cty_{MST}_{KY}_v3.xlsx",
    "kq3": OUT / f"KQ3_Ket_luan_danh_gia_{MST}_{KY}_v3.xlsx",
}

for k, src in v2.items():
    if not src.exists():
        raise SystemExit(f"missing {src}")
    shutil.copy2(lp(src), lp(v3[k]))
    print("copied", v3[k].name)

# ---------------------------------------------------------------------------
# GĐ1: bump labels + dates
# ---------------------------------------------------------------------------
for key in ("kq11", "kq12", "kq13"):
    wb = load_workbook(lp(v3[key]))
    for ws in wb.worksheets:
        fix_dates_font(ws)
        autowidth(ws)
    wb.save(lp(v3[key]))
    wb.close()
print("GĐ1 v3 ok")

# ---------------------------------------------------------------------------
# KQ2.1
# ---------------------------------------------------------------------------
wb = load_workbook(lp(v3["kq21"]))

# IN07: keep 1 NNT row
ws = wb["IN07_NCC_CanhBao"]
nnt_rows = []
for r in range(4, min(10, (ws.max_row or 4) + 1)):
    mst = str(ws.cell(r, 1).value or "").strip()
    ten = str(ws.cell(r, 2).value or "")
    if mst == MST or (mst.startswith("3603883505") and "ĐỐI CHIẾU" not in ten):
        nnt_rows.append(r)
if len(nnt_rows) > 1:
    # keep first, clear extras
    for r in nnt_rows[1:]:
        for c in range(1, 17):
            ws.cell(r, c).value = None
    print("IN07 dedup NNT", nnt_rows)

# NCC invoice last data row (before Tong cong / Top-5)
ws = wb["KQ2_NCC_mua_vao"]
last_inv = 8
for r in range(8, (ws.max_row or 8) + 1):
    v = ws.cell(r, 1).value
    if v == "Tổng cộng":
        last_inv = r - 2  # blank row before total
        break
    if isinstance(v, (int, float)) and v >= 1:
        last_inv = r
print("NCC last invoice row", last_inv, "stt", ws.cell(last_inv, 1).value)

# HD02 / HD05 extra columns
ws.cell(7, 13).value = "HĐ sau ngày hiệu lực ĐKKD bất thường (Có | Không | Chưa đủ nguồn)"
ws.cell(7, 14).value = "NCC mới cấp MST < 12 tháng (Có | Không | Chưa đủ nguồn)"
for c in (13, 14):
    ws.cell(7, c).font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    ws.cell(7, c).fill = PatternFill("solid", fgColor="1F4E79")
    ws.cell(7, c).alignment = Alignment(wrap_text=True, vertical="center")
    ws.cell(7, c).border = THIN

# IN07 map MST -> DKKD date / Nhom_F / Ngay_cap
in07 = wb["IN07_NCC_CanhBao"]
flag = {}
for r in range(3, (in07.max_row or 3) + 1):
    mst = str(in07.cell(r, 1).value or "").strip()
    if not mst or not mst[0].isdigit() or mst == MST:
        continue
    flag[mst] = {
        "nhom": in07.cell(r, 4).value,
        "dkkd": in07.cell(r, 10).value,
        "hl": in07.cell(r, 11).value,
        "cap": in07.cell(r, 12).value,
        "loai": in07.cell(r, 3).value,
    }

n_f0 = 0
n_one_src = 0
n_hd02 = 0
n_hd05 = 0
n_eval = 0
for r in range(8, last_inv + 1):
    mst = str(ws.cell(r, 2).value or "").strip()
    if not mst:
        continue
    n_eval += 1
    nhom = str(ws.cell(r, 11).value or "")
    if nhom in {"F0", "F1"}:
        n_f0 += 1
    src = str(ws.cell(r, 9).value or "")
    if src in {"Bảng kê", "HĐĐT"}:
        n_one_src += 1
        # keep source label; Tong_hop_co will count these
    ngay = ws.cell(r, 5).value
    if isinstance(ngay, datetime):
        ngay = ngay.date()
        ws.cell(r, 5).value = ngay
        ws.cell(r, 5).number_format = DATE_FMT
    rec = flag.get(mst, {})
    hl = rec.get("hl")
    if isinstance(hl, datetime):
        hl = hl.date()
    dkkd = str(rec.get("dkkd") or rec.get("loai") or "")
    abnormal = any(
        k in dkkd.lower()
        for k in ["ngừng", "ngung", "giải thể", "giai the", "không hoạt động", "khong hoat dong", "chấm dứt", "cham dut"]
    ) or str(rec.get("loai") or "") == "DKKD"
    if ngay and hl and abnormal and ngay > hl:
        ws.cell(r, 13).value = "Có"
        n_hd02 += 1
    elif not hl:
        ws.cell(r, 13).value = "Chưa đủ nguồn" if abnormal else "Không"
    else:
        ws.cell(r, 13).value = "Không"
    ws.cell(r, 13).font = FONT
    ws.cell(r, 13).border = THIN
    cap = rec.get("cap")
    if isinstance(cap, datetime):
        cap = cap.date()
    if ngay and cap:
        months = (ngay.year - cap.year) * 12 + (ngay.month - cap.month)
        if months < 12:
            ws.cell(r, 14).value = "Có"
            n_hd05 += 1
        else:
            ws.cell(r, 14).value = "Không"
    else:
        ws.cell(r, 14).value = "Chưa đủ nguồn"
    ws.cell(r, 14).font = FONT
    ws.cell(r, 14).border = THIN

print("flags HD01 F0/F1 rows", n_f0, "one-src", n_one_src, "HD02", n_hd02, "HD05", n_hd05, "eval", n_eval)

# Tong_hop_co: point formulas at real columns + last row
th = wb["KQ2_Tong_hop_co"]
# HD01 nhóm F = col K
th["F12"] = f'=(SUMPRODUCT(--(LEN(KQ2_NCC_mua_vao!$K$8:$K${last_inv})>0)))-(COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"Chưa đủ nguồn"))'
th["G12"] = f'=COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"F0")+COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"F1")'
th["H12"] = f'=COUNTIF(KQ2_NCC_mua_vao!$K$8:$K${last_inv},"Chưa đủ nguồn")'
# one source = col I
th["F13"] = f'=(SUMPRODUCT(--(LEN(KQ2_NCC_mua_vao!$I$8:$I${last_inv})>0)))-(COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"Chưa đủ nguồn"))'
th["G13"] = f'=COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"Bảng kê")+COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"HĐĐT")'
th["H13"] = f'=COUNTIF(KQ2_NCC_mua_vao!$I$8:$I${last_inv},"Chưa đủ nguồn")'
# HD02 col M
th["F14"] = f'=(SUMPRODUCT(--(LEN(KQ2_NCC_mua_vao!$M$8:$M${last_inv})>0)))-(COUNTIF(KQ2_NCC_mua_vao!$M$8:$M${last_inv},"Chưa đủ nguồn"))'
th["G14"] = f'=COUNTIF(KQ2_NCC_mua_vao!$M$8:$M${last_inv},"Có")'
th["H14"] = f'=COUNTIF(KQ2_NCC_mua_vao!$M$8:$M${last_inv},"Chưa đủ nguồn")'
# HD05 col N
th["F17"] = f'=(SUMPRODUCT(--(LEN(KQ2_NCC_mua_vao!$N$8:$N${last_inv})>0)))-(COUNTIF(KQ2_NCC_mua_vao!$N$8:$N${last_inv},"Chưa đủ nguồn"))'
th["G17"] = f'=COUNTIF(KQ2_NCC_mua_vao!$N$8:$N${last_inv},"Có")'
th["H17"] = f'=COUNTIF(KQ2_NCC_mua_vao!$N$8:$N${last_inv},"Chưa đủ nguồn")'

# GTGT: delete empty month rows T05-T12 (Q19)
ws = wb["KQ2_Doi_chieu_GTGT"]
unmerge_all(ws)
to_del = []
for r in range(1, (ws.max_row or 1) + 1):
    a = str(ws.cell(r, 1).value or "").strip().upper()
    if a in {f"T{i:02d}" for i in range(5, 13)} or a in {f"T{i}" for i in range(5, 13)}:
        nums = [ws.cell(r, c).value for c in range(2, 20)]
        if not any(isinstance(v, (int, float)) and v for v in nums):
            to_del.append(r)
for r in reversed(to_del):
    ws.delete_rows(r, 1)
print("GTGT deleted empty month rows", len(to_del), to_del[:12])

for ws in wb.worksheets:
    fix_dates_font(ws)
    autowidth(ws, cap=40)
wb.save(lp(v3["kq21"]))
wb.close()
print("KQ2.1 v3 saved")

# ---------------------------------------------------------------------------
# KQ2.2
# ---------------------------------------------------------------------------
wb = load_workbook(lp(v3["kq22"]))
ws = wb["KQ2_511_632"]
# clarify khóa 3 vs "khớp 287/287"
note = (
    "287/287 dòng ghép được, không thiếu 632. Toàn bộ khóa 3 vì diễn giải 632 ghi "
    "'Xuất bán theo HĐ : n' khác tên hàng 511 - đây là đặc thù sổ, không tự kết luận sai sót. "
    "Lãi gộp giảm mạnh: 2022 27,06% → 2023 30,29% → 2024 8,49% → 2025 1,47% (TC28). "
    "1 HĐ bán dưới giá năm 2024 (HĐ số 6 ngày 23/01/2024, lỗ 745.000 đ) - chọn mẫu chứng từ gốc."
)
ws["G6"] = note
ws["G6"].alignment = Alignment(wrap_text=True, vertical="top")
ws["G6"].font = FONT

th = wb["KQ2_Tong_hop_co"]
# DT02: số dòng không khớp 632/511 = B13 + B14 of yearly table
th["F8"] = "=IF(KQ2_511_632!B10=\"\",\"Chưa đủ nguồn\",4)"
th["G8"] = "=N(KQ2_511_632!B13)+N(KQ2_511_632!B14)"
th["H8"] = 0
th["I8"] = '=IF(G8>0,"Có bất thường","Không")'
# DT01 bán dưới giá = B15
th["F9"] = "=IF(KQ2_511_632!B15=\"\",\"Chưa đủ nguồn\",4)"
th["G9"] = "=N(KQ2_511_632!B15)"
th["H9"] = 0
th["I9"] = '=IF(G9>0,"Có bất thường","Không")'
# TN04 years 21-24
th["F15"] = '=(SUMPRODUCT(--(LEN(KQ2_Gia_thanh_va_Dinh_muc!$I$21:$I$24)>0)))-(COUNTIF(KQ2_Gia_thanh_va_Dinh_muc!$I$21:$I$24,"Chưa đủ nguồn"))'
th["G15"] = '=COUNTIF(KQ2_Gia_thanh_va_Dinh_muc!$I$21:$I$24,"Vượt ngưỡng")'
th["H15"] = '=COUNTIF(KQ2_Gia_thanh_va_Dinh_muc!$I$21:$I$24,"Chưa đủ nguồn")'
# TN17 331
th["F16"] = '=(SUMPRODUCT(--(LEN(KQ2_Cong_no_Tien_mat!$J$8:$J$11)>0)))-(COUNTIF(KQ2_Cong_no_Tien_mat!$J$8:$J$11,"Chưa đủ nguồn"))'
th["G16"] = '=COUNTIF(KQ2_Cong_no_Tien_mat!$J$8:$J$11,"Vượt ngưỡng")'
th["H16"] = '=COUNTIF(KQ2_Cong_no_Tien_mat!$J$8:$J$11,"Chưa đủ nguồn")'

# cash 3a: find "PHẦN CHI TIỀN MẶT" block and point formula there; also fill Cờ đạt ngưỡng
ws = wb["KQ2_Cong_no_Tien_mat"]
cash_header = None
for r in range(1, (ws.max_row or 1) + 1):
    a = str(ws.cell(r, 1).value or "")
    if a.startswith("PHẦN CHI TIỀN MẶT"):
        cash_header = r
        break
if cash_header:
    first = cash_header + 2
    last = first
    for r in range(first, (ws.max_row or first) + 1):
        if ws.cell(r, 1).value in (None, "") and ws.cell(r, 6).value in (None, ""):
            break
        last = r
        # col 8 in this mini table is not K; set a flag col if amount >= threshold
        amt = ws.cell(r, 6).value
        nguong = ws.cell(r, 7).value
        if isinstance(amt, (int, float)) and isinstance(nguong, (int, float)) and amt >= nguong:
            ws.cell(r, 8).value = "Vượt ngưỡng"
            ws.cell(r, 8).font = FONT_B
    th["F17"] = f'=(SUMPRODUCT(--(LEN(KQ2_Cong_no_Tien_mat!$H${first}:$H${last})>0)))-(COUNTIF(KQ2_Cong_no_Tien_mat!$H${first}:$H${last},"Chưa đủ nguồn"))'
    th["G17"] = f'=COUNTIF(KQ2_Cong_no_Tien_mat!$H${first}:$H${last},"Vượt ngưỡng")'
    th["H17"] = f'=COUNTIF(KQ2_Cong_no_Tien_mat!$H${first}:$H${last},"Chưa đủ nguồn")'
    print("cash flag rows", first, last)

for ws in wb.worksheets:
    fix_dates_font(ws)
    autowidth(ws, cap=42)
wb.save(lp(v3["kq22"]))
wb.close()
print("KQ2.2 v3 saved")

# ---------------------------------------------------------------------------
# KQ3
# ---------------------------------------------------------------------------
wb = load_workbook(lp(v3["kq3"]))
ws = wb["KQ3"]
if ws["C4"].value in ("v1", "v2", None):
    ws["C4"] = VER
# ngày
for r in range(1, 8):
    for c in range(1, 16):
        v = ws.cell(r, c).value
        if isinstance(v, str):
            ws.cell(r, c).value = bump_text(v)
        if v == "v2":
            ws.cell(r, c).value = VER
ws["C5"] = TODAY.strftime("%d/%m/%Y")
ws["J5"] = f"KQ1.3 {VER}, KQ2.1 {VER}, KQ2.2 {VER} (chạy lại {TODAY.strftime('%d/%m/%Y')}: HD04, nhận xét module, sửa cờ máy/Q19/IN07)"
# AE thay đổi if column 31
# find header 'Thay đổi'
hdr = [ws.cell(9, c).value for c in range(1, 40)]
print("KQ3 row9", hdr)
for c, h in enumerate(hdr, 1):
    if h and "Thay" in str(h):
        for r in range(10, (ws.max_row or 10) + 1):
            if ws.cell(r, 1).value:
                old = ws.cell(r, c).value or ""
                note_ae = f"v3 {TODAY.strftime('%d/%m/%Y')}: chạy lại form vòng 2b, giữ HD04, mở rộng nhận xét KQ2."
                if note_ae not in str(old):
                    ws.cell(r, c).value = (str(old) + " | " if old else "") + note_ae
                    ws.cell(r, c).font = FONT
        break

ws28 = wb["Danh_gia_28_tieu_chi_rui_ro"]
ws28["A2"] = f"Áp dụng cho {TEN} (MST {MST}) - Kỳ {KY} | {VER} | {TODAY.strftime('%d/%m/%Y')}"
ws28["A2"].font = FONT

for ws in wb.worksheets:
    fix_dates_font(ws)
    autowidth(ws, cap=36)
if "E" in wb["Danh_gia_28_tieu_chi_rui_ro"].column_dimensions:
    wb["Danh_gia_28_tieu_chi_rui_ro"].column_dimensions["E"].width = 60
wb.save(lp(v3["kq3"]))
wb.close()
print("KQ3 v3 saved")

# ---------------------------------------------------------------------------
# Verify
# ---------------------------------------------------------------------------
print("\n===== VERIFY v3 =====")
for key, p in v3.items():
    wb = load_workbook(lp(p), data_only=False)
    print(p.name, "sheets", wb.sheetnames, "size", p.stat().st_size)
    if key == "kq21":
        ws = wb["KQ2_NCC_mua_vao"]
        print("  HD03 D5", ws["D5"].value, "HD04 H5", ws["H5"].value, "F5", ws["F5"].value)
        print("  M7", ws.cell(7, 13).value, "N7", ws.cell(7, 14).value)
        years = []
        ws = wb["KQ2_Doi_chieu_GTGT"]
        months = []
        for rr in range(1, ws.max_row + 1):
            a = ws.cell(rr, 1).value
            if isinstance(a, str) and a.startswith("Năm "):
                years.append(a)
            if isinstance(a, str) and a.upper().startswith("T") and a[1:].isdigit():
                months.append(a)
        print("  GTGT years", years, "month leftovers", months)
        ws = wb["IN07_NCC_CanhBao"]
        nnt = [ws.cell(r, 1).value for r in range(4, 8) if str(ws.cell(r, 1).value or "") == MST]
        print("  IN07 NNT rows", nnt)
        print("  Tong_hop G12", wb["KQ2_Tong_hop_co"]["G12"].value)
        print("  comment NCC last", wb["KQ2_NCC_mua_vao"].cell(wb["KQ2_NCC_mua_vao"].max_row, 1).value)
    if key == "kq22":
        print("  511 G6", str(wb["KQ2_511_632"]["G6"].value)[:120])
        print("  Tong_hop G8", wb["KQ2_Tong_hop_co"]["G8"].value, "G9", wb["KQ2_Tong_hop_co"]["G9"].value)
        print("  kho comment", wb["KQ2_Kho_Gia_thanh"].cell(wb["KQ2_Kho_Gia_thanh"].max_row - 2, 1).value)
        print("  gia thanh comment", wb["KQ2_Gia_thanh_va_Dinh_muc"].cell(wb["KQ2_Gia_thanh_va_Dinh_muc"].max_row - 2, 1).value)
        print("  congno comment", wb["KQ2_Cong_no_Tien_mat"].cell(wb["KQ2_Cong_no_Tien_mat"].max_row - 2, 1).value)
    if key == "kq3":
        print("  C4", wb["KQ3"]["C4"].value, "28TC", wb["Danh_gia_28_tieu_chi_rui_ro"].max_row)
    wb.close()
print("DONE v3")
