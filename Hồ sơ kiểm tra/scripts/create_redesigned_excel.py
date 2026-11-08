# -*- coding: utf-8 -*-
"""Tao lai / do du lieu vao Danh sach kiem tra quyet toan pro.xlsm theo cau truc moi.

Thay doi chinh:
- Bo sheet LICH_SU_2024, LICH_SU_2025, DU_LIEU_IN_AN
- So theo doi chia 2 trang thai / 2 sheet: 01_DA_BAN_HANH, 02_DANG_THUC_HIEN
- Bo cot: Han cong bo QD, Han ky BBKT, Han nop tien, Tinh trang nop, toan bo theo doi nop NSNN
- Cot Thong bao qua han khi qua han ma chua thuc hien
- Font Times New Roman
"""

import re
from datetime import datetime, date, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

SOURCE_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\Danh sach kiem tra quyet toan.xlsx"
TARGET_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm"

FONT_NAME = "Times New Roman"
STATUS_BAN_HANH = "1. Đã ban hành"
STATUS_DANG_THUC_HIEN = "2. Đang thực hiện"
MAX_DATA_ROWS = 200


def parse_vn_date(val):
    """Chuyen chuoi ngay VN thanh date. Tra ve None neu khong parse duoc."""
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    text = str(val).strip()
    if not text:
        return None
    text = text.replace(".", "/").replace("-", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def to_number(val):
    if val is None or val == "":
        return 0
    if isinstance(val, (int, float)):
        return float(val)
    text = str(val).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return 0


def workday_add(start, days):
    """Cong so ngay lam viec (bo Sat/Sun). Don gian, khong tinh nghi le."""
    if not start or days is None:
        return None
    cur = start
    added = 0
    step = 1 if days >= 0 else -1
    target = abs(int(days))
    while added < target:
        cur = cur + timedelta(days=step)
        if cur.weekday() < 5:
            added += 1
    return cur


def classify_status(row):
    """Da ban hanh = co So QDKT; Dang thuc hien = chua co QDKT."""
    qdkt = str(row.get("QĐKT") or "").strip()
    if qdkt:
        return STATUS_BAN_HANH
    return STATUS_DANG_THUC_HIEN


def build_overdue_message(row, today=None):
    """Sinh thong bao qua han khi chua thuc hien dung han."""
    today = today or date(2026, 9, 30)
    messages = []

    ngay_qd = parse_vn_date(row.get("ngày QĐ"))
    ngay_cb = parse_vn_date(row.get("Ngày công bố"))
    den_ngay = parse_vn_date(row.get("Đến ngày"))
    ngay_bbkt = parse_vn_date(row.get("Ngày ký Biên bản kiểm tra"))
    qdxphc = str(row.get("QĐXPHC") or "").strip()
    ngay_qdxphc = parse_vn_date(row.get("ngày QĐ3"))

    # Cong bo QD: cham nhat 10 ngay lam viec tu ngay ban hanh QD
    if ngay_qd and not ngay_cb:
        han_cb = workday_add(ngay_qd, 10)
        if han_cb and today > han_cb:
            messages.append("QUÁ HẠN CÔNG BỐ QĐ")
        elif han_cb and today >= workday_add(ngay_qd, 7):
            messages.append("CẬN HẠN CÔNG BỐ QĐ")

    # Ky BBKT: trong 05 ngay lam viec ke tu ngay ket thuc KT
    if den_ngay and not ngay_bbkt:
        han_bb = workday_add(den_ngay, 5)
        if han_bb and today > han_bb:
            messages.append("QUÁ HẠN KÝ BBKT")
        elif han_bb and today >= workday_add(den_ngay, 3):
            messages.append("CẬN HẠN KÝ BBKT")

    # Ban hanh QDXPHC: 07 ngay lam viec ke tu ngay ky BBKT
    if ngay_bbkt and not qdxphc and not ngay_qdxphc:
        han_xp = workday_add(ngay_bbkt, 7)
        if han_xp and today > han_xp:
            messages.append("QUÁ HẠN BAN HÀNH QĐXPHC")
        elif han_xp and today >= workday_add(ngay_bbkt, 5):
            messages.append("CẬN HẠN BAN HÀNH QĐXPHC")

    # Thoi han KT tai tru so NNT: khong qua 20 ngay lam viec
    thoi_han = row.get("Thời hạn kiểm tra ")
    try:
        thoi_han_num = int(str(thoi_han).strip()) if thoi_han not in (None, "") else None
    except ValueError:
        thoi_han_num = None
    if thoi_han_num is not None and thoi_han_num > 20:
        messages.append("VƯỢT QUÁ 20 NGÀY KT")

    if not messages:
        if qdxphc or ngay_qdxphc:
            return "Đã hoàn thành các mốc chính"
        if ngay_qd or ngay_cb or den_ngay or ngay_bbkt:
            return "Trong hạn"
        return "Chưa phát sinh tiến trình"

    return " | ".join(messages)


def load_source_rows():
    print("Loading source workbook...")
    wb_src = openpyxl.load_workbook(SOURCE_PATH, data_only=True)
    ws = wb_src["2026"]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    rows = []
    for r in range(2, ws.max_row + 1):
        mst = ws.cell(r, 3).value
        ten = ws.cell(r, 4).value
        if not (mst or ten):
            continue
        data = {}
        for c in range(1, ws.max_column + 1):
            h = headers[c - 1]
            if h:
                data[h] = ws.cell(r, c).value
        rows.append(data)

    print(f"Found {len(rows)} enterprise records for 2026.")
    return rows


def style_pack():
    f_title = Font(name=FONT_NAME, size=16, bold=True, color="1F4E78")
    f_subtitle = Font(name=FONT_NAME, size=10, italic=True, color="595959")
    f_col_header = Font(name=FONT_NAME, size=9, bold=True, color="1F497D")
    f_bold = Font(name=FONT_NAME, size=10, bold=True)
    f_regular = Font(name=FONT_NAME, size=10)
    f_italic = Font(name=FONT_NAME, size=10, italic=True)
    f_white_bold = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
    f_kpi = Font(name=FONT_NAME, size=16, bold=True, color="1F497D")

    fills = {
        "navy": PatternFill("solid", fgColor="1F4E78"),
        "slate": PatternFill("solid", fgColor="2F5597"),
        "amber": PatternFill("solid", fgColor="C65911"),
        "green": PatternFill("solid", fgColor="375623"),
        "darkred": PatternFill("solid", fgColor="C00000"),
        "gray": PatternFill("solid", fgColor="595959"),
        "card": PatternFill("solid", fgColor="F2F4F8"),
        "sub_blue": PatternFill("solid", fgColor="D9E1F2"),
        "sub_amber": PatternFill("solid", fgColor="FCE4D6"),
        "sub_green": PatternFill("solid", fgColor="E2EFDA"),
        "sub_red": PatternFill("solid", fgColor="F8CBAD"),
        "sub_gray": PatternFill("solid", fgColor="F2F2F2"),
        "alert": PatternFill("solid", fgColor="FFC7CE"),
        "ok": PatternFill("solid", fgColor="C6EFCE"),
    }

    thin = Side(border_style="thin", color="B4B4B4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    return {
        "f_title": f_title,
        "f_subtitle": f_subtitle,
        "f_col_header": f_col_header,
        "f_bold": f_bold,
        "f_regular": f_regular,
        "f_italic": f_italic,
        "f_white_bold": f_white_bold,
        "f_kpi": f_kpi,
        "fills": fills,
        "border": border,
    }


def col_definitions(fills):
    """Danh sach cot moi: da bo han cong bo, han ky BB, han nop, tinh trang nop, toan bo NSNN."""
    return [
        # A. Doanh nghiep
        (1, "STT", 6, fills["sub_blue"]),
        (2, "Trạng thái", 18, fills["sub_blue"]),
        (3, "Loại hình", 10, fills["sub_blue"]),
        (4, "Mã số thuế (MST)", 16, fills["sub_blue"]),
        (5, "Tên Người nộp thuế", 38, fills["sub_blue"]),
        (6, "Địa chỉ trụ sở", 36, fills["sub_blue"]),
        (7, "Địa bàn Phường/Xã", 28, fills["sub_blue"]),
        (8, "Mã CQT", 12, fills["sub_blue"]),
        # B. Doan KT
        (9, "Trưởng đoàn (TĐ)", 20, fills["sub_blue"]),
        (10, "Chức vụ TĐ", 14, fills["sub_blue"]),
        (11, "Thành viên (TV)", 20, fills["sub_blue"]),
        (12, "Chức vụ TV", 14, fills["sub_blue"]),
        (13, "Nội dung / Phân công", 30, fills["sub_blue"]),
        # C. Tien trinh (khong con cot han)
        (14, "Số QĐKT", 16, fills["sub_amber"]),
        (15, "Số QĐGS", 16, fills["sub_amber"]),
        (16, "Ngày ban hành QĐ", 14, fills["sub_amber"]),
        (17, "Ngày công bố thực tế", 14, fills["sub_amber"]),
        (18, "Thời hạn KT (ngày)", 12, fills["sub_amber"]),
        (19, "Ngày bắt đầu", 13, fills["sub_amber"]),
        (20, "Ngày kết thúc", 13, fills["sub_amber"]),
        (21, "Ngày ký BBKT", 13, fills["sub_amber"]),
        (22, "Ngày báo cáo", 13, fills["sub_amber"]),
        # D. Xu ly VPHC (khong han nop / tinh trang nop)
        (23, "Số QĐXPHC", 16, fills["sub_red"]),
        (24, "Ngày ban hành QĐXPHC", 14, fills["sub_red"]),
        # E. Ket qua tai chinh
        (25, "Truy thu Thuế TNDN", 15, fills["sub_green"]),
        (26, "Truy thu Thuế GTGT", 15, fills["sub_green"]),
        (27, "Tổng cộng truy thu", 15, fills["sub_green"]),
        (28, "Tiền chậm nộp GTGT (0.03%)", 16, fills["sub_green"]),
        (29, "Tiền chậm nộp TNDN (0.03%)", 16, fills["sub_green"]),
        (30, "Tổng tiền chậm nộp", 15, fills["sub_green"]),
        (31, "Phạt khai sai 20%", 14, fills["sub_green"]),
        (32, "Phạt thủ tục VPHC", 14, fills["sub_green"]),
        (33, "Tổng tiền phạt", 14, fills["sub_green"]),
        (34, "Tổng cộng phát sinh", 16, fills["sub_green"]),
        (35, "Giảm lỗ", 12, fills["sub_green"]),
        (36, "Giảm khấu trừ", 13, fills["sub_green"]),
        # F. Canh bao
        (37, "Thông báo quá hạn", 42, fills["sub_gray"]),
    ]


def write_track_sheet(wb, title, rows, styles, status_value):
    fills = styles["fills"]
    border = styles["border"]
    ws = wb.create_sheet(title=title)
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "A4"

    # Tieu de
    ws.merge_cells("A1:AK1")
    ws["A1"] = f"SỔ THEO DÕI KIỂM TRA THUẾ 2026 - {status_value.upper()}"
    ws["A1"].font = styles["f_title"]
    ws["A1"].alignment = Alignment(vertical="center")

    ws.merge_cells("A2:AK2")
    ws["A2"] = (
        "QĐ 1161/QĐ-CT & Luật Quản lý thuế số 108/2025/QH15 | "
        "Font Times New Roman | Cảnh báo tự động khi quá hạn mà chưa thực hiện"
    )
    ws["A2"].font = styles["f_subtitle"]

    # Nhom cot hang 3
    groups = [
        (1, 8, "THÔNG TIN DOANH NGHIỆP & PHÁP LÝ", fills["navy"]),
        (9, 13, "ĐOÀN KIỂM TRA & PHÂN CÔNG", fills["slate"]),
        (14, 22, "TIẾN TRÌNH KIỂM TRA", fills["amber"]),
        (23, 24, "BAN HÀNH QĐXPHC", fills["darkred"]),
        (25, 36, "KẾT QUẢ TÀI CHÍNH & NGHĨA VỤ THUẾ", fills["green"]),
        (37, 37, "CẢNH BÁO", fills["gray"]),
    ]
    for c1, c2, label, fill in groups:
        ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c2)
        cell = ws.cell(3, c1, label)
        cell.font = styles["f_white_bold"]
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(c1, c2 + 1):
            ws.cell(3, c).border = border
            ws.cell(3, c).fill = fill

    # Header cot hang 4
    col_defs = col_definitions(fills)
    for c_idx, name, width, fill in col_defs:
        cell = ws.cell(4, c_idx, name)
        cell.font = styles["f_col_header"]
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(c_idx)].width = width
    ws.row_dimensions[4].height = 36

    # Dropdown trang thai
    dv = DataValidation(
        type="list",
        formula1=f'"{STATUS_BAN_HANH},{STATUS_DANG_THUC_HIEN}"',
        allow_blank=True,
    )
    dv.error = "Chỉ chọn 1 trong 2 trạng thái"
    dv.errorTitle = "Trạng thái không hợp lệ"
    ws.add_data_validation(dv)
    dv.add(f"B5:B{4 + MAX_DATA_ROWS}")

    money_cols = {25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36}
    center_cols = {1, 2, 3, 4, 8, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24}
    date_cols = {16, 17, 19, 20, 21, 22, 24}

    for idx, row in enumerate(rows, start=1):
        r = 4 + idx
        ngay_qd = parse_vn_date(row.get("ngày QĐ"))
        ngay_cb = parse_vn_date(row.get("Ngày công bố"))
        den_ngay = parse_vn_date(row.get("Đến ngày"))
        ngay_bbkt = parse_vn_date(row.get("Ngày ký Biên bản kiểm tra"))
        ngay_bc = parse_vn_date(row.get("Ngày báo cáo"))
        ngay_qdxphc = parse_vn_date(row.get("ngày QĐ3"))
        thoi_han = row.get("Thời hạn kiểm tra ")
        try:
            thoi_han = int(str(thoi_han).strip()) if thoi_han not in (None, "") else ""
        except ValueError:
            thoi_han = str(thoi_han or "")

        tndn = to_number(row.get("Thuế TNDN"))
        gtgt = to_number(row.get("Thuế GTGT"))
        phat_gtgt = to_number(row.get("Số tiền phạt 0,03% GTGT"))
        phat_tndn = to_number(row.get("Số tiền phạt 0,03% TNDN"))
        phat_20 = to_number(row.get("Số tiền phạt 20%"))
        phat_vphc = to_number(row.get("Phạt KK sai VPHC"))
        giam_lo = to_number(row.get("Giảm lỗ"))
        giam_kt = to_number(row.get("giảm khấu trừ"))

        values = {
            1: idx,
            2: status_value,
            3: str(row.get("TH") or "KH").strip(),
            4: str(row.get("MST") or "").strip(),
            5: str(row.get("Tên công ty") or "").strip(),
            6: str(row.get("địa chỉ") or "").strip(),
            7: str(row.get("CQT") or "").strip(),
            8: f'=IFERROR(VLOOKUP(G{r},DANH_MUC!$A$3:$B$15,2,FALSE),"")',
            9: str(row.get("TĐ") or "").strip(),
            10: str(row.get("CVTĐ") or "").strip(),
            11: str(row.get("TV") or "").strip(),
            12: str(row.get("CVTV") or "").strip(),
            13: str(row.get("Nội dung kiểm tra") or "").strip(),
            14: str(row.get("QĐKT") or "").strip(),
            15: str(row.get("QĐGS") or "").strip(),
            16: ngay_qd,
            17: ngay_cb,
            18: thoi_han,
            19: ngay_cb,  # bat dau = ngay cong bo
            20: den_ngay,
            21: ngay_bbkt,
            22: ngay_bc or ngay_bbkt,
            23: str(row.get("QĐXPHC") or "").strip(),
            24: ngay_qdxphc,
            25: tndn,
            26: gtgt,
            27: f"=Y{r}+Z{r}",
            28: phat_gtgt,
            29: phat_tndn,
            30: f"=AB{r}+AC{r}",
            31: phat_20,
            32: phat_vphc,
            33: f"=AE{r}+AF{r}",
            34: f"=AA{r}+AD{r}+AG{r}",
            35: giam_lo,
            36: giam_kt,
            37: build_overdue_message(row),
        }

        for c in range(1, 38):
            cell = ws.cell(r, c, values.get(c, ""))
            cell.font = styles["f_regular"]
            cell.border = border
            if c in center_cols:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif c in money_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0"
            elif c == 37:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            if c in date_cols and isinstance(values.get(c), date):
                cell.number_format = "DD/MM/YYYY"

            # To mau canh bao
            msg = str(values.get(37) or "")
            if c == 37:
                if "QUÁ HẠN" in msg:
                    cell.fill = fills["alert"]
                    cell.font = Font(name=FONT_NAME, size=10, bold=True, color="9C0006")
                elif "CẬN HẠN" in msg:
                    cell.fill = PatternFill("solid", fgColor="FFEB9C")
                    cell.font = Font(name=FONT_NAME, size=10, bold=True, color="9C5700")
                elif "Đã hoàn thành" in msg:
                    cell.fill = fills["ok"]
                    cell.font = Font(name=FONT_NAME, size=10, color="006100")

    # Dong trong de nhap them
    start_empty = 5 + len(rows)
    for r in range(start_empty, start_empty + 20):
        for c in range(1, 38):
            cell = ws.cell(r, c, "")
            cell.font = styles["f_regular"]
            cell.border = border
        ws.cell(r, 1, f'=IF(D{r}<>"",ROW()-4,"")')
        ws.cell(r, 2, status_value)
        ws.cell(r, 8, f'=IFERROR(VLOOKUP(G{r},DANH_MUC!$A$3:$B$15,2,FALSE),"")')
        ws.cell(r, 27, f'=IF(OR(Y{r}<>"",Z{r}<>""),Y{r}+Z{r},"")')
        ws.cell(r, 30, f'=IF(OR(AB{r}<>"",AC{r}<>""),AB{r}+AC{r},"")')
        ws.cell(r, 33, f'=IF(OR(AE{r}<>"",AF{r}<>""),AE{r}+AF{r},"")')
        ws.cell(r, 34, f'=IF(OR(AA{r}<>"",AD{r}<>"",AG{r}<>""),IF(AA{r}="",0,AA{r})+IF(AD{r}="",0,AD{r})+IF(AG{r}="",0,AG{r}),"")')
        # Cong thuc canh bao dong trong (dung TODAY)
        ws.cell(
            r,
            37,
            (
                f'=IF(D{r}="","",'
                f'IF(AND(P{r}<>"",Q{r}="",TODAY()>WORKDAY(P{r},10)),"QUÁ HẠN CÔNG BỐ QĐ",'
                f'IF(AND(T{r}<>"",U{r}="",TODAY()>WORKDAY(T{r},5)),"QUÁ HẠN KÝ BBKT",'
                f'IF(AND(U{r}<>"",W{r}="",TODAY()>WORKDAY(U{r},7)),"QUÁ HẠN BAN HÀNH QĐXPHC",'
                f'IF(AND(R{r}<>"",R{r}>20),"VƯỢT QUÁ 20 NGÀY KT","Trong hạn")))))'
            ),
        )
        for c in money_cols:
            ws.cell(r, c).number_format = "#,##0"
        for c in date_cols:
            ws.cell(r, c).number_format = "DD/MM/YYYY"

    # Conditional formatting cho cot thong bao
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    red_font = Font(name=FONT_NAME, color="9C0006", bold=True)
    ws.conditional_formatting.add(
        f"AK5:AK{4 + MAX_DATA_ROWS}",
        FormulaRule(formula=['ISNUMBER(SEARCH("QUÁ HẠN",AK5))'], fill=red_fill, font=red_font),
    )

    ws.auto_filter.ref = f"A4:AK{4 + len(rows)}"
    ws.row_dimensions[1].height = 24
    return ws


def write_dashboard(wb, styles, n_ban_hanh, n_dang_th, n_overdue, total_money):
    fills = styles["fills"]
    border = styles["border"]
    ws = wb.create_sheet(title="00_DASHBOARD", index=0)
    ws.sheet_view.showGridLines = True

    ws.merge_cells("B2:J2")
    ws["B2"] = "BẢNG ĐIỀU HÀNH KIỂM TRA THUẾ NĂM 2026"
    ws["B2"].font = styles["f_title"]

    ws.merge_cells("B3:J3")
    ws["B3"] = (
        "Theo dõi 2 trạng thái (Đã ban hành / Đang thực hiện) | "
        "QĐ 1161/QĐ-CT & Luật Quản lý thuế số 108/2025/QH15 | Cảnh báo quá hạn tự động"
    )
    ws["B3"].font = styles["f_subtitle"]

    cards = [
        ("TỔNG SỐ HỒ SƠ", f"={n_ban_hanh + n_dang_th}", "B5", "C6", fills["navy"]),
        ("ĐÃ BAN HÀNH", f"={n_ban_hanh}", "D5", "E6", fills["green"]),
        ("ĐANG THỰC HIỆN", f"={n_dang_th}", "F5", "G6", fills["amber"]),
        ("CẢNH BÁO QUÁ HẠN", f"={n_overdue}", "H5", "I6", fills["darkred"]),
    ]
    for title, formula, top_left, bot_right, fill in cards:
        c1, r1 = top_left[0], int(top_left[1:])
        c2, r2 = bot_right[0], int(bot_right[1:])
        ws.merge_cells(f"{c1}{r1}:{c2}{r1}")
        ws[f"{c1}{r1}"] = title
        ws[f"{c1}{r1}"].font = styles["f_white_bold"]
        ws[f"{c1}{r1}"].fill = fill
        ws[f"{c1}{r1}"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(f"{c1}{r2}:{c2}{r2}")
        ws[f"{c1}{r2}"] = formula
        ws[f"{c1}{r2}"].font = styles["f_kpi"]
        ws[f"{c1}{r2}"].fill = fills["card"]
        ws[f"{c1}{r2}"].alignment = Alignment(horizontal="center", vertical="center")
        for row in ws[f"{c1}{r1}:{c2}{r2}"]:
            for cell in row:
                cell.border = border

    ws.merge_cells("B8:D8")
    ws["B8"] = "TỔNG KIẾN NGHỊ XỬ LÝ (VNĐ)"
    ws["B8"].font = styles["f_white_bold"]
    ws["B8"].fill = fills["slate"]
    ws["B8"].alignment = Alignment(horizontal="center")
    ws.merge_cells("B9:D9")
    ws["B9"] = total_money
    ws["B9"].font = styles["f_kpi"]
    ws["B9"].number_format = "#,##0"
    ws["B9"].fill = fills["card"]
    ws["B9"].alignment = Alignment(horizontal="center")
    for row in ws["B8:D9"]:
        for cell in row:
            cell.border = border

    # Bang 2 trang thai
    ws["B11"] = "1. PHÂN LOẠI THEO 2 TRẠNG THÁI"
    ws["B11"].font = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
    headers = ["Trạng thái", "Số hồ sơ", "Tỷ lệ (%)", "Ghi chú"]
    for i, h in enumerate(headers):
        cell = ws.cell(12, 2 + i, h)
        cell.font = styles["f_white_bold"]
        cell.fill = fills["slate"]
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    status_rows = [
        (STATUS_BAN_HANH, n_ban_hanh, "Có Số QĐKT - quyết định kiểm tra đã ban hành"),
        (STATUS_DANG_THUC_HIEN, n_dang_th, "Chưa có Số QĐKT - đang chuẩn bị / thực hiện"),
    ]
    total = max(n_ban_hanh + n_dang_th, 1)
    for i, (name, count, note) in enumerate(status_rows):
        r = 13 + i
        ws.cell(r, 2, name).font = styles["f_regular"]
        ws.cell(r, 3, count).font = styles["f_bold"]
        ws.cell(r, 3).alignment = Alignment(horizontal="center")
        ws.cell(r, 4, count / total).number_format = "0.0%"
        ws.cell(r, 5, note).font = styles["f_italic"]
        for c in range(2, 6):
            ws.cell(r, c).border = border

    # SLA
    ws["G11"] = "2. HẠN ĐỊNH NGHIỆP VỤ (SLA) - CẢNH BÁO KHI QUÁ HẠN"
    ws["G11"].font = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
    sla_headers = ["Hạng mục", "Căn cứ", "Thời hạn chuẩn", "Khi quá hạn chưa làm"]
    for i, h in enumerate(sla_headers):
        cell = ws.cell(12, 7 + i, h)
        cell.font = styles["f_white_bold"]
        cell.fill = fills["amber"]
        cell.border = border
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    sla_data = [
        ("Công bố QĐ kiểm tra", "Điều 22 Luật 108 & QĐ 1161", "Chậm nhất 10 ngày làm việc từ ngày ký QĐ", "Hiện thông báo QUÁ HẠN CÔNG BỐ QĐ"),
        ("Thời hạn kiểm tra tại NNT", "Điều 22 Luật 108 & QĐ 1161", "Không quá 20 ngày làm việc (gia hạn 01 lần ≤ 20 ngày)", "Hiện thông báo VƯỢT QUÁ 20 NGÀY KT"),
        ("Ký Biên bản kiểm tra", "Quy trình 1161/QĐ-CT", "Trong 05 ngày làm việc từ ngày công khai dự thảo", "Hiện thông báo QUÁ HẠN KÝ BBKT"),
        ("Ban hành QĐXPHC", "Luật XLVPHC, NĐ 125 & QĐ 1161", "07 ngày làm việc (tối đa 30-60 ngày vụ phức tạp)", "Hiện thông báo QUÁ HẠN BAN HÀNH QĐXPHC"),
    ]
    for i, row in enumerate(sla_data):
        r = 13 + i
        for j, val in enumerate(row):
            cell = ws.cell(r, 7 + j, val)
            cell.font = styles["f_bold"] if j == 0 else styles["f_regular"]
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Huong dan
    ws["B16"] = "3. CẤU TRÚC FILE HIỆN TẠI"
    ws["B16"].font = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
    notes = [
        "Sheet 00_DASHBOARD: KPI điều hành và bảng hạn định pháp lý.",
        "Sheet 01_DA_BAN_HANH: sổ theo dõi trạng thái Đã ban hành (có Số QĐKT).",
        "Sheet 02_DANG_THUC_HIEN: sổ theo dõi trạng thái Đang thực hiện (chưa có Số QĐKT).",
        "Sheet DANH_MUC: danh mục Mã CQT và tiểu mục kho bạc.",
        "Đã bỏ: sheet lịch sử 2024/2025, sheet dữ liệu in ấn.",
        "Đã bỏ cột: Hạn công bố QĐ, Hạn ký BBKT, Hạn nộp tiền, Tình trạng nộp, toàn bộ theo dõi nộp NSNN.",
        "Cột Thông báo quá hạn tự hiện khi quá hạn mà chưa thực hiện. Font toàn bộ: Times New Roman.",
    ]
    for i, note in enumerate(notes):
        ws.cell(17 + i, 2, f"- {note}").font = styles["f_regular"]
        ws.merge_cells(start_row=17 + i, start_column=2, end_row=17 + i, end_column=10)

    widths = {"A": 3, "B": 28, "C": 12, "D": 12, "E": 40, "F": 4, "G": 26, "H": 28, "I": 34, "J": 36}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in range(13, 17):
        ws.row_dimensions[r].height = 32


def write_danh_muc(wb, styles):
    fills = styles["fills"]
    border = styles["border"]
    ws = wb.create_sheet(title="DANH_MUC")

    ws["A1"] = "DANH MỤC CƠ QUAN THUẾ ĐỊA BÀN PHƯỜNG / XÃ"
    ws["A1"].font = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
    ws["A2"] = "Tên cơ quan thu"
    ws["B2"] = "Mã CQT"
    for c in ("A2", "B2"):
        ws[c].fill = fills["slate"]
        ws[c].font = styles["f_white_bold"]
        ws[c].border = border

    cqt_data = [
        ("Xã Trị An - Thuế cơ sở 3 Thành phố Đồng Nai", "1140972"),
        ("Xã Tân An - Thuế cơ sở 3 Thành phố Đồng Nai", "1140973"),
        ("Phường Tân Triều - Thuế cơ sở 3 Thành phố Đồng Nai", "1140974"),
        ("Phường Biên Hòa - Thuế cơ sở 3 Thành phố Đồng Nai", "1140962"),
        ("Phường Trấn Biên - Thuế cơ sở 3 Thành phố Đồng Nai", "1140963"),
        ("Phường Tam Hiệp - Thuế cơ sở 3 Thành phố Đồng Nai", "1140964"),
        ("Phường Long Bình - Thuế cơ sở 3 Thành phố Đồng Nai", "1140965"),
        ("Phường Trảng Dài - Thuế cơ sở 3 Thành phố Đồng Nai", "1140966"),
        ("Phường Hố Nai - Thuế cơ sở 3 Thành phố Đồng Nai", "1140967"),
        ("Phường Long Hưng - Thuế cơ sở 3 Thành phố Đồng Nai", "1140968"),
        ("Phường Phước Tân - Thuế cơ sở 3 Thành phố Đồng Nai", "1140969"),
        ("Phường Tam Phước - Thuế cơ sở 3 Thành phố Đồng Nai", "1140970"),
    ]
    for i, (name, code) in enumerate(cqt_data, start=3):
        ws.cell(i, 1, name).font = styles["f_regular"]
        ws.cell(i, 2, code).font = styles["f_bold"]
        ws.cell(i, 2).alignment = Alignment(horizontal="center")
        ws.cell(i, 1).border = border
        ws.cell(i, 2).border = border

    ws["D1"] = "DANH MỤC TIỂU MỤC KHO BẠC"
    ws["D1"].font = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
    for col, val in (("D2", "Mã tiểu mục"), ("E2", "Nội dung kinh tế"), ("F2", "Ghi chú")):
        ws[col] = val
        ws[col].fill = fills["amber"]
        ws[col].font = styles["f_white_bold"]
        ws[col].border = border

    tm_data = [
        ("1052", "Thuế thu nhập doanh nghiệp (TNDN)", "Truy thu kiểm tra"),
        ("1701", "Thuế giá trị gia tăng (GTGT)", "Truy thu kiểm tra"),
        ("4254", "Tiền phạt VPHC về thuế, hóa đơn", "QĐXPHC"),
        ("4944", "Tiền chậm nộp thuế TNDN", "0.03%/ngày"),
        ("4931", "Tiền chậm nộp thuế GTGT", "0.03%/ngày"),
        ("4299", "Tiền phạt VPHC khác", "Phạt thủ tục"),
    ]
    for i, (code, name, note) in enumerate(tm_data, start=3):
        ws.cell(i, 4, code).font = styles["f_bold"]
        ws.cell(i, 5, name).font = styles["f_regular"]
        ws.cell(i, 6, note).font = styles["f_italic"]
        for c in range(4, 7):
            ws.cell(i, c).border = border

    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 20


def create_redesigned_workbook():
    rows = load_source_rows()
    styles = style_pack()

    ban_hanh = [r for r in rows if classify_status(r) == STATUS_BAN_HANH]
    dang_th = [r for r in rows if classify_status(r) == STATUS_DANG_THUC_HIEN]

    n_overdue = 0
    total_money = 0.0
    for r in rows:
        msg = build_overdue_message(r)
        if "QUÁ HẠN" in msg:
            n_overdue += 1
        total_money += (
            to_number(r.get("Thuế TNDN"))
            + to_number(r.get("Thuế GTGT"))
            + to_number(r.get("Số tiền phạt 0,03% GTGT"))
            + to_number(r.get("Số tiền phạt 0,03% TNDN"))
            + to_number(r.get("Số tiền phạt 20%"))
            + to_number(r.get("Phạt KK sai VPHC"))
        )

    wb = openpyxl.Workbook()
    default = wb.active
    wb.remove(default)

    write_dashboard(wb, styles, len(ban_hanh), len(dang_th), n_overdue, total_money)
    write_track_sheet(wb, "01_DA_BAN_HANH", ban_hanh, styles, STATUS_BAN_HANH)
    write_track_sheet(wb, "02_DANG_THUC_HIEN", dang_th, styles, STATUS_DANG_THUC_HIEN)
    write_danh_muc(wb, styles)

    # Dam bao khong con sheet lich su / in an
    forbidden = {"DU_LIEU_IN_AN", "LICH_SU_2024", "LICH_SU_2025", "SO_THEO_DOI_2026"}
    for name in list(wb.sheetnames):
        if name in forbidden:
            del wb[name]

    ws_dash = wb["00_DASHBOARD"]
    ws_dash["B24"] = (
        "4. TỰ CHUYỂN DÒNG KHI ĐỔI TRẠNG THÁI: tại sheet 02_DANG_THUC_HIEN, "
        'chọn "1. Đã ban hành" sẽ tự chuyển hồ sơ sang 01_DA_BAN_HANH '
        "(chiều ngược lại cũng vậy). Dùng file .xlsm và bật Enable Content / macro."
    )
    ws_dash["B24"].font = Font(name=FONT_NAME, size=10)

    print(f"Saving redesigned workbook to: {TARGET_PATH}")
    wb.save(TARGET_PATH)
    print("Done!")
    print(f"Sheets: {wb.sheetnames}")
    print(f"Đã ban hành: {len(ban_hanh)} | Đang thực hiện: {len(dang_th)} | Quá hạn: {n_overdue}")

    try:
        from inject_status_move_vba import inject
        inject()
    except Exception as exc:
        print(f"WARNING: chua gan duoc macro tu chuyen dong: {exc}")
        print("Chay rieng: python scripts/inject_status_move_vba.py")

    return TARGET_PATH


if __name__ == "__main__":
    create_redesigned_workbook()
