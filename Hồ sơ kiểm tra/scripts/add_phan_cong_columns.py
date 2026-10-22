# -*- coding: utf-8 -*-
"""Cap nhat bo sung 2 cot: Phan cong truong doan va Phan cong thanh vien vao JAVIS PRO."""

import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OLD_FILE = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan.xlsx"
XLSX_FILE = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan - JAVIS PRO.xlsx"
XLSM_FILE = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan - JAVIS PRO.xlsm"

# 1. Load data tu file cu
wb_old = openpyxl.load_workbook(OLD_FILE, data_only=True)
ws_old = wb_old["2026"]
old_dict = {}
for r in range(2, ws_old.max_row + 1):
    mst = ws_old.cell(r, 3).value
    if mst:
        mst_clean = str(mst).strip()
        old_dict[mst_clean] = {
            "pc_td": ws_old.cell(r, 17).value,
            "pc_tv": ws_old.cell(r, 18).value,
        }

# Special cases da chot trong memory/facts:
# Vạn Phát Thành (MST 3601409441)
old_dict["3601409441"] = {
    "pc_td": "- Chi phí quản lý doanh nghiệp\n- Chi phí bán hàng",
    "pc_tv": "- Chi phí tài chính",
}
# Tâm Hoàng Phát (MST 3603721832)
old_dict["3603721832"] = {
    "pc_td": "- Chi phí quản lý doanh nghiệp\n- Chi phí bán hàng",
    "pc_tv": "- Chi phí tài chính",
}

print(f"Loaded {len(old_dict)} MST assignment mappings.")

def shift_cell_reference(formula_str, offset=2):
    """Dich cac toa do o co cot tu AA tro di them offset cot."""
    pattern = r'(?<![A-Za-z0-9_!])(\$?[A-Z]{1,3})(\$?[0-9]+)(?![A-Za-z0-9_])'
    def repl(m):
        full_col = m.group(1)
        row_str = m.group(2)
        col_clean = full_col.replace('$', '')
        is_abs = full_col.startswith('$')
        try:
            col_idx = openpyxl.utils.column_index_from_string(col_clean)
        except ValueError:
            return m.group(0)
        if col_idx >= 27:  # tu AA tro di
            new_col = ('$' if is_abs else '') + openpyxl.utils.get_column_letter(col_idx + offset)
            return new_col + row_str
        return m.group(0)
    return re.sub(pattern, repl, formula_str)

def process_sheet(ws):
    print(f"Processing sheet: {ws.title}")

    # A. Unmerge merged cells anh huong boi viec chen cot
    ranges_to_remove = []
    for m in list(ws.merged_cells.ranges):
        if m.min_row <= 4:
            ranges_to_remove.append(m)

    for m in ranges_to_remove:
        ws.merged_cells.remove(m)

    # Convert any MergedCell objects in rows 1-3 to regular Cell
    from openpyxl.cell.cell import Cell, MergedCell
    for (r, c), cell in list(ws._cells.items()):
        if r <= 3 and isinstance(cell, MergedCell):
            ws._cells[(r, c)] = Cell(ws, row=r, column=c)

    # B. Dời các cột từ 50 xuống 27 sang phải 2 cột (c -> c + 2)
    max_r = ws.max_row
    for r in range(4, max_r + 1):
        for c in range(50, 26, -1):
            src_cell = ws.cell(r, c)
            dest_cell = ws.cell(r, c + 2)
            val = src_cell.value

            # Neu la cong thuc, dich toa do
            if isinstance(val, str) and val.startswith("="):
                # Rieng cong thuc H o hang 14 bi loi [1]DANH_MUC thi sua luon
                if "[1]DANH_MUC" in val:
                    val = val.replace("[1]DANH_MUC", "DANH_MUC")
                dest_cell.value = shift_cell_reference(val, offset=2)
            else:
                dest_cell.value = val

            # Sao chep format
            dest_cell.font = Font(
                name=src_cell.font.name,
                size=src_cell.font.size,
                bold=src_cell.font.bold,
                italic=src_cell.font.italic,
                underline=src_cell.font.underline,
                color=src_cell.font.color,
            ) if src_cell.font else None

            dest_cell.fill = PatternFill(
                fill_type=src_cell.fill.fill_type,
                start_color=src_cell.fill.start_color,
                end_color=src_cell.fill.end_color,
            ) if src_cell.fill and src_cell.fill.fill_type else PatternFill(fill_type=None)

            dest_cell.border = Border(
                left=Side(style=src_cell.border.left.style, color=src_cell.border.left.color) if src_cell.border and src_cell.border.left else Side(),
                right=Side(style=src_cell.border.right.style, color=src_cell.border.right.color) if src_cell.border and src_cell.border.right else Side(),
                top=Side(style=src_cell.border.top.style, color=src_cell.border.top.color) if src_cell.border and src_cell.border.top else Side(),
                bottom=Side(style=src_cell.border.bottom.style, color=src_cell.border.bottom.color) if src_cell.border and src_cell.border.bottom else Side(),
            ) if src_cell.border else None

            dest_cell.alignment = Alignment(
                horizontal=src_cell.alignment.horizontal,
                vertical=src_cell.alignment.vertical,
                wrap_text=src_cell.alignment.wrap_text,
            ) if src_cell.alignment else None

            dest_cell.number_format = src_cell.number_format

    # C. Thiet lap do rong cot cho cac cot da dich va 2 cot moi
    # Do rong cu tu AA(27) den AX(50) -> AC(29) den AZ(52)
    old_widths = {
        27: 16.0, 28: 13.0, 29: 14.0, 30: 13.0, 31: 12.0, 32: 13.0, 33: 13.0,
        34: 13.0, 35: 13.0, 36: 16.0, 37: 14.0, 38: 15.0, 39: 13.0, 40: 13.0,
        41: 16.0, 42: 13.0, 43: 15.0, 44: 14.0, 45: 13.0, 46: 13.0, 47: 16.0,
        48: 12.0, 49: 13.0, 50: 42.0
    }
    for old_c, w in old_widths.items():
        new_c_let = get_column_letter(old_c + 2)
        ws.column_dimensions[new_c_let].width = w

    # Do rong cho 2 cot moi AA va AB: 35.0
    ws.column_dimensions['AA'].width = 35.0
    ws.column_dimensions['AB'].width = 35.0

    # D. Dien tieu de va format cho 2 cot moi: Col 27 (AA) va Col 28 (AB)
    thin_border = Border(
        left=Side(style="thin", color="FFB4B4B4"),
        right=Side(style="thin", color="FFB4B4B4"),
        top=Side(style="thin", color="FFB4B4B4"),
        bottom=Side(style="thin", color="FFB4B4B4"),
    )
    col_header_font = Font(name="Times New Roman", size=9.0, bold=True, color="FF1F497D")
    sub_blue_fill = PatternFill(fill_type="solid", fgColor="FFD9E1F2")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_font = Font(name="Times New Roman", size=10.0)
    data_align = Alignment(vertical="center", wrap_text=True)

    # Header row 4
    cell_aa4 = ws.cell(4, 27, "Phân công trưởng đoàn")
    cell_aa4.font = col_header_font
    cell_aa4.fill = sub_blue_fill
    cell_aa4.border = thin_border
    cell_aa4.alignment = header_align

    cell_ab4 = ws.cell(4, 28, "Phân công thành viên")
    cell_ab4.font = col_header_font
    cell_ab4.fill = sub_blue_fill
    cell_ab4.border = thin_border
    cell_ab4.alignment = header_align

    # E. Dien du lieu cac hang 5 tro di cho 2 cot moi
    for r in range(5, max_r + 1):
        mst = ws.cell(r, 4).value
        ten = ws.cell(r, 5).value
        pc_td_val = None
        pc_tv_val = None
        if mst or ten:
            mst_clean = str(mst or "").strip()
            item = old_dict.get(mst_clean)
            if item:
                pc_td_val = item.get("pc_td")
                pc_tv_val = item.get("pc_tv")

        c_aa = ws.cell(r, 27, pc_td_val)
        c_aa.font = data_font
        c_aa.border = thin_border
        c_aa.alignment = data_align

        c_ab = ws.cell(r, 28, pc_tv_val)
        c_ab.font = data_font
        c_ab.border = thin_border
        c_ab.alignment = data_align

    # F. Merge lai cac dai o dong 1, 2, 3
    # Hang 1 va 2: A1:AZ1, A2:AZ2
    ws.merge_cells("A1:AZ1")
    ws.merge_cells("A2:AZ2")

    # Hang 3:
    # Nhom Thong tin DN: A3:H3
    ws.merge_cells("A3:H3")
    # Nhom Thong tin phap ly: I3:U3
    ws.merge_cells("I3:U3")

    # Nhom Doan KT: V3:AB3
    ws.merge_cells("V3:AB3")
    amber_fill = PatternFill(fill_type="solid", fgColor="FFC65911")
    slate_fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
    group_font = Font(name="Times New Roman", size=9.0, bold=True, color="FFFFFFFF")
    center_align = Alignment(horizontal="center", vertical="center")

    cell_v3 = ws.cell(3, 22)
    cell_v3.value = "ĐOÀN KIỂM TRA & PHÂN CÔNG"
    cell_v3.font = group_font
    cell_v3.fill = amber_fill
    cell_v3.alignment = center_align
    for c in range(22, 29):
        ws.cell(3, c).border = thin_border
        ws.cell(3, c).fill = amber_fill

    # Nhom Tien trinh: AC3:AK3 (truoc day la AA3:AI3)
    ws.merge_cells("AC3:AK3")
    cell_ac3 = ws.cell(3, 29)
    cell_ac3.value = "TIẾN TRÌNH KIỂM TRA"
    cell_ac3.font = group_font
    cell_ac3.fill = amber_fill
    cell_ac3.alignment = center_align
    for c in range(29, 38):
        ws.cell(3, c).border = thin_border
        ws.cell(3, c).fill = amber_fill

    # Nhom Ban hanh QDXPHC: AL3:AM3 (truoc day la AJ3:AK3)
    ws.merge_cells("AL3:AM3")
    darkred_fill = PatternFill(fill_type="solid", fgColor="FFC00000")
    cell_al3 = ws.cell(3, 38)
    cell_al3.value = "BAN HÀNH QĐXPHC"
    cell_al3.font = group_font
    cell_al3.fill = darkred_fill
    cell_al3.alignment = center_align
    for c in range(38, 40):
        ws.cell(3, c).border = thin_border
        ws.cell(3, c).fill = darkred_fill

    # Nhom Ket qua tai chinh: AN3:AY3 (truoc day la AL3:AW3)
    ws.merge_cells("AN3:AY3")
    green_fill = PatternFill(fill_type="solid", fgColor="FF375623")
    cell_an3 = ws.cell(3, 40)
    cell_an3.value = "KẾT QUẢ TÀI CHÍNH & NGHĨA VỤ THUẾ"
    cell_an3.font = group_font
    cell_an3.fill = green_fill
    cell_an3.alignment = center_align
    for c in range(40, 52):
        ws.cell(3, c).border = thin_border
        ws.cell(3, c).fill = green_fill

    # O Canh bao: AZ3 (truoc day la AX3)
    gray_fill = PatternFill(fill_type="solid", fgColor="FF595959")
    ws.merge_cells("AZ3:AZ3")
    cell_az3 = ws.cell(3, 52)
    cell_az3.value = "CẢNH BÁO"
    cell_az3.font = group_font
    cell_az3.fill = gray_fill
    cell_az3.alignment = center_align
    cell_az3.border = thin_border

    # G. Cap nhat Conditional Formatting (AZ5:AZ<max_r>)
    from openpyxl.formatting.rule import FormulaRule
    ws.conditional_formatting._cf_rules.clear()
    red_fill = PatternFill(fill_type="solid", start_color="FFFFC7CE", end_color="FFFFC7CE")
    red_font = Font(name="Times New Roman", size=10, bold=True, color="FF9C0006")
    cf_rule = FormulaRule(formula=['ISNUMBER(SEARCH("QUÁ HẠN",AZ5))'], fill=red_fill, font=red_font)
    ws.conditional_formatting.add(f"AZ5:AZ{max_r}", cf_rule)

    # Sua loi [1]DANH_MUC o bat ky hang nao con sot
    for r in range(5, max_r + 1):
        h_val = str(ws.cell(r, 8).value or "")
        if "[1]DANH_MUC" in h_val:
            ws.cell(r, 8).value = h_val.replace("[1]DANH_MUC", "DANH_MUC")

def update_dashboard(ws_dash):
    # Cap nhat mo ta Dashboard
    ws_dash["B12"] = "- Nhóm đoàn KT & nội dung: Trưởng đoàn, CVTĐ, Thành viên, CVTV, Nội dung kiểm tra,"
    ws_dash["B13"] = "- Đã bổ sung 2 cột: Phân công trưởng đoàn, Phân công thành viên (theo mẫu file cũ)."
    print("Updated Dashboard notes.")

# Thuc hien cap nhat tren ca 2 file
for path, is_xlsm in [(XLSX_FILE, False), (XLSM_FILE, True)]:
    print(f"\n==========================================")
    print(f"Opening workbook: {path}")
    wb = openpyxl.load_workbook(path, keep_vba=is_xlsm)
    process_sheet(wb["01_DA_BAN_HANH"])
    process_sheet(wb["02_DANG_THUC_HIEN"])
    update_dashboard(wb["00_DASHBOARD"])
    print(f"Saving workbook: {path}...")
    wb.save(path)
    print(f"Successfully updated: {path}")

print("\nALL FILES UPDATED SUCCESSFULLY!")
