# -*- coding: utf-8 -*-
"""
Script tong hop HDDT va Bang ke 4 sheet cho CONG TY TNHH THANH THANH PHUONG
MST: 3600840461
Ky: 2020-2025

Cap nhat bo sung 2 cot:
1. Tình trạng hoạt động (lay tu cot 45/42 cua Danh ba NNT: Trạng thái ĐKT tổ chức/hoạt động)
2. Cơ quan thuế quản lý (lay tu cot 2 cua Danh ba NNT: Tên cơ quan thuế)
"""
import os
import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import xml.etree.ElementTree as ET

MST = "3600840461"
KY = "2020-2025"
TEN_CTY = "CÔNG TY TNHH THANH THANH PHƯƠNG"

DIR_HDDT = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\1. DU LIEU THO NGANH THUE\1.2. Hoa don dien tu & Bang ke\HDDT.xlsx")
DIR_BK = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\2. HO SO NNT CUNG CAP\2.5. Bang ke mua vao ban ra\Bang_ke_mua_vao_ban_ra_tong_hop_2020-2025.xlsx")
DIR_DANHBA = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\3. DU LIEU BO SUNG & TRA CUU QLT\3.2. Danh ba trang thai NNT\Danh ba NNT.xls")

OUT_VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\cau-thanh\Tong_hop_HDDT_BangKe_3600840461_2020-2025.xlsx")
OUT_G = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\4. KET QUA PHAN TICH\cau-thanh\Tong_hop_HDDT_BangKe_3600840461_2020-2025.xlsx")

def chuan_hoa_mst(val):
    if val is None:
        return ""
    s = str(val).strip().replace(" ", "")
    if not s or s.lower() == "none":
        return ""
    if len(s) == 10 and s.isdigit():
        return s
    if len(s) == 9 and s.isdigit():
        return "0" + s
    if len(s) == 12 and s.isdigit():
        return "0" + s[:9] + "-" + s[9:]
    if len(s) == 13 and s.isdigit():
        return s[:10] + "-" + s[10:]
    if len(s) == 14 and "-" in s:
        return s
    return s

def sanitize_val(val):
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip()
        if s.startswith("="):
            return "'" + s
        return s
    return val

def load_danhba_map(p_danhba: Path):
    if not p_danhba.exists():
        print(f"Warning: Danh ba file not found at {p_danhba}")
        return {}
    
    print(f"Loading Danh ba NNT from {p_danhba}...")
    tree = ET.parse(p_danhba)
    root = tree.getroot()

    def ss_index(cell, default):
        for key, val in cell.attrib.items():
            if key == 'Index' or key.endswith('}Index'):
                return int(val)
        return default

    def ss_text(cell):
        for child in list(cell):
            if child.tag.endswith('}Data') or child.tag == 'Data':
                return (child.text or '').strip()
        return ''

    danhba = {}
    for row in root.iter():
        if row.tag.endswith('}Row') or row.tag == 'Row':
            cells = {}
            idx = 1
            for cell in list(row):
                if not (cell.tag.endswith('}Cell') or cell.tag == 'Cell'):
                    continue
                idx = ss_index(cell, idx)
                cells[idx] = ss_text(cell)
                idx += 1
            mst = cells.get(4, '').strip()
            if mst and mst != 'Mã số thuế':
                cqt = cells.get(2, '').strip()
                # Col 45 là Trạng thái ĐKT tổ chức (cột thứ 42 tính từ cột Mã số thuế Col 4: 45 - 4 + 1 = 42)
                tt = cells.get(45, '').strip() or cells.get(44, '').strip() or cells.get(53, '').strip()
                
                clean_m = mst.replace('-', '').strip()
                rec = {'tt': tt, 'cqt': cqt}
                danhba[mst] = rec
                danhba[clean_m] = rec
                if len(clean_m) >= 10:
                    danhba[clean_m[:10]] = rec
    print(f"Loaded {len(danhba)} entries from Danh ba NNT.")
    return danhba

def get_danhba_info(danhba_map, mst_str):
    if not mst_str:
        return "", ""
    mst_s = str(mst_str).strip()
    clean_m = mst_s.replace('-', '').strip()
    
    hit = danhba_map.get(mst_s) or danhba_map.get(clean_m)
    if not hit and len(clean_m) >= 10:
        hit = danhba_map.get(clean_m[:10])
    
    if hit:
        return hit.get('tt', ''), hit.get('cqt', '')
    return "", ""

def build_4_sheet():
    danhba_map = load_danhba_map(DIR_DANHBA)

    print(f"Reading HDDT from {DIR_HDDT}...")
    wb_hddt = openpyxl.load_workbook(DIR_HDDT, data_only=True)
    
    print(f"Reading Bang ke from {DIR_BK}...")
    wb_bk = openpyxl.load_workbook(DIR_BK, data_only=True)

    out_wb = openpyxl.Workbook()
    default_sheet = out_wb.active

    font_regular = Font(name="Times New Roman", size=10)
    font_bold = Font(name="Times New Roman", size=10, bold=True)
    font_header = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
    font_title = Font(name="Times New Roman", size=12, bold=True, color="1F497D")

    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_total = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    total_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000")
    )

    num_format_money = "#,##0"

    # ==========================================
    # SHEET 1: HDDT_Mua_vao
    # ==========================================
    print("Processing HDDT_Mua_vao...")
    ws1 = out_wb.create_sheet(title="HDDT_Mua_vao")
    src_ws1 = wb_hddt["Mua vao"]
    rows_h_mv = list(src_ws1.iter_rows(values_only=True))
    header_h_mv = list(rows_h_mv[0])
    # Cột MST người bán ở index 5
    # Thêm cột MST_chuan vào sau cột 5 (vị trí 6)
    # Thêm 2 cột cuối: Tình trạng hoạt động, Cơ quan thuế quản lý
    new_header_h_mv = header_h_mv[:6] + ["MST_chuan"] + header_h_mv[6:] + ["Tình trạng hoạt động", "Cơ quan thuế quản lý"]

    ws1.cell(row=1, column=1, value=f"TỔNG HỢP HÓA ĐƠN ĐIỆN TỬ MUA VÀO - {TEN_CTY} (MST: {MST}) - KỲ {KY}").font = font_title
    for col_idx, h_text in enumerate(new_header_h_mv, start=1):
        c = ws1.cell(row=3, column=col_idx, value=h_text)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    curr_row = 4
    money_cols_ws1 = [15, 16, 17, 18, 19]

    stt_count = 1
    for r in rows_h_mv[1:]:
        tt = str(r[20] or "")
        kq = str(r[21] or "")
        if "bị xóa bỏ" in tt or "bị hủy" in tt or "bị thay thế" in tt:
            continue
        
        mst_goc = r[5]
        mst_c = chuan_hoa_mst(mst_goc)
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)

        new_row = list(r[:6]) + [mst_c] + list(r[6:]) + [tt_hd, cqt_ql]
        new_row[0] = stt_count
        stt_count += 1

        for col_idx, val in enumerate(new_row, start=1):
            cell = ws1.cell(row=curr_row, column=col_idx, value=sanitize_val(val))
            cell.font = font_regular
            cell.border = thin_border
            if col_idx in money_cols_ws1:
                cell.number_format = num_format_money
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 5, 6, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [24, 25]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    # Dòng tổng cộng
    total_row_ws1 = curr_row
    ws1.cell(row=total_row_ws1, column=1, value="TỔNG CỘNG").font = font_bold
    ws1.cell(row=total_row_ws1, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_h_mv) + 1):
        cell = ws1.cell(row=total_row_ws1, column=col_idx)
        cell.fill = fill_total
        cell.border = total_border
        if col_idx in money_cols_ws1:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws1-1})"
            cell.font = font_bold
            cell.number_format = num_format_money
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ==========================================
    # SHEET 2: HDDT_Ban_ra
    # ==========================================
    print("Processing HDDT_Ban_ra...")
    ws2 = out_wb.create_sheet(title="HDDT_Ban_ra")
    src_ws2 = wb_hddt["Ban ra"]
    rows_h_br = list(src_ws2.iter_rows(values_only=True))
    header_h_br = list(rows_h_br[0])
    # Cột MST người mua ở index 7
    new_header_h_br = header_h_br[:8] + ["MST_chuan"] + header_h_br[8:] + ["Tình trạng hoạt động", "Cơ quan thuế quản lý"]

    ws2.cell(row=1, column=1, value=f"TỔNG HỢP HÓA ĐƠN ĐIỆN TỬ BÁN RA - {TEN_CTY} (MST: {MST}) - KỲ {KY}").font = font_title
    for col_idx, h_text in enumerate(new_header_h_br, start=1):
        c = ws2.cell(row=3, column=col_idx, value=h_text)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    curr_row = 4
    money_cols_ws2 = [15, 16, 17, 18, 19]

    stt_count = 1
    for r in rows_h_br[1:]:
        tt = str(r[20] or "")
        kq = str(r[21] or "")
        if "bị xóa bỏ" in tt or "bị hủy" in tt or "bị thay thế" in tt or "không đủ điều kiện" in kq:
            continue
        
        mst_goc = r[7]
        mst_c = chuan_hoa_mst(mst_goc)
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)

        new_row = list(r[:8]) + [mst_c] + list(r[8:]) + [tt_hd, cqt_ql]
        new_row[0] = stt_count
        stt_count += 1

        for col_idx, val in enumerate(new_row, start=1):
            cell = ws2.cell(row=curr_row, column=col_idx, value=sanitize_val(val))
            cell.font = font_regular
            cell.border = thin_border
            if col_idx in money_cols_ws2:
                cell.number_format = num_format_money
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 5, 6, 8, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [24, 25]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    total_row_ws2 = curr_row
    ws2.cell(row=total_row_ws2, column=1, value="TỔNG CỘNG").font = font_bold
    ws2.cell(row=total_row_ws2, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_h_br) + 1):
        cell = ws2.cell(row=total_row_ws2, column=col_idx)
        cell.fill = fill_total
        cell.border = total_border
        if col_idx in money_cols_ws2:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws2-1})"
            cell.font = font_bold
            cell.number_format = num_format_money
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ==========================================
    # SHEET 3: Bang_ke_Mua_vao
    # ==========================================
    print("Processing Bang_ke_Mua_vao...")
    ws3 = out_wb.create_sheet(title="Bang_ke_Mua_vao")
    src_ws3 = wb_bk["mua vao"]
    rows_bk_mv = list(src_ws3.iter_rows(values_only=True))
    header_bk_mv = list(rows_bk_mv[2]) # Header tại row 3
    # MST_doi_tac ở index 10
    new_header_bk_mv = header_bk_mv[:11] + ["MST_chuan"] + header_bk_mv[11:] + ["Tình trạng hoạt động", "Cơ quan thuế quản lý"]

    ws3.cell(row=1, column=1, value=f"TỔNG HỢP BẢNG KÊ MUA VÀO (NNT NỘP) - {TEN_CTY} (MST: {MST}) - KỲ {KY}").font = font_title
    for col_idx, h_text in enumerate(new_header_bk_mv, start=1):
        c = ws3.cell(row=3, column=col_idx, value=h_text)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    curr_row = 4
    money_cols_ws3 = [14, 16, 17]

    stt_count = 1
    for r in rows_bk_mv[3:]:
        stt_val = str(r[0] or "").strip().upper()
        so_hd_val = str(r[7] or "").strip()
        if stt_val in ["TONG", "TỔNG CỘNG", "TỔNG", "TOTAL"] or "(2)" in so_hd_val:
            continue
        
        mst_goc = r[10]
        mst_c = chuan_hoa_mst(mst_goc)
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)

        new_row = list(r[:11]) + [mst_c] + list(r[11:]) + [tt_hd, cqt_ql]
        new_row[0] = stt_count
        stt_count += 1

        for col_idx, val in enumerate(new_row, start=1):
            cell = ws3.cell(row=curr_row, column=col_idx, value=sanitize_val(val))
            cell.font = font_regular
            cell.border = thin_border
            if col_idx in money_cols_ws3:
                try:
                    if val is not None and str(val).strip() != "":
                        cell.value = float(val)
                except:
                    pass
                cell.number_format = num_format_money
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 3, 6, 8, 9, 11, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [19, 20]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    total_row_ws3 = curr_row
    ws3.cell(row=total_row_ws3, column=1, value="TỔNG CỘNG").font = font_bold
    ws3.cell(row=total_row_ws3, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_bk_mv) + 1):
        cell = ws3.cell(row=total_row_ws3, column=col_idx)
        cell.fill = fill_total
        cell.border = total_border
        if col_idx in [14, 16]:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws3-1})"
            cell.font = font_bold
            cell.number_format = num_format_money
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ==========================================
    # SHEET 4: Bang_ke_Ban_ra
    # ==========================================
    print("Processing Bang_ke_Ban_ra...")
    ws4 = out_wb.create_sheet(title="Bang_ke_Ban_ra")
    src_ws4 = wb_bk["ban ra"]
    rows_bk_br = list(src_ws4.iter_rows(values_only=True))
    header_bk_br = list(rows_bk_br[2])
    # MST_doi_tac ở index 9
    new_header_bk_br = header_bk_br[:10] + ["MST_chuan"] + header_bk_br[10:] + ["Tình trạng hoạt động", "Cơ quan thuế quản lý"]

    ws4.cell(row=1, column=1, value=f"TỔNG HỢP BẢNG KÊ BÁN RA (NNT NỘP) - {TEN_CTY} (MST: {MST}) - KỲ {KY}").font = font_title
    for col_idx, h_text in enumerate(new_header_bk_br, start=1):
        c = ws4.cell(row=3, column=col_idx, value=h_text)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    curr_row = 4
    money_cols_ws4 = [13, 15, 16]

    stt_count = 1
    for r in rows_bk_br[3:]:
        stt_val = str(r[0] or "").strip().upper()
        so_hd_val = str(r[6] or "").strip()
        if stt_val in ["TONG", "TỔNG CỘNG", "TỔNG", "TOTAL"] or "(2)" in so_hd_val:
            continue
        
        mst_goc = r[9]
        mst_c = chuan_hoa_mst(mst_goc)
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)

        new_row = list(r[:10]) + [mst_c] + list(r[10:]) + [tt_hd, cqt_ql]
        new_row[0] = stt_count
        stt_count += 1

        for col_idx, val in enumerate(new_row, start=1):
            cell = ws4.cell(row=curr_row, column=col_idx, value=sanitize_val(val))
            cell.font = font_regular
            cell.border = thin_border
            if col_idx in money_cols_ws4:
                try:
                    if val is not None and str(val).strip() != "":
                        cell.value = float(val)
                except:
                    pass
                cell.number_format = num_format_money
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 3, 5, 7, 8, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [18, 19]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    total_row_ws4 = curr_row
    ws4.cell(row=total_row_ws4, column=1, value="TỔNG CỘNG").font = font_bold
    ws4.cell(row=total_row_ws4, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_bk_br) + 1):
        cell = ws4.cell(row=total_row_ws4, column=col_idx)
        cell.fill = fill_total
        cell.border = total_border
        if col_idx in [13, 15]:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws4-1})"
            cell.font = font_bold
            cell.number_format = num_format_money
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # Auto fit column width
    print("Auto-fitting column widths...")
    for ws in [ws1, ws2, ws3, ws4]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in list(col)[:100]:
                val_str = str(cell.value or "")
                if cell.row == 1:
                    continue
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

    # Xóa default sheet
    out_wb.remove(default_sheet)

    # Save to OUT_VAULT
    OUT_VAULT.parent.mkdir(parents=True, exist_ok=True)
    print(f"Saving to {OUT_VAULT}...")
    out_wb.save(OUT_VAULT)

    # Save to OUT_G
    OUT_G.parent.mkdir(parents=True, exist_ok=True)
    print(f"Saving to {OUT_G}...")
    out_wb.save(OUT_G)
    print("Done successfully!")

if __name__ == "__main__":
    build_4_sheet()
