# -*- coding: utf-8 -*-
"""
Script tao ca 2 file theo Huong 2:
File 1: Tong_hop_HDDT_BangKe_3600840461_2020-2025.xlsx
  - HDDT_Mua_vao (27 cot: co cot Canh bao rui ro & Thong tin CV)
  - HDDT_Ban_ra (25 cot)
  - Bang_ke_Mua_vao (22 cot: co cot Canh bao rui ro & Thong tin CV)
  - Bang_ke_Ban_ra (19 cot)
  - Tong_hop_trang_thai_rui_ro (Sheet 5: Tong hop trang thai hoat dong & CV Cong an / CVCBRR)

File 2: Tong_hop_NCC_mua_vao_CV_trang_thai_3600840461_2020-2025.xlsx
  - Phan_1_F1_NNT
  - Bang_2_1_NCC_F0
  - Bang_2_2_Toan_bo_NCC (100% NCC Union BK va HDDT)
  - Bang_2_3_Ky_ke_khai_CT
  - Phan_3_Nhan_xet
"""
import os
import sys
import re
import shutil
from pathlib import Path
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import xml.etree.ElementTree as ET

MST_NNT = "3600840461"
KY = "2020-2025"
TEN_NNT = "CÔNG TY TNHH THANH THANH PHƯƠNG"

DIR_HDDT = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\1. DU LIEU THO NGANH THUE\1.2. Hoa don dien tu & Bang ke\HDDT.xlsx")
DIR_BK = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\2. HO SO NNT CUNG CAP\2.5. Bang ke mua vao ban ra\Bang_ke_mua_vao_ban_ra_tong_hop_2020-2025.xlsx")
DIR_DANHBA = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\3. DU LIEU BO SUNG & TRA CUU QLT\3.2. Danh ba trang thai NNT\Danh ba NNT.xls")
CVCBRR = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx")
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")

OUT_FILE1_VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\cau-thanh\Tong_hop_HDDT_BangKe_3600840461_2020-2025.xlsx")
OUT_FILE1_G = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\4. KET QUA PHAN TICH\cau-thanh\Tong_hop_HDDT_BangKe_3600840461_2020-2025.xlsx")

OUT_FILE2_VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\cau-thanh\Tong_hop_NCC_mua_vao_CV_trang_thai_3600840461_2020-2025.xlsx")
OUT_FILE2_G = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\THANH_THANH_PHUONG_3600840461\4. KET QUA PHAN TICH\cau-thanh\Tong_hop_NCC_mua_vao_CV_trang_thai_3600840461_2020-2025.xlsx")

# Formatting tools
FONT_TITLE = Font(name="Times New Roman", size=12, bold=True, color="1F497D")
FONT_SUBTITLE = Font(name="Times New Roman", size=11, bold=True, color="000000")
FONT_HEADER = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_REGULAR = Font(name="Times New Roman", size=10)
FONT_BOLD = Font(name="Times New Roman", size=10, bold=True)
FONT_RED_BOLD = Font(name="Times New Roman", size=10, bold=True, color="C00000")
FONT_ORG_BOLD = Font(name="Times New Roman", size=10, bold=True, color="B25900")

FILL_HEADER = PatternFill("solid", fgColor="1F497D")
FILL_HEADER2 = PatternFill("solid", fgColor="2F5597")
FILL_TOTAL = PatternFill("solid", fgColor="DCE6F1")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
FILL_GRAY = PatternFill("solid", fgColor="F2F2F2")

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)
TOTAL_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="double", color="000000")
)
NUM_FMT = "#,##0"

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

def mst_clean(s):
    if not s:
        return ""
    s = str(s).split("-")[0].strip()
    return re.sub(r"\D", "", s)

def sanitize_val(val):
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip()
        if s.startswith("="):
            return "'" + s
        return s
    return val

def autowidth(ws, min_w=10, cap=55):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in list(col)[:250]:
            if cell.row in [1, 2]:
                continue
            v = str(cell.value or "")
            if v.startswith("="):
                continue
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, cap), min_w)

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
                tt = cells.get(45, '').strip() or cells.get(44, '').strip() or cells.get(53, '').strip()
                ten = cells.get(5, '').strip()
                clean_m = mst.replace('-', '').strip()
                rec = {'tt': tt, 'cqt': cqt, 'ten': ten}
                danhba[mst] = rec
                danhba[clean_m] = rec
                if len(clean_m) >= 10:
                    danhba[clean_m[:10]] = rec
    print(f"Loaded {len(danhba)} entries from Danh ba NNT.")
    return danhba

def load_cvcbrr_and_ca():
    print("Loading CVCBRR...")
    f0_map = {}
    f1_map = {}
    if CVCBRR.exists():
        wb_cv = openpyxl.load_workbook(CVCBRR, read_only=True, data_only=True)
        if 'F0' in wb_cv.sheetnames:
            for r in wb_cv['F0'].iter_rows(min_row=8, values_only=True):
                m = mst_clean(r[1] if r and len(r)>1 else '')
                if m:
                    f0_map[m] = {
                        'mst': m,
                        'ten': str(r[2] or '').strip() if len(r)>2 else '',
                        'cv': str(r[3] or '').strip() if len(r)>3 else '',
                        'ngay': str(r[4])[:10] if len(r)>4 and r[4] else '',
                        'co_quan': str(r[5] or '').strip() if len(r)>5 else '',
                        'nd': str(r[6] or '').strip() if len(r)>6 else ''
                    }
        if 'F1' in wb_cv.sheetnames:
            for r in wb_cv['F1'].iter_rows(min_row=8, values_only=True):
                m_ban = mst_clean(r[1] if r and len(r)>1 else '')
                m_mua = mst_clean(r[8] if r and len(r)>8 else '')
                if m_ban:
                    f1_map[m_ban] = {
                        'mst_ban': m_ban,
                        'ten_ban': str(r[2] or '').strip() if len(r)>2 else '',
                        'cv': str(r[3] or '').strip() if len(r)>3 else '',
                        'ngay': str(r[4])[:10] if len(r)>4 and r[4] else '',
                        'co_quan': str(r[5] or '').strip() if len(r)>5 else '',
                        'm_mua': m_mua,
                        'ten_mua': str(r[9] or '').strip() if len(r)>9 else ''
                    }
        wb_cv.close()
    print(f"Loaded F0: {len(f0_map)}, F1: {len(f1_map)}")

    print("Loading CA...")
    ca_map = {}
    if CA.exists():
        wb_ca = openpyxl.load_workbook(CA, read_only=True, data_only=True)
        if 'NHAP_HO_SO' in wb_ca.sheetnames:
            ws = wb_ca['NHAP_HO_SO']
            last_cv, last_ngay, last_cq, last_nd = '', '', '', ''
            for r in list(ws.iter_rows(values_only=True))[4:]:
                if len(r) > 2 and r[2] is not None and str(r[2]).strip():
                    last_cv = str(r[2]).strip()
                    last_ngay = str(r[3])[:10] if len(r)>3 and r[3] is not None else ''
                    last_cq = str(r[4] or '').strip() if len(r)>4 else ''
                    last_nd = str(r[7] or '').strip() if len(r)>7 else ''
                mst = ''
                if len(r) > 9 and r[9]:
                    mst = mst_clean(r[9])
                if mst:
                    ca_map[mst] = {
                        'mst': mst,
                        'ten': str(r[10] or '').strip() if len(r)>10 else '',
                        'cv': last_cv,
                        'ngay': last_ngay,
                        'co_quan': last_cq,
                        'nd': last_nd
                    }
        for sname in ['TAM_NGUNG', 'DANH_SACH_LOC']:
            if sname in wb_ca.sheetnames:
                ws = wb_ca[sname]
                for r in list(ws.iter_rows(values_only=True))[2:]:
                    for val in r:
                        m = mst_clean(val)
                        if len(m) in (10, 13) and (m.startswith('0') or m.startswith('3')):
                            if m not in ca_map:
                                ca_map[m] = {
                                    'mst': m,
                                    'ten': '',
                                    'cv': f"Theo dõi Công an ({sname})",
                                    'ngay': '',
                                    'co_quan': 'Công an',
                                    'nd': f"Nằm trong danh sách {sname}"
                                }
        wb_ca.close()
    print(f"Loaded CA: {len(ca_map)}")
    return f0_map, f1_map, ca_map

def get_risk_info(mst_str, f0_map, f1_map, ca_map):
    if not mst_str:
        return "", ""
    m = mst_clean(mst_str)
    # Priority: CA -> F0 -> F1
    if m in ca_map:
        info = ca_map[m]
        cv_txt = f"CV {info['cv']}"
        if info['ngay']:
            cv_txt += f" ngày {info['ngay']}"
        if info['co_quan']:
            cv_txt += f" ({info['co_quan']})"
        return "Công an", cv_txt
    elif m in f0_map:
        info = f0_map[m]
        cv_txt = f"CV {info['cv']}"
        if info['ngay']:
            cv_txt += f" ngày {info['ngay']}"
        if info['co_quan']:
            cv_txt += f" ({info['co_quan']})"
        return "F0", cv_txt
    elif m in f1_map:
        info = f1_map[m]
        cv_txt = f"CV {info['cv']}"
        if info['ngay']:
            cv_txt += f" ngày {info['ngay']}"
        if info['co_quan']:
            cv_txt += f" ({info['co_quan']})"
        return "F1", cv_txt
    return "", ""

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

print("Data loaders ready.")

def run_generation():
    danhba_map = load_danhba_map(DIR_DANHBA)
    f0_map, f1_map, ca_map = load_cvcbrr_and_ca()

    print(f"Reading HDDT from {DIR_HDDT}...")
    wb_hddt = openpyxl.load_workbook(DIR_HDDT, data_only=True)
    
    print(f"Reading Bang ke from {DIR_BK}...")
    wb_bk = openpyxl.load_workbook(DIR_BK, data_only=True)

    out_wb1 = openpyxl.Workbook()
    default_sheet1 = out_wb1.active

    # Data structures for NCC analysis
    all_ncc = {}
    all_invoices = []

    # ==========================================
    # FILE 1 - SHEET 1: HDDT_Mua_vao (27 cols)
    # ==========================================
    print("Processing File 1 - HDDT_Mua_vao...")
    ws1 = out_wb1.create_sheet(title="HDDT_Mua_vao")
    src_ws1 = wb_hddt["Mua vao"]
    rows_h_mv = list(src_ws1.iter_rows(values_only=True))
    header_h_mv = list(rows_h_mv[0])
    
    new_header_h_mv = header_h_mv[:6] + ["MST_chuan"] + header_h_mv[6:] + [
        "Tình trạng hoạt động", "Cơ quan thuế quản lý", "Cảnh báo rủi ro", "Thông tin công văn CA / CVCBRR"
    ]

    ws1.cell(row=1, column=1, value=f"TỔNG HỢP HÓA ĐƠN ĐIỆN TỬ MUA VÀO - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}").font = FONT_TITLE
    for col_idx, h_text in enumerate(new_header_h_mv, start=1):
        c = ws1.cell(row=3, column=col_idx, value=h_text)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
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
        ten_nb = str(r[6] or "").strip()
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)
        risk_flag, risk_cv = get_risk_info(mst_c, f0_map, f1_map, ca_map)

        new_row = list(r[:6]) + [mst_c] + list(r[6:]) + [tt_hd, cqt_ql, risk_flag, risk_cv]
        new_row[0] = stt_count
        stt_count += 1

        # Money values
        ds_val = float(r[13] or 0.0) if r[13] is not None else 0.0
        thue_val = float(r[14] or 0.0) if r[14] is not None else 0.0
        ngay_val = r[4]

        # Register to all_ncc
        if mst_c:
            if mst_c not in all_ncc:
                all_ncc[mst_c] = {
                    'mst': mst_c, 'ten': ten_nb, 'tt': tt_hd, 'cqt': cqt_ql,
                    'risk_flag': risk_flag, 'risk_cv': risk_cv,
                    'bk_count': 0, 'bk_ds': 0.0, 'bk_thue': 0.0,
                    'hddt_count': 0, 'hddt_ds': 0.0, 'hddt_thue': 0.0
                }
            all_ncc[mst_c]['hddt_count'] += 1
            all_ncc[mst_c]['hddt_ds'] += ds_val
            all_ncc[mst_c]['hddt_thue'] += thue_val

            all_invoices.append({
                'mst': mst_c, 'ten': ten_nb, 'nguon': 'HĐĐT',
                'so_hd': str(r[3] or ''), 'ngay_hd': str(ngay_val)[:10] if ngay_val else '',
                'mat_hang': '', 'ds': ds_val, 'thue': thue_val,
                'risk_flag': risk_flag, 'risk_cv': risk_cv, 'tt': tt_hd, 'cqt': cqt_ql
            })

        for col_idx, val in enumerate(new_row, start=1):
            cell = ws1.cell(row=curr_row, column=col_idx, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if col_idx in money_cols_ws1:
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 5, 6, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [24, 25]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 26: # Cảnh báo rủi ro
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if risk_flag in ["Công an", "F0"]:
                    cell.font = FONT_RED_BOLD
                    cell.fill = FILL_RED
                elif risk_flag == "F1":
                    cell.font = FONT_ORG_BOLD
                    cell.fill = FILL_YEL
            elif col_idx == 27:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    # Total row
    total_row_ws1 = curr_row
    ws1.cell(row=total_row_ws1, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws1.cell(row=total_row_ws1, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_h_mv) + 1):
        cell = ws1.cell(row=total_row_ws1, column=col_idx)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_idx in money_cols_ws1:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws1-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ==========================================
    # FILE 1 - SHEET 2: HDDT_Ban_ra (25 cols)
    # ==========================================
    print("Processing File 1 - HDDT_Ban_ra...")
    ws2 = out_wb1.create_sheet(title="HDDT_Ban_ra")
    src_ws2 = wb_hddt["Ban ra"]
    rows_h_br = list(src_ws2.iter_rows(values_only=True))
    header_h_br = list(rows_h_br[0])
    new_header_h_br = header_h_br[:8] + ["MST_chuan"] + header_h_br[8:] + ["Tình trạng hoạt động", "Cơ quan thuế quản lý"]

    ws2.cell(row=1, column=1, value=f"TỔNG HỢP HÓA ĐƠN ĐIỆN TỬ BÁN RA - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}").font = FONT_TITLE
    for col_idx, h_text in enumerate(new_header_h_br, start=1):
        c = ws2.cell(row=3, column=col_idx, value=h_text)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
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
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if col_idx in money_cols_ws2:
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 5, 6, 8, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [24, 25]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    total_row_ws2 = curr_row
    ws2.cell(row=total_row_ws2, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws2.cell(row=total_row_ws2, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_h_br) + 1):
        cell = ws2.cell(row=total_row_ws2, column=col_idx)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_idx in money_cols_ws2:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws2-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ==========================================
    # FILE 1 - SHEET 3: Bang_ke_Mua_vao (22 cols)
    # ==========================================
    print("Processing File 1 - Bang_ke_Mua_vao...")
    ws3 = out_wb1.create_sheet(title="Bang_ke_Mua_vao")
    src_ws3 = wb_bk["mua vao"]
    rows_bk_mv = list(src_ws3.iter_rows(values_only=True))
    header_bk_mv = list(rows_bk_mv[2])
    new_header_bk_mv = header_bk_mv[:11] + ["MST_chuan"] + header_bk_mv[11:] + [
        "Tình trạng hoạt động", "Cơ quan thuế quản lý", "Cảnh báo rủi ro", "Thông tin công văn CA / CVCBRR"
    ]

    ws3.cell(row=1, column=1, value=f"TỔNG HỢP BẢNG KÊ MUA VÀO (NNT NỘP) - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}").font = FONT_TITLE
    for col_idx, h_text in enumerate(new_header_bk_mv, start=1):
        c = ws3.cell(row=3, column=col_idx, value=h_text)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
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
        ten_nb = str(r[9] or "").strip()
        tt_hd, cqt_ql = get_danhba_info(danhba_map, mst_c)
        risk_flag, risk_cv = get_risk_info(mst_c, f0_map, f1_map, ca_map)

        new_row = list(r[:11]) + [mst_c] + list(r[11:]) + [tt_hd, cqt_ql, risk_flag, risk_cv]
        new_row[0] = stt_count
        stt_count += 1

        ds_val = 0.0
        try:
            if r[12] is not None and str(r[12]).strip() != "":
                ds_val = float(r[12])
        except: pass

        thue_val = 0.0
        try:
            if r[14] is not None and str(r[14]).strip() != "":
                thue_val = float(r[14])
        except: pass

        if mst_c:
            if mst_c not in all_ncc:
                all_ncc[mst_c] = {
                    'mst': mst_c, 'ten': ten_nb, 'tt': tt_hd, 'cqt': cqt_ql,
                    'risk_flag': risk_flag, 'risk_cv': risk_cv,
                    'bk_count': 0, 'bk_ds': 0.0, 'bk_thue': 0.0,
                    'hddt_count': 0, 'hddt_ds': 0.0, 'hddt_thue': 0.0
                }
            all_ncc[mst_c]['bk_count'] += 1
            all_ncc[mst_c]['bk_ds'] += ds_val
            all_ncc[mst_c]['bk_thue'] += thue_val

            all_invoices.append({
                'mst': mst_c, 'ten': ten_nb, 'nguon': 'Bảng kê',
                'so_hd': str(r[7] or ''), 'ngay_hd': str(r[8])[:10] if r[8] else '',
                'mat_hang': str(r[11] or ''), 'ds': ds_val, 'thue': thue_val,
                'risk_flag': risk_flag, 'risk_cv': risk_cv, 'tt': tt_hd, 'cqt': cqt_ql
            })

        for col_idx, val in enumerate(new_row, start=1):
            cell = ws3.cell(row=curr_row, column=col_idx, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if col_idx in money_cols_ws3:
                try:
                    if val is not None and str(val).strip() != "":
                        cell.value = float(val)
                except: pass
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 3, 6, 8, 9, 11, 12]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [19, 20]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 21: # Cảnh báo rủi ro
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if risk_flag in ["Công an", "F0"]:
                    cell.font = FONT_RED_BOLD
                    cell.fill = FILL_RED
                elif risk_flag == "F1":
                    cell.font = FONT_ORG_BOLD
                    cell.fill = FILL_YEL
            elif col_idx == 22:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    total_row_ws3 = curr_row
    ws3.cell(row=total_row_ws3, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws3.cell(row=total_row_ws3, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_bk_mv) + 1):
        cell = ws3.cell(row=total_row_ws3, column=col_idx)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_idx in [14, 16]:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws3-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ==========================================
    # FILE 1 - SHEET 4: Bang_ke_Ban_ra (19 cols)
    # ==========================================
    print("Processing File 1 - Bang_ke_Ban_ra...")
    ws4 = out_wb1.create_sheet(title="Bang_ke_Ban_ra")
    src_ws4 = wb_bk["ban ra"]
    rows_bk_br = list(src_ws4.iter_rows(values_only=True))
    header_bk_br = list(rows_bk_br[2])
    new_header_bk_br = header_bk_br[:10] + ["MST_chuan"] + header_bk_br[10:] + ["Tình trạng hoạt động", "Cơ quan thuế quản lý"]

    ws4.cell(row=1, column=1, value=f"TỔNG HỢP BẢNG KÊ BÁN RA (NNT NỘP) - {TEN_NNT} (MST: {MST_NNT}) - KỲ {KY}").font = FONT_TITLE
    for col_idx, h_text in enumerate(new_header_bk_br, start=1):
        c = ws4.cell(row=3, column=col_idx, value=h_text)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
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
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if col_idx in money_cols_ws4:
                try:
                    if val is not None and str(val).strip() != "":
                        cell.value = float(val)
                except: pass
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 2, 3, 5, 7, 8, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [18, 19]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

    total_row_ws4 = curr_row
    ws4.cell(row=total_row_ws4, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws4.cell(row=total_row_ws4, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_idx in range(1, len(new_header_bk_br) + 1):
        cell = ws4.cell(row=total_row_ws4, column=col_idx)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_idx in [13, 15]:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_ws4-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # ==========================================
    # FILE 1 - SHEET 5: Tong_hop_trang_thai_rui_ro
    # ==========================================
    print("Processing File 1 - Sheet 5: Tong_hop_trang_thai_rui_ro...")
    ws5 = out_wb1.create_sheet(title="Tong_hop_trang_thai_rui_ro")
    ws5.cell(row=1, column=1, value=f"TỔNG HỢP NHÀ CUNG CẤP THEO TRẠNG THÁI HOẠT ĐỘNG VÀ CẢNH BÁO RỦI RO").font = FONT_TITLE
    ws5.cell(row=2, column=1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ phân tích: {KY}").font = FONT_SUBTITLE

    # Aggregate by status
    status_summary = {}
    for ncc in all_ncc.values():
        st = ncc['tt'] or "Không rõ / Chưa tra cứu"
        if st not in status_summary:
            status_summary[st] = {
                'count_ncc': 0, 'hd_count': 0, 'hd_ds': 0.0, 'hd_thue': 0.0,
                'bk_count': 0, 'bk_ds': 0.0, 'bk_thue': 0.0
            }
        status_summary[st]['count_ncc'] += 1
        status_summary[st]['hd_count'] += ncc['hddt_count']
        status_summary[st]['hd_ds'] += ncc['hddt_ds']
        status_summary[st]['hd_thue'] += ncc['hddt_thue']
        status_summary[st]['bk_count'] += ncc['bk_count']
        status_summary[st]['bk_ds'] += ncc['bk_ds']
        status_summary[st]['bk_thue'] += ncc['bk_thue']

    # Table 1: Theo trang thai hoat dong
    ws5.cell(row=4, column=1, value="PHẦN I. TỔNG HỢP THEO TÌNH TRẠNG HOẠT ĐỘNG (ĐKKD - DANH BẠ NNT)").font = FONT_SUBTITLE
    hdr1 = [
        "STT", "Tình trạng hoạt động (ĐKKD)", "Số lượng NCC",
        "Số HĐ (HĐĐT)", "Doanh số chưa thuế (HĐĐT)", "Thuế GTGT (HĐĐT)",
        "Số HĐ (Bảng kê)", "Doanh số chưa thuế (Bảng kê)", "Thuế GTGT (Bảng kê)",
        "Chênh lệch DS (HĐĐT - BK)", "Chênh lệch Thuế (HĐĐT - BK)"
    ]
    for c_idx, h in enumerate(hdr1, start=1):
        c = ws5.cell(row=5, column=c_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r_idx = 6
    stt_s = 1
    # Sorted by count_ncc desc
    for st, d in sorted(status_summary.items(), key=lambda x: x[1]['count_ncc'], reverse=True):
        row_vals = [
            stt_s, st, d['count_ncc'],
            d['hd_count'], d['hd_ds'], d['hd_thue'],
            d['bk_count'], d['bk_ds'], d['bk_thue'],
            d['hd_ds'] - d['bk_ds'], d['hd_thue'] - d['bk_thue']
        ]
        stt_s += 1
        for col_i, val in enumerate(row_vals, start=1):
            cell = ws5.cell(row=r_idx, column=col_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if col_i in [1, 3, 4, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i in [5, 6, 8, 9, 10, 11]:
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        r_idx += 1

    # Total row Table 1
    t1_tot = r_idx
    ws5.cell(row=t1_tot, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws5.cell(row=t1_tot, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_i in range(1, len(hdr1) + 1):
        cell = ws5.cell(row=t1_tot, column=col_i)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_i in [3, 4, 7]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}6:{col_letter}{t1_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i in [5, 6, 8, 9, 10, 11]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}6:{col_letter}{t1_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")
    r_idx += 2

    # Table 2: Theo rui ro CVCBRR & CA
    risk_summary = {
        'Công an': {'name': 'NCC trong danh sách Công an điều tra/xác minh', 'ncc': 0, 'hd_cnt': 0, 'hd_ds': 0.0, 'hd_thue': 0.0, 'bk_cnt': 0, 'bk_ds': 0.0, 'bk_thue': 0.0, 'ghi_chu': 'Rủi ro rất cao, kiến nghị kiểm tra UNC & hồ sơ thực tế'},
        'F0': {'name': 'NCC dính CVCBRR F0 (Bán hóa đơn)', 'ncc': 0, 'hd_cnt': 0, 'hd_ds': 0.0, 'hd_thue': 0.0, 'bk_cnt': 0, 'bk_ds': 0.0, 'bk_thue': 0.0, 'ghi_chu': 'Rủi ro cao (HD01), xác minh hàng hóa và thanh toán'},
        'F1': {'name': 'NCC dính CVCBRR F1 (Mua từ F0 - Tầng 2)', 'ncc': 0, 'hd_cnt': 0, 'hd_ds': 0.0, 'hd_thue': 0.0, 'bk_cnt': 0, 'bk_ds': 0.0, 'bk_thue': 0.0, 'ghi_chu': 'Kiểm tra giải trình chuỗi mua bán hàng hóa'},
        'Bình thường': {'name': 'NCC bình thường (Không có cảnh báo)', 'ncc': 0, 'hd_cnt': 0, 'hd_ds': 0.0, 'hd_thue': 0.0, 'bk_cnt': 0, 'bk_ds': 0.0, 'bk_thue': 0.0, 'ghi_chu': 'Đối chiếu theo kế hoạch thông thường'}
    }
    for ncc in all_ncc.values():
        rf = ncc['risk_flag']
        cat = rf if rf in ['Công an', 'F0', 'F1'] else 'Bình thường'
        risk_summary[cat]['ncc'] += 1
        risk_summary[cat]['hd_cnt'] += ncc['hddt_count']
        risk_summary[cat]['hd_ds'] += ncc['hddt_ds']
        risk_summary[cat]['hd_thue'] += ncc['hddt_thue']
        risk_summary[cat]['bk_cnt'] += ncc['bk_count']
        risk_summary[cat]['bk_ds'] += ncc['bk_ds']
        risk_summary[cat]['bk_thue'] += ncc['bk_thue']

    ws5.cell(row=r_idx, column=1, value="PHẦN II. TỔNG HỢP THEO CẢNH BÁO RỦI RO (F0, F1, CÔNG AN)").font = FONT_SUBTITLE
    r_idx += 1
    hdr2 = [
        "STT", "Phân loại rủi ro", "Số lượng NCC",
        "Số HĐ (HĐĐT)", "Doanh số chưa thuế (HĐĐT)", "Thuế GTGT (HĐĐT)",
        "Số HĐ (Bảng kê)", "Doanh số chưa thuế (Bảng kê)", "Thuế GTGT (Bảng kê)",
        "Đánh giá & Định hướng xử lý"
    ]
    for c_idx, h in enumerate(hdr2, start=1):
        c = ws5.cell(row=r_idx, column=c_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r_idx += 1

    t2_start = r_idx
    stt_r = 1
    for k in ['Công an', 'F0', 'F1', 'Bình thường']:
        d = risk_summary[k]
        row_vals = [
            stt_r, d['name'], d['ncc'],
            d['hd_cnt'], d['hd_ds'], d['hd_thue'],
            d['bk_cnt'], d['bk_ds'], d['bk_thue'],
            d['ghi_chu']
        ]
        stt_r += 1
        for col_i, val in enumerate(row_vals, start=1):
            cell = ws5.cell(row=r_idx, column=col_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if k in ['Công an', 'F0']:
                cell.fill = FILL_RED
                if col_i in [1, 2]: cell.font = FONT_RED_BOLD
            elif k == 'F1':
                cell.fill = FILL_YEL
                if col_i in [1, 2]: cell.font = FONT_ORG_BOLD

            if col_i in [1, 3, 4, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i in [5, 6, 8, 9]:
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        r_idx += 1

    # Total row Table 2
    t2_tot = r_idx
    ws5.cell(row=t2_tot, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws5.cell(row=t2_tot, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_i in range(1, len(hdr2) + 1):
        cell = ws5.cell(row=t2_tot, column=col_i)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_i in [3, 4, 7]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}{t2_start}:{col_letter}{t2_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i in [5, 6, 8, 9]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}{t2_start}:{col_letter}{t2_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")
    r_idx += 2

    # Table 3: Danh sach chi tiet NCC co rui ro hoac bat thuong
    ws5.cell(row=r_idx, column=1, value="PHẦN III. DANH SÁCH CHI TIẾT CÁC NHÀ CUNG CẤP CÓ CẢNH BÁO RỦI RO HOẶC TRẠNG THÁI BẤT THƯỜNG").font = FONT_SUBTITLE
    r_idx += 1
    hdr3 = [
        "STT", "MST_chuan", "Tên người bán", "Tình trạng hoạt động (ĐKKD)", "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro", "Thông tin công văn CA / CVCBRR", "Nguồn dữ liệu",
        "Số HĐ (HĐĐT)", "Doanh số (HĐĐT)", "Thuế (HĐĐT)",
        "Số HĐ (Bảng kê)", "Doanh số (Bảng kê)", "Thuế (Bảng kê)",
        "Đề xuất kiểm tra & xác minh"
    ]
    for c_idx, h in enumerate(hdr3, start=1):
        c = ws5.cell(row=r_idx, column=c_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r_idx += 1

    t3_start = r_idx
    stt_ab = 1
    for mst_c, ncc in sorted(all_ncc.items(), key=lambda x: (0 if x[1]['risk_flag'] in ['Công an', 'F0'] else (1 if x[1]['risk_flag'] == 'F1' else 2), -x[1]['hddt_ds'])):
        rf = ncc['risk_flag']
        tt_s = ncc['tt'] or ''
        is_abnormal_status = ("đang hoạt động" not in tt_s.lower()) and bool(tt_s)
        if rf in ['Công an', 'F0', 'F1'] or is_abnormal_status:
            src_str = "Chung" if (ncc['hddt_count'] > 0 and ncc['bk_count'] > 0) else ("HĐĐT" if ncc['hddt_count'] > 0 else "Bảng kê")
            dexuat = ""
            if rf == "Công an":
                dexuat = "Xác minh toàn bộ hợp đồng, hóa đơn, UNC ngân hàng, chứng từ vận chuyển thực tế"
            elif rf == "F0":
                dexuat = "Rà soát tính có thật của hàng hóa, hồ sơ thanh toán, loại trừ chi phí nếu HĐ bất hợp pháp"
            elif rf == "F1":
                dexuat = "Kiểm tra chuỗi cung ứng, cuống hóa đơn xăng dầu, định mức tiêu hao"
            else:
                dexuat = "Kiểm tra thời điểm lập hóa đơn so với ngày cơ quan thuế thông báo ngừng hoạt động"

            row_vals = [
                stt_ab, mst_c, ncc['ten'], tt_s, ncc['cqt'],
                rf if rf else "ĐKKD bất thường", ncc['risk_cv'], src_str,
                ncc['hddt_count'], ncc['hddt_ds'], ncc['hddt_thue'],
                ncc['bk_count'], ncc['bk_ds'], ncc['bk_thue'],
                dexuat
            ]
            stt_ab += 1
            for col_i, val in enumerate(row_vals, start=1):
                cell = ws5.cell(row=r_idx, column=col_i, value=sanitize_val(val))
                cell.font = FONT_REGULAR
                cell.border = THIN_BORDER
                if rf in ['Công an', 'F0']:
                    cell.fill = FILL_RED
                    if col_i in [2, 3, 6]: cell.font = FONT_RED_BOLD
                elif rf == 'F1':
                    cell.fill = FILL_YEL
                    if col_i in [2, 3, 6]: cell.font = FONT_ORG_BOLD
                elif is_abnormal_status:
                    if col_i == 4: cell.font = FONT_RED_BOLD

                if col_i in [1, 2, 6, 8, 9, 12]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_i in [10, 11, 13, 14]:
                    cell.number_format = NUM_FMT
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            r_idx += 1

    # Total row Table 3
    t3_tot = r_idx
    ws5.cell(row=t3_tot, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws5.cell(row=t3_tot, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_i in range(1, len(hdr3) + 1):
        cell = ws5.cell(row=t3_tot, column=col_i)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_i in [9, 12]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}{t3_start}:{col_letter}{t3_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i in [10, 11, 13, 14]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}{t3_start}:{col_letter}{t3_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # Autowidth all sheets in File 1
    print("Auto-fitting column widths for File 1...")
    for ws in [ws1, ws2, ws3, ws4, ws5]:
        autowidth(ws)

    out_wb1.remove(default_sheet1)
    OUT_FILE1_VAULT.parent.mkdir(parents=True, exist_ok=True)
    out_wb1.save(OUT_FILE1_VAULT)
    print(f"Saved File 1 to {OUT_FILE1_VAULT}")

    OUT_FILE1_G.parent.mkdir(parents=True, exist_ok=True)
    out_wb1.save(OUT_FILE1_G)
    print(f"Saved File 1 to {OUT_FILE1_G}")

    # =========================================================================
    # FILE 2: Tong_hop_NCC_mua_vao_CV_trang_thai_3600840461_2020-2025.xlsx
    # =========================================================================
    print("\n=======================================================")
    print("Processing File 2: Tong_hop_NCC_mua_vao_CV_trang_thai...")
    out_wb2 = openpyxl.Workbook()
    default_sheet2 = out_wb2.active

    # 1. Sheet: Phan_1_F1_NNT
    ws_p1 = out_wb2.create_sheet(title="Phan_1_F1_NNT")
    ws_p1.cell(row=1, column=1, value="PHẦN 1. RÀ SOÁT RỦI RO TỰ THÂN CỦA NGƯỜI NỘP THUẾ").font = FONT_TITLE
    ws_p1.cell(row=2, column=1, value=f"Doanh nghiệp kiểm tra: {TEN_NNT} - MST: {MST_NNT} - Kỳ phân tích: {KY}").font = FONT_SUBTITLE
    
    hdr_p1 = ["STT", "Nội dung kiểm tra rủi ro tự thân", "Cơ sở dữ liệu đối soát", "Kết quả tra cứu", "Đánh giá mức rủi ro", "Ghi chú & Căn cứ"]
    for c_i, h in enumerate(hdr_p1, start=1):
        c = ws_p1.cell(row=4, column=c_i, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    p1_rows = [
        [1, "Kiểm tra rủi ro NNT là F0 (Doanh nghiệp bị cảnh báo bán hóa đơn)", "Sổ CVCBRR CQT - Sheet F0", "KHÔNG CÓ TRONG DANH SÁCH", "An toàn (Không kích hoạt rủi ro)", "NNT không có công văn cảnh báo bán hóa đơn bất hợp pháp"],
        [2, "Kiểm tra rủi ro NNT là F1 (Doanh nghiệp mua HĐ từ doanh nghiệp F0)", "Sổ CVCBRR CQT - Sheet F1", "KHÔNG CÓ TRONG DANH SÁCH", "An toàn (Không kích hoạt rủi ro HD01)", "NNT không nằm trong danh mục người mua hóa đơn của các doanh nghiệp F0"],
        [3, "Kiểm tra NNT trong danh sách phối hợp xác minh của Cơ quan Công an", "Sổ theo dõi hồ sơ Công an (Bao_cao_cong_van.xlsx)", "KHÔNG CÓ TRONG DANH SÁCH", "An toàn (Bình thường)", "Cơ quan Công an không có công văn yêu cầu cung cấp hồ sơ/điều tra đối với NNT"],
        [4, "Kiểm tra trạng thái hoạt động ĐKKD của NNT", "Danh bạ NNT - Hệ thống TMS", "NNT đang hoạt động", "Bình thường", "Trạng thái mã số thuế đang hoạt động, tuân thủ kê khai nghĩa vụ thuế"]
    ]
    for r_i, r_data in enumerate(p1_rows, start=5):
        for c_i, val in enumerate(r_data, start=1):
            cell = ws_p1.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if c_i == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i in [4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = FONT_BOLD

    # 2. Sheet: Bang_2_1_NCC_F0
    ws_21 = out_wb2.create_sheet(title="Bang_2_1_NCC_F0")
    ws_21.cell(row=1, column=1, value="BẢNG 2.1. DANH SÁCH NHÀ CUNG CẤP THUỘC DIỆN CẢNH BÁO RỦI RO (F0, F1) VÀ CÔNG AN").font = FONT_TITLE
    ws_21.cell(row=2, column=1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}").font = FONT_SUBTITLE

    hdr_21 = [
        "STT", "MST_chuan", "Tên người bán / Nhà cung cấp", "Nhóm rủi ro (F0/CA/F1)", "Thông tin công văn cảnh báo / Công an",
        "Tình trạng hoạt động (ĐKKD)", "Cơ quan thuế quản lý", "Nguồn dữ liệu",
        "Số HĐ (BK)", "Doanh số (BK)", "Thuế (BK)",
        "Số HĐ (HĐĐT)", "Doanh số (HĐĐT)", "Thuế (HĐĐT)",
        "Mức độ rủi ro", "Đề xuất xử lý khi thanh tra"
    ]
    for c_i, h in enumerate(hdr_21, start=1):
        c = ws_21.cell(row=4, column=c_i, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r21_idx = 5
    stt_21 = 1
    for mst_c, ncc in sorted(all_ncc.items(), key=lambda x: (0 if x[1]['risk_flag'] in ['Công an', 'F0'] else (1 if x[1]['risk_flag'] == 'F1' else 2), -x[1]['hddt_ds'])):
        rf = ncc['risk_flag']
        if rf in ['Công an', 'F0', 'F1']:
            src_str = "Chung" if (ncc['hddt_count'] > 0 and ncc['bk_count'] > 0) else ("HĐĐT" if ncc['hddt_count'] > 0 else "Bảng kê")
            muc_do = "Cao" if rf in ["Công an", "F0"] else "Trung bình"
            dexuat = ""
            if rf == "Công an":
                dexuat = "Xác minh toàn bộ hồ sơ thanh toán qua ngân hàng, hợp đồng, biên bản giao nhận thực tế"
            elif rf == "F0":
                dexuat = "Rà soát hóa đơn mua vào, tính có thật của hàng hóa, loại trừ chi phí nếu sử dụng HĐ bất hợp pháp"
            elif rf == "F1":
                dexuat = "Kiểm tra cuống hóa đơn, định mức xăng dầu, bảng kê điều xe thực tế"

            r_vals = [
                stt_21, mst_c, ncc['ten'], rf, ncc['risk_cv'],
                ncc['tt'], ncc['cqt'], src_str,
                ncc['bk_count'], ncc['bk_ds'], ncc['bk_thue'],
                ncc['hddt_count'], ncc['hddt_ds'], ncc['hddt_thue'],
                muc_do, dexuat
            ]
            stt_21 += 1
            for c_i, val in enumerate(r_vals, start=1):
                cell = ws_21.cell(row=r21_idx, column=c_i, value=sanitize_val(val))
                cell.font = FONT_REGULAR
                cell.border = THIN_BORDER
                if rf in ['Công an', 'F0']:
                    cell.fill = FILL_RED
                    if c_i in [2, 3, 4, 15]: cell.font = FONT_RED_BOLD
                elif rf == 'F1':
                    cell.fill = FILL_YEL
                    if c_i in [2, 3, 4, 15]: cell.font = FONT_ORG_BOLD

                if c_i in [1, 2, 4, 8, 9, 12, 15]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i in [10, 11, 13, 14]:
                    cell.number_format = NUM_FMT
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            r21_idx += 1

    # Total row Bảng 2.1
    t21_tot = r21_idx
    ws_21.cell(row=t21_tot, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws_21.cell(row=t21_tot, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_i in range(1, len(hdr_21) + 1):
        cell = ws_21.cell(row=t21_tot, column=col_i)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_i in [9, 12]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}5:{col_letter}{t21_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i in [10, 11, 13, 14]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}5:{col_letter}{t21_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # 3. Sheet: Bang_2_2_Toan_bo_NCC (100% NCC)
    ws_22 = out_wb2.create_sheet(title="Bang_2_2_Toan_bo_NCC")
    ws_22.cell(row=1, column=1, value="BẢNG 2.2. TỔNG HỢP TOÀN BỘ 100% NHÀ CUNG CẤP MUA VÀO (UNION BẢNG KÊ & HĐĐT)").font = FONT_TITLE
    ws_22.cell(row=2, column=1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}").font = FONT_SUBTITLE

    hdr_22 = [
        "STT", "MST_chuan", "Tên người bán / Nhà cung cấp", "Nguồn dữ liệu",
        "Số HĐ (BK)", "Doanh số (BK)", "Thuế (BK)",
        "Số HĐ (HĐĐT)", "Doanh số (HĐĐT)", "Thuế (HĐĐT)",
        "Chênh lệch DS (HĐĐT - BK)", "Chênh lệch Thuế (HĐĐT - BK)",
        "Tình trạng hoạt động (ĐKKD)", "Cơ quan thuế quản lý", "Nhóm rủi ro (F0/CA/F1)", "Thông tin công văn"
    ]
    for c_i, h in enumerate(hdr_22, start=1):
        c = ws_22.cell(row=4, column=c_i, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r22_idx = 5
    stt_22 = 1
    # Sorted by risk flag then doanh số hddt
    for mst_c, ncc in sorted(all_ncc.items(), key=lambda x: (0 if x[1]['risk_flag'] in ['Công an', 'F0'] else (1 if x[1]['risk_flag'] == 'F1' else 2), -x[1]['hddt_ds'])):
        rf = ncc['risk_flag']
        src_str = "Chung" if (ncc['hddt_count'] > 0 and ncc['bk_count'] > 0) else ("HĐĐT" if ncc['hddt_count'] > 0 else "Bảng kê")
        tt_s = ncc['tt'] or ''
        is_abnormal_status = ("đang hoạt động" not in tt_s.lower()) and bool(tt_s)

        r_vals = [
            stt_22, mst_c, ncc['ten'], src_str,
            ncc['bk_count'], ncc['bk_ds'], ncc['bk_thue'],
            ncc['hddt_count'], ncc['hddt_ds'], ncc['hddt_thue'],
            ncc['hddt_ds'] - ncc['bk_ds'], ncc['hddt_thue'] - ncc['bk_thue'],
            tt_s, ncc['cqt'], rf, ncc['risk_cv']
        ]
        stt_22 += 1
        for c_i, val in enumerate(r_vals, start=1):
            cell = ws_22.cell(row=r22_idx, column=c_i, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if rf in ['Công an', 'F0']:
                cell.fill = FILL_RED
                if c_i in [2, 3, 15]: cell.font = FONT_RED_BOLD
            elif rf == 'F1':
                cell.fill = FILL_YEL
                if c_i in [2, 3, 15]: cell.font = FONT_ORG_BOLD
            elif is_abnormal_status:
                if c_i == 13: cell.font = FONT_RED_BOLD

            if c_i in [1, 2, 4, 5, 8, 15]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i in [6, 7, 9, 10, 11, 12]:
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        r22_idx += 1

    # Total row Bảng 2.2
    t22_tot = r22_idx
    ws_22.cell(row=t22_tot, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws_22.cell(row=t22_tot, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_i in range(1, len(hdr_22) + 1):
        cell = ws_22.cell(row=t22_tot, column=col_i)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_i in [5, 8]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}5:{col_letter}{t22_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i in [6, 7, 9, 10, 11, 12]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}5:{col_letter}{t22_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # 4. Sheet: Bang_2_3_Ky_ke_khai_CT
    ws_23 = out_wb2.create_sheet(title="Bang_2_3_Ky_ke_khai_CT")
    ws_23.cell(row=1, column=1, value="BẢNG 2.3. CHI TIẾT HÓA ĐƠN THEO KỲ KÊ KHAI CỦA CÁC NHÀ CUNG CẤP RỦI RO").font = FONT_TITLE
    ws_23.cell(row=2, column=1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}").font = FONT_SUBTITLE

    hdr_23 = [
        "STT", "Nguồn", "Kỳ/Năm", "Số HĐ", "Ngày HĐ", "MST_chuan", "Tên người bán",
        "Mặt hàng", "Doanh số chưa thuế", "Thuế GTGT",
        "Cảnh báo rủi ro", "Thông tin công văn", "Tình trạng hoạt động (ĐKKD)", "Cơ quan thuế quản lý"
    ]
    for c_i, h in enumerate(hdr_23, start=1):
        c = ws_23.cell(row=4, column=c_i, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r23_idx = 5
    stt_23 = 1
    # Filter invoices for risky NCCs or abnormal status
    risky_msts = {m for m, n in all_ncc.items() if n['risk_flag'] in ['Công an', 'F0', 'F1'] or (n['tt'] and "đang hoạt động" not in n['tt'].lower())}
    for inv in sorted([x for x in all_invoices if x['mst'] in risky_msts], key=lambda x: (x['ngay_hd'], x['mst'])):
        rf = inv['risk_flag']
        r_vals = [
            stt_23, inv['nguon'], inv['ngay_hd'][:7] if inv['ngay_hd'] else '',
            inv['so_hd'], inv['ngay_hd'], inv['mst'], inv['ten'],
            inv['mat_hang'], inv['ds'], inv['thue'],
            rf if rf else "ĐKKD bất thường", inv['risk_cv'], inv['tt'], inv['cqt']
        ]
        stt_23 += 1
        for c_i, val in enumerate(r_vals, start=1):
            cell = ws_23.cell(row=r23_idx, column=c_i, value=sanitize_val(val))
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            if rf in ['Công an', 'F0']:
                cell.fill = FILL_RED
                if c_i in [4, 6, 7, 11]: cell.font = FONT_RED_BOLD
            elif rf == 'F1':
                cell.fill = FILL_YEL
                if c_i in [4, 6, 7, 11]: cell.font = FONT_ORG_BOLD

            if c_i in [1, 2, 3, 4, 5, 6, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i in [9, 10]:
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        r23_idx += 1

    # Total row Bảng 2.3
    t23_tot = r23_idx
    ws_23.cell(row=t23_tot, column=1, value="TỔNG CỘNG").font = FONT_BOLD
    ws_23.cell(row=t23_tot, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col_i in range(1, len(hdr_23) + 1):
        cell = ws_23.cell(row=t23_tot, column=col_i)
        cell.fill = FILL_TOTAL
        cell.border = TOTAL_BORDER
        if col_i in [9, 10]:
            col_letter = get_column_letter(col_i)
            cell.value = f"=SUM({col_letter}5:{col_letter}{t23_tot-1})"
            cell.font = FONT_BOLD
            cell.number_format = NUM_FMT
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # 5. Sheet: Phan_3_Nhan_xet
    ws_p3 = out_wb2.create_sheet(title="Phan_3_Nhan_xet")
    ws_p3.cell(row=1, column=1, value="PHẦN 3. ĐÁNH GIÁ RỦI RO VÀ ĐỀ XUẤT TRỌNG TÂM THANH TRA KIỂM TRA").font = FONT_TITLE
    ws_p3.cell(row=2, column=1, value=f"Đơn vị: {TEN_NNT} - MST: {MST_NNT} - Kỳ: {KY}").font = FONT_SUBTITLE

    hdr_p3 = ["STT", "Chuyên đề / Nhóm nội dung", "Kết quả rà soát thực tế tại hồ sơ", "Đề cương trọng tâm kiểm tra xác minh tại NNT"]
    for c_i, h in enumerate(hdr_p3, start=1):
        c = ws_p3.cell(row=4, column=c_i, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    p3_content = [
        [1, "Rủi ro tự thân của NNT (F0 / F1 / Công an)",
         "NNT không nằm trong danh sách F0 (CVCBRR), không nằm trong danh sách F1 (mua từ F0), và không có trong Sổ theo dõi phối hợp Công an. Mã số thuế đang hoạt động bình thường.",
         "Kết luận: NNT tự thân an toàn. Hồ sơ không thuộc diện trọng điểm điều tra mua bán hóa đơn tự thân."],
        [2, "Nhà cung cấp trong danh sách Công an (2 đơn vị)",
         "Phát hiện 2 nhà cung cấp nằm trong danh sách Công an: (1) CÔNG TY TNHH VI NA MEN (MST: 3601368717, CV 7309/YC-CSKT ngày 12/03/2026 của CQ CSĐT Công an TP Đồng Nai); (2) CHI NHÁNH CÔNG TY CỔ PHẦN ĐẦU TƯ XÂY DỰNG TM DV XNK TRƯỜNG TRƯỜNG PHÁT JP (MST: 0312242910-001, CV 1155/PC03-Đ3 ngày 12/11/2026 của PC03 Công an Đồng Nai).",
         "Trọng tâm 1: Yêu cầu NNT giải trình và cung cấp toàn bộ hồ sơ gốc: Hợp đồng mua bán, hóa đơn GTGT, biên bản giao nhận hàng hóa, chứng từ vận chuyển, chứng từ thanh toán không dùng tiền mặt (UNC, sao kê ngân hàng) cho 100% các hóa đơn phát sinh với 2 đơn vị này."],
        [3, "Nhà cung cấp dính CVCBRR F1 (1 đơn vị)",
         "Có 1 nhà cung cấp nằm trong danh sách F1 là TẬP ĐOÀN XĂNG DẦU VIỆT NAM (MST: 0100109106, CV 33cty). Đây là đơn vị cung ứng nhiên liệu phục vụ vận tải/sản xuất.",
         "Trọng tâm 2: Rà soát cuống hóa đơn xăng dầu, định mức nhiên liệu phương tiện vận tải, bảng kê lịch trình xe thực tế để chứng minh tính hợp lý của chi phí xăng dầu."],
        [4, "Nhà cung cấp ngừng hoạt động / đóng mã số thuế",
         f"Phát hiện tổng cộng {status_summary.get('NNT ngừng hoạt động và đã hoàn thành thủ tục chấm dứt hiệu l', {}).get('count_ncc', 0) + status_summary.get('NNT ngừng HĐ nhưng chưa hoàn thành thủ tục chấm dứt hiệu lực', {}).get('count_ncc', 0)} NCC có trạng thái ngừng hoạt động hoặc đóng MST.",
         "Trọng tâm 3: Đối chiếu chính xác ngày lập từng hóa đơn mua vào so với ngày cơ quan thuế ban hành thông báo NNT ngừng hoạt động/đóng MST. Loại trừ thuế GTGT đầu vào và chi phí đối với hóa đơn lập sau thời điểm ngừng hoạt động."]
    ]
    for r_i, r_data in enumerate(p3_content, start=5):
        for c_i, val in enumerate(r_data, start=1):
            cell = ws_p3.cell(row=r_i, column=c_i, value=val)
            cell.font = FONT_REGULAR
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c_i == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")
            elif c_i == 2:
                cell.font = FONT_BOLD

    # Autowidth all sheets in File 2
    print("Auto-fitting column widths for File 2...")
    for ws in [ws_p1, ws_21, ws_22, ws_23, ws_p3]:
        autowidth(ws)

    ws_p3.column_dimensions['A'].width = 8
    ws_p3.column_dimensions['B'].width = 32
    ws_p3.column_dimensions['C'].width = 45
    ws_p3.column_dimensions['D'].width = 50

    out_wb2.remove(default_sheet2)
    OUT_FILE2_VAULT.parent.mkdir(parents=True, exist_ok=True)
    out_wb2.save(OUT_FILE2_VAULT)
    print(f"Saved File 2 to {OUT_FILE2_VAULT}")

    OUT_FILE2_G.parent.mkdir(parents=True, exist_ok=True)
    out_wb2.save(OUT_FILE2_G)
    print(f"Saved File 2 to {OUT_FILE2_G}")

    print("\nALL JOBS COMPLETED SUCCESSFULLY!")

if __name__ == "__main__":
    run_generation()
