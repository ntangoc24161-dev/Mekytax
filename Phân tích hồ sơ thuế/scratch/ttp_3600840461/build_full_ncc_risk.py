# -*- coding: utf-8 -*-
"""
Script tong hop toan dien cho CONG TY TNHH THANH THANH PHUONG (MST: 3600840461, Ky: 2020-2025)
Theo Huong 2:
1. File 1: Tong_hop_HDDT_BangKe_3600840461_2020-2025.xlsx
   - HDDT_Mua_vao: them cot Canh bao rui ro & Thong tin CV CA/CVCBRR
   - HDDT_Ban_ra: giu nguyen
   - Bang_ke_Mua_vao: them cot Canh bao rui ro & Thong tin CV CA/CVCBRR
   - Bang_ke_Ban_ra: giu nguyen
   - Sheet 5: Tong_hop_trang_thai_rui_ro (Tong hop theo trang thai hoat dong & CV Cong an / CVCBRR)
2. File 2: Tong_hop_NCC_mua_vao_CV_trang_thai_3600840461_2020-2025.xlsx
   - Phan_1_F1_NNT
   - Bang_2_1_NCC_F0
   - Bang_2_2_Toan_bo_NCC (100% NCC)
   - Bang_2_3_Ky_ke_khai_CT
   - Phan_3_Nhan_xet
"""
import os
import sys
import re
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
        # Parse NHAP_HO_SO
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
        # Parse other sheets if any missed
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
                                    'cv': 'Theo dõi Công an (' + sname + ')',
                                    'ngay': '',
                                    'co_quan': 'Công an',
                                    'nd': 'Nằm trong danh sách ' + sname
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

print("Setup completed.")
