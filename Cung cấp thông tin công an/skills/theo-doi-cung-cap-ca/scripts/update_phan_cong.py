# -*- coding: utf-8 -*-
"""
Script ho tro doc, doi chieu va cap nhat sheet 'Phan cong' trong file Excel
theo doi ho so cung cap Cong an tai Thue co so / Co quan Thue.
File Excel: I:\\My Drive\\HS CONG AN-2025-2026\\HS CUNG CAP CA 2026\\File theo doi hs CUNG CAP CA-01.04.2025.xls
"""

import os
import sys
import json
import argparse
from datetime import datetime

DEFAULT_EXCEL_PATH = r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\File theo doi hs CUNG CAP CA-01.04.2025.xls"
SHEET_NAME = "Phan cong"

# Anh xa cot (1-based cho Excel COM)
COL_STT = 1
COL_SO_CV = 2
COL_NGAY_CV = 3
COL_CO_QUAN = 4
COL_SO_CV_DEN = 5
COL_NGAY_CV_DEN = 6
COL_CBT_CC = 7
COL_MST = 8
COL_TEN_NNT = 9
COL_DIA_CHI = 10
COL_TO_KT = 11
COL_CAN_BO_XL = 12
COL_NGAY_GIAO = 13
COL_TINH_TRANG = 14
COL_GHI_CHU = 15
COL_NGAY_GP = 16
COL_DAI_DIEN_PL = 17
COL_NOI_DUNG = 18

def normalize_mst(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.replace(" ", "").replace("-", "")

def get_excel_app():
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel

def find_rows_by_so_cv(excel_path, query_so_cv):
    """Tim kiem cac dong khop voi So cong van"""
    import xlrd
    wb = xlrd.open_workbook(excel_path)
    s = wb.sheet_by_name(SHEET_NAME)
    
    query = str(query_so_cv).strip().lower()
    matches = []
    
    current_stt = None
    group_rows = []
    
    for r in range(5, s.nrows):
        stt_val = s.cell_value(r, 0)
        cv_val = str(s.cell_value(r, 1)).strip()
        mst_val = normalize_mst(s.cell_value(r, 7))
        ten_nnt = str(s.cell_value(r, 8)).strip()
        so_den = str(s.cell_value(r, 4)).strip()
        noi_dung = str(s.cell_value(r, 17)).strip()
        
        row_info = {
            "excel_row": r + 1,
            "stt": stt_val,
            "so_cv": cv_val,
            "so_den": so_den,
            "mst": mst_val,
            "ten_nnt": ten_nnt,
            "noi_dung": noi_dung,
            "to_kt": str(s.cell_value(r, 10)).strip(),
            "can_bo_xl": str(s.cell_value(r, 11)).strip()
        }
        
        if cv_val and query in cv_val.lower():
            # Tim thay dong dau cua cum
            target_stt = stt_val
            # Gom tat ca cac dong cung STT
            matches.append(row_info)
            for sub_r in range(r + 1, min(r + 20, s.nrows)):
                if s.cell_value(sub_r, 0) == target_stt:
                    matches.append({
                        "excel_row": sub_r + 1,
                        "stt": s.cell_value(sub_r, 0),
                        "so_cv": str(s.cell_value(sub_r, 1)).strip(),
                        "so_den": str(s.cell_value(sub_r, 4)).strip(),
                        "mst": normalize_mst(s.cell_value(sub_r, 7)),
                        "ten_nnt": str(s.cell_value(sub_r, 8)).strip(),
                        "noi_dung": str(s.cell_value(sub_r, 17)).strip(),
                        "to_kt": str(s.cell_value(sub_r, 10)).strip(),
                        "can_bo_xl": str(s.cell_value(sub_r, 11)).strip()
                    })
                else:
                    break
            break
            
    return matches

def find_rows_by_mst(excel_path, query_mst):
    """Tim kiem dong theo Ma so thue"""
    import xlrd
    wb = xlrd.open_workbook(excel_path)
    s = wb.sheet_by_name(SHEET_NAME)
    
    q_mst = normalize_mst(query_mst)
    matches = []
    
    for r in range(5, s.nrows):
        mst_val = normalize_mst(s.cell_value(r, 7))
        if q_mst in mst_val:
            matches.append({
                "excel_row": r + 1,
                "stt": s.cell_value(r, 0),
                "so_cv": str(s.cell_value(r, 1)).strip(),
                "so_den": str(s.cell_value(r, 4)).strip(),
                "mst": mst_val,
                "ten_nnt": str(s.cell_value(r, 8)).strip(),
                "noi_dung": str(s.cell_value(r, 17)).strip(),
                "to_kt": str(s.cell_value(r, 10)).strip(),
                "can_bo_xl": str(s.cell_value(r, 11)).strip()
            })
    return matches

def update_cells(excel_path, update_data):
    """
    Cap nhat cac o tren sheet Phan cong qua COM Excel de giu nguyen dinh dang.
    update_data: list of dict, moi dict gom:
      - row: so dong Excel (1-based)
      - col: so cot Excel (1-based) hoac ten cot
      - value: gia tri moi
    """
    excel = get_excel_app()
    try:
        wb = excel.Workbooks.Open(excel_path)
        ws = wb.Sheets(SHEET_NAME)
        
        applied = []
        for item in update_data:
            r = item["row"]
            c = item["col"]
            val = item["value"]
            
            old_val = ws.Cells(r, c).Value
            ws.Cells(r, c).Value = val
            applied.append({
                "row": r,
                "col": c,
                "old_val": str(old_val),
                "new_val": str(val)
            })
            
        wb.Save()
        wb.Close(True)
        return {"status": "success", "applied": applied}
    finally:
        excel.Quit()

def main():
    parser = argparse.ArgumentParser(description="Cong cu ho tro theo doi cung cap ho so Cong an")
    parser.add_argument("--file", default=DEFAULT_EXCEL_PATH, help="Duong dan file Excel")
    parser.add_argument("--find-cv", help="Tim kiem theo so cong van")
    parser.add_argument("--find-mst", help="Tim kiem theo ma so thue")
    parser.add_argument("--update-json", help="File JSON chua danh sach cap nhat")
    
    args = parser.parse_args()
    
    if args.find_cv:
        res = find_rows_by_so_cv(args.file, args.find_cv)
        print(json.dumps(res, ensure_ascii=False, indent=2))
    elif args.find_mst:
        res = find_rows_by_mst(args.file, args.find_mst)
        print(json.dumps(res, ensure_ascii=False, indent=2))
    elif args.update_json:
        with open(args.update_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        res = update_cells(args.file, data)
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print("Vui long chi dinh tham so --find-cv, --find-mst hoac --update-json")

if __name__ == "__main__":
    main()
