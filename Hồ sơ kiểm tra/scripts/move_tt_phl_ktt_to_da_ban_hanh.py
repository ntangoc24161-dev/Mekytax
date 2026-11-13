# -*- coding: utf-8 -*-
import os
import shutil
import win32com.client as win32

excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'

# Ensure backup exists
backup_dir = r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\backups'
os.makedirs(backup_dir, exist_ok=True)
bak_file = os.path.join(backup_dir, 'Danh sach kiem tra quyet toan pro.bak-truoc-chuyen-3hs-20261113.xlsm')
if not os.path.exists(bak_file):
    shutil.copy2(excel_path, bak_file)
    print('Backup created at:', bak_file)

excel = win32.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False

try:
    wb = excel.Workbooks.Open(excel_path)
    s1 = wb.Sheets('01_DA_BAN_HANH')
    s2 = wb.Sheets('02_DANG_THUC_HIEN')

    # Order: TT (3603883505), Phat Hoang Long (3600974507), Kien Thanh Tan (3603216935)
    mst_order = ['3603883505', '3600974507', '3603216935']
    
    # Find rows in s2
    s2_rows = {}
    for r in range(5, s2.UsedRange.Rows.Count + 1):
        mst = str(s2.Cells(r, 4).Value or '').strip()
        if mst in mst_order:
            s2_rows[mst] = r
            
    print('Found in s2:', s2_rows)
    
    dest_row = 18
    for mst in mst_order:
        src_r = s2_rows[mst]
        ten_nnt = s2.Cells(src_r, 5).Value
        print(f'Moving MST {mst} ({ten_nnt}) from s2 row {src_r} to s1 row {dest_row}...')
        
        # Copy data cells from col 3 to 51
        for c in range(3, 52):
            if c in (8, 42, 45, 48, 49):
                continue
            s1.Cells(dest_row, c).Value = s2.Cells(src_r, c).Value
            s1.Cells(dest_row, c).NumberFormat = s2.Cells(src_r, c).NumberFormat
            
        # Set status
        s1.Cells(dest_row, 2).Value = '1. Đã ban hành'
        
        # Ensure standard formulas in s1
        s1.Cells(dest_row, 1).Formula = f'=IF(D{dest_row}<>"",ROW()-5,"")'
        s1.Cells(dest_row, 8).Formula = f'=IFERROR(VLOOKUP(G{dest_row},DANH_MUC!$A$3:$B$15,2,FALSE),"")'
        s1.Cells(dest_row, 42).Formula = f'=IF(OR(AN{dest_row}<>"",AO{dest_row}<>""),AN{dest_row}+AO{dest_row},"")'
        s1.Cells(dest_row, 45).Formula = f'=IF(OR(AQ{dest_row}<>"",AR{dest_row}<>""),AQ{dest_row}+AR{dest_row},"")'
        s1.Cells(dest_row, 48).Formula = f'=IF(OR(AT{dest_row}<>"",AU{dest_row}<>""),AT{dest_row}+AU{dest_row},"")'
        s1.Cells(dest_row, 49).Formula = f'=IF(OR(AP{dest_row}<>"",AS{dest_row}<>"",AV{dest_row}<>""),IF(AP{dest_row}="",0,AP{dest_row})+IF(AS{dest_row}="",0,AS{dest_row})+IF(AV{dest_row}="",0,AV{dest_row}),"")'
        s1.Cells(dest_row, 52).Formula = f'=IF(D{dest_row}="","",IF(AND(AE{dest_row}<>"",AF{dest_row}="",TODAY()>WORKDAY(AE{dest_row},10)),"QUÁ HẠN CÔNG BỐ QĐ",IF(AND(AI{dest_row}<>"",AJ{dest_row}="",TODAY()>WORKDAY(AI{dest_row},5)),"QUÁ HẠN KÝ BBKT",IF(AND(AJ{dest_row}<>"",AL{dest_row}="",TODAY()>WORKDAY(AJ{dest_row},7)),"QUÁ HẠN BAN HÀNH QĐXPHC",IF(AND(AG{dest_row}<>"",AG{dest_row}>20),"VƯỢT QUÁ 20 NGÀY KT","Trong hạn")))))'
        s1.Cells(dest_row, 53).Formula = f'=DocSo(AP{dest_row})'
        s1.Cells(dest_row, 54).Formula = f'=DocSo(AS{dest_row})'
        s1.Cells(dest_row, 55).Formula = f'=DocSo(AV{dest_row})'
        s1.Cells(dest_row, 56).Formula = f'=DocSo(AW{dest_row})'
        s1.Cells(dest_row, 57).Formula = f'=DocSo(AX{dest_row})'
        s1.Cells(dest_row, 58).Formula = f'=DocSo(AY{dest_row})'
        
        dest_row += 1

    # Delete moved rows in s2 in reverse order so row indices don't shift
    for mst in sorted(s2_rows.keys(), key=lambda k: s2_rows[k], reverse=True):
        src_r = s2_rows[mst]
        print(f'Deleting s2 row {src_r} (MST {mst})...')
        s2.Rows(src_r).Delete()

    wb.Save()
    print('Workbook saved successfully!')
    wb.Close()
finally:
    excel.Quit()
