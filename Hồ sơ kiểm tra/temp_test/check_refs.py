import openpyxl
import re

excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
wb = openpyxl.load_workbook(excel_path, keep_vba=True, data_only=False)

# Check all formulas across all sheets that refer to 01_DA_BAN_HANH
refs = []
for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or '')
            if '01_DA_BAN_HANH' in val:
                refs.append((sheetname, ws.cell(r, c).coordinate, val))

print(f"Total formula references to 01_DA_BAN_HANH across workbook: {len(refs)}")
for ref in refs[:20]:
    print(ref)

# Also check 02_DANG_THUC_HIEN header row to see if it's identical
ws2 = wb['02_DANG_THUC_HIEN']
print(f"02_DANG_THUC_HIEN: max_row={ws2.max_row}, max_col={ws2.max_column}")
