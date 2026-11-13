import openpyxl

excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
wb = openpyxl.load_workbook(excel_path, keep_vba=True, data_only=False)

ws = wb['01_DA_BAN_HANH']

# 1. Check current columns
print("Current max col in 01_DA_BAN_HANH:", ws.max_column)
print("Col 52 header:", ws.cell(4, 52).value)

# 2. Check DuLieuMerge
ws_merge = wb['DuLieuMerge']
print("DuLieuMerge max_row:", ws_merge.max_row, "max_col:", ws_merge.max_column)
for c in range(1, 10):
    print(ws_merge.cell(1, c).value, ws_merge.cell(2, c).value)
