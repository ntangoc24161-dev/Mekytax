import openpyxl

excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
wb = openpyxl.load_workbook(excel_path, keep_vba=True, data_only=False)

ws_merge = wb['DuLieuMerge']
ws_banhanh = wb['01_DA_BAN_HANH']

print("Mapping in DuLieuMerge row 1 (Header) and row 2 (Formula):")
for col in range(1, ws_merge.max_column + 1):
    h = ws_merge.cell(1, col).value
    f = ws_merge.cell(2, col).value
    print(f"Col {col:2d}: {h:20s} -> {str(f)}")
