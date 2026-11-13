import openpyxl

src = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
wb = openpyxl.load_workbook(src, keep_vba=True, data_only=False)

ws = wb['01_DA_BAN_HANH']
print("01_DA_BAN_HANH max_column:", ws.max_column)
print("Headers in row 4 from col 50 to 58:")
for c in range(50, 59):
    print(c, ws.cell(4, c).value)

print("\nFormulas in row 5 (first data row) col 53 to 58:")
for c in range(53, 59):
    print(c, ws.cell(5, c).value)

print("\nDuLieuMerge max_row:", wb['DuLieuMerge'].max_row)
ws_m = wb['DuLieuMerge']
for r in range(2, 6):
    print(f"Row {r}: Ten_NNT={ws_m.cell(r, 1).value}, MST={ws_m.cell(r, 2).value}, Tong_truy_thu_chu={ws_m.cell(r, 40).value}")
