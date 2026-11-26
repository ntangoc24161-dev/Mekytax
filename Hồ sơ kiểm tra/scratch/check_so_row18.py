import openpyxl

path_so = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm"
wb = openpyxl.load_workbook(path_so, data_only=False)
ws = wb["01_DA_BAN_HANH"]
print("Row 18 columns AN to BC:")
for col in range(35, 55):
    header = ws.cell(row=4, column=col).value
    val = ws.cell(row=18, column=col).value
    col_letter = openpyxl.utils.get_column_letter(col)
    print(f"Col {col} ({col_letter}) - {header}: {val}")

print("\nRow 18 columns Z, AA, AB:")
for col in [26, 27, 28]: # Z, AA, AB
    header = ws.cell(row=4, column=col).value
    val = ws.cell(row=18, column=col).value
    col_letter = openpyxl.utils.get_column_letter(col)
    print(f"Col {col} ({col_letter}) - {header}: {val}")
