import openpyxl

path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm"

wb = openpyxl.load_workbook(path, data_only=True)
print("Sheets:", wb.sheetnames)
ws = wb["01_DA_BAN_HANH"]
for r in range(1, 30):
    val_mst = str(ws.cell(r, 4).value or "")
    if "3603859012" in val_mst:
        print(f"Found Son Tien Hai at row {r}:")
        for c in range(1, 55):
            val = ws.cell(r, c).value
            header = ws.cell(4, c).value
            if val is not None:
                print(f"  Col {c} ({header}): {repr(val)}")
        break
