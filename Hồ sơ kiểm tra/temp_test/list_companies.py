import openpyxl

excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
wb = openpyxl.load_workbook(excel_path, keep_vba=True, data_only=True)
ws = wb['01_DA_BAN_HANH']

print("Rows in 01_DA_BAN_HANH:")
for r in range(5, ws.max_row + 1):
    stt = ws.cell(r, 1).value
    mst = ws.cell(r, 4).value
    ten = ws.cell(r, 5).value
    if mst or ten:
        print(f"Row {r:2d}: STT={stt}, MST={mst}, Tên={ten}")
