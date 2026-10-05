import os
import openpyxl
import xlrd

target = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\9. TRIQUYEN"

# Check BANG KE BAN RA
br_path = os.path.join(target, "2024 Trí Quyên", "BANG KE BAN RA QUY 1.2024.xlsx")
if os.path.exists(br_path):
    wb = openpyxl.load_workbook(br_path, data_only=True)
    ws = wb.active
    print("=== BAN RA 2024 Q1 ===")
    for row in list(ws.iter_rows(values_only=True))[:10]:
        if any(row):
            print(" ", [str(x) for x in row if x is not None])

nkc_path = os.path.join(target, "2024 Trí Quyên", "SO NHAT KY CHUNG 2024.xlsx")
if os.path.exists(nkc_path):
    wb = openpyxl.load_workbook(nkc_path, data_only=True)
    ws = wb.active
    print("=== NKC 2024 ===")
    for row in list(ws.iter_rows(values_only=True))[:10]:
        if any(row):
            print(" ", [str(x) for x in row if x is not None])
