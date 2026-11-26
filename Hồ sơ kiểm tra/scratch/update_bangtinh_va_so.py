import shutil
from datetime import datetime
import openpyxl

path_bt = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\Bang tinh.xlsx"
path_bt_bak = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\Bang tinh_backup_before_30112026.xlsx"
shutil.copyfile(path_bt, path_bt_bak)

wb_bt = openpyxl.load_workbook(path_bt, data_only=False)
ws_ct = wb_bt["TINH PHAT + CHAM NOP (C T)"]
ws_ct['D47'] = datetime(2026, 11, 30, 0, 0, 0)
wb_bt.save(path_bt)
print("Updated Bang tinh.xlsx: D47 = 2026-11-30")

path_so = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm"
path_so_bak = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro_bak.xlsm"
shutil.copyfile(path_so, path_so_bak)

wb_so = openpyxl.load_workbook(path_so, keep_vba=True)
ws_so = wb_so["01_DA_BAN_HANH"]

# Row 18
# AN (col 40): Truy thu TNDN = 14023597
ws_so.cell(row=18, column=40, value=14023597)
# AO (col 41): Truy thu GTGT = 0
ws_so.cell(row=18, column=41, value=0)
# AQ (col 43): Cham nop GTGT = 0
ws_so.cell(row=18, column=43, value=0)
# AR (col 44): Cham nop TNDN = 1022320
ws_so.cell(row=18, column=44, value=1022320)
# AT (col 46): Phat 20% = 2804719
ws_so.cell(row=18, column=46, value=2804719)
# AU (col 47): Phat thu tuc = 0
ws_so.cell(row=18, column=47, value=0)
# AX (col 50): Giam lo = 0
ws_so.cell(row=18, column=50, value=0)
# AY (col 51): Giam khau tru = 0
ws_so.cell(row=18, column=51, value=0)

# Verify Z, AA, AB
print("Z18:", ws_so.cell(row=18, column=26).value)
print("AA18:", ws_so.cell(row=18, column=27).value)
print("AB18:", ws_so.cell(row=18, column=28).value)

wb_so.save(path_so)
print("Updated Danh sach kiem tra quyet toan pro.xlsm successfully!")
