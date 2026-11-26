import openpyxl
from datetime import datetime

path_bt = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\Bang tinh.xlsx"
wb = openpyxl.load_workbook(path_bt, data_only=False)
ws = wb["TINH PHAT + CHAM NOP (C T)"]

print("D47 hiện tại:", ws['D47'].value)
print("Các ô từ D47 đến D66 tham chiếu như thế nào:")
for r in range(47, 67):
    print(f"D{r} = {ws[f'D{r}'].value}")
