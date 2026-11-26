import openpyxl

path_bt = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\Bang tinh.xlsx"
wb = openpyxl.load_workbook(path_bt, data_only=False)
for s in wb.sheetnames:
    if "TINH PHAT" in s:
        ws = wb[s]
        print("Sheet:", s)
        for r in range(1, 40):
            vals = [f"{c}{r}={ws[f'{c}{r}'].value}" for c in ['A','B','C','D','E','F','G','H','I','J','K','L','M'] if ws[f'{c}{r}'].value is not None]
            if vals:
                print(" | ".join(vals[:6]))
