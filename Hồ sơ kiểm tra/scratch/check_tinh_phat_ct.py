import openpyxl

path_bt = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\Bang tinh.xlsx"
wb = openpyxl.load_workbook(path_bt, data_only=False)
ws = wb["TINH PHAT + CHAM NOP (C T)"]

print("--- TIM CAC DONG CO NAM 2022-2025 VA CONG THUC ---")
for r in range(40, 80):
    val_a = ws[f'A{r}'].value
    if val_a is not None:
        row_str = f"A{r}={val_a}"
        for col in ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q']:
            v = ws[f'{col}{r}'].value
            if v is not None:
                row_str += f" | {col}{r}={v}"
        print(row_str)
