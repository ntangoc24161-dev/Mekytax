#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
import os

base_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH'
kq21_path = os.path.join(base_path, 'KQ2.1_Phan_tich_CQT_3603883505_2022-2025_v4_FINAL.xlsx')

wb = openpyxl.load_workbook(kq21_path, data_only=True)

# Xem KQ2_NCC_mua_vao
print("=" * 100)
print("Sheet: KQ2_NCC_mua_vao")
print("=" * 100)
ws = wb['KQ2_NCC_mua_vao']

# In 5 dòng đầu
for row_num in range(1, 6):
    print(f"\nDòng {row_num}:")
    for col in range(1, 15):
        val = ws.cell(row_num, col).value
        if val or col <= 8:
            print(f"  Cột {col}: {val}")

wb.close()
