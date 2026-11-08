#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
import os

base_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH'
kq21_path = os.path.join(base_path, 'KQ2.1_Phan_tich_CQT_3603883505_2022-2025_v4_FINAL.xlsx')

wb = openpyxl.load_workbook(kq21_path, data_only=True)

# Xem KQ2_NCC_mua_vao
print("=" * 100)
print("Sheet: KQ2_NCC_mua_vao - Dòng 6-15")
print("=" * 100)
ws = wb['KQ2_NCC_mua_vao']

# In dòng 6-15
for row_num in range(6, 16):
    row_data = []
    for col in range(1, 10):
        val = ws.cell(row_num, col).value
        if val:
            row_data.append(f"C{col}:{val}")
    if row_data or row_num == 6:
        print(f"Dòng {row_num}: {' | '.join(row_data) if row_data else 'TRỐNG'}")

# Tìm dòng "Tổng cộng" hoặc header
print("\n" + "=" * 100)
print("Tìm dòng header...")
print("=" * 100)

for row_num in range(1, 50):
    val_c1 = ws.cell(row_num, 1).value
    if val_c1 and ('STT' in str(val_c1) or 'Tên' in str(val_c1) or 'MST' in str(val_c1)):
        print(f"\nDòng {row_num} là HEADER:")
        for col in range(1, 15):
            h = ws.cell(row_num, col).value
            if h:
                print(f"  C{col}: {h}")

        # Hiển thị 2 dòng dữ liệu sau header
        print(f"\nDữ liệu (dòng {row_num+1}-{row_num+2}):")
        for data_row in range(row_num+1, row_num+3):
            data = []
            for col in range(1, 10):
                d = ws.cell(data_row, col).value
                data.append(f"{col}:{d}")
            print(f"  {' | '.join(data)}")
        break

wb.close()
