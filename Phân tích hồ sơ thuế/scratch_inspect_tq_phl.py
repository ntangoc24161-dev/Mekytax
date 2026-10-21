import os
import openpyxl
import xlrd
import csv
import re

target = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"

# 1. Inspect 9. TRIQUYEN
tq_path = os.path.join(target, "9. TRIQUYEN", "2025 Trí Quyên", "PHU LUC 01-1, 02-1 TRIQUYEN Q1.2025.xlsx")
if not os.path.exists(tq_path):
    # look for any xlsx in 9. TRIQUYEN
    for r, d, f in os.walk(os.path.join(target, "9. TRIQUYEN")):
        for file in f:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                tq_path = os.path.join(r, file)
                break
        if tq_path: break

print("Checking Tri Quyen file:", tq_path)
try:
    if tq_path.endswith('.xlsx'):
        wb = openpyxl.load_workbook(tq_path, data_only=True)
        ws = wb.active
        for row in list(ws.iter_rows(values_only=True))[:15]:
            if any(row):
                line = " | ".join([str(x) for x in row if x is not None])
                print(" ", line)
except Exception as e:
    print("Error reading Tri Quyen xlsx:", e)

# 2. Inspect . Phat hoang long
phl_path = os.path.join(target, ". Phat hoang long", "PTDN01_20260924094404.xlsx")
print("\nChecking Phat Hoang Long file:", phl_path)
try:
    if os.path.exists(phl_path):
        wb = openpyxl.load_workbook(phl_path, data_only=True)
        ws = wb.active
        for row in list(ws.iter_rows(values_only=True))[:15]:
            if any(row):
                line = " | ".join([str(x) for x in row if x is not None])
                print(" ", line)
except Exception as e:
    print("Error reading Phat Hoang Long xlsx:", e)

# 3. Check CSV in Phat Hoang Long
csv_path = os.path.join(target, ". Phat hoang long", "tieu_de.csv")
print("\nChecking CSV:", csv_path)
try:
    if os.path.exists(csv_path):
        with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
            for _ in range(5):
                print(" ", f.readline().strip())
except Exception as e:
    print("Error reading CSV:", e)
