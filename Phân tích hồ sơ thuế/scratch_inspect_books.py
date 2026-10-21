import os, glob
from pathlib import Path
import openpyxl
import xlrd # for .xls if installed or openpyxl

base_dir = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505")

print("=== INSPECTING ACCOUNTING BOOKS ===")
for root, dirs, files in os.walk(base_dir / "2. HO SO NNT CUNG CAP"):
    for f in files:
        if f.endswith('.xls') or f.endswith('.xlsx'):
            full_p = os.path.join(root, f)
            sz = os.path.getsize(full_p)
            print(f"{f:40s} | {sz:8d} bytes | {os.path.relpath(root, base_dir)}")
