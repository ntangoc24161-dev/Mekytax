import json
import os
import zipfile
import re

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Inspect 1. SON TIEN HAI
print("=== 1. SON TIEN HAI ===")
for f in data.get("1. SON TIEN HAI", {}).get("files", []):
    print(" ", f['rel'])

# Inspect 2. T&T BHOA
print("\n=== 2. T&T BHOA ===")
for f in data.get("2. T&T BHOA", {}).get("files", [])[:15]:
    print(" ", f['rel'])

# Inspect 8. HUY THANH DAT (inside zip)
zip_path = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\8. HUY THANH DAT\tiliuphcvcngtckimtrathunm2026.zip"
try:
    with zipfile.ZipFile(zip_path, 'r') as z:
        print("\n=== 8. HUY THANH DAT (ZIP contents) ===")
        for name in z.namelist()[:20]:
            print(" ", name)
        print(f" Total files in zip: {len(z.namelist())}")
except Exception as e:
    print(" Error reading zip:", e)

# Inspect . Phat hoang long
print("\n=== . Phat hoang long ===")
for f in data.get(". Phat hoang long", {}).get("files", []):
    print(" ", f['rel'])
