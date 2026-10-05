import os
import re
from pathlib import Path

to_khai_dir = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\To khai"

mst_info = {}

for root, dirs, files in os.walk(to_khai_dir):
    for f in files:
        if f.lower().endswith('.xml'):
            full_p = os.path.join(root, f)
            try:
                with open(full_p, 'r', encoding='utf-8', errors='ignore') as fp:
                    head = fp.read(4000)
                
                mst_m = re.search(r'<mst>([^<]+)</mst>', head, re.IGNORECASE) or re.search(r'<mstNNT>([^<]+)</mstNNT>', head, re.IGNORECASE)
                ten_m = re.search(r'<tenNNT>([^<]+)</tenNNT>', head, re.IGNORECASE) or re.search(r'<ten_nnt>([^<]+)</ten_nnt>', head, re.IGNORECASE) or re.search(r'<tenNnt>([^<]+)</tenNnt>', head, re.IGNORECASE)
                
                if mst_m and ten_m:
                    mst = mst_m.group(1).strip()
                    ten = ten_m.group(1).strip()
                    if mst not in mst_info:
                        mst_info[mst] = set()
                    mst_info[mst].add(ten)
            except Exception as e:
                pass

print("MST AND COMPANY NAMES FOUND IN TO KHAI XML:")
for mst, names in sorted(mst_info.items()):
    print(f"MST: {mst} -> {list(names)}")
