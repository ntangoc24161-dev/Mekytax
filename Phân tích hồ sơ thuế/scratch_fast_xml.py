import os
import re

to_khai_dir = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\To khai"

subdirs = [d for d in os.listdir(to_khai_dir) if os.path.isdir(os.path.join(to_khai_dir, d))]

for sub in sorted(subdirs):
    sub_path = os.path.join(to_khai_dir, sub)
    found = False
    for root, dirs, files in os.walk(sub_path):
        for f in files:
            if f.lower().endswith('.xml'):
                full_p = os.path.join(root, f)
                try:
                    with open(full_p, 'r', encoding='utf-8', errors='ignore') as fp:
                        head = fp.read(4000)
                    mst_m = re.search(r'<mst>([^<]+)</mst>', head, re.IGNORECASE) or re.search(r'<mstNNT>([^<]+)</mstNNT>', head, re.IGNORECASE)
                    ten_m = re.search(r'<tenNNT>([^<]+)</tenNNT>', head, re.IGNORECASE) or re.search(r'<ten_nnt>([^<]+)</ten_nnt>', head, re.IGNORECASE) or re.search(r'<tenNnt>([^<]+)</tenNnt>', head, re.IGNORECASE)
                    mst = mst_m.group(1).strip() if mst_m else "N/A"
                    ten = ten_m.group(1).strip() if ten_m else "N/A"
                    print(f"Folder: {sub:35} | MST: {mst:15} | NNT: {ten}")
                    found = True
                    break
                except Exception as e:
                    pass
        if found:
            break
