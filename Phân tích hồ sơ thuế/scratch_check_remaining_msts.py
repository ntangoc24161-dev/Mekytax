import os
import re
import zipfile

target = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"

# 1. Check Son Tien Hai XMLs
sth_dir = os.path.join(target, "1. SON TIEN HAI")
print("=== SON TIEN HAI ===")
for root, dirs, files in os.walk(sth_dir):
    for f in files:
        if f.lower().endswith('.xml'):
            try:
                with open(os.path.join(root, f), 'r', encoding='utf-8', errors='ignore') as fp:
                    head = fp.read(3000)
                mst_m = re.search(r'<mst>([^<]+)</mst>', head, re.IGNORECASE) or re.search(r'<mstNNT>([^<]+)</mstNNT>', head, re.IGNORECASE)
                ten_m = re.search(r'<tenNNT>([^<]+)</tenNNT>', head, re.IGNORECASE) or re.search(r'<ten_nnt>([^<]+)</ten_nnt>', head, re.IGNORECASE)
                if mst_m and ten_m:
                    print(f"  MST: {mst_m.group(1)} | Ten: {ten_m.group(1)}")
                    break
            except: pass
    if mst_m: break

# 2. Check Tam Hoang Phat
thp_dir = os.path.join(target, "4. TAM HOANG PHAT")
print("\n=== TAM HOANG PHAT ===")
for root, dirs, files in os.walk(thp_dir):
    for f in files:
        if "bien ban" in f.lower() or "thp" in f.lower() or f.endswith('.xlsx'):
            print("  Sample file:", f)
            break

# 3. Check Hung Anh Phat
hap_dir = os.path.join(target, "5. HUNG ANH PHAT")
print("\n=== HUNG ANH PHAT ===")
for root, dirs, files in os.walk(hap_dir):
    for f in files:
        if f.lower().endswith('.pdf') or f.lower().endswith('.xls'):
            print("  Sample file:", f)
            break

# 4. Check Tri Quyen
tq_dir = os.path.join(target, "9. TRIQUYEN")
print("\n=== TRI QUYEN ===")
for root, dirs, files in os.walk(tq_dir):
    for f in files:
        if f.endswith('.xls') or f.endswith('.xlsx'):
            print("  Sample file:", f)
            break

# 5. Check Phat Hoang Long
phl_dir = os.path.join(target, ". Phat hoang long")
print("\n=== PHAT HOANG LONG ===")
for root, dirs, files in os.walk(phl_dir):
    for f in files:
        if f.endswith('.csv') or f.endswith('.xlsx'):
            print("  Sample file:", f)
            break
