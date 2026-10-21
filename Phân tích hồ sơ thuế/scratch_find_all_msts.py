import os
import re

target = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"

def find_xml_info(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower().endswith('.xml'):
                full_p = os.path.join(root, f)
                try:
                    with open(full_p, 'r', encoding='utf-8', errors='ignore') as fp:
                        head = fp.read(4000)
                    mst_m = re.search(r'<mst>([^<]+)</mst>', head, re.IGNORECASE) or re.search(r'<mstNNT>([^<]+)</mstNNT>', head, re.IGNORECASE)
                    ten_m = re.search(r'<tenNNT>([^<]+)</tenNNT>', head, re.IGNORECASE) or re.search(r'<ten_nnt>([^<]+)</ten_nnt>', head, re.IGNORECASE) or re.search(r'<tenNnt>([^<]+)</tenNnt>', head, re.IGNORECASE)
                    if mst_m and ten_m:
                        return mst_m.group(1).strip(), ten_m.group(1).strip()
                except:
                    pass
    return None, None

def find_text_in_files(folder_path):
    # look in file names or sample files for MST (10 digits)
    mst_candidates = set()
    for root, dirs, files in os.walk(folder_path):
        for f in files:
            m = re.findall(r'360\d{7}', f)
            for x in m:
                mst_candidates.add(x)
    return list(mst_candidates)

dirs = [d for d in os.listdir(target) if os.path.isdir(os.path.join(target, d))]

for d in sorted(dirs):
    p = os.path.join(target, d)
    mst, ten = find_xml_info(p)
    msts_in_names = find_text_in_files(p)
    print(f"\nFolder: {d}")
    print(f"  XML info: MST={mst}, Name={ten}")
    print(f"  MSTs in filenames: {msts_in_names}")
