import os
import sys
import json
from pathlib import Path

base_dir = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"

def get_dir_summary(path):
    all_files = []
    for root, dirs, files in os.walk(path):
        for f in files:
            full_f = os.path.join(root, f)
            rel_f = os.path.relpath(full_f, path)
            try:
                size = os.path.getsize(full_f)
            except Exception as e:
                size = -1
            all_files.append({"rel": rel_f, "size": size, "ext": os.path.splitext(f)[1].lower()})
    return all_files

top_items = os.listdir(base_dir)
result = {}

for item in top_items:
    item_path = os.path.join(base_dir, item)
    if os.path.isdir(item_path):
        files = get_dir_summary(item_path)
        result[item] = {
            "type": "dir",
            "file_count": len(files),
            "files": files
        }
    else:
        try:
            size = os.path.getsize(item_path)
        except:
            size = -1
        result[item] = {
            "type": "file",
            "size": size
        }

with open("scan_result.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"Scanned {len(top_items)} top level items. Summary:")
for k, v in result.items():
    if v["type"] == "dir":
        print(f"[DIR] {k}: {v['file_count']} files")
    else:
        print(f"[FILE] {k}: {v['size']} bytes")
