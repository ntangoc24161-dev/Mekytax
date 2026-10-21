import json
import os
import xml.etree.ElementTree as ET

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Inspect To khai subfolders
to_khai_files = data.get('To khai', {}).get('files', [])
to_khai_subdirs = {}
for f in to_khai_files:
    parts = f['rel'].split('\\')
    top = parts[0]
    to_khai_subdirs.setdefault(top, []).append(f['rel'])

print("TO KHAI SUBFOLDERS (MSTs):")
for mst_folder, flist in sorted(to_khai_subdirs.items()):
    print(f"- {mst_folder}: {len(flist)} files")

# Inspect HO SO KE KOACH subfolders
hs_kk_files = data.get('HO SO KE KOACH', {}).get('files', [])
hs_kk_subdirs = {}
for f in hs_kk_files:
    parts = f['rel'].split('\\')
    top = parts[0]
    hs_kk_subdirs.setdefault(top, []).append(f['rel'])

print("\nHO SO KE KOACH SUBFOLDERS:")
for sf, flist in sorted(hs_kk_subdirs.items()):
    print(f"- {sf}: {len(flist)} files")
