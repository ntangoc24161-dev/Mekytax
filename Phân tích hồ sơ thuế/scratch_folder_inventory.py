import json
import os
import re

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("INVENTORY OF ALL FOLDERS:")
for k, v in data.items():
    if v['type'] == 'dir':
        print(f"\n=======================================================")
        print(f"FOLDER: {k} (Total files: {v['file_count']})")
        print(f"=======================================================")
        files = v['files']
        # Show all file names
        for f in sorted(files, key=lambda x: x['rel'])[:30]:
            print(f"  {f['rel']} ({f['size']:,} bytes)")
        if len(files) > 30:
            print(f"  ... and {len(files) - 30} more files")
