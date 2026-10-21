import json
import re

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("ALL DIRECTORIES AND TOP-LEVEL FILES:")
for k, v in data.items():
    if v['type'] == 'file':
        print(f"[FILE] {k} ({v['size']} bytes)")
    else:
        print(f"\n=======================================================")
        print(f"[DIR] {k} ({v['file_count']} files)")
        print(f"=======================================================")
        # Group by top subfolder
        subfolders = {}
        for f in v['files']:
            parts = f['rel'].split('\\')
            top_sub = parts[0] if len(parts) > 1 else "(root)"
            subfolders.setdefault(top_sub, []).append(f['rel'])
        
        for sf, flist in sorted(subfolders.items()):
            print(f"  --> Subfolder: [{sf}] ({len(flist)} files)")
            for ex in flist[:5]:
                print(f"       * {ex}")
            if len(flist) > 5:
                print(f"       ... and {len(flist) - 5} more")
