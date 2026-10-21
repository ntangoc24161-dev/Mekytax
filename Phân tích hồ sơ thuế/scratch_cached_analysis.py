import json

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("COMPANIES / FOLDERS IN KE HOACH:")
for k in sorted(data.keys()):
    if data[k]['type'] == 'dir':
        print(f"\nFolder: {k} ({data[k]['file_count']} files)")
        exts = {}
        for f in data[k]['files']:
            e = f['ext']
            exts[e] = exts.get(e, 0) + 1
        print(f"  Extensions: {exts}")
