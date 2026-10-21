import json

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

to_khai = data.get('To khai', {}).get('files', [])
print(f"Total files in To khai: {len(to_khai)}")

subfolders = {}
for f in to_khai:
    parts = f['rel'].split('\\')
    sub = parts[0]
    subfolders.setdefault(sub, []).append(parts[-1])

for sub, files in sorted(subfolders.items()):
    print(f"\nSubfolder: {sub} ({len(files)} files)")
    for sample in files[:5]:
        print(f"  - {sample}")
