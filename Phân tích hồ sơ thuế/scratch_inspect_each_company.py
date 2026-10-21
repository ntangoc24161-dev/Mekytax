import json
import re

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

companies = [
    "1. SON TIEN HAI",
    "2. T&T BHOA",
    "3. PHUC NGUYEN TIEN",
    "4. TAM HOANG PHAT",
    "5. HUNG ANH PHAT",
    "6. DAI HUY HOANG",
    "7. NHAT VINH-PASS LAM",
    "8. HUY THANH DAT",
    "9. TRIQUYEN",
    ". Phat hoang long",
    "HO SO KE KOACH"
]

for c in companies:
    if c in data:
        files = [f['rel'] for f in data[c]['files']]
        print(f"\n=======================================================")
        print(f"COMPANY / FOLDER: {c} ({len(files)} files)")
        print(f"=======================================================")
        for f in files[:20]:
            print(f"  - {f}")
        if len(files) > 20:
            print(f"  ... and {len(files) - 20} more files")
