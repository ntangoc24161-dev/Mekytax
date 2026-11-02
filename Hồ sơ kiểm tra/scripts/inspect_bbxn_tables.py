import docx

doc = docx.Document(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx")

print("=== BBXN TRUONG DOAN (Tables 18 to 24) ===")
for ti in range(18, 25):
    t = doc.tables[ti]
    print(f"\n--- Table {ti} ({len(t.rows)} rows, {len(t.columns)} cols) ---")
    for ri, r in enumerate(t.rows):
        cells = [c.text.strip().replace("\n", " ") for c in r.cells]
        print(f"r{ri}: {cells}")

print("\n=== BBXN THANH VIEN (Tables 25 to 31) ===")
for ti in range(25, 32):
    t = doc.tables[ti]
    print(f"\n--- Table {ti} ({len(t.rows)} rows, {len(t.columns)} cols) ---")
    for ri, r in enumerate(t.rows):
        cells = [c.text.strip().replace("\n", " ") for c in r.cells]
        print(f"r{ri}: {cells}")
