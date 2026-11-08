import docx

doc = docx.Document(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx")

print("=== TABLE 1 ===")
for r in doc.tables[1].rows:
    for c in r.cells:
        print(repr(c.text.strip()))

print("=== TABLE 3 ===")
for r in doc.tables[3].rows:
    for c in r.cells:
        print(repr(c.text.strip()))

print("=== ALL TCS3 PARAGRAPHS FULL TEXT ===")
for i, p in enumerate(doc.paragraphs):
    if "/qđ-tcs3" in p.text.lower() or "qđ-tcs3" in p.text.lower():
        print(f"P{i}: {p.text}")
