import docx

doc = docx.Document(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx")

print("=== TABLE 42 ===")
for ri, r in enumerate(doc.tables[42].rows):
    row_txt = [c.text.strip().replace("\n", " ") for c in r.cells]
    print(f"r{ri}: {row_txt}")

print("=== TABLE 50 ===")
for ri, r in enumerate(doc.tables[50].rows):
    row_txt = [c.text.strip().replace("\n", " ") for c in r.cells]
    print(f"r{ri}: {row_txt}")
