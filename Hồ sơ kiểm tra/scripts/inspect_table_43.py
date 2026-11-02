import docx

doc = docx.Document(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx")

print("=== TABLE 43 ===")
for ri, r in enumerate(doc.tables[43].rows):
    row_txt = [c.text.strip().replace("\n", " ") for c in r.cells]
    print(f"r{ri}: {row_txt}")
