import docx

doc = docx.Document(r"output\3603859012_HO-SO-VE_son-tien-hai\QD XPHC Cong ty Son Tien Hai.docx")
print("Total paragraphs:", len(doc.paragraphs))
print("Total tables:", len(doc.tables))
for i, p in enumerate(doc.paragraphs):
    if any(k in p.text.lower() for k in ["tcs3", "kinh doanh", "doanh nghiệp", "lê hoài phương", "6722949824"]):
        print(f"P{i}: {p.text}")

for ti, t in enumerate(doc.tables):
    for ri, r in enumerate(t.rows):
        for ci, c in enumerate(r.cells):
            if any(k in c.text.lower() for k in ["tcs3", "kinh doanh", "doanh nghiệp", "lê hoài phương", "6722949824"]):
                print(f"T{ti}[r{ri},c{ci}]: {repr(c.text.strip())}")
