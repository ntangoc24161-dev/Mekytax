import docx

doc = docx.Document(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx")

for ti in [17, 24, 32]:
    print(f"=== TABLE {ti} ===")
    t = doc.tables[ti]
    for ri, r in enumerate(t.rows):
        for ci, c in enumerate(r.cells):
            print(f"r{ri} c{ci}: {repr(c.text)}")
