import docx

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

def extract_fields(elem):
    instrs = elem.xpath('.//w:instrText')
    return [i.text for i in instrs if i.text]

print("=== PARAGRAPHS ===")
for i, p in enumerate(doc.paragraphs):
    flds = extract_fields(p._element)
    if flds:
        print(f"P{i} [L={len(flds)}]: {flds} | text: {p.text[:80]}")

print("\n=== TABLES ===")
t_count = 0
for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                flds = extract_fields(p._element)
                if flds:
                    t_count += 1
                    print(f"T{ti} R{ri} C{ci} P{pi}: {flds} | text: {p.text[:80]}")

print(f"\nTotal: paragraphs={len(doc.paragraphs)}, table cells with fields={t_count}")
