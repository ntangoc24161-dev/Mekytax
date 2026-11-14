import docx

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

def check_plain_guillemets(p, name):
    # Check if there is any « or » outside of fldChar separate..end
    in_field_preview = False
    plain = []
    for child in p._element.iter():
        tag = child.tag.split('}')[-1]
        if tag == 'fldChar':
            fld_type = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType')
            if fld_type == 'separate':
                in_field_preview = True
            elif fld_type == 'end':
                in_field_preview = False
        elif tag == 't':
            if not in_field_preview and child.text and ('«' in child.text or '»' in child.text):
                plain.append(child.text)
    if plain:
        print(f"Plain guillemets in {name}: {plain}")

for i, p in enumerate(doc.paragraphs):
    check_plain_guillemets(p, f"P{i}")

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                check_plain_guillemets(p, f"T{ti}_R{ri}_C{ci}_P{pi}")

print("Check plain guillemets done.")
