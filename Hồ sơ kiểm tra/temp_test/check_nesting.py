import docx

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

nested = 0
def check_p_nesting(p, name):
    global nested
    depth = 0
    for child in p._element.iter():
        tag = child.tag.split('}')[-1]
        if tag == 'fldChar':
            fld_type = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType')
            if fld_type == 'begin':
                depth += 1
                if depth > 1:
                    print(f"Nested field in {name} (depth={depth})")
                    nested += 1
            elif fld_type == 'end':
                depth -= 1

for i, p in enumerate(doc.paragraphs):
    check_p_nesting(p, f"P{i}")

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                check_p_nesting(p, f"T{ti}_R{ri}_C{ci}_P{pi}")

print(f"Total nested fields: {nested}")
