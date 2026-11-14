import docx
import re

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

def find_plain_guillemets(p, loc):
    # check full text
    full_text = p.text
    # find all matches of «...»
    matches = re.findall(r'«([^»]+)»', full_text)
    if not matches:
        return
    # check which of these matches are inside actual field preview
    # We can inspect the runs
    in_field_preview = False
    field_previews = []
    for child in p._element.iter():
        tag = child.tag.split('}')[-1]
        if tag == 'fldChar':
            fld_type = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType')
            if fld_type == 'separate':
                in_field_preview = True
            elif fld_type == 'end':
                in_field_preview = False
        elif tag == 't' and in_field_preview:
            field_previews.append(child.text or '')
    combined_field_previews = ''.join(field_previews)
    
    # If the paragraph has matches not accounted for in combined_field_previews
    for m in matches:
        token = f"«{m}»"
        if token not in combined_field_previews:
            print(f"PLAIN FIELD: {loc} -> {token} | in p.text: {full_text[:80]}")

for i, p in enumerate(doc.paragraphs):
    find_plain_guillemets(p, f"P{i}")

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                find_plain_guillemets(p, f"T{ti}_R{ri}_C{ci}_P{pi}")
