import docx
import re

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

def find_fields(elem):
    # Walk through children to reconstruct fields
    # A field starts with fldChar begin, has instrTexts, then fldChar separate, then runs with t, then fldChar end
    fields = []
    curr_instr = []
    curr_text = []
    in_instr = False
    in_separate = False

    for child in elem.iter():
        tag = child.tag.split('}')[-1]
        if tag == 'fldChar':
            fld_type = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType')
            if fld_type == 'begin':
                curr_instr = []
                curr_text = []
                in_instr = True
                in_separate = False
            elif fld_type == 'separate':
                in_instr = False
                in_separate = True
            elif fld_type == 'end':
                in_instr = False
                in_separate = False
                fields.append((''.join(curr_instr).strip(), ''.join(curr_text)))
        elif tag == 'instrText' and in_instr:
            curr_instr.append(child.text or '')
        elif tag == 't' and in_separate:
            curr_text.append(child.text or '')
    return fields

all_fields = []
for p in doc.paragraphs:
    all_fields.extend(find_fields(p._element))

for t in doc.tables:
    for row in t.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                all_fields.extend(find_fields(p._element))

print(f"Total reconstructed fields: {len(all_fields)}")
unique_fields = set()
for instr, text in all_fields:
    unique_fields.add((instr, text))

for instr, text in sorted(unique_fields):
    print(f"INSTR: [{instr}]  -->  TEXT: [{text}]")
