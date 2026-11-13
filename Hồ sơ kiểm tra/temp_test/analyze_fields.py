import docx
from collections import defaultdict

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

def find_fields(elem, container_name):
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
                fields.append((container_name, ''.join(curr_instr).strip(), ''.join(curr_text)))
        elif tag == 'instrText' and in_instr:
            curr_instr.append(child.text or '')
        elif tag == 't' and in_separate:
            curr_text.append(child.text or '')
    return fields

all_fields = []
for i, p in enumerate(doc.paragraphs):
    all_fields.extend(find_fields(p._element, f"P{i}"))

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                all_fields.extend(find_fields(p._element, f"T{ti}_R{ri}_C{ci}_P{pi}"))

grouped = defaultdict(list)
for loc, instr, text in all_fields:
    grouped[instr].append(loc)

for instr, locs in sorted(grouped.items(), key=lambda x: len(x[1]), reverse=True):
    print(f"{len(locs):3d} x [{instr}] (e.g. {locs[0]})")
