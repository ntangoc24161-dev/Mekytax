import docx
from collections import Counter
from standardize_template import W_NS

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\temp_test\MAUTONG_HOP_standardized.docx")

# 1. Balance check
mismatches = 0
total_fields = 0
all_instrs = []

def inspect_p(p, name):
    global mismatches, total_fields
    begins = 0
    seps = 0
    ends = 0
    curr_instr = []
    in_instr = False
    
    for c in p._element:
        for elem in c.iter():
            tag = elem.tag.split('}')[-1]
            if tag == 'fldChar':
                fc_type = elem.get(f"{{{W_NS}}}fldCharType") or elem.get("w:fldCharType") or elem.get("fldCharType")
                if fc_type == 'begin':
                    begins += 1
                    in_instr = True
                    curr_instr = []
                elif fc_type == 'separate':
                    seps += 1
                    in_instr = False
                    all_instrs.append(''.join(curr_instr).strip())
                elif fc_type == 'end':
                    ends += 1
            elif tag == 'instrText' and in_instr:
                curr_instr.append(elem.text or '')
                
    if begins != seps or begins != ends:
        print(f"Mismatch in {name}: begins={begins}, seps={seps}, ends={ends}")
        mismatches += 1
    total_fields += begins

for i, p in enumerate(doc.paragraphs):
    inspect_p(p, f"P{i}")

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                inspect_p(p, f"T{ti}_R{ri}_C{ci}_P{pi}")

print(f"Total fields in doc: {total_fields}")
print(f"Mismatches: {mismatches}")

print("\n=== FIELD COUNTS ===")
counts = Counter(all_instrs)
for k, v in sorted(counts.items(), key=lambda x: x[1], reverse=True):
    print(f"{v:3d} x [{k}]")

# Check plain guillemets outside fields
plain_outside = []
def check_outside(p, name):
    in_field_preview = False
    for c in p._element:
        for elem in c.iter():
            tag = elem.tag.split('}')[-1]
            if tag == 'fldChar':
                fc_type = elem.get(f"{{{W_NS}}}fldCharType") or elem.get("w:fldCharType") or elem.get("fldCharType")
                if fc_type == 'separate':
                    in_field_preview = True
                elif fc_type == 'end':
                    in_field_preview = False
            elif tag == 't':
                if not in_field_preview and elem.text and ('«' in elem.text or '»' in elem.text):
                    plain_outside.append((name, elem.text))

for i, p in enumerate(doc.paragraphs):
    check_outside(p, f"P{i}")

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                check_outside(p, f"T{ti}_R{ri}_C{ci}_P{pi}")

print(f"\nPlain guillemets outside fields: {len(plain_outside)}")
if plain_outside:
    for name, text in plain_outside:
        print(f"  {name}: {repr(text)}")
else:
    print("Zero plain guillemets outside fields! 100% of placeholders are now real MERGEFIELDs!")
