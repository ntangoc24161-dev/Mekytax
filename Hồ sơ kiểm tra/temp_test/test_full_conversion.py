import docx
import re
import copy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from standardize_template import get_field_mapping, make_mergefield_elements, process_paragraph_existing_fields, W_NS

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

# Step 1: Process all existing fields in all paragraphs and table cells
total_existing_updated = 0
for i, p in enumerate(doc.paragraphs):
    cnt = process_paragraph_existing_fields(p._element)
    total_existing_updated += cnt

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                cnt = process_paragraph_existing_fields(p._element)
                total_existing_updated += cnt

print(f"Step 1 completed: Updated {total_existing_updated} existing fields.")

plain_mapping = {
    'TĐ': ('MERGEFIELD Truong_doan', '«Truong_doan»'),
    'GĐ': ('MERGEFIELD Giam_doc', '«Giam_doc»'),
    'KT': ('MERGEFIELD Ke_toan_truong', '«Ke_toan_truong»'),
    'TV': ('MERGEFIELD Thanh_vien', '«Thanh_vien»'),
    'Phân_công_trưởng_đoàn': ('MERGEFIELD Phan_cong_TD', '«Phan_cong_TD»'),
    'Phân_công_thành_viên': ('MERGEFIELD Phan_cong_TV', '«Phan_cong_TV»'),
}

def has_fldChar(node):
    for e in node.iter():
        if e.tag.split('}')[-1] == 'fldChar':
            return True
    return False

def get_node_text(node):
    return ''.join(e.text or '' for e in node.iter() if e.tag.split('}')[-1] == 't')

def get_node_rPr(node):
    for e in node.iter():
        if e.tag.split('}')[-1] == 'rPr':
            return e
    return None

def convert_plain_guillemets_in_p(p):
    p_elem = p._element
    children = list(p_elem)
    in_field = False
    runs_outside = []
    for c in children:
        for elem in c.iter():
            tag = elem.tag.split('}')[-1]
            if tag == 'fldChar':
                fc_type = elem.get(f"{{{W_NS}}}fldCharType") or elem.get("w:fldCharType") or elem.get("fldCharType")
                if fc_type == 'begin':
                    in_field = True
                elif fc_type == 'end':
                    in_field = False
        if not in_field and c.tag.split('}')[-1] == 'r':
            runs_outside.append(c)

    full_outside_text = ''.join(get_node_text(r) for r in runs_outside)
    if '«' not in full_outside_text:
        return 0

    converted = 0
    
    # Check runs outside for whole tokens or multi tokens
    for r in list(runs_outside):
        text = get_node_text(r)
        if not text:
            continue
        rPr_elem = get_node_rPr(r)
        
        # Check if single token like «TĐ»
        m_single = re.fullmatch(r'«([^»]+)»', text.strip())
        if m_single and m_single.group(1) in plain_mapping:
            tag = m_single.group(1)
            instr, disp = plain_mapping[tag]
            new_elems = make_mergefield_elements(instr, disp, rPr_elem)
            idx = p_elem.index(r)
            p_elem.remove(r)
            for ne in reversed(new_elems):
                p_elem.insert(idx, ne)
            converted += 1
            continue
            
        # Check if multi token like '«GĐ»            «KT»'
        if '«' in text and '»' in text:
            parts = re.split(r'(«[^»]+»)', text)
            if len(parts) > 1 and any(p.startswith('«') for p in parts):
                idx = p_elem.index(r)
                p_elem.remove(r)
                insert_pos = idx
                for part in parts:
                    m = re.match(r'«([^»]+)»', part)
                    if m and m.group(1) in plain_mapping:
                        tag = m.group(1)
                        instr, disp = plain_mapping[tag]
                        new_elems = make_mergefield_elements(instr, disp, rPr_elem)
                        for ne in new_elems:
                            p_elem.insert(insert_pos, ne)
                            insert_pos += 1
                        converted += 1
                    elif part:
                        r_new = OxmlElement('w:r')
                        if rPr_elem is not None:
                            r_new.append(copy.deepcopy(rPr_elem))
                        t_new = OxmlElement('w:t')
                        t_new.text = part
                        if part.startswith(' ') or part.endswith(' '):
                            t_new.set(qn('xml:space'), 'preserve')
                        r_new.append(t_new)
                        p_elem.insert(insert_pos, r_new)
                        insert_pos += 1
                continue

    # Now check for split runs (e.g. run '«', run 'Phân_công_trưởng_đoàn', run '»')
    runs_outside = [c for c in p_elem if c.tag.split('}')[-1] == 'r' and not has_fldChar(c)]
    i = 0
    while i < len(runs_outside) - 2:
        r1, r2, r3 = runs_outside[i], runs_outside[i+1], runs_outside[i+2]
        t1 = get_node_text(r1)
        t2 = get_node_text(r2)
        t3 = get_node_text(r3)
        if t1.strip() == '«' and t2.strip() in plain_mapping and t3.strip().startswith('»'):
            tag = t2.strip()
            instr, disp = plain_mapping[tag]
            rPr_elem = get_node_rPr(r2)
            new_elems = make_mergefield_elements(instr, disp, rPr_elem)
            
            idx = p_elem.index(r1)
            p_elem.remove(r1)
            p_elem.remove(r2)
            
            trailing = t3.strip()[1:]
            if trailing:
                for elem in r3.iter():
                    if elem.tag.split('}')[-1] == 't':
                        elem.text = trailing
            else:
                p_elem.remove(r3)
                
            for ne in reversed(new_elems):
                p_elem.insert(idx, ne)
                
            converted += 1
            runs_outside = [c for c in p_elem if c.tag.split('}')[-1] == 'r' and not has_fldChar(c)]
            i = 0
            continue
        i += 1
        
    return converted

total_plain_converted = 0
for i, p in enumerate(doc.paragraphs):
    cnt = convert_plain_guillemets_in_p(p)
    total_plain_converted += cnt

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                cnt = convert_plain_guillemets_in_p(p)
                total_plain_converted += cnt

print(f"Step 2 completed: Converted {total_plain_converted} plain guillemets into real MERGEFIELDs.")

# Fix P172 duplicate field
p172 = doc.paragraphs[172]
p172_fields = []
curr = None
for i, c in enumerate(p172._element):
    for fc in c.iter():
        if fc.tag.split('}')[-1] == 'fldChar':
            fc_type = fc.get(f"{{{W_NS}}}fldCharType") or fc.get("w:fldCharType") or fc.get("fldCharType")
            if fc_type == 'begin':
                curr = [i]
            elif fc_type == 'end' and curr is not None:
                curr.append(i)
                p172_fields.append(curr)
                curr = None

if len(p172_fields) == 2:
    print("Fixing duplicate field in P172...")
    f2 = p172_fields[1]
    for idx in range(f2[1], f2[0] - 1, -1):
        elem = p172._element[idx]
        p172._element.remove(elem)
    print("P172 duplicate field removed successfully.")

# Save to test output
doc.save(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\temp_test\MAUTONG_HOP_standardized.docx")
print("Saved to temp_test/MAUTONG_HOP_standardized.docx successfully!")
