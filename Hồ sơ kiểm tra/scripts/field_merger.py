import docx
import re
import copy
from docx.oxml.ns import qn
from field_formatter import format_value

TAG_R = qn('w:r')
TAG_FLDCHAR = qn('w:fldChar')
TAG_INSTRTEXT = qn('w:instrText')
TAG_T = qn('w:t')
TAG_RPR = qn('w:rPr')
ATTR_FLDCHARTYPE = qn('w:fldCharType')

def parse_instr(instr_str):
    instr_clean = instr_str.strip()
    m = re.match(r'MERGEFIELD\s+([^\s\\]+)(.*)', instr_clean)
    if m:
        field_name = m.group(1).strip()
        switch = m.group(2).strip()
        return field_name, switch
    return instr_clean, ""

def merge_fields_in_element(parent_elm, data_dict):
    i = 0
    merged_count = 0
    
    while i < len(parent_elm):
        child = parent_elm[i]
        
        has_begin = False
        if child.tag == TAG_R:
            for sub in child:
                if sub.tag == TAG_FLDCHAR and sub.get(ATTR_FLDCHARTYPE) == 'begin':
                    has_begin = True
                    break
                    
        if has_begin:
            begin_idx = i
            sep_idx = -1
            end_idx = -1
            instr_texts = []
            result_runs = []
            
            j = i + 1
            while j < len(parent_elm):
                curr = parent_elm[j]
                if curr.tag == TAG_R:
                    for sub in curr:
                        if sub.tag == TAG_INSTRTEXT and sep_idx == -1:
                            if sub.text:
                                instr_texts.append(sub.text)
                        elif sub.tag == TAG_FLDCHAR:
                            ctype = sub.get(ATTR_FLDCHARTYPE)
                            if ctype == 'separate' and sep_idx == -1:
                                sep_idx = j
                            elif ctype == 'end':
                                end_idx = j
                                break
                    if end_idx != -1:
                        break
                    if sep_idx != -1 and j != sep_idx:
                        result_runs.append(curr)
                j += 1
                
            if end_idx != -1 and instr_texts:
                full_instr = "".join(instr_texts)
                field_name, switch = parse_instr(full_instr)
                
                val = data_dict.get(field_name, None)
                if val is None:
                    for k, v in data_dict.items():
                        if k.lower() == field_name.lower():
                            val = v
                            break
                            
                formatted_str = format_value(field_name, switch, val)
                
                # Create clean new run
                r_new = docx.oxml.OxmlElement('w:r')
                
                # Copy rPr from result_run if available, else from begin run
                source_run = result_runs[0] if result_runs else child
                rprs = [sub for sub in source_run if sub.tag == TAG_RPR]
                if rprs:
                    r_new.append(copy.deepcopy(rprs[0]))
                    
                t_elm = docx.oxml.OxmlElement('w:t')
                t_elm.text = formatted_str
                # preserve whitespace if string has spaces
                if formatted_str.startswith(' ') or formatted_str.endswith(' '):
                    t_elm.set(qn('xml:space'), 'preserve')
                r_new.append(t_elm)
                
                # Insert r_new at begin_idx
                parent_elm.insert(begin_idx, r_new)
                
                # Remove all original elements from begin_idx to end_idx
                num_to_remove = end_idx - begin_idx + 1
                for _ in range(num_to_remove):
                    parent_elm.remove(parent_elm[begin_idx + 1])
                    
                i = begin_idx + 1
                merged_count += 1
                continue
                
        i += 1
        
    return merged_count
