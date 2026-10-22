import re
import json
import time

t0 = time.time()
with open('distinct_mst.txt', 'r', encoding='utf-8') as f:
    needed_msts = set(line.strip() for line in f if line.strip())

print(f"Loaded {len(needed_msts)} MSTs to search.")

matched_nnt = {}

# Regex pattern to match MST data: <Data ss:Type="String">360...</Data>
pattern = re.compile(r'<Data ss:Type=[\'"]String[\'"]>([0-9]{10}(?:-[0-9]{3})?|[0-9]{13})</Data>')

with open(r'I:/My Drive/DB NNT.xls', 'r', encoding='utf-8', errors='ignore') as f:
    row_lines = []
    in_row = False
    in_nnt_sheet = False
    row_count = 0
    
    for line in f:
        if '<Worksheet ss:Name="NNT">' in line:
            in_nnt_sheet = True
        elif '<Worksheet' in line and in_nnt_sheet:
            break
            
        if not in_nnt_sheet:
            continue
            
        if '<Row' in line:
            in_row = True
            row_lines = [line]
            continue
            
        if in_row:
            row_lines.append(line)
            if '</Row>' in line:
                in_row = False
                row_count += 1
                row_str = ''.join(row_lines)
                
                # Fast regex check for any MST in this row
                found_msts = pattern.findall(row_str)
                hit = False
                for m in found_msts:
                    if m in needed_msts:
                        hit = True
                        break
                        
                if hit:
                    cells = re.findall(r'<Cell([^>]*)>(.*?)</Cell>', row_str, re.DOTALL)
                    col = 1
                    row_data = {}
                    for attrs, content in cells:
                        m_idx = re.search(r'ss:Index=[\'"](\d+)[\'"]', attrs)
                        if m_idx:
                            col = int(m_idx.group(1))
                        m_data = re.search(r'<Data[^>]*>(.*?)</Data>', content, re.DOTALL)
                        val = m_data.group(1).strip() if m_data else ''
                        row_data[col] = val
                        col += 1
                    
                    mst_val = row_data.get(4, '').replace(' ', '').replace('-', '')
                    # also check original with dash if any
                    orig_mst = row_data.get(4, '').strip()
                    target_key = orig_mst if orig_mst in needed_msts else mst_val
                    
                    if target_key in needed_msts or mst_val in needed_msts:
                        matched_nnt[target_key] = {
                            'mst': orig_mst,
                            'cqt': row_data.get(2, ''),
                            'ten_nnt': row_data.get(5, ''),
                            'dia_chi': row_data.get(8, ''),
                            'to_kt': row_data.get(9, ''),
                            'phuong_xa': row_data.get(10, ''),
                            'cbql': row_data.get(11, ''),
                            'quan_huyen': row_data.get(12, ''),
                            'tinh_tp': row_data.get(14, ''),
                            'trang_thai': row_data.get(45, '')
                        }
                row_lines = []

print(f"Matched {len(matched_nnt)} out of {len(needed_msts)} MSTs in DB NNT. Time: {time.time()-t0:.2f}s")
with open('db_nnt_matched.json', 'w', encoding='utf-8') as f:
    json.dump(matched_nnt, f, ensure_ascii=False, indent=2)

unmatched = needed_msts - set(matched_nnt.keys())
print(f"Unmatched MSTs count: {len(unmatched)}")
with open('unmatched_mst.txt', 'w', encoding='utf-8') as f:
    for u in sorted(unmatched):
        f.write(u + '\n')
