import openpyxl

wb = openpyxl.load_workbook(r'sources\DB NNT 03.06.2026.xlsx', data_only=True)
ws_nnt = wb['NNT']
ws_s1 = wb['Sheet1']

headers_nnt = [c.value for c in ws_nnt[3]]
headers_s1 = [c.value for c in ws_s1[2]]

with open(r'sources\4020.md', 'r', encoding='utf-8') as f:
    content = f.read()

rows_4020 = []
for line in content.splitlines():
    if '|' in line and not ':-:' in line and not 'Tên doanh nghiệp' in line and not '**TỔNG**' in line:
        parts = [p.strip() for p in line.split('|')[1:-1]]
        if len(parts) >= 8:
            rows_4020.append({
                'stt': parts[0],
                'name': parts[1],
                'mst': parts[2],
                'so_hd': parts[3],
                'ds_chua_thue': parts[4],
                'thue_vat': parts[5],
                'ds_co_thue': parts[6],
                'dia_chi': parts[7]
            })

# Load NNT map
map_nnt = {}
for r in range(4, ws_nnt.max_row+1):
    mst = str(ws_nnt.cell(r, 4).value or '').strip().replace('-', '')
    if mst.endswith('.0'): mst = mst[:-2]
    if mst:
        map_nnt[mst] = {headers_nnt[c-1]: ws_nnt.cell(r, c).value for c in range(1, len(headers_nnt)+1)}
        if mst.startswith('0') and len(mst) == 10:
            map_nnt[mst[1:]] = map_nnt[mst]
        elif len(mst) == 9:
            map_nnt['0' + mst] = map_nnt[mst]

# Load Sheet1 map
map_s1 = {}
for r in range(3, ws_s1.max_row+1):
    mst = str(ws_s1.cell(r, 1).value or '').strip().replace('-', '')
    if mst.endswith('.0'): mst = mst[:-2]
    if mst:
        map_s1[mst] = {headers_s1[c-1]: ws_s1.cell(r, c).value for c in range(1, len(headers_s1)+1)}
        if mst.startswith('0') and len(mst) == 10:
            map_s1[mst[1:]] = map_s1[mst]
        elif len(mst) == 9:
            map_s1['0' + mst] = map_s1[mst]

matched_26 = [2, 5, 6, 8, 9, 11, 12, 13, 15, 18, 19, 20, 21, 22, 23, 24, 27, 29, 30, 32, 35, 37, 38, 40, 41, 42]

for idx, item in enumerate(rows_4020):
    stt_int = int(item['stt'])
    if stt_int in matched_26:
        mst = item['mst'].strip().replace('-', '')
        d_nnt = map_nnt.get(mst) or map_nnt.get(mst[1:] if mst.startswith('0') else '0'+mst)
        d_s1 = map_s1.get(mst) or map_s1.get(mst[1:] if mst.startswith('0') else '0'+mst)
        
        st_dkt = d_nnt.get('Trạng thái ĐKT tổ chức') if d_nnt else None
        st_s1 = d_s1.get('Trạng thái') if d_s1 else None
        
        print(f"STT 4020: {item['stt']:>2} | {item['name'][:35]:<35} | MST: {item['mst']:<13} | Col 45 (Trạng thái ĐKT tổ chức): {repr(st_dkt)}")
