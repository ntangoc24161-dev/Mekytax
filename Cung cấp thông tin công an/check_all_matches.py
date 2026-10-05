import openpyxl

wb_nnt = openpyxl.load_workbook(r'sources\DB NNT 03.06.2026.xlsx', data_only=True)
ws_nnt = wb_nnt['NNT']
headers = [cell.value for cell in ws_nnt[3]]

xlsx_nnt = {}
for r in range(4, ws_nnt.max_row + 1):
    mst = str(ws_nnt.cell(r, 4).value or '').strip()
    if mst.endswith('.0'):
        mst = mst[:-2]
    mst_clean = mst.replace('-', '').strip()
    if mst_clean:
        row_dict = {headers[c-1]: ws_nnt.cell(r, c).value for c in range(1, len(headers)+1)}
        xlsx_nnt[mst_clean] = row_dict
        if mst_clean.startswith('0') and len(mst_clean) == 10:
            xlsx_nnt[mst_clean[1:]] = row_dict
        elif len(mst_clean) == 9:
            xlsx_nnt['0' + mst_clean] = row_dict

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
                'dia_chi': parts[7]
            })

bh_tcs3_matched = []
other_matched = []

for item in rows_4020:
    mst = item['mst'].strip().replace('-', '')
    nnt = xlsx_nnt.get(mst)
    if not nnt:
        if mst.startswith('0') and len(mst) == 10:
            nnt = xlsx_nnt.get(mst[1:])
        elif len(mst) == 9:
            nnt = xlsx_nnt.get('0' + mst)
    
    # Also check if address is Bien Hoa
    is_bh = 'Biên Hoà' in item['dia_chi'] or 'Biên Hòa' in item['dia_chi'] or (nnt is not None and ('Biên Hòa' in str(nnt.get('Tên Quận Huyện')) or 'Biên Hoà' in str(nnt.get('Tên Quận Huyện'))))
    
    if is_bh or nnt is not None:
        bh_tcs3_matched.append((item, nnt))
        print(f"TCS3 STT {item['stt']}: {item['name']} | MST: {item['mst']} | In NNT: {nnt is not None}")
    else:
        other_matched.append(item)
        print(f"OTHER STT {item['stt']}: {item['name']} | MST: {item['mst']} | Addr: {item['dia_chi']}")

print(f"\nTotal TCS3: {len(bh_tcs3_matched)}")
print(f"Total Other: {len(other_matched)}")
