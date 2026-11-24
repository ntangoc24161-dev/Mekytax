import openpyxl
import os
import re

excel_path = r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx'
wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
ws = wb['NHAP_HO_SO']

check_folders = [12, 2, 46, 48, 49, 54, 66, 68, 9, 74, 105, 159, 160, 172, 183, 187, 200, 205, 202, 249, 250, 252, 258, 259, 280]

excel_status = {}
for row in ws.iter_rows(min_row=5, values_only=True):
    try:
        f = int(float(str(row[1]).strip()))
    except:
        continue
    if f in check_folders:
        if f not in excel_status:
            excel_status[f] = []
        excel_status[f].append({
            'cv': row[2],
            'dn': row[10],
            'to_kt': row[12],
            'cb': row[13],
            'tt': row[14],
            'so_tl': row[22]
        })

for f in check_folders:
    rows = excel_status.get(f, [])
    print(f"Folder {f}: {len(rows)} entries in Excel")
    for r in rows[:2]:
        print(f"   CV: {r['cv']} | To: {r['to_kt']} | CB: {r['cb']} | TT: {r['tt']} | So TL: {r['so_tl']}")
