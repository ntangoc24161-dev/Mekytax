import openpyxl
import os
import re

excel_path = r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx'
wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

excel_replies = {}

ws = wb['NHAP_HO_SO']
for row in ws.iter_rows(min_row=5, values_only=True):
    folder = row[1]
    so_cv_tl = row[22]
    ngay_tl = row[23]
    status = row[14]
    can_bo = row[13]
    so_cv = row[2]
    if folder is not None:
        try:
            f_num = int(float(str(folder).strip()))
        except:
            f_num = str(folder).strip()
        if f_num not in excel_replies:
            excel_replies[f_num] = []
        excel_replies[f_num].append({
            'so_cv': so_cv,
            'so_cv_tl': so_cv_tl,
            'ngay_tl': ngay_tl,
            'status': status,
            'can_bo': can_bo
        })

dir_1 = r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026'
folders_on_disk = os.listdir(dir_1)

r_pattern = re.compile(r'(-R(\b|\s|$)|_R(\b|\s|$))', re.IGNORECASE)

matched_to_rename = []
already_r = []

for d in folders_on_disk:
    full_path = os.path.join(dir_1, d)
    if not os.path.isdir(full_path):
        continue
    is_r = bool(r_pattern.search(d))
    m = re.match(r'^(\d+)', d)
    f_num = int(m.group(1)) if m else None
    
    files = os.listdir(full_path)
    reply_files = [f for f in files if f.lower().endswith('.pdf') and any(k in f.lower() for k in ['tra loi', 'trả lời', 'signed', 'phuc dap', 'cung cap thong tin', 'cung cap'])]
    actual_reply_files = [f for f in reply_files if not re.match(r'^(công văn|cv)\s+\d+.*(yc|anđt|cshs)', f, re.IGNORECASE)]
    
    excel_info = excel_replies.get(f_num, [])
    has_tl_in_excel = any(item['so_cv_tl'] or 'hoàn thành' in str(item['status']).lower() for item in excel_info)
    
    if is_r:
        already_r.append(d)
    else:
        if has_tl_in_excel or actual_reply_files:
            matched_to_rename.append({
                'dir': d,
                'f_num': f_num,
                'excel_tl': [(i['so_cv_tl'], i['ngay_tl'], i['status']) for i in excel_info if i['so_cv_tl']],
                'reply_files': reply_files
            })

print(f"Total folders on disk: {len(folders_on_disk)}")
print(f"Already has -R: {len(already_r)}")
print(f"Candidates to check: {len(matched_to_rename)}")
for c in matched_to_rename:
    print("-------------------------------------------")
    print(f"Folder: {c['dir']}")
    print(f"  Excel TL: {c['excel_tl']}")
    print(f"  Files: {c['reply_files']}")
