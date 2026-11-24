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
            'status': status
        })

dir_1 = r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026'
folders_on_disk = os.listdir(dir_1)
r_pattern = re.compile(r'(-R(\b|\s|$)|_R(\b|\s|$))', re.IGNORECASE)

# Group 1: Folders where Excel has SO CV TRA LOI and status 'Đã hoàn thành'
group1_excel_completed = []
# Group 2: Folders not completed in Excel, but have actual reply signed PDF inside disk folder
group2_file_signed = []
# Group 3: Pending issues
group3_issues = []

for d in folders_on_disk:
    full_path = os.path.join(dir_1, d)
    if not os.path.isdir(full_path):
        continue
    is_r = bool(r_pattern.search(d))
    m = re.match(r'^(\d+)', d)
    f_num = int(m.group(1)) if m else None
    
    excel_info = excel_replies.get(f_num, [])
    has_tl_in_excel = any(item['so_cv_tl'] for item in excel_info)
    all_completed_in_excel = len(excel_info) > 0 and all('hoàn thành' in str(item['status']).lower() for item in excel_info)
    
    # check files
    files = os.listdir(full_path)
    reply_files = [f for f in files if f.lower().endswith('.pdf') and any(k in f.lower() for k in ['tra loi', 'trả lời', 'signed', 'phuc dap', 'cung cap thong tin'])]
    actual_reply_files = [f for f in reply_files if not re.match(r'^(công văn|cv)\s+\d+.*(yc|anđt|cshs)', f, re.IGNORECASE)]

    if not is_r:
        if has_tl_in_excel:
            group1_excel_completed.append((d, f_num, [i['so_cv_tl'] for i in excel_info if i['so_cv_tl']], actual_reply_files))
        elif actual_reply_files:
            group2_file_signed.append((d, f_num, actual_reply_files))

print(f"--- NHOM 1: Co so CV tra loi tren Excel ({len(group1_excel_completed)} thu muc) ---")
for d, f_num, tls, files in group1_excel_completed:
    print(f"Folder: {d} | CV TL: {tls} | Files tren o dia: {files}")

print(f"\n--- NHOM 2: Chua co so tren Excel nhung co file tra loi/signed tren o dia ({len(group2_file_signed)} thu muc) ---")
for d, f_num, files in group2_file_signed:
    print(f"Folder: {d} | Files: {files}")
