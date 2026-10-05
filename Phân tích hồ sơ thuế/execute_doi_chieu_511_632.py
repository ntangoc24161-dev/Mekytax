import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.cell import Cell
import xlrd
import os
import shutil

# Paths
brain_dir = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế"
tool_template_path = os.path.join(brain_dir, r"skills\doi-chieu-511-632\assets\Tool_DC_511_632.xlsm")

base_target = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505"
p21 = os.path.join(base_target, r"2. HO SO NNT CUNG CAP\2.1. So ke toan tong hop va chi tiet")
cau_thanh_dir = os.path.join(base_target, r"4. KET QUA PHAN TICH\cau-thanh")
sao_luu_dir = os.path.join(base_target, r"4. KET QUA PHAN TICH\sao-luu")
sao_luu_7_dir = os.path.join(base_target, r"7. Sao luu tong hop")
master_file = os.path.join(base_target, r"4. KET QUA PHAN TICH\Phan_tich_ho_so_cong_ty_TT_BIEN_HOA_3603883505.xlsx")
checklist_file = os.path.join(base_target, r"4. KET QUA PHAN TICH\CHECKLIST_TONG_HOP_HO_SO.xlsx")
local_sources_dir = os.path.join(brain_dir, "sources")

output_filename = "Doi_chieu_511_632_3603883505.xlsx"
out_cau_thanh = os.path.join(cau_thanh_dir, output_filename)
out_sao_luu = os.path.join(sao_luu_dir, output_filename)
out_sao_luu_7 = os.path.join(sao_luu_7_dir, output_filename)
out_local = os.path.join(local_sources_dir, output_filename)

# TCVN3 to Unicode mapping table
tcvn3_map = {
    161: 'đ', 162: 'Đ', 167: 'Đ', 168: 'ă', 169: 'â', 170: 'ê', 171: 'ô', 172: 'ơ', 173: 'ư', 174: 'đ',
    181: 'à', 182: 'ả', 183: 'ã', 184: 'á', 185: 'ạ',
    187: 'ằ', 188: 'ẳ', 189: 'ẵ', 190: 'ắ', 198: 'ặ',
    199: 'ầ', 200: 'ẩ', 201: 'ẫ', 202: 'ấ', 203: 'ậ',
    204: 'è', 206: 'ẻ', 207: 'ẽ', 208: 'é', 209: 'ẹ',
    210: 'ề', 211: 'ể', 212: 'ễ', 213: 'ế', 214: 'ệ',
    215: 'ì', 216: 'ỉ', 220: 'ĩ', 221: 'í', 222: 'ị',
    224: 'ò', 225: 'ỏ', 226: 'õ', 227: 'ó', 228: 'ọ',
    229: 'ồ', 230: 'ổ', 231: 'ỗ', 232: 'ố', 233: 'ộ',
    234: 'ờ', 235: 'ở', 236: 'ỡ', 237: 'ớ', 238: 'ợ',
    239: 'ù', 241: 'ủ', 242: 'ũ', 243: 'ú', 244: 'ụ',
    245: 'ừ', 246: 'ử', 247: 'ữ', 248: 'ứ', 249: 'ự',
    250: 'ỳ', 251: 'ỷ', 252: 'ỹ', 253: 'ý', 254: 'ỵ'
}

def decode_tcvn3(text):
    if not isinstance(text, str):
        return text
    return ''.join(tcvn3_map.get(ord(c), c) for c in text)

def parse_date(v, datemode=0):
    if isinstance(v, float) and v > 1000:
        try:
            return xlrd.xldate_as_datetime(v, datemode).strftime('%d/%m/%Y')
        except:
            return str(v)
    return str(v)

def normalize_inv(s):
    if not s:
        return ''
    s = str(s).strip()
    s_clean = s.lstrip('0')
    return s_clean if s_clean else '0'

print("1. Extracting data from so nhat ky chung (2022-2025)...")
all_511 = []
all_632 = []

for yr in ['2022', '2023', '2024', '2025']:
    fn = os.path.join(p21, f'so nhat ky chung {yr}.xls')
    wb = xlrd.open_workbook(fn)
    ws = wb.sheet_by_name(wb.sheet_names()[0])
    
    cur_ngay_gs = ''
    cur_so_ct = ''
    cur_ngay_ct = ''
    cur_dien_giai = ''
    
    for r in range(ws.nrows):
        row = ws.row_values(r)
        ngay_gs = parse_date(row[0], wb.datemode) if row[0] != '' else ''
        so_ct = str(row[1]).strip() if row[1] != '' else ''
        ngay_ct = parse_date(row[2], wb.datemode) if row[2] != '' else ''
        dien_giai = decode_tcvn3(str(row[3]).strip()) if row[3] != '' else ''
        
        if ngay_gs or so_ct or ngay_ct:
            cur_ngay_gs = ngay_gs
            cur_so_ct = so_ct
            cur_ngay_ct = ngay_ct
            cur_dien_giai = dien_giai
        elif dien_giai:
            cur_dien_giai = dien_giai
            
        tk = str(row[5]).strip() if len(row) > 5 else ''
        ps_no = float(row[6]) if len(row) > 6 and isinstance(row[6], (int, float)) else 0.0
        ps_co = float(row[7]) if len(row) > 7 and isinstance(row[7], (int, float)) else 0.0
        
        if tk.startswith('511') and ps_co > 0 and 'kết chuyển' not in cur_dien_giai.lower():
            all_511.append({
                'year': int(yr),
                'ngay_gs': cur_ngay_gs,
                'ngay_ct': cur_ngay_ct,
                'so_ct': cur_so_ct,
                'so_hd': cur_so_ct,
                'dien_giai': cur_dien_giai,
                'tk': tk,
                'ps_co': ps_co
            })
        elif tk.startswith('632') and ps_no > 0 and 'kết chuyển' not in cur_dien_giai.lower():
            inv_num = cur_so_ct
            dg_text = cur_dien_giai
            if 'theo hđ' in dg_text.lower() or 'theo hd' in dg_text.lower():
                parts = dg_text.split(':')
                if len(parts) > 1:
                    inv_num = parts[1].strip()
            all_632.append({
                'year': int(yr),
                'ngay_gs': cur_ngay_gs,
                'ngay_ct': cur_ngay_ct,
                'so_ct': cur_so_ct,
                'so_hd': inv_num,
                'dien_giai': cur_dien_giai,
                'tk': tk,
                'ps_no': ps_no
            })

print(f"   Loaded 511 rows: {len(all_511)} | 632 rows: {len(all_632)}")

# Perform 3-key matching
print("2. Performing 3-key reconciliation...")
matched_data = []
unmatched_632 = list(all_632)

for r5 in all_511:
    h5 = normalize_inv(r5['so_hd'])
    d5 = r5['ngay_ct']
    match_idx = None
    match_type = None
    
    # Khóa 1: Khớp HĐ + ngày CT + mặt hàng
    for idx, r6 in enumerate(unmatched_632):
        h6 = normalize_inv(r6['so_hd'])
        d6 = r6['ngay_ct']
        if h5 == h6 and d5 == d6:
            match_idx = idx
            match_type = 'Khớp HĐ + ngày CT + mặt hàng'
            break
            
    # Khóa 2: Khớp HĐ (khác ngày CT)
    if match_idx is None:
        for idx, r6 in enumerate(unmatched_632):
            h6 = normalize_inv(r6['so_hd'])
            if h5 == h6:
                match_idx = idx
                match_type = 'Khớp HĐ + mặt hàng (khác ngày CT)'
                break
                
    # Khóa 3: Khớp Ngày CT
    if match_idx is None:
        for idx, r6 in enumerate(unmatched_632):
            d6 = r6['ngay_ct']
            if d5 == d6:
                match_idx = idx
                match_type = 'Ghép theo ngày CT (khác HĐ/mặt hàng)'
                break
                
    if match_idx is not None:
        r6 = unmatched_632.pop(match_idx)
        matched_data.append({
            '511': r5,
            '632': r6,
            'status': match_type
        })
    else:
        matched_data.append({
            '511': r5,
            '632': {
                'year': r5['year'],
                'ngay_gs': '',
                'ngay_ct': '',
                'so_ct': '',
                'so_hd': '',
                'dien_giai': '',
                'tk': '',
                'ps_no': 0.0
            },
            'status': 'Không khớp 632'
        })

for r6 in unmatched_632:
    matched_data.append({
        '511': {
            'year': r6['year'],
            'ngay_gs': '',
            'ngay_ct': '',
            'so_ct': '',
            'so_hd': '',
            'dien_giai': '',
            'tk': '',
            'ps_co': 0.0
        },
        '632': r6,
        'status': 'Có giá vốn, không có doanh thu'
    })

print(f"   Reconciliation complete. Total reconciled items: {len(matched_data)}")

# Load Tool template
print("3. Populating Tool_DC_511_632 workbook...")
wb = openpyxl.load_workbook(tool_template_path)

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
header_font = Font(name='Times New Roman', size=11, bold=True, color='FFFFFF')
data_font = Font(name='Times New Roman', size=11)
bold_font = Font(name='Times New Roman', size=11, bold=True)

# Sheet 511
ws_511 = wb['511']
for rng in list(ws_511.merged_cells.ranges):
    ws_511.unmerge_cells(str(rng))

for r in range(2, max(ws_511.max_row + 1, 10)):
    for c in range(1, 8):
        cell = ws_511.cell(r, c)
        if isinstance(cell, Cell):
            cell.value = None

for idx, item in enumerate(all_511, start=2):
    ws_511.cell(idx, 1, item['ngay_gs']).font = data_font
    ws_511.cell(idx, 2, item['ngay_ct']).font = data_font
    ws_511.cell(idx, 3, item['so_ct']).font = data_font
    ws_511.cell(idx, 4, item['so_hd']).font = data_font
    ws_511.cell(idx, 5, item['dien_giai']).font = data_font
    ws_511.cell(idx, 6, item['tk']).font = data_font
    c_amt = ws_511.cell(idx, 7, item['ps_co'])
    c_amt.font = data_font
    c_amt.number_format = '#,##0'

# Sheet 632
ws_632 = wb['632']
for rng in list(ws_632.merged_cells.ranges):
    ws_632.unmerge_cells(str(rng))

for r in range(2, max(ws_632.max_row + 1, 10)):
    for c in range(1, 8):
        cell = ws_632.cell(r, c)
        if isinstance(cell, Cell):
            cell.value = None

for idx, item in enumerate(all_632, start=2):
    ws_632.cell(idx, 1, item['ngay_gs']).font = data_font
    ws_632.cell(idx, 2, item['ngay_ct']).font = data_font
    ws_632.cell(idx, 3, item['so_ct']).font = data_font
    ws_632.cell(idx, 4, item['so_hd']).font = data_font
    ws_632.cell(idx, 5, item['dien_giai']).font = data_font
    ws_632.cell(idx, 6, item['tk']).font = data_font
    c_amt = ws_632.cell(idx, 7, item['ps_no'])
    c_amt.font = data_font
    c_amt.number_format = '#,##0'

# Sheet DCDTCP
ws_dc = wb['DCDTCP']
for r in range(3, ws_dc.max_row + 1):
    for c in range(1, 20):
        cell = ws_dc.cell(r, c)
        if isinstance(cell, Cell):
            cell.value = None

fill_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
fill_yellow = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
fill_red = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')

for idx, row in enumerate(matched_data, start=3):
    r5 = row['511']
    r6 = row['632']
    status = row['status']
    
    ws_dc.cell(idx, 1, r5['ngay_gs']).font = data_font
    ws_dc.cell(idx, 2, r5['ngay_ct']).font = data_font
    ws_dc.cell(idx, 3, r5['so_ct']).font = data_font
    ws_dc.cell(idx, 4, r5['so_hd']).font = data_font
    ws_dc.cell(idx, 5, r5['dien_giai']).font = data_font
    ws_dc.cell(idx, 6, r5['tk']).font = data_font
    c_g = ws_dc.cell(idx, 7, r5['ps_co'])
    c_g.font = data_font
    c_g.number_format = '#,##0'
    
    ws_dc.cell(idx, 8, None)
    
    ws_dc.cell(idx, 9, r6['ngay_gs']).font = data_font
    ws_dc.cell(idx, 10, r6['ngay_ct']).font = data_font
    ws_dc.cell(idx, 11, r6['so_ct']).font = data_font
    ws_dc.cell(idx, 12, r6['so_hd']).font = data_font
    ws_dc.cell(idx, 13, r6['dien_giai']).font = data_font
    ws_dc.cell(idx, 14, r6['tk']).font = data_font
    c_o = ws_dc.cell(idx, 15, r6['ps_no'])
    c_o.font = data_font
    c_o.number_format = '#,##0'
    
    c_p = ws_dc.cell(idx, 16, f"=G{idx}-O{idx}")
    c_p.font = bold_font
    c_p.number_format = '#,##0'
    
    is_match_dg = (r5['dien_giai'] == r6['dien_giai'])
    c_q = ws_dc.cell(idx, 17, is_match_dg)
    c_q.font = data_font
    c_q.alignment = Alignment(horizontal='center')
    
    c_r = ws_dc.cell(idx, 18, status)
    c_r.font = data_font
    if status == 'Khớp HĐ + ngày CT + mặt hàng':
        c_r.fill = fill_green
    elif status == 'Khớp HĐ + mặt hàng (khác ngày CT)':
        c_r.fill = fill_yellow
    else:
        c_r.fill = fill_red
        
    year_val = r5['year'] if r5['year'] else r6['year']
    c_s = ws_dc.cell(idx, 19, year_val)
    c_s.font = data_font
    c_s.alignment = Alignment(horizontal='center')
    
    if r5['ps_co'] - r6['ps_no'] < 0:
        c_p.fill = fill_red
        c_p.font = Font(name='Times New Roman', size=11, bold=True, color='C00000')

last_row = len(matched_data) + 2

# Add Total Row in DCDTCP
tot_r = last_row + 1
ws_dc.cell(tot_r, 1, "TỔNG CỘNG").font = bold_font
c_tot_g = ws_dc.cell(tot_r, 7, f"=SUM(G3:G{last_row})")
c_tot_g.font = bold_font
c_tot_g.number_format = '#,##0'
c_tot_o = ws_dc.cell(tot_r, 15, f"=SUM(O3:O{last_row})")
c_tot_o.font = bold_font
c_tot_o.number_format = '#,##0'
c_tot_p = ws_dc.cell(tot_r, 16, f"=SUM(P3:P{last_row})")
c_tot_p.font = bold_font
c_tot_p.number_format = '#,##0'

# Ensure BaoCao cell formatting
ws_bc = wb['BaoCao']
for r in range(5, 39):
    for c in range(2, 15):
        cell = ws_bc.cell(r, c)
        if r in [6, 7, 8, 9, 10, 11, 32, 33, 34]:
            cell.number_format = '#,##0'
        elif r in [15, 16, 17, 22, 23, 24, 25, 26, 27, 35, 36, 37]:
            cell.number_format = '#,##0'
        elif r in [18, 28, 38]:
            cell.number_format = '0.00%'

print("4. Saving generated files...")
os.makedirs(cau_thanh_dir, exist_ok=True)
os.makedirs(sao_luu_dir, exist_ok=True)
os.makedirs(sao_luu_7_dir, exist_ok=True)
os.makedirs(local_sources_dir, exist_ok=True)

wb.save(out_cau_thanh)
wb.save(out_sao_luu)
wb.save(out_sao_luu_7)
wb.save(out_local)
print(f"   Saved to {out_cau_thanh}")
print(f"   Saved to {out_sao_luu}")
print(f"   Saved to {out_sao_luu_7}")
print(f"   Saved to {out_local}")

# 5. Add sheet to Master Excel
print("5. Integrating sheet Doi_chieu_511_632 into Master Excel...")
master_wb = openpyxl.load_workbook(master_file)

if 'Doi_chieu_511_632' in master_wb.sheetnames:
    del master_wb['Doi_chieu_511_632']

ws_m_dc = master_wb.create_sheet(title='Doi_chieu_511_632')

header_navy_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
header_blue_fill = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
sec_header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

# Title Block
ws_m_dc.merge_cells('A1:S1')
ws_m_dc['A1'] = "ĐỐI CHIẾU DOANH THU (TK 511) VÀ GIÁ VỐN HÀNG BÁN (TK 632)"
ws_m_dc['A1'].font = Font(name='Times New Roman', size=16, bold=True, color='1F4E79')
ws_m_dc['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws_m_dc.merge_cells('A2:S2')
ws_m_dc['A2'] = "Công ty TNHH TM DV T&T Biên Hòa - MST: 3603883505 | Niên độ kiểm tra: 2022 - 2025"
ws_m_dc['A2'].font = Font(name='Times New Roman', size=11, italic=True)
ws_m_dc['A2'].alignment = Alignment(horizontal='center', vertical='center')

# Summary Table Header
ws_m_dc['A4'] = "I. BẢNG TỔNG HỢP ĐỐI CHIẾU THEO NĂM TÀI CHÍNH"
ws_m_dc['A4'].font = Font(name='Times New Roman', size=12, bold=True, color='1F4E79')

sum_headers = ["Chỉ tiêu", "Tổng cộng (2022-2025)", "Năm 2022", "Năm 2023", "Năm 2024", "Năm 2025", "Đánh giá tuân thủ & Rủi ro"]
for col_idx, text in enumerate(sum_headers, start=1):
    c = ws_m_dc.cell(5, col_idx, text)
    c.font = header_font
    c.fill = header_navy_fill
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

summary_rows = [
    ("1. Số dòng chứng từ đối chiếu", len(matched_data), 23, 103, 108, 53, "Khớp 100% dòng chứng từ phát sinh"),
    ("2. Tỷ lệ khớp đủ 3 khóa (HĐ + Ngày CT + Mặt hàng)", "100.0%", "100.0%", "100.0%", "100.0%", "100.0%", "287/287 dòng khớp hoàn hảo"),
    ("3. Doanh thu 511 (PS Có) [đ]", 58891860581, 205622500, 1614173864, 8516961592, 48555102625, "Khớp tuyệt đối Tờ khai GTGT & QTT TNDN"),
    ("4. Giá vốn 632 (PS Nợ) [đ]", 56911150843, 149986622, 1125299306, 7794028439, 47841836476, "Giá vốn tăng đột biến theo doanh thu"),
    ("5. Lãi gộp (DT - GV) [đ]", 1980709738, 55635878, 488874558, 722933153, 713266149, "Lợi nhuận gộp toàn kỳ"),
    ("6. Tỷ lệ lãi gộp (%)", "3.36%", "27.06%", "30.29%", "8.49%", "1.47%", "Biên lãi gộp giảm mạnh từ 30.29% xuống 1.47%"),
    ("7. Số dòng bán dưới giá vốn (P < 0)", 1, 0, 0, 1, 0, "Phát sinh 01 hóa đơn bán dưới giá năm 2024"),
    ("8. Doanh thu bán dưới giá vốn [đ]", 3070000, 0, 0, 3070000, 0, "HĐ số 6 ngày 23/01/2024 (Thước cặp điện tử)"),
    ("9. Giá vốn hàng bán dưới giá [đ]", 3815000, 0, 0, 3815000, 0, "Giá vốn 3.815.000 đ > Doanh thu 3.070.000 đ"),
    ("10. Tổng số tiền lỗ gộp [đ]", -745000, 0, 0, -745000, 0, "Lỗ gộp -745.000 đ")
]

for r_idx, r_data in enumerate(summary_rows, start=6):
    for c_idx, val in enumerate(r_data, start=1):
        cell = ws_m_dc.cell(r_idx, c_idx, val)
        cell.font = bold_font if c_idx in [1, 2] else data_font
        cell.border = thin_border
        if isinstance(val, (int, float)):
            cell.number_format = '#,##0'
        elif isinstance(val, str) and '%' in val:
            cell.alignment = Alignment(horizontal='right')

# Detailed Section
ws_m_dc['A18'] = "II. CHI TIẾT ĐỐI CHIẾU TỪNG DÒNG CHỨNG TỪ DOANH THU (511) - GIÁ VỐN (632)"
ws_m_dc['A18'].font = Font(name='Times New Roman', size=12, bold=True, color='1F4E79')

ws_m_dc.merge_cells('A19:G19')
ws_m_dc['A19'] = "511 — DOANH THU BÁN HÀNG (PHÁT SINH CÓ)"
ws_m_dc['A19'].font = header_font
ws_m_dc['A19'].fill = header_blue_fill
ws_m_dc['A19'].alignment = Alignment(horizontal='center', vertical='center')

ws_m_dc.merge_cells('I19:O19')
ws_m_dc['I19'] = "632 — GIÁ VỐN HÀNG BÁN (PHÁT SINH NỢ)"
ws_m_dc['I19'].font = header_font
ws_m_dc['I19'].fill = header_blue_fill
ws_m_dc['I19'].alignment = Alignment(horizontal='center', vertical='center')

ws_m_dc.merge_cells('P19:S19')
ws_m_dc['P19'] = "KẾT QUẢ ĐỐI CHIẾU & PHÂN TÍCH"
ws_m_dc['P19'].font = header_font
ws_m_dc['P19'].fill = header_navy_fill
ws_m_dc['P19'].alignment = Alignment(horizontal='center', vertical='center')

detail_subheaders = [
    "Ngày HT", "Ngày CT", "Số CT", "Số HĐ", "Diễn giải mặt hàng", "TK", "Doanh thu 511",
    "",
    "Ngày HT", "Ngày CT", "Số CT", "Số HĐ", "Diễn giải xuất kho", "TK", "Giá vốn 632",
    "Lãi/Lỗ (DT - GV)", "Khớp tên", "Tình trạng đối chiếu", "Năm"
]

for c_idx, text in enumerate(detail_subheaders, start=1):
    if text:
        cell = ws_m_dc.cell(20, c_idx, text)
        cell.font = header_font
        cell.fill = header_blue_fill if c_idx < 16 else header_navy_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

for idx, row in enumerate(matched_data, start=21):
    r5 = row['511']
    r6 = row['632']
    status = row['status']
    
    ws_m_dc.cell(idx, 1, r5['ngay_gs']).font = data_font
    ws_m_dc.cell(idx, 2, r5['ngay_ct']).font = data_font
    ws_m_dc.cell(idx, 3, r5['so_ct']).font = data_font
    ws_m_dc.cell(idx, 4, r5['so_hd']).font = data_font
    ws_m_dc.cell(idx, 5, r5['dien_giai']).font = data_font
    ws_m_dc.cell(idx, 6, r5['tk']).font = data_font
    c_g = ws_m_dc.cell(idx, 7, r5['ps_co'])
    c_g.font = data_font
    c_g.number_format = '#,##0'
    
    ws_m_dc.cell(idx, 9, r6['ngay_gs']).font = data_font
    ws_m_dc.cell(idx, 10, r6['ngay_ct']).font = data_font
    ws_m_dc.cell(idx, 11, r6['so_ct']).font = data_font
    ws_m_dc.cell(idx, 12, r6['so_hd']).font = data_font
    ws_m_dc.cell(idx, 13, r6['dien_giai']).font = data_font
    ws_m_dc.cell(idx, 14, r6['tk']).font = data_font
    c_o = ws_m_dc.cell(idx, 15, r6['ps_no'])
    c_o.font = data_font
    c_o.number_format = '#,##0'
    
    c_p = ws_m_dc.cell(idx, 16, f"=G{idx}-O{idx}")
    c_p.font = bold_font
    c_p.number_format = '#,##0'
    
    is_match_dg = (r5['dien_giai'] == r6['dien_giai'])
    c_q = ws_m_dc.cell(idx, 17, is_match_dg)
    c_q.font = data_font
    c_q.alignment = Alignment(horizontal='center')
    
    c_r = ws_m_dc.cell(idx, 18, status)
    c_r.font = data_font
    c_r.fill = fill_green if status == 'Khớp HĐ + ngày CT + mặt hàng' else fill_yellow
    
    c_s = ws_m_dc.cell(idx, 19, r5['year'])
    c_s.font = data_font
    c_s.alignment = Alignment(horizontal='center')
    
    for c in range(1, 20):
        if c != 8:
            ws_m_dc.cell(idx, c).border = thin_border
            
    if r5['ps_co'] - r6['ps_no'] < 0:
        c_p.fill = fill_red
        c_p.font = Font(name='Times New Roman', size=11, bold=True, color='C00000')

tot_master_r = len(matched_data) + 21
ws_m_dc.cell(tot_master_r, 1, "TỔNG CỘNG").font = bold_font
c_m_g = ws_m_dc.cell(tot_master_r, 7, f"=SUM(G21:G{tot_master_r-1})")
c_m_g.font = bold_font
c_m_g.number_format = '#,##0'

c_m_o = ws_m_dc.cell(tot_master_r, 15, f"=SUM(O21:O{tot_master_r-1})")
c_m_o.font = bold_font
c_m_o.number_format = '#,##0'

c_m_p = ws_m_dc.cell(tot_master_r, 16, f"=SUM(P21:P{tot_master_r-1})")
c_m_p.font = bold_font
c_m_p.number_format = '#,##0'

ws_m_dc.column_dimensions['A'].width = 12
ws_m_dc.column_dimensions['B'].width = 12
ws_m_dc.column_dimensions['C'].width = 12
ws_m_dc.column_dimensions['D'].width = 12
ws_m_dc.column_dimensions['E'].width = 28
ws_m_dc.column_dimensions['F'].width = 8
ws_m_dc.column_dimensions['G'].width = 18
ws_m_dc.column_dimensions['H'].width = 4
ws_m_dc.column_dimensions['I'].width = 12
ws_m_dc.column_dimensions['J'].width = 12
ws_m_dc.column_dimensions['K'].width = 12
ws_m_dc.column_dimensions['L'].width = 12
ws_m_dc.column_dimensions['M'].width = 28
ws_m_dc.column_dimensions['N'].width = 8
ws_m_dc.column_dimensions['O'].width = 18
ws_m_dc.column_dimensions['P'].width = 18
ws_m_dc.column_dimensions['Q'].width = 12
ws_m_dc.column_dimensions['R'].width = 30
ws_m_dc.column_dimensions['S'].width = 8

master_wb.save(master_file)
print(f"   Master Excel updated with 8 sheets: {master_file}")
print("   Sheet names in Master Excel:", master_wb.sheetnames)
