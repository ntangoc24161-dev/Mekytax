import xlrd, openpyxl, os, datetime
from collections import Counter

p21 = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\2. HO SO NNT CUNG CAP\2.1. So ke toan tong hop va chi tiet'

tcvn3_dict = {
    0xb5: 'à', 0xb8: 'á', 0xb6: 'ả', 0xb7: 'ã', 0xb9: 'ạ',
    0xa8: 'ă', 0xbb: 'ằ', 0xbe: 'ắ', 0xbc: 'ẳ', 0xbd: 'ẵ', 0xc6: 'ặ',
    0xa9: 'â', 0xc7: 'ầ', 0xcb: 'ấ', 0xc8: 'ẩ', 0xc9: 'ẫ', 0xcf: 'ậ',
    0xe1: 'è', 0xe3: 'é', 0xe2: 'ẻ', 0xe5: 'ẽ', 0xe4: 'ẹ',
    0xaa: 'ê', 0xe6: 'ề', 0xe9: 'ế', 0xe7: 'ể', 0xe8: 'ễ', 0xed: 'ệ',
    0xef: 'ì', 0xf3: 'í', 0xf1: 'ỉ', 0xf2: 'ĩ', 0xf4: 'ị',
    0xf5: 'ò', 0xf8: 'ó', 0xf6: 'ỏ', 0xf7: 'õ', 0xf9: 'ọ',
    0xab: 'ô', 0xfb: 'ồ', 0xfa: 'ố', 0xfc: 'ổ', 0xfe: 'ỗ', 0xff: 'ộ',
    0xac: 'ơ', 0xd2: 'ờ', 0xd5: 'ớ', 0xd3: 'ở', 0xd4: 'ỡ', 0xd6: 'ợ',
    0xd7: 'ù', 0xd9: 'ú', 0xd8: 'ủ', 0xdc: 'ũ', 0xda: 'ụ',
    0xad: 'ư', 0xdd: 'ừ', 0xe0: 'ứ', 0xde: 'ử', 0xdf: 'ữ', 0xea: 'ự',
    0xf7: 'ỳ', 0xfb: 'ý', 0xf8: 'ỷ', 0xf9: 'ỹ', 0xfa: 'ỵ',
    0xae: 'đ', 0xa1: 'đ', 0xa2: 'Đ',
    0x80: 'À', 0x81: 'Á', 0x82: 'Ả', 0x83: 'Ã', 0x84: 'Ạ',
    0x85: 'Ă', 0x86: 'Ằ', 0x87: 'Ắ', 0x88: 'Ẳ', 0x89: 'Ẵ', 0x8a: 'Ặ',
    0x8b: 'Â', 0x8c: 'Ầ', 0x8d: 'Ấ', 0x8e: 'Ẩ', 0x8f: 'Ẫ', 0x90: 'Ậ',
    0x91: 'È', 0x92: 'É', 0x93: 'Ẻ', 0x94: 'Ẽ', 0x95: 'Ẹ',
    0x96: 'Ê', 0x97: 'Ề', 0x98: 'Ế', 0x99: 'Ể', 0x9a: 'Ễ', 0x9b: 'Ệ',
    0x9c: 'Ì', 0x9d: 'Í', 0x9e: 'Ỉ', 0x9f: 'Ĩ', 0xa0: 'Ị',
    0xc0: 'Ò', 0xc1: 'Ó', 0xc2: 'Ỏ', 0xc3: 'Õ', 0xc4: 'Ọ',
    0xc5: 'Ô', 0xca: 'Ồ', 0xcb: 'Ố', 0xcc: 'Ổ', 0xcd: 'Ỗ', 0xce: 'Ộ',
    0xd0: 'Ơ', 0xd1: 'Ờ', 0xd2: 'Ớ', 0xd3: 'Ở', 0xd4: 'Ỡ', 0xd5: 'Ợ',
    0xd6: 'Ù', 0xd7: 'Ú', 0xd8: 'Ủ', 0xd9: 'Ũ', 0xda: 'Ụ',
    0xdb: 'Ư', 0xdc: 'Ừ', 0xdd: 'Ứ', 0xde: 'Ử', 0xdf: 'Ữ', 0xea: 'Ự',
    0xe1: 'Ỳ', 0xe2: 'Ý', 0xe3: 'Ỷ', 0xe4: 'Ỹ', 0xe5: 'Ỵ',
}

def to_uni(text):
    if not isinstance(text, str):
        return text
    return ''.join(tcvn3_dict.get(ord(c), c) for c in text)

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
        dien_giai = to_uni(str(row[3]).strip()) if row[3] != '' else ''
        
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
        
        if tk.startswith('511') and ps_co > 0 and 'kết chuyển' not in cur_dien_giai.lower() and 'kÕt chuyÓn' not in cur_dien_giai.lower():
            all_511.append({
                'year': yr,
                'ngay_gs': cur_ngay_gs,
                'ngay_ct': cur_ngay_ct,
                'so_ct': cur_so_ct,
                'so_hd': cur_so_ct,
                'dien_giai': cur_dien_giai,
                'tk': tk,
                'ps_co': ps_co
            })
        elif tk.startswith('632') and ps_no > 0 and 'kết chuyển' not in cur_dien_giai.lower() and 'kÕt chuyÓn' not in cur_dien_giai.lower():
            inv_num = cur_so_ct
            dg_text = cur_dien_giai
            if 'theo hđ' in dg_text.lower() or 'theo hd' in dg_text.lower():
                parts = dg_text.split(':')
                if len(parts) > 1:
                    inv_num = parts[1].strip()
            all_632.append({
                'year': yr,
                'ngay_gs': cur_ngay_gs,
                'ngay_ct': cur_ngay_ct,
                'so_ct': cur_so_ct,
                'so_hd': inv_num,
                'dien_giai': cur_dien_giai,
                'tk': tk,
                'ps_no': ps_no
            })

# Match
matched = []
unmatched_632 = list(all_632)

for r5 in all_511:
    h5 = normalize_inv(r5['so_hd'])
    d5 = r5['ngay_ct']
    match_idx = None
    match_type = None
    
    for idx, r6 in enumerate(unmatched_632):
        h6 = normalize_inv(r6['so_hd'])
        d6 = r6['ngay_ct']
        if h5 == h6 and d5 == d6:
            match_idx = idx
            match_type = 'Khớp HĐ + ngày CT + mặt hàng'
            break
            
    if match_idx is None:
        for idx, r6 in enumerate(unmatched_632):
            h6 = normalize_inv(r6['so_hd'])
            if h5 == h6:
                match_idx = idx
                match_type = 'Khớp HĐ + mặt hàng (khác ngày CT)'
                break
                
    if match_idx is None:
        for idx, r6 in enumerate(unmatched_632):
            d6 = r6['ngay_ct']
            if d5 == d6:
                match_idx = idx
                match_type = 'Ghép theo ngày CT (khác HĐ/mặt hàng)'
                break
                
    if match_idx is not None:
        r6 = unmatched_632.pop(match_idx)
        dt = r5['ps_co']
        gv = r6['ps_no']
        lai_lo = dt - gv
        matched.append({
            'year': r5['year'],
            'ngay_ct': r5['ngay_ct'],
            'so_hd': r5['so_hd'],
            'dien_giai_511': r5['dien_giai'],
            'dien_giai_632': r6['dien_giai'],
            'dt': dt,
            'gv': gv,
            'lai_lo': lai_lo,
            'status': match_type
        })

print('=== PHÂN TÍCH THEO NĂM ===')
for yr in ['2022', '2023', '2024', '2025']:
    items = [m for m in matched if m['year'] == yr]
    tot_dt = sum(m['dt'] for m in items)
    tot_gv = sum(m['gv'] for m in items)
    tot_lai = tot_dt - tot_gv
    margin = (tot_lai / tot_dt * 100) if tot_dt > 0 else 0
    
    below_cost = [m for m in items if m['lai_lo'] < 0]
    breakeven = [m for m in items if m['lai_lo'] == 0]
    profit = [m for m in items if m['lai_lo'] > 0]
    
    print(f'NĂM {yr}:')
    print(f'  Số dòng: {len(items)} | DT 511: {tot_dt:,.0f} đ | GV 632: {tot_gv:,.0f} đ | Lãi gộp: {tot_lai:,.0f} đ ({margin:.2f}%)')
    print(f'  Lãi (P > 0): {len(profit)} dòng | Hòa vốn (P = 0): {len(breakeven)} dòng | Lỗ / Dưới giá (P < 0): {len(below_cost)} dòng')
    if below_cost:
        print('  Chi tiết bán dưới giá:')
        for b in below_cost:
            print(f"    HĐ {b['so_hd']} ngày {b['ngay_ct']}: Mặt hàng: {b['dien_giai_511']} | DT: {b['dt']:,.0f} đ | GV: {b['gv']:,.0f} đ | Lỗ: {b['lai_lo']:,.0f} đ")

tot_all_dt = sum(m['dt'] for m in matched)
tot_all_gv = sum(m['gv'] for m in matched)
tot_all_lai = tot_all_dt - tot_all_gv
all_below = [m for m in matched if m['lai_lo'] < 0]
all_even = [m for m in matched if m['lai_lo'] == 0]
print('=== TỔNG CỘNG TOÀN KỲ (2022-2025) ===')
print(f'Tổng số dòng: {len(matched)}')
print(f'Doanh thu 511: {tot_all_dt:,.0f} đ')
print(f'Giá vốn 632  : {tot_all_gv:,.0f} đ')
print(f'Lãi gộp       : {tot_all_lai:,.0f} đ (Tỷ lệ: {tot_all_lai/tot_all_dt*100:.2f}%)')
print(f'Bán dưới giá (P < 0): {len(all_below)} dòng | Hòa vốn (P = 0): {len(all_even)} dòng | Bán có lãi: {len(matched)-len(all_below)-len(all_even)} dòng')
if all_below:
    tot_below_dt = sum(b['dt'] for b in all_below)
    tot_below_gv = sum(b['gv'] for b in all_below)
    tot_below_loss = sum(b['lai_lo'] for b in all_below)
    print(f'Tổng DT bán dưới giá: {tot_below_dt:,.0f} đ ({tot_below_dt/tot_all_dt*100:.2f}% tổng DT)')
    print(f'Tổng GV bán dưới giá: {tot_below_gv:,.0f} đ')
    print(f'Tổng lỗ bán dưới giá: {tot_below_loss:,.0f} đ')
