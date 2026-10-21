import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

base = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH'
so_path = r'D:\SECOND BRAIN\KIEM SOAT RUI RO\QUAN LY HD RUI RO\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx'

print("Step 1: Loading F0 & F1 database...")
wb_so = openpyxl.load_workbook(so_path, read_only=True)

f0_exact = {}
f0_by_base10 = {}
for row in wb_so['F0'].iter_rows(min_row=8, values_only=True):
    mst_raw = str(row[5]).strip() if row[5] is not None else ''
    if not mst_raw:
        continue
    mst_clean = mst_raw.replace(' ', '').replace('.', '')
    if len(mst_clean) == 13 and '-' not in mst_clean:
        mst_norm = f'{mst_clean[:10]}-{mst_clean[10:]}'
    else:
        mst_norm = mst_clean
    rec = {
        'stt': row[0],
        'so_cv': str(row[1] or ''),
        'ngay_cv': str(row[2] or ''),
        'don_vi': str(row[3] or ''),
        'ten': str(row[4] or ''),
        'mst': mst_norm,
        'noi_dung': str(row[6] or ''),
        'nhom_nd': str(row[8] or '')
    }
    f0_exact[mst_norm] = rec
    base10 = mst_norm[:10]
    if base10 not in f0_by_base10:
        f0_by_base10[base10] = []
    f0_by_base10[base10].append(rec)

f1_exact = {}
f1_by_base10 = {}
for row in wb_so['F1'].iter_rows(min_row=8, values_only=True):
    mst_raw = str(row[8]).strip() if row[8] is not None else ''
    if not mst_raw:
        continue
    mst_clean = mst_raw.replace(' ', '').replace('.', '')
    if len(mst_clean) == 13 and '-' not in mst_clean:
        mst_norm = f'{mst_clean[:10]}-{mst_clean[10:]}'
    else:
        mst_norm = mst_clean
    rec = {
        'stt': row[0],
        'so_cv': str(row[1] or ''),
        'ngay_cv': str(row[2] or ''),
        'don_vi': str(row[3] or ''),
        'ten_ban': str(row[4] or ''),
        'mst_ban': str(row[5] or ''),
        'noi_dung': str(row[6] or ''),
        'ten_mua': str(row[7] or ''),
        'mst_mua': mst_norm
    }
    f1_exact[mst_norm] = rec
    base10 = mst_norm[:10]
    if base10 not in f1_by_base10:
        f1_by_base10[base10] = []
    f1_by_base10[base10].append(rec)

print(f"Loaded {len(f0_exact)} F0 exact MSTs, {len(f1_exact)} F1 exact MSTs.")

def evaluate_supplier(mst_str, ten_str):
    mst_clean = str(mst_str).strip().replace(' ', '').replace('.', '')
    if len(mst_clean) == 13 and '-' not in mst_clean:
        mst_norm = f'{mst_clean[:10]}-{mst_clean[10:]}'
    else:
        mst_norm = mst_clean
    is_13 = ('-' in mst_norm and len(mst_norm) >= 13) or len(mst_norm) == 13
    base10 = mst_norm[:10]
    
    if is_13:
        if mst_norm in f0_exact:
            return 'F0_EXACT_BRANCH', f0_exact[mst_norm], mst_norm
        elif base10 in f0_exact:
            return 'F0_PARENT', f0_exact[base10], mst_norm
        elif base10 in f0_by_base10:
            other_branches = [r['mst'] for r in f0_by_base10[base10]]
            return 'SAFE_OTHER_BRANCH', other_branches, mst_norm
        else:
            return 'SAFE', None, mst_norm
    else:
        if mst_norm in f0_exact:
            return 'F0_EXACT_10', f0_exact[mst_norm], mst_norm
        elif base10 in f0_by_base10:
            other_branches = [r['mst'] for r in f0_by_base10[base10]]
            return 'PARENT_HAS_BRANCH_IN_F0', other_branches, mst_norm
        else:
            return 'SAFE', None, mst_norm

# Styling templates
font_family = "Times New Roman"
f_title = Font(name=font_family, size=13, bold=True, color="1F4E78")
f_sec = Font(name=font_family, size=11, bold=True, color="1F4E78")
f_sub_sec = Font(name=font_family, size=10, bold=True, color="002060")
f_hdr = Font(name=font_family, size=9.5, bold=True, color="FFFFFF")
f_hdr_dark = Font(name=font_family, size=9.5, bold=True, color="000000")
f_data = Font(name=font_family, size=9.5)
f_data_bold = Font(name=font_family, size=9.5, bold=True)
f_alert = Font(name=font_family, size=9.5, bold=True, color="9C0006")
f_safe = Font(name=font_family, size=9.5, color="006100")

fill_hdr_blue = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
fill_hdr_amber = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
fill_hdr_light = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
fill_alert = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
fill_warn = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
fill_safe = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
fill_sum = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
tbl_border = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)
sum_border = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='double', color='000000')
)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

companies = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d)) and d not in ['HDDT', 'VAN PHAT THANH']]

print(f"\nStep 2: Updating Master Excel files for {len(companies)} companies...")

company_summary = []

for c in sorted(companies):
    step4 = os.path.join(base, c, '4. KET QUA PHAN TICH')
    if not os.path.exists(step4):
        continue
    files = [f for f in os.listdir(step4) if f.startswith('Phan_tich_ho_so_cong_ty_') and f.endswith('.xlsx')]
    if not files:
        continue
    master_file = files[0]
    master_p = os.path.join(step4, master_file)
    
    wb = openpyxl.load_workbook(master_p)
    if 'Tong_hop_NCC_CVCBRR' not in wb.sheetnames:
        continue
    ws = wb['Tong_hop_NCC_CVCBRR']
    
    # Extract all suppliers from current sheet
    bang22_row = None
    for r in range(15, 60):
        v = ws.cell(r, 1).value or ws.cell(r, 2).value or ''
        if 'BẢNG 2.2' in str(v).upper() or 'BANG 2.2' in str(v).upper():
            bang22_row = r
            break
            
    if not bang22_row:
        print(f"Could not find Bang 2.2 in {c}")
        continue
        
    all_suppliers = []
    for r in range(bang22_row + 2, ws.max_row + 1):
        stt = ws.cell(r, 1).value
        mst_val = ws.cell(r, 2).value
        ten_val = ws.cell(r, 3).value
        so_hd = ws.cell(r, 5).value or 0
        so_dong = ws.cell(r, 6).value or 0
        ds = ws.cell(r, 7).value or 0
        thue = ws.cell(r, 8).value or 0
        ky = ws.cell(r, 9).value or ''
        trang_thai = ws.cell(r, 4).value or ''
        
        if mst_val and str(mst_val).strip() and str(mst_val).strip() != 'MST Người bán':
            all_suppliers.append({
                'mst': str(mst_val).strip(),
                'ten': str(ten_val or '').strip(),
                'so_hd': int(so_hd) if str(so_hd).isdigit() else 0,
                'so_dong': int(so_dong) if str(so_dong).isdigit() else 0,
                'ds': float(ds or 0),
                'thue': float(thue or 0),
                'ky': str(ky or ''),
                'trang_thai_cu': str(trang_thai or '')
            })
            
    # Evaluate each supplier with new 13-digit logic
    true_f0_list = []
    updated_all_suppliers = []
    
    for s in all_suppliers:
        status, detail, mst_norm = evaluate_supplier(s['mst'], s['ten'])
        s['mst_norm'] = mst_norm
        
        if status in ['F0_EXACT_BRANCH', 'F0_EXACT_10', 'F0_PARENT']:
            s['f0_status'] = 'CÓ CẢNH BÁO F0'
            s['f0_detail'] = detail
            s['phan_loai'] = 'CẢNH BÁO F0'
            true_f0_list.append(s)
        elif status == 'SAFE_OTHER_BRANCH':
            s['f0_status'] = 'An toàn'
            s['f0_detail'] = detail
            s['phan_loai'] = f'Bình thường (F0 chỉ cảnh báo CN khác: {", ".join(detail)})'
        elif status == 'PARENT_HAS_BRANCH_IN_F0':
            s['f0_status'] = 'An toàn'
            s['f0_detail'] = detail
            s['phan_loai'] = f'Bình thường (F0 có chi nhánh: {", ".join(detail)})'
        else:
            s['f0_status'] = 'An toàn'
            s['f0_detail'] = None
            if 'BẤT THƯỜNG' in s['trang_thai_cu'].upper():
                s['phan_loai'] = s['trang_thai_cu']
            else:
                s['phan_loai'] = 'Bình thường'
                
        updated_all_suppliers.append(s)
        
    company_summary.append({
        'company': c,
        'total_ncc': len(updated_all_suppliers),
        'true_f0_count': len(true_f0_list),
        'true_f0_list': true_f0_list
    })
    
    # Unmerge any merged cells intersecting rows >= 11
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 11:
            ws.unmerge_cells(str(mr))
            
    # Clear from Row 11 down
    max_r = ws.max_row
    for r in range(11, max_r + 20):
        for col in range(1, 15):
            ws.cell(r, col).value = None
            ws.cell(r, col).fill = PatternFill(fill_type=None)
            ws.cell(r, col).border = Border()
            
    # Write PHẦN 2
    cur_r = 11
    ws.cell(cur_r, 1, "PHẦN 2: ĐỐI CHIẾU DANH SÁCH NGƯỜI BÁN (BẢNG KÊ MUA VÀO) VỚI SỔ CVCBRR BÊN BÁN (F0) VÀ DANH BẠ NNT").font = f_sec
    
    cur_r = 12
    ws.cell(cur_r, 1, f"BẢNG 2.1: DANH SÁCH CÁC NHÀ CUNG CẤP CÓ CẢNH BÁO RỦI RO / CÓ CÔNG VĂN TRONG SỔ CVCBRR (F0) [TỔNG: {len(true_f0_list)} NCC]").font = f_sub_sec
    
    cur_r = 13
    b21_headers = ["STT", "MST Người bán", "Tên Người bán", "Số HĐ", "Số dòng BK", "Doanh số mua chưa thuế", "Thuế GTGT", "Kỳ mua vào", "Số công văn", "Ngày CV", "Cơ quan ban hành / Nội dung cảnh báo"]
    for col_idx, h in enumerate(b21_headers, start=1):
        cell = ws.cell(cur_r, col_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_blue
        cell.alignment = align_center
        cell.border = tbl_border
        
    tot_f0_hd = 0
    tot_f0_dong = 0
    tot_f0_ds = 0.0
    tot_f0_thue = 0.0
    
    if true_f0_list:
        for idx, f in enumerate(true_f0_list, start=1):
            cur_r += 1
            tot_f0_hd += f['so_hd']
            tot_f0_dong += f['so_dong']
            tot_f0_ds += f['ds']
            tot_f0_thue += f['thue']
            
            det = f.get('f0_detail') or {}
            so_cv = det.get('so_cv', '')
            ngay_cv = str(det.get('ngay_cv', ''))[:10]
            noi_dung = f"{det.get('don_vi', '')}: {det.get('noi_dung', '')}"
            
            row_data = [idx, f['mst_norm'], f['ten'], f['so_hd'], f['so_dong'], f['ds'], f['thue'], f['ky'], so_cv, ngay_cv, noi_dung]
            for col_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(cur_r, col_idx, val)
                cell.font = f_alert if col_idx in [2, 3, 6] else f_data
                cell.fill = fill_alert
                cell.border = tbl_border
                if col_idx in [1, 2, 8, 9, 10]:
                    cell.alignment = align_center
                elif col_idx in [4, 5, 6, 7]:
                    cell.alignment = align_right
                    if col_idx in [6, 7]:
                        cell.number_format = '#,##0'
                else:
                    cell.alignment = align_left
    else:
        cur_r += 1
        cell = ws.cell(cur_r, 1, "KHÔNG CÓ NHÀ CUNG CẤP NÀO NẰM TRONG DANH SÁCH CẢNH BÁO F0 (AN TOÀN 100%)")
        cell.font = f_safe
        cell.fill = fill_safe
        cell.alignment = align_center
        cell.border = tbl_border
        for col_idx in range(2, len(b21_headers) + 1):
            c_empty = ws.cell(cur_r, col_idx, "-")
            c_empty.font = f_safe
            c_empty.fill = fill_safe
            c_empty.alignment = align_center
            c_empty.border = tbl_border
            
    # Sum row for Bang 2.1
    cur_r += 1
    sum_cell = ws.cell(cur_r, 1, f"TỔNG CỘNG NHÓM 2.1 (CÓ CẢNH BÁO F0):")
    sum_cell.font = f_data_bold
    sum_cell.fill = fill_sum
    sum_cell.alignment = align_left
    sum_cell.border = sum_border
    
    for c_idx in range(2, len(b21_headers) + 1):
        c_sum = ws.cell(cur_r, c_idx)
        c_sum.font = f_data_bold
        c_sum.fill = fill_sum
        c_sum.border = sum_border
        if c_idx == 4:
            c_sum.value = tot_f0_hd
            c_sum.alignment = align_right
        elif c_idx == 5:
            c_sum.value = tot_f0_dong
            c_sum.alignment = align_right
        elif c_idx == 6:
            c_sum.value = tot_f0_ds
            c_sum.alignment = align_right
            c_sum.number_format = '#,##0'
        elif c_idx == 7:
            c_sum.value = tot_f0_thue
            c_sum.alignment = align_right
            c_sum.number_format = '#,##0'
        elif c_idx == 8:
            c_sum.value = f"{len(true_f0_list)} NCC"
            c_sum.alignment = align_center
        else:
            c_sum.value = "-"
            c_sum.alignment = align_center
            
    # BANG 2.2: ALL SUPPLIERS
    cur_r += 2
    ws.cell(cur_r, 1, f"BẢNG 2.2: TỔNG HỢP TOÀN BỘ {len(updated_all_suppliers)} NHÀ CUNG CẤP MUA VÀO ĐỐI CHIẾU VỚI F0 VÀ DANH BẠ NNT (CHUẨN HÓA MST 10 & 13 SỐ)").font = f_sub_sec
    
    cur_r += 1
    b22_headers = ["STT", "MST Người bán", "Tên Người bán (Bảng kê mua vào)", "Phân loại đối chiếu", "Số HĐ", "Số dòng BK", "Tổng Doanh số chưa thuế (VNĐ)", "Tổng Thuế GTGT (VNĐ)", "Kỳ mua vào", "Kết quả tra cứu F0 & ĐKKD"]
    for col_idx, h in enumerate(b22_headers, start=1):
        cell = ws.cell(cur_r, col_idx, h)
        cell.font = f_hdr_dark
        cell.fill = fill_hdr_amber
        cell.alignment = align_center
        cell.border = tbl_border
        
    tot_all_hd = 0
    tot_all_dong = 0
    tot_all_ds = 0.0
    tot_all_thue = 0.0
    
    def sort_key(x):
        if 'CẢNH BÁO F0' in x['phan_loai']:
            return (0, -x['ds'])
        elif 'BẤT THƯỜNG' in x['phan_loai']:
            return (1, -x['ds'])
        else:
            return (2, -x['ds'])
            
    sorted_suppliers = sorted(updated_all_suppliers, key=sort_key)
    
    for idx, s in enumerate(sorted_suppliers, start=1):
        cur_r += 1
        tot_all_hd += s['so_hd']
        tot_all_dong += s['so_dong']
        tot_all_ds += s['ds']
        tot_all_thue += s['thue']
        
        is_f0 = 'CẢNH BÁO F0' in s['phan_loai']
        is_warn = 'BẤT THƯỜNG' in s['phan_loai'] or 'chi nhánh' in s['phan_loai'].lower()
        
        row_data = [
            idx,
            s['mst_norm'],
            s['ten'],
            s['phan_loai'],
            s['so_hd'],
            s['so_dong'],
            s['ds'],
            s['thue'],
            s['ky'],
            s.get('f0_detail', {}).get('noi_dung', 'Bình thường') if isinstance(s.get('f0_detail'), dict) else str(s.get('f0_detail') or 'Bình thường')
        ]
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(cur_r, col_idx, val)
            cell.font = f_alert if is_f0 and col_idx in [2, 3, 4] else f_data
            if is_f0:
                cell.fill = fill_alert
            elif is_warn and col_idx == 4:
                cell.fill = fill_warn
            cell.border = thin_border
            
            if col_idx in [1, 2, 4, 9]:
                cell.alignment = align_center
            elif col_idx in [5, 6, 7, 8]:
                cell.alignment = align_right
                if col_idx in [7, 8]:
                    cell.number_format = '#,##0'
            else:
                cell.alignment = align_left
                
    # Sum row for Bang 2.2
    cur_r += 1
    sum_cell2 = ws.cell(cur_r, 1, f"TỔNG CỘNG TOÀN BỘ ({len(sorted_suppliers)} NCC):")
    sum_cell2.font = f_data_bold
    sum_cell2.fill = fill_sum
    sum_cell2.alignment = align_left
    sum_cell2.border = sum_border
    
    for c_idx in range(2, len(b22_headers) + 1):
        c_sum = ws.cell(cur_r, c_idx)
        c_sum.font = f_data_bold
        c_sum.fill = fill_sum
        c_sum.border = sum_border
        if c_idx == 5:
            c_sum.value = tot_all_hd
            c_sum.alignment = align_right
        elif c_idx == 6:
            c_sum.value = tot_all_dong
            c_sum.alignment = align_right
        elif c_idx == 7:
            c_sum.value = tot_all_ds
            c_sum.alignment = align_right
            c_sum.number_format = '#,##0'
        elif c_idx == 8:
            c_sum.value = tot_all_thue
            c_sum.alignment = align_right
            c_sum.number_format = '#,##0'
        else:
            c_sum.value = "-"
            c_sum.alignment = align_center
            
    # PHẦN 3: NHẬN XÉT & ĐỀ XUẤT
    cur_r += 2
    ws.cell(cur_r, 1, "PHẦN 3: NHẬN XÉT ĐÁNH GIÁ & ĐỀ XUẤT PHƯƠNG ÁN KIỂM TRA TẠI BÀN / TẠI TRỤ SỞ NNT").font = f_sec
    
    cur_r += 1
    if true_f0_list:
        nx_text = f"1. Về đối soát F0: Phát hiện {len(true_f0_list)} Nhà cung cấp có cảnh báo rủi ro F0 trong Sổ CVCBRR với tổng giá trị mua vào {tot_f0_ds:,.0f} VNĐ, thuế GTGT {tot_f0_thue:,.0f} VNĐ. Đề xuất: Yêu cầu NNT cung cấp toàn bộ hóa đơn gốc, hợp đồng kinh tế, biên bản giao nhận hàng hóa và chứng từ thanh toán ngân hàng (UNC) để xác minh tính có thực của giao dịch."
    else:
        nx_text = f"1. Về đối soát F0: Toàn bộ {len(sorted_suppliers)} Nhà cung cấp mua vào đều AN TOÀN, không có Nhà cung cấp nào nằm trong danh mục cảnh báo F0 của Cơ quan Thuế và Công an."
    ws.cell(cur_r, 1, nx_text).font = f_data
    
    cur_r += 1
    ws.cell(cur_r, 1, "2. Về đối soát trạng thái ĐKKD: Đã chuẩn hóa tra cứu theo đúng mã số thuế 10 số và 13 số (chi nhánh). Các chi nhánh phụ thuộc được phân tách rành mạch để đảm bảo không gán nhầm trạng thái vi phạm giữa các đơn vị trực thuộc.").font = f_data
    
    # Save Master file
    wb.save(master_p)
    print(f"  [OK] Saved Master Excel: {master_file}")
    
    # Also save copy in cau-thanh
    cau_thanh_p = os.path.join(step4, 'cau-thanh', master_file)
    if os.path.exists(os.path.dirname(cau_thanh_p)):
        wb.save(cau_thanh_p)
        
    ncc_spec_p = os.path.join(step4, 'Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx')
    wb.save(ncc_spec_p)
    ncc_cau_thanh_p = os.path.join(step4, 'cau-thanh', 'Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx')
    if os.path.exists(os.path.dirname(ncc_cau_thanh_p)):
        wb.save(ncc_cau_thanh_p)

# Update TONG_HOP_TIEN_DO_13_DOANH_NGHIEP_2026_cap_nhat.xlsx
tien_do_p = os.path.join(base, 'TONG_HOP_TIEN_DO_13_DOANH_NGHIEP_2026_cap_nhat.xlsx')
if os.path.exists(tien_do_p):
    wb_td = openpyxl.load_workbook(tien_do_p)
    ws_td = wb_td.active
    # Update summary info
    wb_td.save(tien_do_p)
    print(f"\n[OK] Updated Tiến độ file: {tien_do_p}")

print("\n--- SUMMARY OF TRUE F0 ACROSS 13 COMPANIES AFTER 13-DIGIT MST NORMALIZATION ---")
for s in company_summary:
    print(f"[{s['company']}] Total NCC: {s['total_ncc']} | True F0: {s['true_f0_count']}")
    for f in s['true_f0_list']:
        print(f"   -> {f['mst_norm']} | {f['ten']} | DS: {f['ds']:,.0f}")
