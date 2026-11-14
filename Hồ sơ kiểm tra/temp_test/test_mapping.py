import docx
import re

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

# Let's see the full list of old instruction strings and how to map them
# Standard fields:
# Group 1:
# Ten_NNT, MST, Dia_chi, Phuong_xa, Nganh_nghe, Giam_doc, CCCD_GD, Ngay_cap_CCCD, Noi_cap_CCCD,
# Ke_toan_truong, CCCD_KT, So_DKKD, Ngay_cap_DKKD, Tai_khoan_NH, Chi_nhanh, Ma_CQT, Ten_CQT
# Group 2:
# Truong_doan, Chuc_vu_TD, Thanh_vien, Chuc_vu_TV, Noi_dung_KT, Phan_cong_TD, Phan_cong_TV
# Group 3:
# So_QDKT, Ngay_QDKT, So_QDGS, Ngay_QDGS, Thoi_han_KT, Thoi_ky_KT,
# Ngay_cong_bo, Ngay_bat_dau, Ngay_ket_thuc, Ngay_ky_BBKT, Ngay_bao_cao, So_QDXPHC, Ngay_QDXPHC
# Group 4:
# Truy_thu_TNDN, Truy_thu_GTGT, Tong_truy_thu, Tong_truy_thu_chu,
# Tong_cham_nop, Tong_cham_nop_chu, Cham_nop_TNDN, Cham_nop_GTGT,
# Phat_20_phan_tram, Phat_thu_tuc, Tong_tien_phat, Tong_tien_phat_chu,
# Tong_phat_sinh, Tong_phat_sinh_chu, Giam_lo, Giam_lo_chu, Giam_khau_tru, Giam_khau_tru_chu

def get_field_mapping(instr, context_loc=""):
    raw = instr.strip()
    # Normalize spaces
    clean = re.sub(r'\s+', ' ', raw)
    
    # Check what field it is
    # 1. Ten cong ty
    if 'Tên_công_ty' in clean:
        return 'MERGEFIELD Ten_NNT', '«Ten_NNT»'
    if 'MST' in clean:
        return 'MERGEFIELD MST', '«MST»'
    if 'địa_chỉ' in clean.lower() or 'địa_chỉ' in clean:
        return 'MERGEFIELD Dia_chi', '«Dia_chi»'
    if 'NNKD' in clean:
        return 'MERGEFIELD Nganh_nghe', '«Nganh_nghe»'
    if 'GĐ' in clean:
        return 'MERGEFIELD Giam_doc', '«Giam_doc»'
    if 'CCCDKT' in clean:
        return 'MERGEFIELD CCCD_KT', '«CCCD_KT»'
    if 'CCCD' in clean:
        return 'MERGEFIELD CCCD_GD', '«CCCD_GD»'
    if 'ngày_cấp2' in clean.lower():
        return 'MERGEFIELD Ngay_cap_DKKD \\@ "dd/MM/yyyy"', '«Ngay_cap_DKKD»'
    if 'ngày_cấp' in clean.lower():
        return 'MERGEFIELD Ngay_cap_CCCD \\@ "dd/MM/yyyy"', '«Ngay_cap_CCCD»'
    if 'Nơi_cấp' in clean:
        return 'MERGEFIELD Noi_cap_CCCD', '«Noi_cap_CCCD»'
    if 'KT' in clean:
        return 'MERGEFIELD Ke_toan_truong', '«Ke_toan_truong»'
    if 'GGCNĐKKD' in clean:
        return 'MERGEFIELD So_DKKD', '«So_DKKD»'
    if 'Tài_khoản_số' in clean:
        return 'MERGEFIELD Tai_khoan_NH', '«Tai_khoan_NH»'
    if 'Chi_nhánh' in clean:
        return 'MERGEFIELD Chi_nhanh', '«Chi_nhanh»'
    if 'Ma_CQT' in clean:
        return 'MERGEFIELD Ma_CQT', '«Ma_CQT»'
    if 'CQT' in clean:
        return 'MERGEFIELD Ten_CQT', '«Ten_CQT»'
        
    # Group 2: Doan kiem tra
    if 'CVTĐ' in clean:
        return 'MERGEFIELD Chuc_vu_TD', '«Chuc_vu_TD»'
    if 'TĐ' in clean:
        return 'MERGEFIELD Truong_doan', '«Truong_doan»'
    if 'CVTV' in clean:
        return 'MERGEFIELD Chuc_vu_TV', '«Chuc_vu_TV»'
    if 'TV' in clean:
        return 'MERGEFIELD Thanh_vien', '«Thanh_vien»'
    if 'Nội_dung_kiểm_tra' in clean:
        return 'MERGEFIELD Noi_dung_KT', '«Noi_dung_KT»'
    if 'Phân_công_trưởng_đoàn' in clean:
        return 'MERGEFIELD Phan_cong_TD', '«Phan_cong_TD»'
    if 'Phân_công_thành_viên' in clean:
        return 'MERGEFIELD Phan_cong_TV', '«Phan_cong_TV»'
        
    # Group 3: Tien trinh & Moc van ban
    if 'QĐKT' in clean:
        return 'MERGEFIELD So_QDKT', '«So_QDKT»'
    if 'ngày_qđ' in clean.lower() or 'ngay_qđ' in clean.lower():
        return 'MERGEFIELD Ngay_QDKT \\@ "dd/MM/yyyy"', '«Ngay_QDKT»'
    if 'QĐGS' in clean:
        return 'MERGEFIELD So_QDGS', '«So_QDGS»'
    if 'Thời_hạn_kiểm_tra' in clean:
        return 'MERGEFIELD Thoi_han_KT', '«Thoi_han_KT»'
    if 'Thời_kỳ_kiểm_tra' in clean:
        return 'MERGEFIELD Thoi_ky_KT', '«Thoi_ky_KT»'
    if 'Ngày_công_bố' in clean:
        return 'MERGEFIELD Ngay_cong_bo \\@ "dd/MM/yyyy"', '«Ngay_cong_bo»'
    if 'Từ_ngày' in clean:
        return 'MERGEFIELD Ngay_bat_dau \\@ "dd/MM/yyyy"', '«Ngay_bat_dau»'
    if 'Đến_ngày' in clean:
        return 'MERGEFIELD Ngay_ket_thuc \\@ "dd/MM/yyyy"', '«Ngay_ket_thuc»'
    if 'Ngày_ký_Biên_bản_kiểm_tra' in clean:
        return 'MERGEFIELD Ngay_ky_BBKT \\@ "dd/MM/yyyy"', '«Ngay_ky_BBKT»'
    if 'Ngày_báo_cáo' in clean:
        if '\\@ "MM"' in clean or '@ "MM"' in clean:
            return 'MERGEFIELD Ngay_bao_cao \\@ "MM"', '«Ngay_bao_cao»'
        if '\\@ "dd"' in clean or '@ "dd"' in clean:
            return 'MERGEFIELD Ngay_bao_cao \\@ "dd"', '«Ngay_bao_cao»'
        return 'MERGEFIELD Ngay_bao_cao \\@ "dd/MM/yyyy"', '«Ngay_bao_cao»'

    # Group 4: So lieu tai chinh
    if 'Thuế_TNDN' in clean:
        return 'MERGEFIELD Truy_thu_TNDN \\# "#,##0"', '«Truy_thu_TNDN»'
    if 'Thuế_GTGT' in clean:
        return 'MERGEFIELD Truy_thu_GTGT \\# "#,##0"', '«Truy_thu_GTGT»'
    if 'Tổng_cộng_truy_thu_BC' in clean:
        return 'MERGEFIELD Tong_truy_thu_chu', '«Tong_truy_thu_chu»'
    if 'Tổng_cộng_truy_thu' in clean or 'Tổng' in clean and 'phạt' not in clean:
        return 'MERGEFIELD Tong_truy_thu \\# "#,##0"', '«Tong_truy_thu»'
    if 'Số_tiền_phạt_20_BC' in clean:
        return 'MERGEFIELD Tong_tien_phat_chu', '«Tong_tien_phat_chu»'
    if 'Tổng_phạt_kk_sai' in clean:
        return 'MERGEFIELD Tong_tien_phat \\# "#,##0"', '«Tong_tien_phat»'
    if 'Tổng_cộng_số_phạt__chậm_nộp' in clean:
        return 'MERGEFIELD Tong_tien_phat \\# "#,##0"', '«Tong_tien_phat»'
    if 'Số_tiền_phạt_003_TNDN' in clean:
        return 'MERGEFIELD Cham_nop_TNDN \\# "#,##0"', '«Cham_nop_TNDN»'
    if 'Số_tiền_phạt_003_GTGT' in clean:
        return 'MERGEFIELD Cham_nop_GTGT \\# "#,##0"', '«Cham_nop_GTGT»'
    if 'Tổng_phạt' in clean:
        return 'MERGEFIELD Tong_cham_nop \\# "#,##0"', '«Tong_cham_nop»'
    if 'Phạt_KK_sai_VPHC' in clean:
        return 'MERGEFIELD Phat_thu_tuc \\# "#,##0"', '«Phat_thu_tuc»'
    if 'Số_tiền_phạt_20' in clean:
        return 'MERGEFIELD Phat_20_phan_tram \\# "#,##0"', '«Phat_20_phan_tram»'
    if 'Tổng_cộng_phạt_và_truy_thu_BC' in clean:
        return 'MERGEFIELD Tong_phat_sinh_chu', '«Tong_phat_sinh_chu»'
    if 'Tổng_cộng_phạt_và_truy_thu' in clean:
        return 'MERGEFIELD Tong_phat_sinh \\# "#,##0"', '«Tong_phat_sinh»'
    if 'Giảm_lỗ' in clean:
        return 'MERGEFIELD Giam_lo \\# "#,##0"', '«Giam_lo»'
    if 'giảm_khấu_trừ' in clean:
        return 'MERGEFIELD Giam_khau_tru \\# "#,##0"', '«Giam_khau_tru»'

    return None, None

# Test on all fields
from analyze_fields import all_fields

unmapped = []
for loc, instr, text in all_fields:
    target_instr, target_text = get_field_mapping(instr, loc)
    if not target_instr:
        unmapped.append((loc, instr, text))

print(f"Total fields: {len(all_fields)}")
print(f"Unmapped fields count: {len(unmapped)}")
if unmapped:
    for loc, instr, text in unmapped:
        print(f"UNMAPPED: at {loc}: [{instr}] text=[{text}]")
else:
    print("ALL 309 FIELDS SUCCESSFULLY MAPPED!")
