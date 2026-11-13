import shutil
import openpyxl

src = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
bak = r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\temp_test\Danh sach kiem tra quyet toan pro.bak.xlsm'

# 1. Backup outside 00. Theo doi chung
shutil.copy2(src, bak)
print("Backup created at:", bak)

wb = openpyxl.load_workbook(src, keep_vba=True, data_only=False)
ws = wb['01_DA_BAN_HANH']

# 2. Add headers to Row 4 (columns 53 to 58)
new_cols = [
    (53, 'Tong_truy_thu_chu', 'Tổng cộng truy thu bằng chữ', 'AP'),
    (54, 'Tong_cham_nop_chu', 'Tổng tiền chậm nộp bằng chữ', 'AS'),
    (55, 'Tong_tien_phat_chu', 'Tổng tiền phạt bằng chữ', 'AV'),
    (56, 'Tong_phat_sinh_chu', 'Tổng cộng phát sinh bằng chữ', 'AW'),
    (57, 'Giam_lo_chu', 'Giảm lỗ bằng chữ', 'AX'),
    (58, 'Giam_khau_tru_chu', 'Giảm khấu trừ bằng chữ', 'AY')
]

# Set header in row 3 (group header) and row 4 (column header)
for col_idx, col_name, col_vn, src_col in new_cols:
    ws.cell(4, col_idx).value = col_name

# Row 3 header can span or label NGHĨA VỤ TÀI CHÍNH (BẰNG CHỮ)
ws.cell(3, 53).value = "NGHĨA VỤ TÀI CHÍNH (BẰNG CHỮ)"

# Populate formulas for rows 5 to 33
for r in range(5, 34):
    for col_idx, col_name, col_vn, src_col in new_cols:
        ws.cell(r, col_idx).value = f"=DocSo({src_col}{r})"

# Also let's check DuLieuMerge: expand to all 12 companies (rows 5 to 16 in 01_DA_BAN_HANH -> rows 2 to 13 in DuLieuMerge)
ws_m = wb['DuLieuMerge']
# Map of DuLieuMerge columns to 01_DA_BAN_HANH columns:
col_map = {
    1: 'E',   # Ten_NNT
    2: 'D',   # MST
    3: 'F',   # Dia_chi
    4: 'G',   # Phuong_xa
    5: 'J',   # Nganh_nghe
    6: 'K',   # Giam_doc
    7: 'M',   # CCCD_GD
    8: 'N',   # Ngay_cap_CCCD
    9: 'O',   # Noi_cap_CCCD
    10: 'P',  # Ke_toan_truong
    11: 'Q',  # CCCD_KT
    12: 'R',  # So_DKKD
    13: 'S',  # Ngay_cap_DKKD
    14: 'U',  # Tai_khoan_NH
    15: 'H',  # Ma_CQT
    16: 'G',  # Ten_CQT (or G)
    17: 'V',  # Truong_doan
    18: 'W',  # Chuc_vu_TD
    19: 'X',  # Thanh_vien
    20: 'Y',  # Chuc_vu_TV
    21: 'Z',  # Noi_dung_KT
    22: 'AA', # Phan_cong_TD
    23: 'AB', # Phan_cong_TV
    24: 'AC', # So_QDKT
    25: 'AE', # Ngay_QDKT
    26: 'AD', # So_QDGS
    27: 'AE', # Ngay_QDGS
    28: 'AG', # Thoi_han_KT
    29: 'I',  # Thoi_ky_KT
    30: 'AF', # Ngay_cong_bo
    31: 'AH', # Ngay_bat_dau
    32: 'AI', # Ngay_ket_thuc
    33: 'AJ', # Ngay_ky_BBKT
    34: 'AK', # Ngay_bao_cao
    35: 'AL', # So_QDXPHC
    36: 'AM', # Ngay_QDXPHC
    37: 'AN', # Truy_thu_TNDN
    38: 'AO', # Truy_thu_GTGT
    39: 'AP', # Tong_truy_thu
    40: '=DocSo(AM{row})', # Tong_truy_thu_chu
    41: 'AS', # Tong_cham_nop
    42: '=DocSo(AO{row})', # Tong_cham_nop_chu
    43: 'AT', # Phat_20_phan_tram
    44: 'AU', # Phat_thu_tuc
    45: 'AV', # Tong_tien_phat
    46: '=DocSo(AS{row})', # Tong_tien_phat_chu
    47: 'AW', # Tong_phat_sinh
    48: '=DocSo(AU{row})', # Tong_phat_sinh_chu
    49: 'AX', # Giam_lo
    50: '=DocSo(AW{row})', # Giam_lo_chu
    51: 'AY', # Giam_khau_tru
    52: '=DocSo(AY{row})'  # Giam_khau_tru_chu
}

# Fill rows 2 to 13 (corresponding to 01_DA_BAN_HANH rows 5 to 16)
for merge_row in range(2, 14):
    ban_hanh_row = merge_row + 3 # row 2 -> 5, row 3 -> 6, ..., row 13 -> 16
    for c in range(1, 53):
        mapping = col_map[c]
        if str(mapping).startswith('='):
            ws_m.cell(merge_row, c).value = mapping.replace('{row}', str(merge_row))
        else:
            ws_m.cell(merge_row, c).value = f"='01_DA_BAN_HANH'!{mapping}{ban_hanh_row}"

wb.save(src)
print("Successfully updated Danh sach kiem tra quyet toan pro.xlsm!")
