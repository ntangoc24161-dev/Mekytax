#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

base_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH'

# Tạo workbook mới
wb_giai_trinh = openpyxl.Workbook()
wb_giai_trinh.remove(wb_giai_trinh.active)

# ĐỊNH DẠNG CHUNG
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", name="Times New Roman", size=11)
title_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
title_font = Font(bold=True, name="Times New Roman", size=12)
warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
info_font = Font(name="Times New Roman", size=10)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
right_align = Alignment(horizontal="right", vertical="center")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def format_number(value):
    """Format số thành VN format"""
    try:
        return float(value) if value else 0
    except:
        return 0

# ============ SHEET 1: NCC F0 ============
ws1 = wb_giai_trinh.create_sheet("01_NCC_F0_cao_nhat")

ws1.merge_cells('A1:G1')
ws1['A1'] = "DANH SÁCH NCC F0 (CÔNG AN/CVCBRR) - CẦN GIẢI TRÌNH"
ws1['A1'].font = title_font
ws1['A1'].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
ws1['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")

headers1 = ["STT", "Tên NCC", "MST", "Địa chỉ", "Trạng thái ĐKKD", "HĐ số", "Ngày HĐ", "Năm", "Giá trị (đ)", "Ghi chú"]
for col_num, header in enumerate(headers1, start=1):
    cell = ws1.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

# Dữ liệu NCC F0
ncc_f0_data = [
    {
        'stt': 1,
        'ten_ncc': 'Công ty TNHH NNM Group',
        'mst': '0317409384',
        'dia_chi': 'Không hoạt động tại địa chỉ đăng ký',
        'trang_thai': 'F0 (Công an)',
        'hd_so': 'Tuyến tính theo BK',
        'ngay_hd': '2023-2024',
        'nam': '2023, 2024',
        'gia_tri': 2442000,
        'ghi_chu': 'Trong sổ CVCBRR hoặc Công an cảnh báo'
    }
]

for idx, item in enumerate(ncc_f0_data, start=4):
    row = idx
    ws1.cell(row=row, column=1).value = item['stt']
    ws1.cell(row=row, column=2).value = item['ten_ncc']
    ws1.cell(row=row, column=3).value = item['mst']
    ws1.cell(row=row, column=4).value = item['dia_chi']
    ws1.cell(row=row, column=5).value = item['trang_thai']
    ws1.cell(row=row, column=6).value = item['hd_so']
    ws1.cell(row=row, column=7).value = item['ngay_hd']
    ws1.cell(row=row, column=8).value = item['nam']
    ws1.cell(row=row, column=9).value = item['gia_tri']
    ws1.cell(row=row, column=10).value = item['ghi_chu']

    for col in range(1, 11):
        cell = ws1.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border
        if col == 9:
            cell.alignment = right_align
            cell.number_format = '#,##0'

ws1.column_dimensions['A'].width = 5
ws1.column_dimensions['B'].width = 25
ws1.column_dimensions['C'].width = 15
ws1.column_dimensions['D'].width = 25
ws1.column_dimensions['E'].width = 20
ws1.column_dimensions['F'].width = 15
ws1.column_dimensions['G'].width = 15
ws1.column_dimensions['H'].width = 12
ws1.column_dimensions['I'].width = 15
ws1.column_dimensions['J'].width = 25

# ============ SHEET 2: NCC ĐKKD BẤT THƯỜNG ============
ws2 = wb_giai_trinh.create_sheet("02_NCC_DDKD_bat_thuong")

ws2.merge_cells('A1:I1')
ws2['A1'] = "DANH SÁCH 24 NCC CÓ TRẠNG THÁI ĐKKD BẤT THƯỜNG - CẦN GIẢI TRÌNH"
ws2['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws2['A1'].fill = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")

headers2 = ["STT", "Tên NCC", "MST", "Trạng thái ĐKKD hiện tại", "Ngày thay đổi", "Lý do kinh doanh tiếp tục", "Tổng HĐ", "Tổng giá trị (đ)", "HĐ phát sinh sau thay đổi"]
for col_num, header in enumerate(headers2, start=1):
    cell = ws2.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

# Gợi ý dữ liệu (cần quét từ IN07)
ncc_abnormal_data = [
    {
        'stt': 1,
        'ten_ncc': '[Từ IN07_NCC_CanhBao - cột Trạng_thai ĐKKD]',
        'mst': '[MST]',
        'trang_thai': 'Tạm ngừng / Đóng MST / Giải thể',
        'ngay_thay_doi': '[dd/mm/yyyy]',
        'ly_do': '[Xin từ công ty]',
        'tong_hd': '[Số]',
        'tong_gia_tri': 0,
        'hd_sau': '[Có/Không]'
    }
]

for idx, item in enumerate(ncc_abnormal_data, start=4):
    row = idx
    ws2.cell(row=row, column=1).value = item['stt']
    ws2.cell(row=row, column=2).value = item['ten_ncc']
    ws2.cell(row=row, column=3).value = item['mst']
    ws2.cell(row=row, column=4).value = item['trang_thai']
    ws2.cell(row=row, column=5).value = item['ngay_thay_doi']
    ws2.cell(row=row, column=6).value = item['ly_do']
    ws2.cell(row=row, column=7).value = item['tong_hd']
    ws2.cell(row=row, column=8).value = item['tong_gia_tri']
    ws2.cell(row=row, column=9).value = item['hd_sau']

    for col in range(1, 10):
        cell = ws2.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border
        if col == 8:
            cell.alignment = right_align
            cell.number_format = '#,##0'

ws2.column_dimensions['A'].width = 5
ws2.column_dimensions['B'].width = 25
ws2.column_dimensions['C'].width = 15
ws2.column_dimensions['D'].width = 20
ws2.column_dimensions['E'].width = 15
ws2.column_dimensions['F'].width = 30
ws2.column_dimensions['G'].width = 12
ws2.column_dimensions['H'].width = 15
ws2.column_dimensions['I'].width = 20

# ============ SHEET 3: HĐ BÁN DƯỚI GIÁ ============
ws3 = wb_giai_trinh.create_sheet("03_HD_ban_duoi_gia")

ws3.merge_cells('A1:J1')
ws3['A1'] = "DANH SÁCH HÓA ĐƠN BÁN DƯỚI GIÁ VỐN - CẦN GIẢI TRÌNH"
ws3['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws3['A1'].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

headers3 = ["STT", "HĐ số", "Ngày HĐ", "Năm", "Tên khách hàng", "Doanh thu (đ)", "Giá vốn (đ)", "Lỗ (đ)", "Lý do bán dưới giá", "Lãi suất theo ngành"]
for col_num, header in enumerate(headers3, start=1):
    cell = ws3.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

# Dữ liệu HĐ bán dưới giá
hd_duoi_gia_data = [
    {
        'stt': 1,
        'hd_so': '6',
        'ngay_hd': '23/01/2024',
        'nam': '2024',
        'ten_khach': '[Từ KQ2_511_632_Chi_tiet]',
        'doanh_thu': 12500000,
        'gia_von': 13245000,
        'lo': 745000,
        'ly_do': '[Xin từ công ty: clearance, khách VIP, ghi nhầm, v.v.]',
        'lai_suat': '[% ngành so sánh]'
    }
]

for idx, item in enumerate(hd_duoi_gia_data, start=4):
    row = idx
    ws3.cell(row=row, column=1).value = item['stt']
    ws3.cell(row=row, column=2).value = item['hd_so']
    ws3.cell(row=row, column=3).value = item['ngay_hd']
    ws3.cell(row=row, column=4).value = item['nam']
    ws3.cell(row=row, column=5).value = item['ten_khach']
    ws3.cell(row=row, column=6).value = item['doanh_thu']
    ws3.cell(row=row, column=7).value = item['gia_von']
    ws3.cell(row=row, column=8).value = item['lo']
    ws3.cell(row=row, column=9).value = item['ly_do']
    ws3.cell(row=row, column=10).value = item['lai_suat']

    for col in range(1, 11):
        cell = ws3.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border
        if col in [6, 7, 8]:
            cell.alignment = right_align
            cell.number_format = '#,##0'

ws3.column_dimensions['A'].width = 5
for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
    ws3.column_dimensions[col].width = 18

# ============ SHEET 4: TOP-5 NCC ============
ws4 = wb_giai_trinh.create_sheet("04_Top5_NCC")

ws4.merge_cells('A1:G1')
ws4['A1'] = "DANH SÁCH TOP-5 NCC CHIẾM 95,22% MUA VÀO - CẦN GIẢI TRÌNH"
ws4['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws4['A1'].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

headers4 = ["STT", "Tên NCC", "MST", "Tổng giá trị mua (đ)", "% mua vào", "Hợp đồng dài hạn", "Ghi chú"]
for col_num, header in enumerate(headers4, start=1):
    cell = ws4.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

# Gợi ý dữ liệu (cần tính từ KQ2_NCC_mua_vao)
top5_data = [
    {
        'stt': 1,
        'ten_ncc': '[Top NCC #1 từ KQ2_NCC_mua_vao - sắp xếp giảm theo tổng giá trị]',
        'mst': '[MST]',
        'gia_tri': 0,
        'phan_tram': 0,
        'hop_dong': '[Có/Không - xin từ công ty]',
        'ghi_chu': '[Lý do lựa chọn: chất lượng, giá cạnh tranh, supply chain, v.v.]'
    }
]

for idx, item in enumerate(top5_data, start=4):
    row = idx
    ws4.cell(row=row, column=1).value = item['stt']
    ws4.cell(row=row, column=2).value = item['ten_ncc']
    ws4.cell(row=row, column=3).value = item['mst']
    ws4.cell(row=row, column=4).value = item['gia_tri']
    ws4.cell(row=row, column=5).value = item['phan_tram']
    ws4.cell(row=row, column=6).value = item['hop_dong']
    ws4.cell(row=row, column=7).value = item['ghi_chu']

    for col in range(1, 8):
        cell = ws4.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border
        if col == 4:
            cell.alignment = right_align
            cell.number_format = '#,##0'
        elif col == 5:
            cell.alignment = right_align
            cell.number_format = '0.00%'

ws4.column_dimensions['A'].width = 5
ws4.column_dimensions['B'].width = 25
ws4.column_dimensions['C'].width = 15
ws4.column_dimensions['D'].width = 18
ws4.column_dimensions['E'].width = 12
ws4.column_dimensions['F'].width = 20
ws4.column_dimensions['G'].width = 30

# ============ SHEET 5: CHÊNH LỆCH HĐĐT - BẢNG KÊ ============
ws5 = wb_giai_trinh.create_sheet("05_Lech_HDDT_BangKe")

ws5.merge_cells('A1:F1')
ws5['A1'] = "DANH SÁCH HÓA ĐƠN LỆCH GIỮA HĐĐT VÀ BẢNG KÊ - CẦN GIẢI TRÌNH"
ws5['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws5['A1'].fill = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")

headers5 = ["STT", "HĐ số", "Ngày HĐĐT", "Ngày Bảng kê", "Lệch nội dung", "Lý do lệch / Sửa chữa"]
for col_num, header in enumerate(headers5, start=1):
    cell = ws5.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

ws5.merge_cells('A4:A6')
ws5['A4'] = "Ghi chú"
ws5['A4'].font = info_font

ws5.merge_cells('B4:F6')
ws5['B4'] = "Dữ liệu từ sheet KQ2_BK_HDDT_Lech (KQ2.1). Liệt kê từng HĐ lệch, so sánh HĐĐT (IN01) và Bảng kê (IN02). Với HĐ lệch: xác minh chứng từ gốc, xác định lý do (ghi nhầm, ngày sai, v.v.)"
ws5['B4'].font = info_font
ws5['B4'].alignment = left_wrap

for col in range(1, 7):
    ws5.column_dimensions[chr(64 + col)].width = 20

# ============ SHEET 6: LỆCH TK 131 - BẢNG KÊ ============
ws6 = wb_giai_trinh.create_sheet("06_Lech_TK131_BangKe")

ws6.merge_cells('A1:F1')
ws6['A1'] = "DANH SÁCH KHÁCH HÀNG CÓ LỆCH GIỮA TK 131 VÀ BẢNG KÊ - CẦN GIẢI TRÌNH"
ws6['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws6['A1'].fill = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")

headers6 = ["STT", "Tên khách hàng", "Nợ theo TK 131 (đ)", "Nợ theo BK (đ)", "Lệch (đ)", "Lý do lệch / Dự kiến thu"]
for col_num, header in enumerate(headers6, start=1):
    cell = ws6.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

ws6.merge_cells('A4:A6')
ws6['A4'] = "Ghi chú"
ws6['A4'].font = info_font

ws6.merge_cells('B4:F6')
ws6['B4'] = "Tính tổng công nợ từ Bảng kê (doanh thu chưa thanh toán) rồi so sánh với số dư TK 131. Với khách lệch: tìm HĐ chưa thu hoặc thu không đủ, xác minh sao kê NH"
ws6['B4'].font = info_font
ws6['B4'].alignment = left_wrap

for col in range(1, 7):
    ws6.column_dimensions[chr(64 + col)].width = 20

# ============ SHEET 7: NLĐ NHIỀU NGUỒN ============
ws7 = wb_giai_trinh.create_sheet("07_NLD_nhieu_nguon")

ws7.merge_cells('A1:H1')
ws7['A1'] = "DANH SÁCH NLĐ CÓ THU NHẬP TỪ NHIỀU NGUỒN - CẦN KIỂM TRA"
ws7['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws7['A1'].fill = PatternFill(start_color="0099FF", end_color="0099FF", fill_type="solid")

headers7 = ["STT", "Tên NLĐ", "CCCD", "Lương cơ bản (đ)", "Bình quân về (đ)", "Thu nhập khác (đ)", "Tổng thu nhập (đ)", "Tình trạng QTT TNCN"]
for col_num, header in enumerate(headers7, start=1):
    cell = ws7.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

ws7.merge_cells('A4:A6')
ws7['A4'] = "Ghi chú"
ws7['A4'].font = info_font

ws7.merge_cells('B4:H6')
ws7['B4'] = "Từ IN05_NLD_TNCN và QTT TNCN: tìm các NLĐ có lương + bình quân về hoặc lương + khác. Kiểm tra QTT có tính đúng theo ưu tiên (lương > bình quân > khác), có được trừ 2 tháng lương khi có bình quân về"
ws7['B4'].font = info_font
ws7['B4'].alignment = left_wrap

ws7.column_dimensions['A'].width = 5
for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
    ws7.column_dimensions[col].width = 18

# Lưu file
output_file = os.path.join(base_path, 'KQ3_NNT_Giai_trinh_Chi_tiet_3603883505.xlsx')
wb_giai_trinh.save(output_file)

print("✅ Đã tạo file chi tiết để NNT giải trình!")
print(f"\n📁 File: KQ3_NNT_Giai_trinh_Chi_tiet_3603883505.xlsx")
print("\n📋 Các sheet chi tiết:")
print("  1. 01_NCC_F0_cao_nhat - Công ty NNM Group (1 NCC)")
print("  2. 02_NCC_DDKD_bat_thuong - 24 NCC trạng thái bất thường")
print("  3. 03_HD_ban_duoi_gia - HĐ #6 bán dưới giá (lỗ 745k)")
print("  4. 04_Top5_NCC - Top-5 NCC chiếm 95,22%")
print("  5. 05_Lech_HDDT_BangKe - Chênh lệch HĐĐT & Bảng kê")
print("  6. 06_Lech_TK131_BangKe - Lệch TK 131 & Bảng kê")
print("  7. 07_NLD_nhieu_nguon - NLĐ thu nhiều nguồn")
print("\n✓ Mỗi sheet có header rõ ràng, dữ liệu cần điền từ KQ hoặc xin từ công ty")
