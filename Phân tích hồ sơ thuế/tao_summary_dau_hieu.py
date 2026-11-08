#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

base_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH'

# Tạo workbook mới
wb_new = openpyxl.Workbook()
ws = wb_new.active
ws.title = "Danh_gia_Dau_hieu_Bat_thuong"

# ĐỊNH DẠNG
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", name="Times New Roman", size=11)
warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
warning_font = Font(color="9C0006", bold=True, name="Times New Roman")
info_font = Font(name="Times New Roman", size=10)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# TIÊU ĐỀ
ws.merge_cells('A1:H1')
title_cell = ws['A1']
title_cell.value = "DANH GIÁ CÁC DẤU HIỆU BẤT THƯỜNG - ĐỀ XUẤT HƯỚNG KIỂM TRA ĐỐI CHIẾU"
title_cell.font = Font(bold=True, size=12, name="Times New Roman")
title_cell.alignment = center
title_cell.fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")

ws.merge_cells('A2:H2')
info_cell = ws['A2']
info_cell.value = "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ T&T BIÊN HÒA (MST 3603883505) | Kỳ kiểm tra: 2022-2025"
info_cell.font = info_font
info_cell.alignment = center

ws.row_dimensions[1].height = 25
ws.row_dimensions[2].height = 20

# HEADER CỘT
headers = ["STT", "Loại dấu hiệu", "Chi tiết / Số liệu", "Mức độ", "Vị trí sheet", "Hướng kiểm tra chính", "Tài liệu cần xem", "Trạng thái"]
row_num = 4

for col_num, header in enumerate(headers, start=1):
    cell = ws.cell(row=row_num, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

ws.row_dimensions[4].height = 25

# DỮ LIỆU
data = [
    {
        'stt': 1,
        'loai': 'NCC F0 (Cao nhất)',
        'chi_tiet': 'Công ty TNHH NNM Group (MST 0317409384) - Không hoạt động tại địa chỉ. 2 HĐ (2023-2024): 2.442.000 đ',
        'muc_do': 'Cao',
        'vi_tri': 'IN07_NCC_CanhBao\nKQ2_NCC_mua_vao',
        'huong_kt': '1. Kiểm chứng chứng từ gốc 2 HĐ\n2. Xác minh trạng thái ĐKKD ngày hóa đơn vs ngày kiểm tra\n3. Lý do mua hàng từ NCC lạ',
        'tai_lieu': 'HĐ GTGT + Hóa đơn bán\nĐơn hàng, Hợp đồng\nCMND đại diện NCC\nCV Công an',
        'trang_thai': 'Cần kiểm tra'
    },
    {
        'stt': 2,
        'loai': 'NCC ĐKKD bất thường',
        'chi_tiet': '24 NCC có trạng thái Tạm ngừng/Đóng MST/Giải thể. Tổng giá trị 301.300.000 đ',
        'muc_do': 'Trung bình',
        'vi_tri': 'IN07_NCC_CanhBao\nKQ2_NCC_mua_vao',
        'huong_kt': '1. Danh sách 24 NCC cần lập và xác định ngày bất thường\n2. Với HĐ phát sinh sau ngày bất thường: kiểm tra chứng từ\n3. Xác minh lý do tiếp tục mua sau khi ngừng kinh doanh',
        'tai_lieu': 'Danh bạ ĐKKD 3.2 (thập hạn)\nHĐ GTGT + Hóa đơn bán\nTờ khai ĐKKD NCC',
        'trang_thai': 'Cần kiểm tra'
    },
    {
        'stt': 3,
        'loai': 'Tỷ lệ giá vốn quá cao',
        'chi_tiet': '2025: 98,53% (biên gộp 1,47%). 2022: Lỗ (-3,39M) -> 2024: Lãi 222M -> 2025: Giảm xuống 155M',
        'muc_do': 'Cao',
        'vi_tri': 'KQ2_EBITDA_Phan_tich\nKQ2_Gia_thanh_va_Dinh_muc',
        'huong_kt': '1. Phân tích nguyên nhân tỷ lệ GV tăng vọt từ 2024 sang 2025\n2. Kiểm tra định mức, chi phí SX, khấu hao\n3. Xác minh tồn kho cuối kỳ 2025 (ct 152/155/156 vs BCTC)\n4. So sánh với biên lợi nhuận ngành',
        'tai_lieu': 'Sổ 632 (Giá vốn) + 621-627\nBCTC [02] biên lợi nhuận\nPhụ lục định mức, tỷ lệ\nData ngành so sánh',
        'trang_thai': 'Cần phân tích sâu'
    },
    {
        'stt': 4,
        'loai': 'Hóa đơn bán dưới giá',
        'chi_tiet': 'HĐ #6 (23/01/2024): DT 12.500.000 đ, GV 13.245.000 đ, LỖ 745.000 đ',
        'muc_do': 'Trung bình',
        'vi_tri': 'KQ2_511_632_Chi_tiet\nKQ2_511_632',
        'huong_kt': '1. Xác minh HĐ gốc: tên khách, nội dung, đơn giá\n2. Lý do bán dưới giá (clearance, khách VIP, ghi nhầm)\n3. So sánh biên lợi nhuận với HĐ khác, với ngành\n4. Kiểm tra có quy định bán tối thiểu hay không',
        'tai_lieu': 'HĐ GTGT + Hóa đơn bán gốc\nSổ 511 (Doanh thu) + 632\nChat/Email với khách\nData giá thị trường',
        'trang_thai': 'Cần kiểm tra'
    },
    {
        'stt': 5,
        'loai': 'Không vay nợ ngân hàng',
        'chi_tiet': 'Chi phí lãi vay = 0 đ trong 4 năm (2022-2025). Tài chính tự bảo đảm hoặc vay cổ đông',
        'muc_do': 'Thấp (Tích cực)',
        'vi_tri': 'KQ2_EBITDA_Phan_tich\nKQ2_Tong_hop_co',
        'huong_kt': '1. Xác minh sao kê NH: có hạn mức tín dụng không\n2. Nếu vay cổ đông: hợp đồng + lãi suất vs thị trường\n3. Kiểm tra vốn lưu động từ doanh thu (đủ hay không)\n4. Xác minh nếu có tài trợ ngoài từ mẹ/Chi nhánh',
        'tai_lieu': 'Sao kê NH (IN11)\nSổ 131/331 (Nợ khách/NCC)\nHợp đồng vay cổ đông\nQTT TNCN (chia lợi nhuận)',
        'trang_thai': 'Thông tin tích cực'
    },
    {
        'stt': 6,
        'loai': 'Top-5 NCC chiếm 95%',
        'chi_tiet': 'HD04: Top-5 NCC = 95,22% mua vào. Ngưỡng ≤ 50%. Vượt RẤT CAO',
        'muc_do': 'Cao',
        'vi_tri': 'KQ2_NCC_mua_vao',
        'huong_kt': '1. Liệt kê Top-5 NCC + doanh số từng NCC\n2. Lý do cục bộ hóa: supply chain, chất lượng, đơn giá cạnh tranh\n3. Kiểm tra hợp đồng dài hạn, cam kết giá\n4. Xác minh giao dịch thật: biên bản giao hàng, sử dụng',
        'tai_lieu': 'Danh sách Top-5 + hợp đồng\nHĐ GTGT Top-5\nEmail/hóa đơn dòng chảy\nBiên bản bàn giao hàng',
        'trang_thai': 'Cần kiểm tra'
    }
]

# Điền dữ liệu
for idx, item in enumerate(data, start=5):
    row = idx

    ws.cell(row=row, column=1).value = item['stt']
    ws.cell(row=row, column=2).value = item['loai']
    ws.cell(row=row, column=3).value = item['chi_tiet']
    ws.cell(row=row, column=4).value = item['muc_do']
    ws.cell(row=row, column=5).value = item['vi_tri']
    ws.cell(row=row, column=6).value = item['huong_kt']
    ws.cell(row=row, column=7).value = item['tai_lieu']
    ws.cell(row=row, column=8).value = item['trang_thai']

    # Định dạng
    for col in range(1, 9):
        cell = ws.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap if col > 2 else center
        cell.border = border

        if col == 4:
            if item['muc_do'] == 'Cao':
                cell.fill = warning_fill
                cell.font = warning_font
            elif item['muc_do'] == 'Trung bình':
                cell.fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")

# Điều chỉnh độ rộng
ws.column_dimensions['A'].width = 5
ws.column_dimensions['B'].width = 20
ws.column_dimensions['C'].width = 30
ws.column_dimensions['D'].width = 13
ws.column_dimensions['E'].width = 20
ws.column_dimensions['F'].width = 38
ws.column_dimensions['G'].width = 28
ws.column_dimensions['H'].width = 15

ws.freeze_panes = 'A5'

# Lưu file
output_file = os.path.join(base_path, 'KQ3_Danh_gia_Dau_hieu_Bat_thuong_3603883505.xlsx')
wb_new.save(output_file)

print("✅ Đã tạo sheet tổng hợp dấu hiệu bất thường!")
print(f"📁 File: {output_file}")
print(f"📊 Tổng {len(data)} dấu hiệu được liệt kê")
print(f"✓ Định dạng Times New Roman, số tiền kiểu VN, ngày dd/mm/yyyy")
