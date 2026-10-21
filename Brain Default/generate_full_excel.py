# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

# Initialize workbook
wb = openpyxl.Workbook()

# Setup sheets
ws_detail = wb.active
ws_detail.title = 'TONG_HOP_CHI_TIET'
ws_summary = wb.create_sheet(title='TONG_HOP_THEO_CV')
ws_all31 = wb.create_sheet(title='TOAN_BO_31_HO_SO')
ws_guide = wb.create_sheet(title='DANH_MUC_TIEU_CHI')

# Fonts
font_title = Font(name='Times New Roman', size=15, bold=True, color='1F4E79')
font_subtitle = Font(name='Times New Roman', size=10, italic=True, color='595959')
font_header = Font(name='Times New Roman', size=10, bold=True, color='FFFFFF')
font_header_dark = Font(name='Times New Roman', size=10, bold=True, color='1F4E79')
font_data = Font(name='Times New Roman', size=10, color='000000')
font_data_bold = Font(name='Times New Roman', size=10, bold=True, color='000000')
font_mst = Font(name='Consolas', size=10, bold=True, color='002060')
font_na = Font(name='Times New Roman', size=10, italic=True, color='7F7F7F')

# Fills
fill_header_cv = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')      # Navy: Thông tin CV
fill_header_nnt = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')     # Blue: Thông tin NNT
fill_header_docs = PatternFill(start_color='1B5E20', end_color='1B5E20', fill_type='solid')    # Dark Green: Hồ sơ thuế
fill_header_admin = PatternFill(start_color='C55A11', end_color='C55A11', fill_type='solid')   # Amber/Brown: Xử lý VPHC & Pháp lý
fill_header_other = PatternFill(start_color='595959', end_color='595959', fill_type='solid')   # Gray: Khác & Lưu trữ

fill_zebra = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
fill_group = PatternFill(start_color='E9EEF4', end_color='E9EEF4', fill_type='solid')

fill_status_moi = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')       # Yellow
fill_status_dang_xl = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')   # Soft Blue
fill_status_trinh_ky = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid') # Soft Green

# Borders
border_thin = Side(style='thin', color='D9D9D9')
border_thick_bottom = Side(style='medium', color='1F4E79')
cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
header_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thick_bottom)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)

# -------------------------------------------------------------
# 1. SHEET TONG_HOP_CHI_TIET
# -------------------------------------------------------------
headers_detail = [
    ('STT', 6, align_center, fill_header_cv),
    ('Số Công văn CA', 22, align_left, fill_header_cv),
    ('Ngày CV CA', 12, align_center, fill_header_cv),
    ('Cơ quan Công an ban hành', 28, align_left, fill_header_cv),
    ('Mã số thuế', 14, align_center, fill_header_nnt),
    ('Tên Người nộp thuế / Doanh nghiệp', 35, align_left, fill_header_nnt),
    ('Địa chỉ trụ sở / Kinh doanh', 30, align_left, fill_header_nnt),
    ('Tổ KT', 8, align_center, fill_header_nnt),
    ('Cán bộ xử lý', 18, align_left, fill_header_nnt),
    ('Tình trạng xử lý', 14, align_center, fill_header_nnt),
    ('Tờ khai thuế GTGT\n(Kỳ/Thời gian yêu cầu)', 24, align_left, fill_header_docs),
    ('Tờ khai Quyết toán TNDN\n(Năm/Thời gian yêu cầu)', 20, align_left, fill_header_docs),
    ('Tờ khai Quyết toán TNCN\n(Năm/Kỳ khấu trừ)', 20, align_left, fill_header_docs),
    ('Báo cáo tài chính (BCTC)\n(Năm yêu cầu)', 18, align_left, fill_header_docs),
    ('Hóa đơn điện tử / Bảng kê HĐ\n(Chi tiết HĐ, thời gian phát sinh)', 32, align_left, fill_header_docs),
    ('Quyết định Xử phạt VPHC / Cưỡng chế\n(Thời gian rà soát / Số QĐ)', 26, align_left, fill_header_admin),
    ('Thông báo Đóng MST / Bỏ địa chỉ\n(Số văn bản / Tình trạng)', 24, align_left, fill_header_admin),
    ('Thông tin Định danh & Pháp lý\n(ĐKKD, ĐDPL, Kế toán, TK Ngân hàng)', 28, align_left, fill_header_admin),
    ('Nghĩa vụ nộp NSNN / Nợ thuế\n(Số tiền nợ, đã nộp NSNN)', 22, align_left, fill_header_admin),
    ('Hồ sơ tài liệu khác / Nội dung đặc thù', 28, align_left, fill_header_other),
    ('Thư mục lưu trữ Google Drive', 30, align_left, fill_header_other),
    ('Tệp Công văn trả lời (Dự thảo)', 32, align_left, fill_header_other),
    ('Tệp Tờ trình Lãnh đạo', 22, align_left, fill_header_other),
    ('Ghi chú nghiệp vụ', 25, align_left, fill_header_other),
]

ws_detail.row_dimensions[1].height = 28
ws_detail.row_dimensions[2].height = 18
ws_detail.row_dimensions[3].height = 10
ws_detail.row_dimensions[4].height = 38

ws_detail.merge_cells('A1:X1')
ws_detail['A1'] = 'BẢNG TỔNG HỢP CHI TIẾT HỒ SƠ CUNG CẤP THÔNG TIN CÔNG VĂN CÔNG AN NĂM 2026 (TỔ KT1 & HỒ SƠ TRỌNG TÂM)'
ws_detail['A1'].font = font_title
ws_detail['A1'].alignment = Alignment(horizontal='left', vertical='center')

ws_detail.merge_cells('A2:X2')
ws_detail['A2'] = 'Trích xuất tự động từ Sổ theo dõi Bao_cao_cong_van.xlsx và hệ thống hồ sơ dự thảo tại I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026 | Thời điểm lập: 25/09/2026'
ws_detail['A2'].font = font_subtitle
ws_detail['A2'].alignment = Alignment(horizontal='left', vertical='center')

# Header row
for col_idx, (h_name, width, h_align, h_fill) in enumerate(headers_detail, 1):
    cell = ws_detail.cell(row=4, column=col_idx, value=h_name.replace('\n', '
'))
    cell.font = font_header
    cell.fill = h_fill
    cell.alignment = align_center
    cell.border = header_border
    col_letter = get_column_letter(col_idx)
    ws_detail.column_dimensions[col_letter].width = width

print('Detail sheet structure ready')
