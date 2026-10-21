#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script tạo các file Excel Checklist định dạng chuẩn in ấn 1 trang A4
phục vụ công tác kiểm tra hồ sơ thuế theo Quy trình 4 bước chuẩn hóa.
"""

import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def get_styles():
    font_name = 'Times New Roman'
    
    styles = {
        'font_title': Font(name=font_name, size=13, bold=True, color='1F4E79'),
        'font_subtitle': Font(name=font_name, size=10, italic=True, color='595959'),
        'font_org': Font(name=font_name, size=9.5, bold=True),
        'font_national': Font(name=font_name, size=9.5, bold=True),
        'font_motto': Font(name=font_name, size=9, italic=True),
        'font_info_label': Font(name=font_name, size=8.5, bold=True),
        'font_info_val': Font(name=font_name, size=8.5),
        
        'font_sec_hdr': Font(name=font_name, size=9, bold=True, color='FFFFFF'),
        'font_tbl_hdr': Font(name=font_name, size=8.5, bold=True, color='FFFFFF'),
        'font_cell_bold': Font(name=font_name, size=8, bold=True),
        'font_cell': Font(name=font_name, size=8),
        'font_cell_italic': Font(name=font_name, size=7.5, italic=True, color='595959'),
        'font_status': Font(name=font_name, size=8, bold=True, color='1F4E79'),
        'font_sign_title': Font(name=font_name, size=9, bold=True),
        'font_sign_sub': Font(name=font_name, size=8, italic=True),
        
        'fill_main_hdr': PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid'),
        'fill_tbl_hdr': PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid'),
        'fill_sec_hdr': PatternFill(start_color='333F48', end_color='333F48', fill_type='solid'),
        'fill_step_1': PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid'),
        'fill_step_2': PatternFill(start_color='1E4E79', end_color='1E4E79', fill_type='solid'),
        'fill_step_3': PatternFill(start_color='203764', end_color='203764', fill_type='solid'),
        'fill_step_4': PatternFill(start_color='0E2841', end_color='0E2841', fill_type='solid'),
        'fill_alt': PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid'),
        'fill_highlight': PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'),
        'fill_info_box': PatternFill(start_color='F9FBFD', end_color='F9FBFD', fill_type='solid'),
        
        'border_thin': Border(
            left=Side(style='thin', color='BFBFBF'),
            right=Side(style='thin', color='BFBFBF'),
            top=Side(style='thin', color='BFBFBF'),
            bottom=Side(style='thin', color='BFBFBF')
        ),
        'border_header': Border(
            left=Side(style='thin', color='FFFFFF'),
            right=Side(style='thin', color='FFFFFF'),
            top=Side(style='medium', color='1F4E79'),
            bottom=Side(style='medium', color='1F4E79')
        ),
        'border_box': Border(
            left=Side(style='medium', color='1F4E79'),
            right=Side(style='medium', color='1F4E79'),
            top=Side(style='medium', color='1F4E79'),
            bottom=Side(style='medium', color='1F4E79')
        ),
        
        'align_center': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'align_left': Alignment(horizontal='left', vertical='center', wrap_text=True),
        'align_right': Alignment(horizontal='right', vertical='center', wrap_text=True),
        'align_center_top': Alignment(horizontal='center', vertical='top', wrap_text=True),
        'align_left_top': Alignment(horizontal='left', vertical='top', wrap_text=True),
    }
    return styles

def setup_page_a4(ws, orientation='landscape', fit_width=1, fit_height=1):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    if orientation == 'landscape':
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.35
        ws.page_margins.bottom = 0.3
        ws.page_margins.header = 0.15
        ws.page_margins.footer = 0.15
    else:
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_margins.left = 0.35
        ws.page_margins.right = 0.35
        ws.page_margins.top = 0.35
        ws.page_margins.bottom = 0.35
        ws.page_margins.header = 0.15
        ws.page_margins.footer = 0.15
        
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_width
    ws.page_setup.fitToHeight = fit_height
    ws.views.sheetView[0].showGridLines = True

def add_header_block(ws, title, subtitle=None, orientation='portrait', max_col=7):
    st = get_styles()
    
    # Left Header: Co quan thue
    half_l = max(2, max_col // 2)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=half_l)
    ws.cell(row=1, column=1, value="CỤC THUẾ TỈNH / TP ....................").font = st['font_org']
    ws.cell(row=1, column=1).alignment = st['align_center']
    
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=half_l)
    ws.cell(row=2, column=1, value="CHI CỤC THUẾ / ĐOÀN KIỂM TRA").font = st['font_org']
    ws.cell(row=2, column=1).alignment = st['align_center']
    
    # Right Header: Quoc hieu
    start_r = half_l + 1
    ws.merge_cells(start_row=1, start_column=start_r, end_row=1, end_column=max_col)
    ws.cell(row=1, column=start_r, value="CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM").font = st['font_national']
    ws.cell(row=1, column=start_r).alignment = st['align_center']
    
    ws.merge_cells(start_row=2, start_column=start_r, end_row=2, end_column=max_col)
    ws.cell(row=2, column=start_r, value="Độc lập - Tự do - Hạnh phúc").font = st['font_motto']
    ws.cell(row=2, column=start_r).alignment = st['align_center']
    
    # Title
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=max_col)
    ws.cell(row=4, column=1, value=title).font = st['font_title']
    ws.cell(row=4, column=1).alignment = st['align_center']
    ws.row_dimensions[4].height = 20
    
    curr_row = 5
    if subtitle:
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=max_col)
        ws.cell(row=curr_row, column=1, value=subtitle).font = st['font_subtitle']
        ws.cell(row=curr_row, column=1).alignment = st['align_center']
        ws.row_dimensions[curr_row].height = 14
        curr_row += 1
        
    return curr_row

def add_enterprise_info_block(ws, start_row, max_col=7):
    st = get_styles()
    r = start_row
    
    half_l = max(3, max_col // 2 + 1)
    # Info row 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=half_l)
    ws.cell(row=r, column=1, value="• Tên doanh nghiệp: .........................................................................").font = st['font_info_label']
    ws.cell(row=r, column=1).alignment = st['align_left']
    
    ws.merge_cells(start_row=r, start_column=half_l+1, end_row=r, end_column=max_col)
    ws.cell(row=r, column=half_l+1, value="• Mã số thuế: ...........................................").font = st['font_info_label']
    ws.cell(row=r, column=half_l+1).alignment = st['align_left']
    ws.row_dimensions[r].height = 14
    r += 1
    
    # Info row 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=half_l)
    ws.cell(row=r, column=1, value="• Niên độ kiểm tra: Từ năm 20... đến năm 20... (Thời kỳ: ..... năm, ..... quý)").font = st['font_info_label']
    ws.cell(row=r, column=1).alignment = st['align_left']
    
    ws.merge_cells(start_row=r, start_column=half_l+1, end_row=r, end_column=max_col)
    ws.cell(row=r, column=half_l+1, value="• Quyết định KT số: ...................................").font = st['font_info_label']
    ws.cell(row=r, column=half_l+1).alignment = st['align_left']
    ws.row_dimensions[r].height = 14
    r += 1
    
    # Info row 3
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=half_l)
    ws.cell(row=r, column=1, value="• Đoàn kiểm tra / CB phân tích: ....................................................").font = st['font_info_label']
    ws.cell(row=r, column=1).alignment = st['align_left']
    
    ws.merge_cells(start_row=r, start_column=half_l+1, end_row=r, end_column=max_col)
    ws.cell(row=r, column=half_l+1, value="• Ngày rà soát: 02/10/2026").font = st['font_info_label']
    ws.cell(row=r, column=half_l+1).alignment = st['align_left']
    ws.row_dimensions[r].height = 14
    r += 1
    
    return r

def add_signatures(ws, start_row, max_col=7):
    st = get_styles()
    r = start_row
    
    # Date line
    ws.merge_cells(start_row=r, start_column=max_col-2, end_row=r, end_column=max_col)
    ws.cell(row=r, column=max_col-2, value="......, ngày ..... tháng ..... năm 2026").font = st['font_sign_sub']
    ws.cell(row=r, column=max_col-2).alignment = st['align_center']
    ws.row_dimensions[r].height = 13
    r += 1
    
    # Title sign
    half_col = max(2, max_col // 2)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=half_col)
    ws.cell(row=r, column=1, value="CÁN BỘ PHÂN TÍCH / RÀ SOÁT").font = st['font_sign_title']
    ws.cell(row=r, column=1).alignment = st['align_center']
    
    ws.merge_cells(start_row=r, start_column=half_col+1, end_row=r, end_column=max_col)
    ws.cell(row=r, column=half_col+1, value="TRƯỞNG ĐOÀN KIỂM TRA").font = st['font_sign_title']
    ws.cell(row=r, column=half_col+1).alignment = st['align_center']
    ws.row_dimensions[r].height = 14
    r += 1
    
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=half_col)
    ws.cell(row=r, column=1, value="(Ký, ghi rõ họ và tên)").font = st['font_sign_sub']
    ws.cell(row=r, column=1).alignment = st['align_center']
    
    ws.merge_cells(start_row=r, start_column=half_col+1, end_row=r, end_column=max_col)
    ws.cell(row=r, column=half_col+1, value="(Ký, ghi rõ họ và tên)").font = st['font_sign_sub']
    ws.cell(row=r, column=half_col+1).alignment = st['align_center']
    ws.row_dimensions[r].height = 13
    
    return r

# ==============================================================================
# 1. SHEET 1: PHIẾU KIỂM SOÁT TỔNG HỢP (1 Trang A4 Portrait)
# ==============================================================================
def create_sheet_summary_portrait(wb):
    ws = wb.active
    ws.title = "Phieu_kiem_soat_tong_hop"
    st = get_styles()
    setup_page_a4(ws, orientation='portrait', fit_width=1, fit_height=1)
    
    curr_r = add_header_block(ws, "PHIẾU KIỂM SOÁT TÌNH TRẠNG HỒ SƠ KIỂM TRA THUẾ", 
                              subtitle="(Rà soát tính đầy đủ hồ sơ theo Quy trình 4 bước chuẩn hóa trước phân tích chuyên sâu)",
                              orientation='portrait', max_col=7)
    
    curr_r = add_enterprise_info_block(ws, curr_r, max_col=7)
    
    # Table Header
    headers = [
        ("STT", 4.5),
        ("Bước quy trình", 13),
        ("Nhóm danh mục hồ sơ kiểm soát", 25),
        ("Thư mục lưu trữ", 17),
        ("Định dạng", 8.5),
        ("Tình trạng", 13),
        ("Ghi chú & Đánh giá rủi ro", 21)
    ]
    
    tbl_hdr_row = curr_r
    for col_idx, (hdr_text, w) in enumerate(headers, 1):
        cell = ws.cell(row=tbl_hdr_row, column=col_idx, value=hdr_text)
        cell.font = st['font_tbl_hdr']
        cell.fill = st['fill_tbl_hdr']
        cell.alignment = st['align_center']
        cell.border = st['border_header']
        ws.column_dimensions[get_column_letter(col_idx)].width = w
        
    ws.row_dimensions[tbl_hdr_row].height = 18
    curr_r += 1
    
    items_summary = [
        ("1", "Bước 1: Dữ liệu thô CQT", "Tờ khai XML (GTGT, TNDN, TNCN, BCTC)", "1.1. To khai XML/", ".xml", "[ ] Đủ  [ ] Thiếu", "Đủ 4 quý/năm trong niên độ (bắt buộc)"),
        ("2", "Bước 1: Dữ liệu thô CQT", "HĐĐT Cổng CQT & Bảng kê mua/bán", "1.2. Hoa don dien tu/", ".xlsx", "[ ] Đủ  [ ] Thiếu", "Gộp 2 sheet mua/bán, lọc nhãn Tổng cộng"),
        ("3", "Bước 1: Dữ liệu thô CQT", "Nghĩa vụ nộp thuế & Thông báo nợ", "1.3. Nghia vu nop/", ".xlsx/.pdf", "[ ] Đủ  [ ] Thiếu", "Chứng từ nộp NSNN, số dư nợ thuế"),
        
        ("4", "Bước 2: Hồ sơ NNT nộp", "Sổ kế toán tổng hợp & Sổ chi tiết TK", "2.1. So ke toan/", ".xlsx", "[ ] Đủ  [ ] Thiếu", "Sổ NKC, Sổ Cái, 111, 112, 131, 331 mở"),
        ("5", "Bước 2: Hồ sơ NNT nộp", "Sổ XNT Kho, Sổ 154, Giá thành, Định mức", "2.2. Kho - Gia thanh/", ".xlsx/.pdf", "[ ] Đủ  [ ] Thiếu", "Khớp BCTC [141] (TC3) và [02] (TC28)"),
        ("6", "Bước 2: Hồ sơ NNT nộp", "Sao kê Ngân hàng (Excel/PDF) & UNC", "2.3. Ngan hang/", ".xlsx/.pdf", "[ ] Đủ  [ ] Thiếu", "Bắt buộc có Excel + PDF dấu đỏ, UNC >= 20tr/5tr"),
        ("7", "Bước 2: Hồ sơ NNT nộp", "Bảng lương, Chấm công, HĐLĐ, Khấu hao", "2.4. Lao dong - TSCD/", ".xlsx/.pdf", "[ ] Đủ  [ ] Thiếu", "Khớp 622, QTT TNCN [23], Khấu hao TT 45"),
        
        ("8", "Bước 3: Dữ liệu bổ sung", "Sổ theo dõi CVCBRR F0/F1 & Công an", "3.1. Canh bao rui ro/", ".xlsx", "[ ] Đủ  [ ] Thiếu", "Quét 100% NCC mua vào với sổ F0/Công an"),
        ("9", "Bước 3: Dữ liệu bổ sung", "Danh bạ trạng thái ĐKKD của NCC", "3.2. Danh ba NNT/", ".xlsx", "[ ] Đủ  [ ] Thiếu", "Trạng thái 100% NCC (đóng MST, bỏ trốn)"),
        ("10", "Bước 3: Dữ liệu bổ sung", "Tra cứu TNCN đa nguồn & Trùng CCCD", "3.3. Tra cuu TNCN/", ".xlsx/.md", "[ ] Đủ  [ ] Thiếu", "Xác minh người có >= 2 nguồn, trùng CCCD"),
        ("11", "Bước 3: Dữ liệu bổ sung", "Giải trình của NNT & Chứng từ bổ sung", "3.4. Giai trinh/", ".pdf/.xlsx", "[ ] Đủ  [ ] K.có", "Văn bản giải trình khoảng trống dữ liệu"),
        
        ("12", "Bước 4: Kết quả phân tích", "Master Excel 7 Sheet chuẩn hóa", "4. KET QUA/", ".xlsx", "[ ] Đã lập", "Phan_tich_ho_so_cong_ty_<Ten>.xlsx"),
        ("13", "Bước 4: Kết quả phân tích", "6 File chuyên đề & Đề cương kiểm tra", "cau-thanh/", ".xlsx/.docx", "[ ] Đã lập", "Báo cáo chi tiết + Đề cương thực địa")
    ]
    
    for item in items_summary:
        stt, step_name, name, folder, fmt, status, note = item
        ws.row_dimensions[curr_r].height = 15.5
        
        c1 = ws.cell(row=curr_r, column=1, value=stt)
        c2 = ws.cell(row=curr_r, column=2, value=step_name)
        c3 = ws.cell(row=curr_r, column=3, value=name)
        c4 = ws.cell(row=curr_r, column=4, value=folder)
        c5 = ws.cell(row=curr_r, column=5, value=fmt)
        c6 = ws.cell(row=curr_r, column=6, value=status)
        c7 = ws.cell(row=curr_r, column=7, value=note)
        
        c1.alignment = st['align_center']
        c2.alignment = st['align_left']
        c3.alignment = st['align_left']
        c4.alignment = st['align_left']
        c5.alignment = st['align_center']
        c6.alignment = st['align_center']
        c7.alignment = st['align_left']
        
        c1.font = st['font_cell_bold']
        c2.font = st['font_cell_bold']
        c3.font = st['font_cell']
        c4.font = st['font_cell_italic']
        c5.font = st['font_cell']
        c6.font = st['font_status']
        c7.font = st['font_cell']
        
        fill = st['fill_alt'] if int(stt) % 2 == 0 else None
        for col_i in range(1, 8):
            cell_i = ws.cell(row=curr_r, column=col_i)
            cell_i.border = st['border_thin']
            if fill:
                cell_i.fill = fill
                
        curr_r += 1
        
    # Summary Box
    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
    summary_text = "KẾT LUẬN ĐÁNH GIÁ TÌNH TRẠNG HỒ SƠ:\n" \
                   "• Tổng số hạng mục kiểm soát: 13 nhóm (Bắt buộc: 09 nhóm, Tùy chọn: 04 nhóm)\n" \
                   "• Đánh giá: [  ] ĐỦ ĐIỀU KIỆN TIẾN HÀNH PHÂN TÍCH BƯỚC 4      [  ] CẦN YÊU CẦU NNT BỔ SUNG TÀI LIỆU\n" \
                   "• Nội dung thiếu trọng yếu: ................................................................................................................................................."
    sum_cell = ws.cell(row=curr_r, column=1, value=summary_text)
    sum_cell.font = st['font_cell_bold']
    sum_cell.alignment = st['align_left_top']
    sum_cell.fill = st['fill_highlight']
    ws.row_dimensions[curr_r].height = 42
    
    for c in range(1, 8):
        ws.cell(row=curr_r, column=c).border = st['border_box']
        
    curr_r += 1
    add_signatures(ws, curr_r, max_col=7)

# ==============================================================================
# 2. SHEET 2: CHECKLIST CHI TIẾT 35 MỤC (1 Trang A4 Landscape)
# ==============================================================================
def create_sheet_master_landscape(wb):
    ws = wb.create_sheet(title="Checklist_toan_bo_35_muc")
    st = get_styles()
    setup_page_a4(ws, orientation='landscape', fit_width=1, fit_height=1)
    
    max_col = 7
    curr_r = add_header_block(ws, "BẢNG CHECKLIST TỔNG HỢP KIỂM SOÁT 35 HẠNG MỤC HỒ SƠ KIỂM TRA THUẾ", 
                              subtitle="(Chuẩn hóa theo Quy trình 4 bước - Định dạng in ấn 01 trang A4 Landscape)",
                              orientation='landscape', max_col=max_col)
    
    curr_r = add_enterprise_info_block(ws, curr_r, max_col=max_col)
    
    headers = [
        ("Mã mục", 7.5),
        ("Tên loại hồ sơ / Tài liệu kiểm tra", 28),
        ("Thư mục lưu trữ", 20),
        ("Định dạng", 9.5),
        ("Tính chất", 10),
        ("Tình trạng", 14),
        ("Mục tiêu kiểm tra & Yêu cầu đối chiếu trọng tâm", 32)
    ]
    
    tbl_hdr_row = curr_r
    for col_idx, (hdr_text, w) in enumerate(headers, 1):
        cell = ws.cell(row=tbl_hdr_row, column=col_idx, value=hdr_text)
        cell.font = st['font_tbl_hdr']
        cell.fill = st['fill_tbl_hdr']
        cell.alignment = st['align_center']
        cell.border = st['border_header']
        ws.column_dimensions[get_column_letter(col_idx)].width = w
        
    ws.row_dimensions[tbl_hdr_row].height = 17
    curr_r += 1
    
    sections = [
        ("1. BƯỚC 1: DỮ LIỆU THÔ NGÀNH THUẾ (Cơ quan thuế chuẩn bị sẵn)", st['fill_step_1'], [
            ("1.1.1", "Tờ khai thuế GTGT các kỳ (Mẫu 842)", "1.1. To khai XML/", ".xml", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Đủ 4 quý/năm, đối chiếu DT bán ra và thuế phát sinh"),
            ("1.1.2", "Tờ khai bổ sung GTGT (KHBS)", "1.1. To khai XML/", ".xml", "Tùy chọn", "[ ] Đủ  [ ] K.có", "Kiểm tra các kỳ có điều chỉnh tăng/giảm thuế"),
            ("1.1.3", "Tờ khai Quyết toán TNDN (Mẫu 892)", "1.1. To khai XML/", ".xml", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Kèm phụ lục KQ HĐKD 03-1A/TNDN"),
            ("1.1.4", "Phụ lục Giao dịch liên kết (NĐ 132)", "1.1. To khai XML/", ".xml", "Tùy chọn", "[ ] Đủ  [ ] K.có", "Bắt buộc nếu có quan hệ liên kết hoặc vay vốn lớn"),
            ("1.1.5", "Tờ khai Quyết toán TNCN (Mẫu 953)", "1.1. To khai XML/", ".xml", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Kèm Bảng kê chi tiết lao động 05-1/BK-TNCN"),
            ("1.1.6", "Tờ khai khấu trừ TNCN quý (Mẫu 05/KK)", "1.1. To khai XML/", ".xml", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Đối chiếu Tiêu chí 22 (Quý vs Quyết toán năm)"),
            ("1.1.7", "Báo cáo tài chính năm (Mẫu 402)", "1.1. To khai XML/", ".xml", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Gồm CĐKT, KQKD, LCTT và Thuyết minh BCTC"),
            ("1.2.1", "Dữ liệu HĐĐT Cổng thông tin CQT", "1.2. Hoa don dien tu/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "File HDDT.xlsx trích xuất đầy đủ mua vào/bán ra"),
            ("1.2.2", "Bảng kê hóa đơn bán ra các kỳ", "1.2. Hoa don dien tu/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Đủ các quý để gộp thành Sheet 'ban ra'"),
            ("1.2.3", "Bảng kê hóa đơn mua vào các kỳ", "1.2. Hoa don dien tu/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Đủ các quý để gộp thành Sheet 'mua vao'"),
            ("1.3.1", "Lịch sử nộp thuế vào NSNN & Nợ thuế", "1.3. Nghia vu nop/", ".xlsx/.pdf", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Chứng từ nộp NSNN, thông báo nợ tiền thuế, tiền phạt")
        ]),
        ("2. BƯỚC 2: HỒ SƠ DO NGƯỜI NỘP THUẾ CUNG CẤP", st['fill_step_2'], [
            ("2.1.1", "Sổ Nhật ký chung & Sổ Cái tổng hợp", "2.1. So ke toan/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Bắt buộc file dữ liệu Excel mở có công thức"),
            ("2.1.2", "Sổ Chi tiết Quỹ tiền mặt (TK 1111)", "2.1. So ke toan/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Quét chi tiền mặt > 20tr (trước 07/25) và >= 5tr (từ 07/25)"),
            ("2.1.3", "Sổ Chi tiết TGNH (TK 1121)", "2.1. So ke toan/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Chi tiết từng tài khoản tại từng ngân hàng"),
            ("2.1.4", "Sổ Chi tiết Phải thu khách hàng (TK 131)", "2.1. So ke toan/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Theo dõi công nợ KH, rà soát DT chưa khai (131 Có)"),
            ("2.1.5", "Sổ Chi tiết Phải trả người bán (TK 331)", "2.1. So ke toan/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Theo dõi công nợ NCC lớn (giải tỏa TC20)"),
            ("2.1.6", "Sổ Vay mượn nội bộ (TK 336, 3388, 341)", "2.1. So ke toan/", ".xlsx/.pdf", "Tùy chọn", "[ ] Đủ  [ ] K.có", "Kèm Hợp đồng mượn tiền cá nhân/nội bộ không lãi"),
            ("2.2.1", "Sổ Tổng hợp & Chi tiết XNT Kho 152", "2.2. Kho - Gia thanh/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Kho NVL, khớp Chỉ tiêu [141] BCTC (giải tỏa TC3)"),
            ("2.2.2", "Sổ Tổng hợp & Chi tiết XNT Kho 155", "2.2. Kho - Gia thanh/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Kho Thành phẩm, khớp Chỉ tiêu [141] BCTC (TC3)"),
            ("2.2.3", "Sổ Chi phí 154 & Giá thành Z SP", "2.2. Kho - Gia thanh/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Bóc tách 621, 622, 627, khớp [02] Giá vốn (TC28)"),
            ("2.2.4", "Bảng định mức tiêu hao nguyên vật liệu", "2.2. Kho - Gia thanh/", ".pdf", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Định mức kỹ thuật có phê duyệt của Giám đốc"),
            ("2.3.1", "File Sao kê Ngân hàng dạng Excel", "2.3. Ngan hang/", ".xlsx/.csv", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Dữ liệu Internet Banking phục vụ chạy đối chiếu tự động"),
            ("2.3.2", "Bản Sao kê Ngân hàng có dấu đỏ", "2.3. Ngan hang/", ".pdf", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Bản scan có mộc tròn ngân hàng lưu trữ pháp lý"),
            ("2.3.3", "Chứng từ thanh toán gốc (UNC, GBN)", "2.3. Ngan hang/", ".pdf", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Ủy nhiệm chi cho HĐ lớn (>= 20tr/5tr) và NCC F0"),
            ("2.4.1", "Bảng thanh toán lương 12 tháng/năm", "2.4. Lao dong - TSCD/", ".xlsx/.pdf", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Có ký nhận hoặc UNC lương, khớp 622 và QTT TNCN"),
            ("2.4.2", "Bảng trích Khấu hao TSCĐ (TK 214)", "2.4. Lao dong - TSCD/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Khung thời gian khấu hao theo Thông tư 45/2013/TT-BTC")
        ]),
        ("3. BƯỚC 3: DỮ LIỆU BỔ SUNG & TRA CỨU QUẢN LÝ THUẾ", st['fill_step_3'], [
            ("3.1.1", "Sổ theo dõi CVCBRR F0 & F1", "3.1. Canh bao rui ro/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Nhận diện NCC phát hành hóa đơn bất hợp pháp"),
            ("3.1.2", "Sổ theo dõi Công văn Công an", "3.1. Canh bao rui ro/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Doanh nghiệp liên quan các vụ án kinh tế/hóa đơn"),
            ("3.2.1", "Danh bạ trạng thái ĐKKD của NCC", "3.2. Danh ba NNT/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Trạng thái 100% NCC (đóng MST, tạm ngừng, bỏ trốn)"),
            ("3.3.1", "Tra cứu TNCN đa nguồn toàn quốc", "3.3. Tra cuu TNCN/", ".xlsx", "Bắt buộc", "[ ] Đủ  [ ] Thiếu", "Tra cứu xem người lao động có thu nhập nơi khác không"),
            ("3.3.2", "Rà soát trùng số CCCD/CMND", "3.3. Tra cuu TNCN/", ".xlsx/.md", "Tùy chọn", "[ ] Đủ  [ ] K.có", "Bóc tách trường hợp 1 CCCD gán cho nhiều người"),
            ("3.4.1", "Giải trình & Hồ sơ bổ sung của NNT", "3.4. Giai trinh/", ".pdf/.xlsx", "Tùy chọn", "[ ] Đủ  [ ] K.có", "Công văn giải trình và chứng từ khắc phục khoảng trống")
        ]),
        ("4. BƯỚC 4: KẾT QUẢ PHÂN TÍCH & BÁO CÁO ĐẦU RA", st['fill_step_4'], [
            ("4.1", "Master Excel tổng hợp chung 7 Sheet", "4. KET QUA/", ".xlsx", "Bắt buộc", "[ ] Đã tạo", "Phan_tich_ho_so_cong_ty_<Ten>.xlsx (7 Sheet chuẩn)"),
            ("4.2", "6 File phân tích chuyên đề cấu thành", "cau-thanh/", ".xlsx", "Bắt buộc", "[ ] Đã tạo", "Bảng kê gộp, 3 chiều GTGT, NCC F0, TNCN, XNT, Tiền mặt"),
            ("4.3", "Đề cương kiểm tra tại trụ sở NNT", "4. KET QUA/", ".docx", "Bắt buộc", "[ ] Đã tạo", "Kế hoạch kiểm tra thực địa tại bàn và trụ sở doanh nghiệp")
        ])
    ]
    
    for sec_title, sec_fill, rows in sections:
        ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=max_col)
        sec_cell = ws.cell(row=curr_r, column=1, value=sec_title)
        sec_cell.font = st['font_sec_hdr']
        sec_cell.fill = sec_fill
        sec_cell.alignment = st['align_left']
        ws.row_dimensions[curr_r].height = 14.5
        for c in range(1, max_col+1):
            ws.cell(row=curr_r, column=c).border = st['border_thin']
        curr_r += 1
        
        for r_idx, (code, name, folder, fmt, nature, status, note) in enumerate(rows):
            ws.row_dimensions[curr_r].height = 13.5
            
            c1 = ws.cell(row=curr_r, column=1, value=code)
            c2 = ws.cell(row=curr_r, column=2, value=name)
            c3 = ws.cell(row=curr_r, column=3, value=folder)
            c4 = ws.cell(row=curr_r, column=4, value=fmt)
            c5 = ws.cell(row=curr_r, column=5, value=nature)
            c6 = ws.cell(row=curr_r, column=6, value=status)
            c7 = ws.cell(row=curr_r, column=7, value=note)
            
            c1.alignment = st['align_center']
            c2.alignment = st['align_left']
            c3.alignment = st['align_left']
            c4.alignment = st['align_center']
            c5.alignment = st['align_center']
            c6.alignment = st['align_center']
            c7.alignment = st['align_left']
            
            c1.font = st['font_cell_bold']
            c2.font = st['font_cell']
            c3.font = st['font_cell_italic']
            c4.font = st['font_cell']
            c5.font = st['font_cell_bold'] if nature == "Bắt buộc" else st['font_cell']
            c6.font = st['font_status']
            c7.font = st['font_cell']
            
            fill = st['fill_alt'] if r_idx % 2 == 1 else None
            for c in range(1, max_col+1):
                cell_i = ws.cell(row=curr_r, column=c)
                cell_i.border = st['border_thin']
                if fill:
                    cell_i.fill = fill
                    
            curr_r += 1
            
    # Add conclusion & signature
    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=max_col)
    sum_cell = ws.cell(row=curr_r, column=1, value="KẾT LUẬN:  [  ] ĐỦ ĐIỀU KIỆN TIẾN HÀNH PHÂN TÍCH CHUYÊN SÂU BƯỚC 4      [  ] CẦN PHÁT HÀNH PHIẾU YÊU CẦU BỔ SUNG TÀI LIỆU")
    sum_cell.font = st['font_cell_bold']
    sum_cell.fill = st['fill_highlight']
    sum_cell.alignment = st['align_left']
    ws.row_dimensions[curr_r].height = 16
    for c in range(1, max_col+1):
        ws.cell(row=curr_r, column=c).border = st['border_box']
    curr_r += 1
    
    add_signatures(ws, curr_r, max_col=max_col)

# ==============================================================================
# 3. SHEET 3..6: CHECKLIST RIÊNG TỪNG BƯỚC (Fit to 1 A4 Page)
# ==============================================================================
def create_step_sheet(wb, sheet_name, step_num, step_title, step_desc, headers, rows, orientation='landscape'):
    ws = wb.create_sheet(title=sheet_name)
    st = get_styles()
    setup_page_a4(ws, orientation=orientation, fit_width=1, fit_height=1)
    
    max_col = len(headers)
    curr_r = add_header_block(ws, f"BẢNG CHECKLIST BƯỚC {step_num}: {step_title.upper()}", 
                              subtitle=f"({step_desc} - Định dạng in 01 trang A4)",
                              orientation=orientation, max_col=max_col)
    
    curr_r = add_enterprise_info_block(ws, curr_r, max_col=max_col)
    
    tbl_hdr_row = curr_r
    for col_idx, (hdr_text, w) in enumerate(headers, 1):
        cell = ws.cell(row=tbl_hdr_row, column=col_idx, value=hdr_text)
        cell.font = st['font_tbl_hdr']
        cell.fill = st['fill_tbl_hdr']
        cell.alignment = st['align_center']
        cell.border = st['border_header']
        ws.column_dimensions[get_column_letter(col_idx)].width = w
        
    ws.row_dimensions[tbl_hdr_row].height = 18
    curr_r += 1
    
    for r_idx, r_data in enumerate(rows):
        ws.row_dimensions[curr_r].height = 16 if orientation == 'landscape' else 18
        fill = st['fill_alt'] if r_idx % 2 == 1 else None
        
        for col_idx, val in enumerate(r_data, 1):
            cell = ws.cell(row=curr_r, column=col_idx, value=val)
            cell.border = st['border_thin']
            if fill:
                cell.fill = fill
                
            if col_idx == 1 or col_idx == 4 or col_idx == 5:
                cell.alignment = st['align_center']
            else:
                cell.alignment = st['align_left']
                
            if col_idx == 1:
                cell.font = st['font_cell_bold']
            elif col_idx == 5:
                cell.font = st['font_status']
            else:
                cell.font = st['font_cell']
                
        curr_r += 1
        
    # Conclusion box
    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=max_col)
    concl_text = f"ĐÁNH GIÁ BƯỚC {step_num}:  [  ] ĐẠT YÊU CẦU KIỂM TRA      [  ] CẦN BỔ SUNG THÊM HỒ SƠ"
    concl_cell = ws.cell(row=curr_r, column=1, value=concl_text)
    concl_cell.font = st['font_cell_bold']
    concl_cell.fill = st['fill_highlight']
    concl_cell.alignment = st['align_left']
    ws.row_dimensions[curr_r].height = 16
    for c in range(1, max_col+1):
        ws.cell(row=curr_r, column=c).border = st['border_box']
    curr_r += 1
    
    add_signatures(ws, curr_r, max_col=max_col)
    return ws

def main():
    base_dir = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\thu-muc-mau-ho-so-doanh-nghiep"
    os.makedirs(base_dir, exist_ok=True)
    
    # 1. Create Master Workbook with all sheets
    wb_master = openpyxl.Workbook()
    create_sheet_summary_portrait(wb_master)
    create_sheet_master_landscape(wb_master)
    
    # Add Step 1 Sheet
    headers_step1 = [
        ("Mã mục", 8), ("Tên loại hồ sơ / Tài liệu", 30), ("Thư mục lưu trữ", 22),
        ("Định dạng", 10), ("Tình trạng", 16), ("Ghi chú & Yêu cầu đối chiếu", 34)
    ]
    rows_step1 = [
        ("1.1.1", "Tờ khai thuế GTGT các kỳ (Mẫu 842)", "1.1. To khai XML/", ".xml", "[ ] Đủ   [ ] Thiếu", "Đủ 4 quý/năm trong niên độ kiểm tra"),
        ("1.1.2", "Tờ khai bổ sung GTGT (KHBS)", "1.1. To khai XML/", ".xml", "[ ] Đủ   [ ] Không có", "Kiểm tra các kỳ có phát sinh điều chỉnh tăng/giảm thuế"),
        ("1.1.3", "Tờ khai Quyết toán TNDN (Mẫu 892)", "1.1. To khai XML/", ".xml", "[ ] Đủ   [ ] Thiếu", "Kèm phụ lục KQ HĐKD 03-1A/TNDN"),
        ("1.1.4", "Phụ lục Giao dịch liên kết (NĐ 132/2020)", "1.1. To khai XML/", ".xml", "[ ] Đủ   [ ] Không có", "Bắt buộc nếu NNT có quan hệ liên kết hoặc vay vốn lớn"),
        ("1.1.5", "Tờ khai Quyết toán TNCN (Mẫu 953)", "1.1. To khai XML/", ".xml", "[ ] Đủ   [ ] Thiếu", "Kèm Bảng kê chi tiết người lao động 05-1/BK-TNCN"),
        ("1.1.6", "Tờ khai khấu trừ TNCN quý (Mẫu 05/KK)", "1.1. To khai XML/", ".xml", "[ ] Đủ   [ ] Thiếu", "Phục vụ đối chiếu Tiêu chí rủi ro 22 (Quý vs Năm)"),
        ("1.1.7", "Báo cáo tài chính năm (Mẫu 402)", "1.1. To khai XML/", ".xml", "[ ] Đủ   [ ] Thiếu", "Gồm CĐKT, KQKD, LCTT và Thuyết minh BCTC"),
        ("1.2.1", "Dữ liệu HĐĐT Cổng thông tin CQT", "1.2. Hoa don dien tu/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "File HDDT.xlsx trích xuất đầy đủ mua vào, bán ra"),
        ("1.2.2", "Bảng kê hóa đơn bán ra các kỳ", "1.2. Hoa don dien tu/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Đủ các quý để gộp thành Sheet 'ban ra'"),
        ("1.2.3", "Bảng kê hóa đơn mua vào các kỳ", "1.2. Hoa don dien tu/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Đủ các quý để gộp thành Sheet 'mua vao'"),
        ("1.3.1", "Lịch sử nộp thuế vào NSNN & Nợ thuế", "1.3. Nghia vu nop/", ".xlsx/.pdf", "[ ] Đủ   [ ] Thiếu", "Chứng từ nộp tiền vào Kho bạc Nhà nước & thông báo nợ")
    ]
    create_step_sheet(wb_master, "Checklist_Buoc_1", "1", "DỮ LIỆU THÔ NGÀNH THUẾ", 
                      "Dữ liệu do Cơ quan thuế chuẩn bị sẵn trước kiểm tra tại bàn",
                      headers_step1, rows_step1, orientation='landscape')
    
    # Add Step 2 Sheet
    headers_step2 = [
        ("Mã mục", 8), ("Tên hồ sơ, sổ sách kế toán", 28), ("Thư mục lưu trữ", 22),
        ("Định dạng", 10), ("Tình trạng", 16), ("Mục tiêu kiểm tra & Yêu cầu giải tỏa rủi ro", 36)
    ]
    rows_step2 = [
        ("2.1.1", "Sổ Nhật ký chung & Sổ Cái tổng hợp", "2.1. So ke toan/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Bắt buộc dữ liệu Excel mở có công thức"),
        ("2.1.2", "Sổ Chi tiết Quỹ tiền mặt (TK 1111)", "2.1. So ke toan/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Quét chi tiền mặt > 20tr (trước 07/2025) và >= 5tr (từ 07/2025)"),
        ("2.1.3", "Sổ Chi tiết TGNH (TK 1121)", "2.1. So ke toan/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Chi tiết từng tài khoản mở tại từng ngân hàng"),
        ("2.1.4", "Sổ Chi tiết Phải thu KH (TK 131)", "2.1. So ke toan/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Chi tiết theo đối tượng khách hàng, rà soát DT chưa khai (131 Có)"),
        ("2.1.5", "Sổ Chi tiết Phải trả người bán (TK 331)", "2.1. So ke toan/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Chi tiết theo từng Nhà cung cấp (giải tỏa Tiêu chí 20)"),
        ("2.1.6", "Sổ Vay mượn nội bộ (TK 336, 3388, 341)", "2.1. So ke toan/", ".xlsx/.pdf", "[ ] Đủ   [ ] Không có", "Kèm theo Hợp đồng mượn tiền cá nhân/nội bộ không tính lãi"),
        ("2.2.1", "Sổ Tổng hợp & Chi tiết XNT Kho 152", "2.2. Kho - Gia thanh/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Kho Nguyên vật liệu, đối chiếu khớp Chỉ tiêu [141] BCTC (TC3)"),
        ("2.2.2", "Sổ Tổng hợp & Chi tiết XNT Kho 155", "2.2. Kho - Gia thanh/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Kho Thành phẩm, đối chiếu khớp Chỉ tiêu [141] BCTC (TC3)"),
        ("2.2.3", "Sổ Chi phí 154 & Bảng tính Giá thành (Z)", "2.2. Kho - Gia thanh/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Bóc tách 621, 622, 627, khớp Chỉ tiêu [02] Giá vốn BCTC (TC28)"),
        ("2.2.4", "Bảng định mức tiêu hao nguyên vật liệu", "2.2. Kho - Gia thanh/", ".pdf", "[ ] Đủ   [ ] Thiếu", "Bản định mức kỹ thuật có phê duyệt của Giám đốc doanh nghiệp"),
        ("2.3.1", "File Sao kê Ngân hàng dạng Excel", "2.3. Ngan hang/", ".xlsx/.csv", "[ ] Đủ   [ ] Thiếu", "Dữ liệu số hóa trích xuất từ Internet Banking đối chiếu tự động"),
        ("2.3.2", "Bản Sao kê Ngân hàng có dấu đỏ", "2.3. Ngan hang/", ".pdf", "[ ] Đủ   [ ] Thiếu", "Bản in scan có mộc tròn của từng ngân hàng để lưu trữ pháp lý"),
        ("2.3.3", "Chứng từ thanh toán gốc (UNC, GBN)", "2.3. Ngan hang/", ".pdf", "[ ] Đủ   [ ] Thiếu", "Ủy nhiệm chi đối với các HĐ >= 20tr (hoặc 5tr) và NCC rủi ro F0"),
        ("2.4.1", "Bảng thanh toán lương 12 tháng/năm", "2.4. Lao dong - TSCD/", ".xlsx/.pdf", "[ ] Đủ   [ ] Thiếu", "Có ký nhận hoặc UNC lương, khớp chi phí 622 và QTT TNCN [23]"),
        ("2.4.2", "Bảng chấm công 12 tháng/năm & HĐLĐ", "2.4. Lao dong - TSCD/", ".xlsx/.pdf", "[ ] Đủ   [ ] Thiếu", "Theo dõi ngày công lao động và hồ sơ nhân sự thực tế"),
        ("2.4.3", "Bảng trích Khấu hao TSCĐ (TK 214)", "2.4. Lao dong - TSCD/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Khung thời gian trích khấu hao theo Thông tư 45/2013/TT-BTC")
    ]
    create_step_sheet(wb_master, "Checklist_Buoc_2", "2", "HỒ SƠ DO NNT CUNG CẤP", 
                      "Hồ sơ, sổ sách kế toán do Doanh nghiệp nộp theo Phiếu yêu cầu",
                      headers_step2, rows_step2, orientation='landscape')
    
    # Add Step 3 Sheet
    headers_step3 = [
        ("Mã mục", 8), ("Tên dữ liệu tra cứu & Bổ sung", 30), ("Thư mục lưu trữ", 22),
        ("Định dạng", 10), ("Tình trạng", 16), ("Mục tiêu tra cứu & Kiểm soát rủi ro", 34)
    ]
    rows_step3 = [
        ("3.1.1", "Sổ theo dõi CVCBRR F0 (Người bán rủi ro)", "3.1. Canh bao rui ro/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Danh sách doanh nghiệp phát hành hóa đơn bất hợp pháp"),
        ("3.1.2", "Sổ theo dõi CVCBRR F1 (Doanh nghiệp mua)", "3.1. Canh bao rui ro/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Danh sách doanh nghiệp mua hàng F0 bị chuyển điều tra"),
        ("3.1.3", "Sổ theo dõi Công văn Công an", "3.1. Canh bao rui ro/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Vụ án kinh tế, mua bán hóa đơn do Cơ quan CSĐT thụ lý"),
        ("3.2.1", "Dữ liệu Danh bạ trạng thái ĐKKD NNT", "3.2. Danh ba NNT/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Trạng thái 100% NCC (đang HĐ, đóng MST, bỏ trốn, giải thể)"),
        ("3.3.1", "Dữ liệu tra cứu TNCN đa nguồn toàn quốc", "3.3. Tra cuu TNCN/", ".xlsx", "[ ] Đủ   [ ] Thiếu", "Tra cứu 100% lao động xem có thu nhập và nộp thuế nơi khác"),
        ("3.3.2", "Kết quả rà soát trùng số CCCD/CMND", "3.3. Tra cuu TNCN/", ".xlsx/.md", "[ ] Đủ   [ ] Không có", "Bóc tách các trường hợp 1 số CCCD gán cho nhiều họ tên"),
        ("3.4.1", "Văn bản giải trình của Doanh nghiệp", "3.4. Giai trinh/", ".pdf", "[ ] Đủ   [ ] Không có", "Công văn giải trình theo Phiếu yêu cầu của Đoàn kiểm tra"),
        ("3.4.2", "Hồ sơ bổ sung theo Phiếu yêu cầu", "3.4. Giai trinh/", ".xlsx/.pdf", "[ ] Đủ   [ ] Không có", "Chứng từ khắc phục khoảng trống dữ liệu sau rà soát Bước 1 & 2")
    ]
    create_step_sheet(wb_master, "Checklist_Buoc_3", "3", "DỮ LIỆU BỔ SUNG & TRA CỨU QLT", 
                      "Dữ liệu tra cứu ngành Thuế, cảnh báo rủi ro & Hồ sơ giải trình bổ sung",
                      headers_step3, rows_step3, orientation='landscape')
    
    # Add Step 4 Sheet
    headers_step4 = [
        ("Mã mục", 8), ("Tên sản phẩm kết xuất", 30), ("Vị trí lưu trữ", 22),
        ("Định dạng", 10), ("Tình trạng", 16), ("Quy chuẩn nghiệm thu & Mục đích sử dụng", 34)
    ]
    rows_step4 = [
        ("4.1", "Master Excel tổng hợp chung 7 Sheet", "4. KET QUA PHAN TICH/", ".xlsx", "[ ] Đã tạo", "Phan_tich_ho_so_cong_ty_<Ten>.xlsx (7 Sheet chuẩn hóa)"),
        ("4.2.1", "Bảng kê gộp mua vào bán ra chuẩn", "cau-thanh/", ".xlsx", "[ ] Đã tạo", "2 sheet Mua vào/Bán ra đã lọc dòng nhãn Tổng cộng"),
        ("4.2.2", "Bảng đối chiếu 3 chiều GTGT - TNDN", "cau-thanh/", ".xlsx", "[ ] Đã tạo", "Khớp đúng 100% doanh thu 10 năm theo chuẩn 5 bảng CVCBRR"),
        ("4.2.3", "Bảng xét NCC mua vào với CBRR & ĐKKD", "cau-thanh/", ".xlsx", "[ ] Đã tạo", "Xét 100% NCC theo F0, F1, Công an và chi tiết hóa đơn"),
        ("4.2.4", "Bảng phân tích TNCN & 28 Tiêu chí rủi ro", "cau-thanh/", ".xlsx", "[ ] Đã tạo", "Chấm điểm 28 tiêu chí rủi ro và rà soát lao động đa nguồn"),
        ("4.2.5", "Bảng phân tích XNT Kho & BCTC [141]", "cau-thanh/", ".xlsx", "[ ] Đã tạo", "Giải tỏa Tiêu chí 3 (Tồn kho) và Tiêu chí 28 (Giá vốn)"),
        ("4.2.6", "Bảng quét chi tiền mặt TK 111 & Ngân hàng", "cau-thanh/", ".xlsx", "[ ] Đã tạo", "Danh sách chi tiền mặt > 20tr / >= 5tr và cụm HĐ trong ngày"),
        ("4.3", "Thư mục sao lưu kỹ thuật", "sao-luu/", ".json/.md", "[ ] Đã lưu", "Lưu trữ dữ liệu thô json và bản nháp phân tích kỹ thuật"),
        ("4.4", "Đề cương kiểm tra tại trụ sở NNT", "4. KET QUA PHAN TICH/", ".docx", "[ ] Đã tạo", "Bản kế hoạch chi tiết các trọng tâm cần kiểm tra thực địa")
    ]
    create_step_sheet(wb_master, "Checklist_Buoc_4", "4", "KẾT QUẢ PHÂN TÍCH & BÁO CÁO", 
                      "Sản phẩm phân tích chuyên sâu, Master Excel 7 Sheet và Đề cương kiểm tra",
                      headers_step4, rows_step4, orientation='landscape')
    
    master_path = os.path.join(base_dir, "CHECKLIST_TONG_HOP_HO_SO.xlsx")
    wb_master.save(master_path)
    print(f"-> Đã tạo Master Workbook: {master_path}")
    
    # 2. Create standalone Excel files for each step folder
    step_folders = [
        ("1. DU LIEU THO NGANH THUE", "CHECKLIST_BUOC_1_DU_LIEU_THO_CQT.xlsx", "1", "DỮ LIỆU THÔ NGÀNH THUẾ", 
         "Dữ liệu do Cơ quan thuế chuẩn bị sẵn trước kiểm tra tại bàn", headers_step1, rows_step1),
        ("2. HO SO NNT CUNG CAP", "CHECKLIST_BUOC_2_HO_SO_NNT_CUNG_CAP.xlsx", "2", "HỒ SƠ DO NNT CUNG CẤP", 
         "Hồ sơ, sổ sách kế toán do Doanh nghiệp nộp theo Phiếu yêu cầu", headers_step2, rows_step2),
        ("3. DU LIEU BO SUNG & TRA CUU QLT", "CHECKLIST_BUOC_3_DU_LIEU_BO_SUNG_QLT.xlsx", "3", "DỮ LIỆU BỔ SUNG & TRA CỨU QLT", 
         "Dữ liệu tra cứu ngành Thuế, cảnh báo rủi ro & Hồ sơ giải trình bổ sung", headers_step3, rows_step3),
        ("4. KET QUA PHAN TICH", "CHECKLIST_BUOC_4_KET_QUA_PHAN_TICH.xlsx", "4", "KẾT QUẢ PHÂN TICH & BÁO CÁO", 
         "Sản phẩm phân tích chuyên sâu, Master Excel 7 Sheet và Đề cương kiểm tra", headers_step4, rows_step4)
    ]
    
    for folder_name, file_name, step_num, title, desc, hdrs, rws in step_folders:
        wb_step = openpyxl.Workbook()
        ws_step = wb_step.active
        ws_step.title = f"Checklist_Buoc_{step_num}"
        st = get_styles()
        setup_page_a4(ws_step, orientation='landscape', fit_width=1, fit_height=1)
        
        max_col = len(hdrs)
        curr_r = add_header_block(ws_step, f"BẢNG CHECKLIST BƯỚC {step_num}: {title.upper()}", 
                                  subtitle=f"({desc} - Định dạng in 01 trang A4)",
                                  orientation='landscape', max_col=max_col)
        curr_r = add_enterprise_info_block(ws_step, curr_r, max_col=max_col)
        
        tbl_hdr_row = curr_r
        for col_idx, (hdr_text, w) in enumerate(hdrs, 1):
            cell = ws_step.cell(row=tbl_hdr_row, column=col_idx, value=hdr_text)
            cell.font = st['font_tbl_hdr']
            cell.fill = st['fill_tbl_hdr']
            cell.alignment = st['align_center']
            cell.border = st['border_header']
            ws_step.column_dimensions[get_column_letter(col_idx)].width = w
            
        ws_step.row_dimensions[tbl_hdr_row].height = 18
        curr_r += 1
        
        for r_idx, r_data in enumerate(rws):
            ws_step.row_dimensions[curr_r].height = 16
            fill = st['fill_alt'] if r_idx % 2 == 1 else None
            
            for col_idx, val in enumerate(r_data, 1):
                cell = ws_step.cell(row=curr_r, column=col_idx, value=val)
                cell.border = st['border_thin']
                if fill:
                    cell.fill = fill
                    
                if col_idx == 1 or col_idx == 4 or col_idx == 5:
                    cell.alignment = st['align_center']
                else:
                    cell.alignment = st['align_left']
                    
                if col_idx == 1:
                    cell.font = st['font_cell_bold']
                elif col_idx == 5:
                    cell.font = st['font_status']
                else:
                    cell.font = st['font_cell']
                    
            curr_r += 1
            
        # Conclusion box
        ws_step.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=max_col)
        concl_text = f"ĐÁNH GIÁ BƯỚC {step_num}:  [  ] ĐẠT YÊU CẦU KIỂM TRA      [  ] CẦN BỔ SUNG THÊM HỒ SƠ"
        concl_cell = ws_step.cell(row=curr_r, column=1, value=concl_text)
        concl_cell.font = st['font_cell_bold']
        concl_cell.fill = st['fill_highlight']
        concl_cell.alignment = st['align_left']
        ws_step.row_dimensions[curr_r].height = 16
        for c in range(1, max_col+1):
            ws_step.cell(row=curr_r, column=c).border = st['border_box']
        curr_r += 1
        
        add_signatures(ws_step, curr_r, max_col=max_col)
        
        file_path = os.path.join(base_dir, folder_name, file_name)
        wb_step.save(file_path)
        print(f"-> Đã tạo file bước: {file_path}")

if __name__ == "__main__":
    main()
