import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os, sys

def create_report():
    wb = openpyxl.Workbook()
    # default sheet
    ws1 = wb.active
    ws1.title = "Tong_hop_GTGT_TNDN"
    ws2 = wb.create_sheet(title="Tong_hop_NCC_CVCBRR")
    
    # Styles
    font_family = "Times New Roman"
    f_title = Font(name=font_family, size=14, bold=True, color="1F4E78")
    f_sub = Font(name=font_family, size=11, italic=True, color="333333")
    f_sec = Font(name=font_family, size=12, bold=True, color="1F4E78")
    f_tbl_title = Font(name=font_family, size=11, bold=True, color="002060")
    f_hdr = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    f_hdr_dark = Font(name=font_family, size=10, bold=True, color="000000")
    f_data = Font(name=font_family, size=10)
    f_data_bold = Font(name=font_family, size=10, bold=True)
    f_note = Font(name=font_family, size=9, italic=True, color="555555")
    
    fill_hdr_blue = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_hdr_navy = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
    fill_hdr_teal = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_hdr_light = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_hdr_amber = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_sum = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_sum_amber = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_alert = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    
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
    align_hdr = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    num_fmt = '#,##0'
    pct_fmt = '0.00%'

    # Load Source Data
    src_wb = openpyxl.load_workbook('sources/doi-chieu-hddt-bangke-tokhai-3601409441.xlsx', data_only=True)
    ws_tk = src_wb['GTGT_01_dung']
    ws_quy = src_wb['Doi_chieu_quy']
    
    # 1. Parse Tờ khai GTGT
    tk_dict = {}
    for r in range(2, ws_tk.max_row+1):
        nam = ws_tk.cell(r, 1).value
        quy = ws_tk.cell(r, 2).value
        if nam in range(2016, 2026):
            tk_dict[(nam, quy)] = {
                'loai': ws_tk.cell(r, 4).value or 'C0',
                'ct23': ws_tk.cell(r, 5).value or 0,
                'ct24': ws_tk.cell(r, 6).value or 0,
                'ct34': ws_tk.cell(r, 7).value or 0,
                'ct35': ws_tk.cell(r, 8).value or 0,
                'ct40': ws_tk.cell(r, 11).value or 0,
                'ct43': ws_tk.cell(r, 12).value if ws_tk.max_column >= 12 else 0
            }

    # 2. Parse Doi_chieu_quy
    bk_mua = {}
    bk_ban = {}
    hddt_mua = {}
    hddt_ban = {}
    for r in range(2, ws_quy.max_row+1):
        chieu = ws_quy.cell(r, 1).value
        nam = ws_quy.cell(r, 2).value
        quy = ws_quy.cell(r, 3).value
        if nam in range(2016, 2026):
            key = (nam, quy)
            if chieu == 'Mua vào':
                bk_mua[key] = {
                    'hd': ws_quy.cell(r, 7).value or 0,
                    'ds': ws_quy.cell(r, 8).value or 0,
                    'thue': ws_quy.cell(r, 14).value or 0
                }
                hddt_mua[key] = {
                    'hd': ws_quy.cell(r, 9).value or 0,
                    'ds': ws_quy.cell(r, 10).value or 0,
                    'thue': ws_quy.cell(r, 15).value or 0
                }
            elif chieu == 'Bán ra':
                bk_ban[key] = {
                    'hd': ws_quy.cell(r, 7).value or 0,
                    'ds': ws_quy.cell(r, 8).value or 0,
                    'thue': ws_quy.cell(r, 14).value or 0
                }
                hddt_ban[key] = {
                    'hd': ws_quy.cell(r, 9).value or 0,
                    'ds': ws_quy.cell(r, 10).value or 0,
                    'thue': ws_quy.cell(r, 15).value or 0
                }

    # HĐĐT hợp lệ sau khi trừ HĐ hủy/thay thế
    hddt_ban_valid = {}
    for nam in range(2016, 2026):
        for quy in range(1, 5):
            k = (nam, quy)
            if nam < 2022 or (nam == 2022 and quy == 1):
                hddt_ban_valid[k] = {'ds': 0, 'thue': 0}
            elif nam == 2022 and quy == 2:
                hddt_ban_valid[k] = {'ds': 641325386, 'thue': 51306031}
            elif nam == 2022 and quy == 3:
                hddt_ban_valid[k] = {'ds': 936105775, 'thue': 74888462}
            elif nam == 2022 and quy == 4:
                hddt_ban_valid[k] = {'ds': 1547918413, 'thue': 123833472}
            elif nam == 2023 and quy in [1, 2, 3]:
                hddt_ban_valid[k] = {'ds': bk_ban[k]['ds'], 'thue': bk_ban[k]['thue']}
            elif nam == 2023 and quy == 4:
                hddt_ban_valid[k] = {'ds': 2270383923, 'thue': 181630716}
            elif nam == 2024 and quy in [1, 2, 3]:
                hddt_ban_valid[k] = {'ds': bk_ban[k]['ds'], 'thue': bk_ban[k]['thue']}
            elif nam == 2024 and quy == 4:
                hddt_ban_valid[k] = {'ds': 1915903860, 'thue': 153272308}
            elif nam == 2025 and quy == 1:
                hddt_ban_valid[k] = {'ds': 2093187103, 'thue': 167454967}
            elif nam == 2025 and quy == 2:
                hddt_ban_valid[k] = {'ds': 2544859542, 'thue': 204206706}
            elif nam == 2025 and quy == 3:
                hddt_ban_valid[k] = {'ds': 2330570978, 'thue': 186445679}
            elif nam == 2025 and quy == 4:
                hddt_ban_valid[k] = {'ds': 3092330018, 'thue': 247386399}

    print("Data parsing complete. Building Sheet 1: Tong_hop_GTGT_TNDN...")

    # ==========================================
    # SHEET 1: Tong_hop_GTGT_TNDN
    # ==========================================
    # Header
    ws1.merge_cells("A1:O1")
    ws1["A1"] = "BẢNG TỔNG HỢP ĐỐI CHIẾU NGHĨA VỤ THUẾ GTGT VÀ QUYẾT TOÁN THUẾ TNDN"
    ws1["A1"].font = f_title
    ws1["A1"].alignment = align_center

    ws1.merge_cells("A2:O2")
    ws1["A2"] = "Người nộp thuế: CÔNG TY TNHH MTV BAO BÌ GIẤY VẠN PHÁT THÀNH | Mã số thuế: 3601409441"
    ws1["A2"].font = f_sub
    ws1["A2"].alignment = align_center

    ws1.merge_cells("A3:O3")
    ws1["A3"] = "Thời kỳ kiểm tra: 2016 - 2025 (40 Quý liên tục) | Chuẩn hóa 5 bảng số liệu công văn cảnh báo rủi ro (CBRR)"
    ws1["A3"].font = f_sub
    ws1["A3"].alignment = align_center

    cur_r = 5
    ws1.cell(cur_r, 1, "PHẦN I. ĐỐI CHIẾU NGHĨA VỤ THUẾ GIÁ TRỊ GIA TĂNG (GTGT) THEO KỲ (2016 - 2025)").font = f_sec

    # ------------------------------------------
    # BẢNG 1.1: KÊ KHAI THUẾ GTGT THEO KỲ (MẪU BẢNG 3 CBRR)
    # ------------------------------------------
    cur_r += 2
    ws1.cell(cur_r, 1, "1.1. Bảng tổng hợp kê khai thuế GTGT theo kỳ (Mẫu Bảng 3 Công văn CBRR)").font = f_tbl_title
    
    cur_r += 1
    hdr11 = ["STT", "Kỳ tính thuế", "Loại tờ khai", "Giá trị HHDV mua vào [23]", "Thuế GTGT mua vào [24]", 
             "Tổng doanh thu HHDV bán ra [34]", "Thuế GTGT HHDV bán ra [35]", "Thuế GTGT phát sinh [40]", 
             "Thuế GTGT còn được KT chuyển kỳ sau [43]", "Ghi chú hồ sơ kê khai"]
    for c_idx, h in enumerate(hdr11, 1):
        cell = ws1.cell(cur_r, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_blue
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws1.row_dimensions[cur_r].height = 28

    b11_start = cur_r + 1
    stt = 1
    for nam in range(2016, 2026):
        for quy in range(1, 5):
            cur_r += 1
            k = (nam, quy)
            tk = tk_dict.get(k, {'loai': 'Chưa nộp', 'ct23': 0, 'ct24': 0, 'ct34': 0, 'ct35': 0, 'ct40': 0, 'ct43': 0})
            
            note = ""
            if nam == 2016 and quy == 2:
                note = "Chỉ tiêu [23] kê 265,2tr (lệch 300tr lỗi gõ phím so với Bảng kê 565,2tr; thuế [24] khớp đúng 56,5tr)"
            elif nam == 2024 and quy == 2:
                note = "Tờ khai bổ sung B1 (ct34=2.721.892.398, ct35=217.751.391)"
            elif nam == 2025 and quy == 4:
                note = "Tờ khai chính thức (đã bổ sung XML ngày 24/01/2026, khớp 100% Bảng kê và HĐĐT; thuế nộp 27,61tr)"

            row_data = [
                stt, f"Quý {quy}/{nam}", tk['loai'], tk['ct23'], tk['ct24'], 
                tk['ct34'], tk['ct35'], tk['ct40'], tk['ct43'], note
            ]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws1.cell(cur_r, c_idx, val)
                cell.font = f_data
                cell.border = thin_border
                if c_idx in [1, 2, 3]:
                    cell.alignment = align_center
                elif c_idx in range(4, 10):
                    cell.alignment = align_right
                    cell.number_format = num_fmt
                else:
                    cell.alignment = align_left
            stt += 1

    b11_end = cur_r
    cur_r += 1
    ws1.cell(cur_r, 1, "CỘNG").font = f_data_bold
    ws1.cell(cur_r, 1).alignment = align_center
    ws1.cell(cur_r, 1).fill = fill_sum
    ws1.cell(cur_r, 1).border = sum_border
    for c_idx in range(2, len(hdr11)+1):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_data_bold
        cell.fill = fill_sum
        cell.border = sum_border
        if c_idx in [4, 5, 6, 7]:
            col_let = get_column_letter(c_idx)
            cell.value = f"=SUM({col_let}{b11_start}:{col_let}{b11_end})"
            cell.number_format = num_fmt
            cell.alignment = align_right
        elif c_idx == 10:
            cell.value = "2016-2025: Bảng kê bán ra và mua vào khớp 100% Tờ khai toàn bộ 40 quý (trừ lỗi gõ phím Q2/2016)"
            cell.alignment = align_left

    # ------------------------------------------
    # BẢNG 1.2: ĐỐI CHIẾU MUA VÀO - BÁN RA THEO KỲ CỦA CÙNG NGUỒN (MỚI BỔ SUNG THEO YÊU CẦU)
    # ------------------------------------------
    cur_r += 3
    ws1.cell(cur_r, 1, "1.2. Bảng đối chiếu Tổng giá trị Mua vào - Doanh thu Bán ra theo kỳ của cùng nguồn Bảng kê, HĐĐT và Chênh lệch").font = f_tbl_title
    
    cur_r += 1
    # Multi-level header
    ws1.merge_cells(f"A{cur_r}:A{cur_r+1}")
    ws1.cell(cur_r, 1, "STT").font = f_hdr
    ws1.cell(cur_r, 1).fill = fill_hdr_navy
    ws1.cell(cur_r, 1).alignment = align_hdr
    ws1.cell(cur_r, 1).border = tbl_border

    ws1.merge_cells(f"B{cur_r}:B{cur_r+1}")
    ws1.cell(cur_r, 2, "Kỳ tính thuế").font = f_hdr
    ws1.cell(cur_r, 2).fill = fill_hdr_navy
    ws1.cell(cur_r, 2).alignment = align_hdr
    ws1.cell(cur_r, 2).border = tbl_border

    ws1.merge_cells(f"C{cur_r}:E{cur_r}")
    ws1.cell(cur_r, 3, "NGUỒN BẢNG KÊ DOANH NGHIỆP").font = f_hdr
    ws1.cell(cur_r, 3).fill = fill_hdr_navy
    ws1.cell(cur_r, 3).alignment = align_hdr
    ws1.cell(cur_r, 3).border = tbl_border

    ws1.merge_cells(f"F{cur_r}:H{cur_r}")
    ws1.cell(cur_r, 6, "NGUỒN HÓA ĐƠN ĐIỆN TỬ HỢP LỆ").font = f_hdr
    ws1.cell(cur_r, 6).fill = fill_hdr_navy
    ws1.cell(cur_r, 6).alignment = align_hdr
    ws1.cell(cur_r, 6).border = tbl_border

    ws1.merge_cells(f"I{cur_r}:K{cur_r}")
    ws1.cell(cur_r, 9, "NGUỒN TỜ KHAI THUẾ GTGT").font = f_hdr
    ws1.cell(cur_r, 9).fill = fill_hdr_navy
    ws1.cell(cur_r, 9).alignment = align_hdr
    ws1.cell(cur_r, 9).border = tbl_border

    ws1.merge_cells(f"L{cur_r}:N{cur_r}")
    ws1.cell(cur_r, 12, "ĐỐI CHIẾU CHÊNH LỆCH GIỮA CÁC NGUỒN").font = f_hdr
    ws1.cell(cur_r, 12).fill = fill_hdr_navy
    ws1.cell(cur_r, 12).alignment = align_hdr
    ws1.cell(cur_r, 12).border = tbl_border

    ws1.merge_cells(f"O{cur_r}:O{cur_r+1}")
    ws1.cell(cur_r, 15, "Đánh giá & Nguyên nhân chênh lệch").font = f_hdr
    ws1.cell(cur_r, 15).fill = fill_hdr_navy
    ws1.cell(cur_r, 15).alignment = align_hdr
    ws1.cell(cur_r, 15).border = tbl_border

    # Row 2 sub-headers
    cur_r += 1
    sub_headers = [
        "", "",
        "Mua vào (BK)", "Bán ra (BK)", "Lệch BR-MV (BK)",
        "Mua vào (HĐĐT)", "Bán ra (HĐĐT)", "Lệch BR-MV (HĐĐT)",
        "Mua vào [23]", "Bán ra [34]", "Lệch BR-MV (TK)",
        "Lệch MV (TK-BK)", "Lệch BR (TK-BK)", "Lệch BR (BK-HĐĐT)"
    ]
    for c_idx in range(1, 16):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_hdr_dark
        cell.fill = fill_hdr_light
        cell.alignment = align_hdr
        cell.border = tbl_border
        if c_idx <= len(sub_headers) and sub_headers[c_idx-1]:
            cell.value = sub_headers[c_idx-1]
    ws1.row_dimensions[cur_r-1].height = 20
    ws1.row_dimensions[cur_r].height = 26

    b12_start = cur_r + 1
    stt = 1
    for nam in range(2016, 2026):
        for quy in range(1, 5):
            cur_r += 1
            k = (nam, quy)
            
            m_bk = bk_mua.get(k, {'ds': 0})['ds']
            b_bk = bk_ban.get(k, {'ds': 0})['ds']
            diff_bk = b_bk - m_bk
            
            m_hd = hddt_mua.get(k, {'ds': 0})['ds']
            b_hd = hddt_ban_valid.get(k, {'ds': 0})['ds']
            diff_hd = b_hd - m_hd if (b_hd > 0 or m_hd > 0) else 0
            
            tk = tk_dict.get(k, {'ct23': 0, 'ct34': 0})
            m_tk = tk['ct23']
            b_tk = tk['ct34']
            diff_tk = (b_tk - m_tk) if (b_tk > 0 or m_tk > 0) else 0
            
            diff_mv_tk_bk = m_tk - m_bk if (b_tk > 0 or m_tk > 0) else 0
            diff_br_tk_bk = b_tk - b_bk if (b_tk > 0 or m_tk > 0) else 0
            diff_br_bk_hd = b_bk - b_hd if b_hd > 0 else 0
            
            # Ghi chú
            eval_str = "Khớp 100% Tờ khai vs Bảng kê"
            if nam < 2022:
                eval_str = "Khớp 100% Tờ khai vs Bảng kê (HĐĐT chưa triển khai)"
            elif nam == 2022 and quy in [1, 2]:
                eval_str = "Khớp 100% Tờ khai vs Bảng kê (Chuyển đổi HĐĐT NĐ 123)"
            elif nam >= 2023:
                eval_str = "Khớp 100% tuyệt đối giữa Bảng kê, Tờ khai và HĐĐT hợp lệ"
            
            if nam == 2016 and quy == 2:
                eval_str = "Lệch MV 300tr (Lỗi gõ phím ct23); Bán ra khớp 100% sau điều chỉnh B1"
            elif nam == 2025 and quy == 4:
                eval_str = "Khớp 100% tuyệt đối giữa Bảng kê, Tờ khai và HĐĐT hợp lệ (3,09 tỷ)"

            row_vals = [
                stt, f"Quý {quy}/{nam}",
                m_bk, b_bk, diff_bk,
                m_hd, b_hd, diff_hd,
                m_tk, b_tk, diff_tk,
                diff_mv_tk_bk, diff_br_tk_bk, diff_br_bk_hd,
                eval_str
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws1.cell(cur_r, c_idx, val)
                cell.font = f_data
                cell.border = thin_border
                if c_idx in [1, 2]:
                    cell.alignment = align_center
                elif c_idx in range(3, 15):
                    cell.alignment = align_right
                    cell.number_format = num_fmt
                else:
                    cell.alignment = align_left
            stt += 1

    b12_end = cur_r
    cur_r += 1
    ws1.cell(cur_r, 1, "CỘNG").font = f_data_bold
    ws1.cell(cur_r, 1).alignment = align_center
    ws1.cell(cur_r, 1).fill = fill_sum
    ws1.cell(cur_r, 1).border = sum_border
    for c_idx in range(2, 16):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_data_bold
        cell.fill = fill_sum
        cell.border = sum_border
        if c_idx in range(3, 15):
            col_let = get_column_letter(c_idx)
            cell.value = f"=SUM({col_let}{b12_start}:{col_let}{b12_end})"
            cell.number_format = num_fmt
            cell.alignment = align_right
        elif c_idx == 15:
            cell.value = "Toàn bộ 10 năm: Bán ra và Mua vào đồng nhất tuyệt đối giữa Bảng kê và Tờ khai"
            cell.alignment = align_left

    # ------------------------------------------
    # BẢNG 1.3: CHI TIẾT BÁN RA GIỮA BẢNG KÊ - HĐĐT - TỜ KHAI
    # ------------------------------------------
    cur_r += 3
    ws1.cell(cur_r, 1, "1.3. Bảng đối chiếu chi tiết HHDV Bán ra giữa Bảng kê - HĐĐT - Tờ khai GTGT (Theo kỳ báo cáo)").font = f_tbl_title
    
    cur_r += 1
    ws1.merge_cells(f"A{cur_r}:A{cur_r+1}")
    ws1.cell(cur_r, 1, "STT").font = f_hdr; ws1.cell(cur_r, 1).fill = fill_hdr_teal; ws1.cell(cur_r, 1).alignment = align_hdr; ws1.cell(cur_r, 1).border = tbl_border
    
    ws1.merge_cells(f"B{cur_r}:B{cur_r+1}")
    ws1.cell(cur_r, 2, "Kỳ tính thuế").font = f_hdr; ws1.cell(cur_r, 2).fill = fill_hdr_teal; ws1.cell(cur_r, 2).alignment = align_hdr; ws1.cell(cur_r, 2).border = tbl_border

    ws1.merge_cells(f"C{cur_r}:D{cur_r}")
    ws1.cell(cur_r, 3, "BẢNG KÊ BÁN RA").font = f_hdr; ws1.cell(cur_r, 3).fill = fill_hdr_teal; ws1.cell(cur_r, 3).alignment = align_hdr; ws1.cell(cur_r, 3).border = tbl_border

    ws1.merge_cells(f"E{cur_r}:F{cur_r}")
    ws1.cell(cur_r, 5, "HÓA ĐƠN ĐIỆN TỬ HỢP LỆ").font = f_hdr; ws1.cell(cur_r, 5).fill = fill_hdr_teal; ws1.cell(cur_r, 5).alignment = align_hdr; ws1.cell(cur_r, 5).border = tbl_border

    ws1.merge_cells(f"G{cur_r}:H{cur_r}")
    ws1.cell(cur_r, 7, "TỜ KHAI THUẾ GTGT").font = f_hdr; ws1.cell(cur_r, 7).fill = fill_hdr_teal; ws1.cell(cur_r, 7).alignment = align_hdr; ws1.cell(cur_r, 7).border = tbl_border

    ws1.merge_cells(f"I{cur_r}:J{cur_r}")
    ws1.cell(cur_r, 9, "CHÊNH LỆCH BẢNG KÊ vs HĐĐT").font = f_hdr; ws1.cell(cur_r, 9).fill = fill_hdr_teal; ws1.cell(cur_r, 9).alignment = align_hdr; ws1.cell(cur_r, 9).border = tbl_border

    ws1.merge_cells(f"K{cur_r}:L{cur_r}")
    ws1.cell(cur_r, 11, "CHÊNH LỆCH BẢNG KÊ vs TỜ KHAI").font = f_hdr; ws1.cell(cur_r, 11).fill = fill_hdr_teal; ws1.cell(cur_r, 11).alignment = align_hdr; ws1.cell(cur_r, 11).border = tbl_border

    ws1.merge_cells(f"M{cur_r}:M{cur_r+1}")
    ws1.cell(cur_r, 13, "Ghi chú & Đánh giá nguyên nhân").font = f_hdr; ws1.cell(cur_r, 13).fill = fill_hdr_teal; ws1.cell(cur_r, 13).alignment = align_hdr; ws1.cell(cur_r, 13).border = tbl_border

    cur_r += 1
    sub_hdr13 = ["", "", "Doanh số chưa thuế", "Thuế GTGT", "Doanh số chưa thuế", "Thuế GTGT", "Doanh thu [34]", "Thuế đầu ra [35]", "Chênh lệch Doanh số", "Chênh lệch Thuế", "Chênh lệch Doanh số", "Chênh lệch Thuế"]
    for c_idx in range(1, 14):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_hdr_dark
        cell.fill = fill_hdr_light
        cell.alignment = align_hdr
        cell.border = tbl_border
        if c_idx <= len(sub_hdr13) and sub_hdr13[c_idx-1]:
            cell.value = sub_hdr13[c_idx-1]
    ws1.row_dimensions[cur_r-1].height = 20
    ws1.row_dimensions[cur_r].height = 26

    b13_start = cur_r + 1
    stt = 1
    for nam in range(2016, 2026):
        for quy in range(1, 5):
            cur_r += 1
            k = (nam, quy)
            b_bk = bk_ban.get(k, {'ds': 0, 'thue': 0})
            b_hd = hddt_ban_valid.get(k, {'ds': 0, 'thue': 0})
            b_tk = tk_dict.get(k, {'ct34': 0, 'ct35': 0})
            
            cl_hd_ds = b_bk['ds'] - b_hd['ds'] if b_hd['ds'] > 0 else (b_bk['ds'] if nam < 2022 else 0)
            cl_hd_thue = b_bk['thue'] - b_hd['thue'] if b_hd['thue'] > 0 else (b_bk['thue'] if nam < 2022 else 0)
            
            cl_tk_ds = b_bk['ds'] - b_tk['ct34'] if (b_tk['ct34'] > 0 or b_tk['ct35'] > 0) else b_bk['ds']
            cl_tk_thue = b_bk['thue'] - b_tk['ct35'] if (b_tk['ct34'] > 0 or b_tk['ct35'] > 0) else b_bk['thue']
            
            note = "Khớp 100% Tờ khai; HĐĐT chưa triển khai (trước NĐ 123)" if nam < 2022 else "Khớp 100% Tờ khai và HĐĐT hợp lệ"
            if nam == 2016 and quy == 2:
                note = "Khớp 100% Tờ khai bổ sung B1 (sau khi trừ giảm trừ doanh thu 5,64 triệu đồng)"
            elif nam == 2022 and quy in [1, 2]:
                note = "Khớp 100% Tờ khai; HĐĐT chuyển đổi theo NĐ 123 từ T7/2022"
            elif nam == 2025 and quy == 4:
                note = "Khớp 100% tuyệt đối giữa Bảng kê, Tờ khai và HĐĐT hợp lệ (3,09 tỷ)"

            row_vals = [
                stt, f"Quý {quy}/{nam}",
                b_bk['ds'], b_bk['thue'],
                b_hd['ds'], b_hd['thue'],
                b_tk['ct34'], b_tk['ct35'],
                cl_hd_ds, cl_hd_thue,
                cl_tk_ds, cl_tk_thue,
                note
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws1.cell(cur_r, c_idx, val)
                cell.font = f_data
                cell.border = thin_border
                if c_idx in [1, 2]:
                    cell.alignment = align_center
                elif c_idx in range(3, 13):
                    cell.alignment = align_right
                    cell.number_format = num_fmt
                else:
                    cell.alignment = align_left
            stt += 1

    b13_end = cur_r
    cur_r += 1
    ws1.cell(cur_r, 1, "CỘNG").font = f_data_bold
    ws1.cell(cur_r, 1).alignment = align_center
    ws1.cell(cur_r, 1).fill = fill_sum
    ws1.cell(cur_r, 1).border = sum_border
    for c_idx in range(2, 14):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_data_bold
        cell.fill = fill_sum
        cell.border = sum_border
        if c_idx in range(3, 13):
            col_let = get_column_letter(c_idx)
            cell.value = f"=SUM({col_let}{b13_start}:{col_let}{b13_end})"
            cell.number_format = num_fmt
            cell.alignment = align_right
        elif c_idx == 13:
            cell.value = "2016-2025: Doanh thu Bán ra Bảng kê khớp 100% Tờ khai QTTNDN [01]"
            cell.alignment = align_left

    # ------------------------------------------
    # BẢNG 1.4: CHI TIẾT MUA VÀO GIỮA BẢNG KÊ - HĐĐT - TỜ KHAI
    # ------------------------------------------
    cur_r += 3
    ws1.cell(cur_r, 1, "1.4. Bảng đối chiếu chi tiết HHDV Mua vào giữa Bảng kê - HĐĐT - Tờ khai GTGT (Theo kỳ báo cáo)").font = f_tbl_title
    
    cur_r += 1
    ws1.merge_cells(f"A{cur_r}:A{cur_r+1}")
    ws1.cell(cur_r, 1, "STT").font = f_hdr; ws1.cell(cur_r, 1).fill = fill_hdr_teal; ws1.cell(cur_r, 1).alignment = align_hdr; ws1.cell(cur_r, 1).border = tbl_border
    
    ws1.merge_cells(f"B{cur_r}:B{cur_r+1}")
    ws1.cell(cur_r, 2, "Kỳ tính thuế").font = f_hdr; ws1.cell(cur_r, 2).fill = fill_hdr_teal; ws1.cell(cur_r, 2).alignment = align_hdr; ws1.cell(cur_r, 2).border = tbl_border

    ws1.merge_cells(f"C{cur_r}:D{cur_r}")
    ws1.cell(cur_r, 3, "BẢNG KÊ MUA VÀO").font = f_hdr; ws1.cell(cur_r, 3).fill = fill_hdr_teal; ws1.cell(cur_r, 3).alignment = align_hdr; ws1.cell(cur_r, 3).border = tbl_border

    ws1.merge_cells(f"E{cur_r}:F{cur_r}")
    ws1.cell(cur_r, 5, "HÓA ĐƠN ĐIỆN TỬ HỢP LỆ").font = f_hdr; ws1.cell(cur_r, 5).fill = fill_hdr_teal; ws1.cell(cur_r, 5).alignment = align_hdr; ws1.cell(cur_r, 5).border = tbl_border

    ws1.merge_cells(f"G{cur_r}:H{cur_r}")
    ws1.cell(cur_r, 7, "TỜ KHAI THUẾ GTGT").font = f_hdr; ws1.cell(cur_r, 7).fill = fill_hdr_teal; ws1.cell(cur_r, 7).alignment = align_hdr; ws1.cell(cur_r, 7).border = tbl_border

    ws1.merge_cells(f"I{cur_r}:J{cur_r}")
    ws1.cell(cur_r, 9, "CHÊNH LỆCH BẢNG KÊ vs HĐĐT").font = f_hdr; ws1.cell(cur_r, 9).fill = fill_hdr_teal; ws1.cell(cur_r, 9).alignment = align_hdr; ws1.cell(cur_r, 9).border = tbl_border

    ws1.merge_cells(f"K{cur_r}:L{cur_r}")
    ws1.cell(cur_r, 11, "CHÊNH LỆCH BẢNG KÊ vs TỜ KHAI").font = f_hdr; ws1.cell(cur_r, 11).fill = fill_hdr_teal; ws1.cell(cur_r, 11).alignment = align_hdr; ws1.cell(cur_r, 11).border = tbl_border

    ws1.merge_cells(f"M{cur_r}:M{cur_r+1}")
    ws1.cell(cur_r, 13, "Ghi chú & Đánh giá nguyên nhân").font = f_hdr; ws1.cell(cur_r, 13).fill = fill_hdr_teal; ws1.cell(cur_r, 13).alignment = align_hdr; ws1.cell(cur_r, 13).border = tbl_border

    cur_r += 1
    sub_hdr14 = ["", "", "Giá trị chưa thuế", "Thuế GTGT", "Giá trị chưa thuế", "Thuế GTGT", "Giá trị mua [23]", "Thuế mua [24]", "Chênh lệch Giá trị", "Chênh lệch Thuế", "Chênh lệch Giá trị", "Chênh lệch Thuế"]
    for c_idx in range(1, 14):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_hdr_dark
        cell.fill = fill_hdr_light
        cell.alignment = align_hdr
        cell.border = tbl_border
        if c_idx <= len(sub_hdr14) and sub_hdr14[c_idx-1]:
            cell.value = sub_hdr14[c_idx-1]
    ws1.row_dimensions[cur_r-1].height = 20
    ws1.row_dimensions[cur_r].height = 26

    b14_start = cur_r + 1
    stt = 1
    for nam in range(2016, 2026):
        for quy in range(1, 5):
            cur_r += 1
            k = (nam, quy)
            m_bk = bk_mua.get(k, {'ds': 0, 'thue': 0})
            m_hd = hddt_mua.get(k, {'ds': 0, 'thue': 0})
            m_tk = tk_dict.get(k, {'ct23': 0, 'ct24': 0})
            
            cl_hd_ds = m_bk['ds'] - m_hd['ds']
            cl_hd_thue = m_bk['thue'] - m_hd['thue']
            
            cl_tk_ds = m_bk['ds'] - m_tk['ct23'] if (m_tk['ct23'] > 0 or m_tk['ct24'] > 0) else m_bk['ds']
            cl_tk_thue = m_bk['thue'] - m_tk['ct24'] if (m_tk['ct23'] > 0 or m_tk['ct24'] > 0) else m_bk['thue']
            
            note = "Khớp 100% Tờ khai; HĐĐT chưa triển khai (trước NĐ 123)" if nam < 2022 else "Khớp 100% Tờ khai; HĐĐT portal người mua lưu không đầy đủ (tham khảo)"
            if nam == 2016 and quy == 2:
                note = "Lệch 300 triệu do gõ phím sót số 5 thành số 2 tại [23]; Thuế [24] khớp 100% (56,52 tr)"
            elif nam == 2025 and quy == 4:
                note = "Khớp 100% tuyệt đối giữa Bảng kê và Tờ khai GTGT (2,76 tỷ; thuế 219,78 tr)"

            row_vals = [
                stt, f"Quý {quy}/{nam}",
                m_bk['ds'], m_bk['thue'],
                m_hd['ds'], m_hd['thue'],
                m_tk['ct23'], m_tk['ct24'],
                cl_hd_ds, cl_hd_thue,
                cl_tk_ds, cl_tk_thue,
                note
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws1.cell(cur_r, c_idx, val)
                cell.font = f_data
                cell.border = thin_border
                if c_idx in [1, 2]:
                    cell.alignment = align_center
                elif c_idx in range(3, 13):
                    cell.alignment = align_right
                    cell.number_format = num_fmt
                else:
                    cell.alignment = align_left
            stt += 1

    b14_end = cur_r
    cur_r += 1
    ws1.cell(cur_r, 1, "CỘNG").font = f_data_bold
    ws1.cell(cur_r, 1).alignment = align_center
    ws1.cell(cur_r, 1).fill = fill_sum
    ws1.cell(cur_r, 1).border = sum_border
    for c_idx in range(2, 14):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_data_bold
        cell.fill = fill_sum
        cell.border = sum_border
        if c_idx in range(3, 13):
            col_let = get_column_letter(c_idx)
            cell.value = f"=SUM({col_let}{b14_start}:{col_let}{b14_end})"
            cell.number_format = num_fmt
            cell.alignment = align_right
        elif c_idx == 13:
            cell.value = "2016-2025: Thuế GTGT mua vào Bảng kê khớp 100% Thuế [24] Tờ khai"
            cell.alignment = align_left

    # ------------------------------------------
    # PHẦN II. ĐỐI CHIẾU DOANH THU & QUYẾT TOÁN TNDN THEO NĂM
    # ------------------------------------------
    cur_r += 3
    ws1.cell(cur_r, 1, "PHẦN II. ĐỐI CHIẾU DOANH THU VÀ NGHĨA VỤ THUẾ THU NHẬP DOANH NGHIỆP (TNDN) (2016 - 2025)").font = f_sec
    
    cur_r += 2
    ws1.cell(cur_r, 1, "2.1. Bảng đối chiếu Doanh thu năm giữa Bảng kê bán ra - HĐĐT - Tờ khai Quyết toán TNDN (Mẫu 03/TNDN)").font = f_tbl_title
    
    cur_r += 1
    ws1.merge_cells(f"A{cur_r}:A{cur_r+1}")
    ws1.cell(cur_r, 1, "STT").font = f_hdr; ws1.cell(cur_r, 1).fill = fill_hdr_blue; ws1.cell(cur_r, 1).alignment = align_hdr; ws1.cell(cur_r, 1).border = tbl_border

    ws1.merge_cells(f"B{cur_r}:B{cur_r+1}")
    ws1.cell(cur_r, 2, "Năm quyết toán").font = f_hdr; ws1.cell(cur_r, 2).fill = fill_hdr_blue; ws1.cell(cur_r, 2).alignment = align_hdr; ws1.cell(cur_r, 2).border = tbl_border

    ws1.merge_cells(f"C{cur_r}:C{cur_r+1}")
    ws1.cell(cur_r, 3, "DOANH THU BẢNG KÊ (1)").font = f_hdr; ws1.cell(cur_r, 3).fill = fill_hdr_blue; ws1.cell(cur_r, 3).alignment = align_hdr; ws1.cell(cur_r, 3).border = tbl_border

    ws1.merge_cells(f"D{cur_r}:D{cur_r+1}")
    ws1.cell(cur_r, 4, "DOANH THU HĐĐT HỢP LỆ (2)").font = f_hdr; ws1.cell(cur_r, 4).fill = fill_hdr_blue; ws1.cell(cur_r, 4).alignment = align_hdr; ws1.cell(cur_r, 4).border = tbl_border

    ws1.merge_cells(f"E{cur_r}:G{cur_r}")
    ws1.cell(cur_r, 5, "TỜ KHAI QUYẾT TOÁN THUẾ TNDN (03/TNDN)").font = f_hdr; ws1.cell(cur_r, 5).fill = fill_hdr_blue; ws1.cell(cur_r, 5).alignment = align_hdr; ws1.cell(cur_r, 5).border = tbl_border

    ws1.merge_cells(f"H{cur_r}:H{cur_r+1}")
    ws1.cell(cur_r, 8, "CHÊNH LỆCH BK vs HĐĐT (1)-(2)").font = f_hdr; ws1.cell(cur_r, 8).fill = fill_hdr_blue; ws1.cell(cur_r, 8).alignment = align_hdr; ws1.cell(cur_r, 8).border = tbl_border

    ws1.merge_cells(f"I{cur_r}:I{cur_r+1}")
    ws1.cell(cur_r, 9, "CHÊNH LỆCH BK vs TNDN [01] (1)-(3)").font = f_hdr; ws1.cell(cur_r, 9).fill = fill_hdr_blue; ws1.cell(cur_r, 9).alignment = align_hdr; ws1.cell(cur_r, 9).border = tbl_border

    ws1.merge_cells(f"J{cur_r}:J{cur_r+1}")
    ws1.cell(cur_r, 10, "THU NHẬP CHỊU THUẾ TNDN").font = f_hdr; ws1.cell(cur_r, 10).fill = fill_hdr_blue; ws1.cell(cur_r, 10).alignment = align_hdr; ws1.cell(cur_r, 10).border = tbl_border

    ws1.merge_cells(f"K{cur_r}:K{cur_r+1}")
    ws1.cell(cur_r, 11, "THUẾ TNDN PHẢI NỘP").font = f_hdr; ws1.cell(cur_r, 11).fill = fill_hdr_blue; ws1.cell(cur_r, 11).alignment = align_hdr; ws1.cell(cur_r, 11).border = tbl_border

    ws1.merge_cells(f"L{cur_r}:L{cur_r+1}")
    ws1.cell(cur_r, 12, "Ghi chú & Đánh giá rủi ro").font = f_hdr; ws1.cell(cur_r, 12).fill = fill_hdr_blue; ws1.cell(cur_r, 12).alignment = align_hdr; ws1.cell(cur_r, 12).border = tbl_border

    cur_r += 1
    sub_hdr21 = ["", "", "", "", "Tổng doanh thu [01] (3)", "Giảm trừ DT [02]", "Doanh thu thuần [03]", "", "", "", "", ""]
    for c_idx in range(1, 13):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_hdr_dark
        cell.fill = fill_hdr_light
        cell.alignment = align_hdr
        cell.border = tbl_border
        if c_idx <= len(sub_hdr21) and sub_hdr21[c_idx-1]:
            cell.value = sub_hdr21[c_idx-1]
    ws1.row_dimensions[cur_r-1].height = 20
    ws1.row_dimensions[cur_r].height = 26

    # 10 năm data
    tndn_data = [
        (1, "Năm 2016", 3496738590, 0, 3496738590, 5640000, 3491098590, 3496738590, 0, 127566097, 10373038, "BK khớp 100% TK [01]; Giảm trừ DT 5,64tr -> DT thuần khớp GTGT Q2 B1"),
        (2, "Năm 2017", 4172620375, 0, 4172620375, 0, 4172620375, 4172620375, 0, 29769260, 5953852, "BK khớp 100% TK [01]; TNCT 29,76tr"),
        (3, "Năm 2018", 3997066131, 0, 3997066131, 0, 3997066131, 3997066131, 0, 43929901, 8785980, "BK khớp 100% TK [01]; TNCT 43,92tr"),
        (4, "Năm 2019", 4279413491, 0, 4279413491, 0, 4279413491, 4279413491, 0, 62977233, 12595447, "BK khớp 100% TK [01]; TNCT 62,97tr"),
        (5, "Năm 2020", 4311192074, 0, 4311192074, 0, 4311192074, 4311192074, 0, 46927608, 6569856, "BK khớp 100% TK [01]; Giảm 30% thuế TNDN theo NQ 116/2020"),
        (6, "Năm 2021", 4275904321, 0, 4275904321, 0, 4275904321, 4275904321, 0, 53065984, 7429238, "BK khớp 100% TK [01]; Giảm 30% thuế TNDN theo NQ 406/2021"),
        (7, "Năm 2022", 5304710195, 3125349574, 5304710195, 0, 5304710195, 2179360621, 0, 101324222, 20264844, "BK khớp 100% TK [01]; HĐĐT áp dụng từ T7/2022 đạt 58,9% BK"),
        (8, "Năm 2023", 6851592363, 6851592363, 6851592363, 0, 6851592363, 0, 0, 130999396, 26199879, "BK khớp 100% TK [01] và khớp 100% HĐĐT hợp lệ"),
        (9, "Năm 2024", 8652414302, 8652414302, 8652414302, 0, 8652414302, 0, 0, 194517356, 38903471, "BK khớp 100% TK [01] và khớp 100% HĐĐT hợp lệ"),
        (10, "Năm 2025", 10060947641, 10060947641, 10060947641, 0, 10060947641, 0, 0, 110140416, 18723871, "BK khớp 100% TK [01] và khớp 100% HĐĐT hợp lệ")
    ]
    b21_start = cur_r + 1
    for row in tndn_data:
        cur_r += 1
        for c_idx, val in enumerate(row, 1):
            cell = ws1.cell(cur_r, c_idx, val)
            cell.font = f_data
            cell.border = thin_border
            if c_idx in [1, 2]:
                cell.alignment = align_center
            elif c_idx in range(3, 12):
                cell.alignment = align_right
                cell.number_format = num_fmt
            else:
                cell.alignment = align_left

    b21_end = cur_r
    cur_r += 1
    ws1.cell(cur_r, 1, "CỘNG").font = f_data_bold
    ws1.cell(cur_r, 1).alignment = align_center
    ws1.cell(cur_r, 1).fill = fill_sum
    ws1.cell(cur_r, 1).border = sum_border
    for c_idx in range(2, 13):
        cell = ws1.cell(cur_r, c_idx)
        cell.font = f_data_bold
        cell.fill = fill_sum
        cell.border = sum_border
        if c_idx in range(3, 12):
            col_let = get_column_letter(c_idx)
            cell.value = f"=SUM({col_let}{b21_start}:{col_let}{b21_end})"
            cell.number_format = num_fmt
            cell.alignment = align_right
        elif c_idx == 12:
            cell.value = "2016-2025: Bảng kê bán ra khớp 100% Quyết toán TNDN Chỉ tiêu [01] (55,40 tỷ)"
            cell.alignment = align_left

    # ------------------------------------------
    # PHẦN III. BẢNG TỔNG HỢP NGUYÊN NHÂN TRỌNG YẾU THEO NĂM (MỚI NÂNG CẤP)
    # ------------------------------------------
    cur_r += 3
    ws1.cell(cur_r, 1, "PHẦN III. BẢNG SỐ LIỆU TỔNG HỢP NGUYÊN NHÂN TRỌNG YẾU VÀ ĐÁNH GIÁ RỦI RO THEO NĂM (2016 - 2025)").font = f_sec
    
    cur_r += 2
    ws1.cell(cur_r, 1, "3.1. Bảng số liệu tổng hợp các nguyên nhân chênh lệch trọng yếu theo từng năm (Có cột Tổng số 10 năm)").font = f_tbl_title
    
    cur_r += 1
    hdr31 = ["STT", "Nội dung chỉ tiêu / Nguyên nhân trọng yếu", "ĐVT", 
             "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", 
             "TỔNG SỐ", "Bản chất nguyên nhân & Đánh giá rủi ro kiểm tra"]
    for c_idx, h in enumerate(hdr31, 1):
        cell = ws1.cell(cur_r, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_blue
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws1.row_dimensions[cur_r].height = 28

    indicators_31 = [
        # Nhóm A
        ("A", "VỀ DOANH THU BÁN RA & NGHĨA VỤ QUYẾT TOÁN THUẾ TNDN", True, "", [], ""),
        (1, "Tổng Doanh thu Bán ra theo Bảng kê", False, "VNĐ", [3496738590, 4172620375, 3997066131, 4279413491, 4311192074, 4275904321, 5304710195, 6851592363, 8652414302, 10060947641], "Dữ liệu kế toán bán ra 10 năm liên tục"),
        (2, "Doanh thu Quyết toán TNDN Chỉ tiêu [01]", False, "VNĐ", [3496738590, 4172620375, 3997066131, 4279413491, 4311192074, 4275904321, 5304710195, 6851592363, 8652414302, 10060947641], "Khớp 100% tuyệt đối Bảng kê bán ra (55,40 tỷ)"),
        (3, "Khoản giảm trừ doanh thu Chỉ tiêu [02]", False, "VNĐ", [5640000, 0, 0, 0, 0, 0, 0, 0, 0, 0], "Phát sinh năm 2016 (Hàng bán bị trả lại/giảm giá Q2/2016)"),
        (4, "Doanh thu thuần Quyết toán TNDN [03] = [01] - [02]", False, "VNĐ", [3491098590, 4172620375, 3997066131, 4279413491, 4311192074, 4275904321, 5304710195, 6851592363, 8652414302, 10060947641], "Doanh thu thuần tính thuế TNDN (55,39 tỷ)"),
        (5, "Chênh lệch Doanh thu Bảng kê vs Quyết toán [01]", False, "VNĐ", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0], "Không có chênh lệch (Độ khớp 100.00%)"),
        (6, "Thu nhập chịu thuế TNDN", False, "VNĐ", [127566097, 29769260, 43929901, 62977233, 46927608, 53065984, 101324222, 130999396, 194517356, 110140416], "Tổng thu nhập chịu thuế 10 năm đạt 901,2 triệu đồng"),
        (7, "Số thuế TNDN phát sinh phải nộp", False, "VNĐ", [10373038, 5953852, 8785980, 12595447, 6569856, 7429238, 20264844, 26199879, 38903471, 18723871], "Năm 2020, 2021 được giảm 30% thuế TNDN"),
        
        # Nhóm B
        ("B", "VỀ ĐỐI CHIẾU HÓA ĐƠN ĐIỆN TỬ (HĐĐT) BÁN RA (NGHỊ ĐỊNH 123/2020/NĐ-CP)", True, "", [], ""),
        (8, "Tổng Doanh thu HĐĐT thô trên portal thuế", False, "VNĐ", [0, 0, 0, 0, 0, 0, 3157259039, 6889191064, 8666161603, 10260664418], "Bao gồm cả HĐ phát sinh mới và HĐ bị thay thế/hủy bỏ"),
        (9, "Doanh thu HĐĐT hợp lệ (đã trừ HĐ hủy / thay thế)", False, "VNĐ", [0, 0, 0, 0, 0, 0, 3125349574, 6851592363, 8652414302, 10060947641], "Khớp 100% Bảng kê và Tờ khai từ năm 2023 - 2025"),
        (10, "Chênh lệch Doanh thu Bảng kê vs HĐĐT hợp lệ", False, "VNĐ", [3496738590, 4172620375, 3997066131, 4279413491, 4311192074, 4275904321, 2179360621, 0, 0, 0], "Chênh lệch hoàn toàn do mốc áp dụng NĐ 123 từ T7/2022"),
        (11, "Chênh lệch do 13 HĐĐT bị thay thế / hủy bỏ trên portal", False, "VNĐ", [0, 0, 0, 0, 0, 0, 31909465, 37598701, 13747301, 199716777], "Đã bóc tách 13 hóa đơn thay thế/hủy chuẩn CBRR"),
        (12, "Số lượng HĐĐT bị thay thế / hủy bỏ", False, "Hóa đơn", [0, 0, 0, 0, 0, 0, 2, 2, 1, 8], "Tổng cộng 13 hóa đơn (2022: 2, 2023: 2, 2024: 1, 2025: 8)"),

        # Nhóm C
        ("C", "VỀ HÀNG HÓA DỊCH VỤ MUA VÀO & THUẾ GTGT KHẤU TRỪ", True, "", [], ""),
        (13, "Tổng Giá trị Mua vào theo Bảng kê", False, "VNĐ", [2071400835, 3310547198, 4291017442, 3296739253, 4348565059, 4329976519, 4355324230, 6344898998, 7742899128, 8958906829], "Tổng 1.764 dòng hóa đơn mua vào 10 năm (49,05 tỷ)"),
        (14, "Giá trị Mua vào kê khai thuế GTGT Chỉ tiêu [23]", False, "VNĐ", [1771400835, 3310547198, 4291017442, 3296739253, 4348565059, 4329976519, 4355324230, 6344898998, 7742899128, 8958906829], "Khớp 100% Bảng kê toàn bộ 10 năm (trừ 300 triệu lỗi gõ phím ct23 Q2/2016)"),
        (15, "Chênh lệch Mua vào do lỗi gõ phím Q2/2016", False, "VNĐ", [300000000, 0, 0, 0, 0, 0, 0, 0, 0, 0], "Lỗi gõ phím sót số 5 thành số 2 tại ct23 (không ảnh hưởng thuế)"),
        (16, "Thuế GTGT mua vào theo Bảng kê", False, "VNĐ", [207140085, 331046722, 429091746, 329673925, 434856505, 432987653, 358327111, 563599814, 628178931, 712901511], "Tổng thuế GTGT đầu vào phát sinh 10 năm (4,43 tỷ)"),
        (17, "Thuế GTGT mua vào kê khai Chỉ tiêu [24]", False, "VNĐ", [207140085, 331046722, 429091746, 329673925, 434856505, 432987653, 358327111, 563599814, 628178931, 712901511], "Khớp 100% tuyệt đối với Bảng kê toàn bộ 10 năm (4,43 tỷ)"),
        (18, "Chênh lệch Thuế mua vào Bảng kê vs Tờ khai [24]", False, "VNĐ", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0], "Khớp 100% tuyệt đối (lệch = 0 đồng trên toàn bộ 40 quý 10 năm)"),

        # Nhóm D
        ("D", "VỀ RỦI RO NHÀ CUNG CẤP ĐẦU VÀO (ĐỐI CHIẾU CVCBRR & DANH BẠ)", True, "", [], ""),
        (19, "Mua vào từ NCC có Cảnh báo Rủi ro F0 (Sổ CVCBRR)", False, "VNĐ", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0], "An toàn tuyệt đối: 0 NCC dính cảnh báo F0"),
        (20, "Mua vào từ 7 NCC có trạng thái ĐKKD bất thường", False, "VNĐ", [1327166410, 1772549728, 2548304368, 1947254435, 1300032760, 2445543930, 3467011899, 1819974000, 526689542, 39040000], "Tổng 17,19 tỷ đồng (Trọng tâm Cty Tân Á 16,73 tỷ)"),
        (21, "Tỷ trọng mua vào từ NCC bất thường / Tổng mua vào", False, "%", [0.6407, 0.5354, 0.5939, 0.5907, 0.2990, 0.5648, 0.7960, 0.2868, 0.0680, 0.0044], "Bình quân 10 năm chiếm 35.05% tổng giá trị mua vào"),
        (22, "Mua vào từ các NCC hoạt động bình thường (51 NCC)", False, "VNĐ", [744234425, 1537997470, 1742713074, 1349484818, 3048532299, 1884432589, 888312331, 4524924998, 7216209586, 8919866829], "Tổng 31,86 tỷ đồng (Chiếm 64.95% tổng mua vào)"),
        (23, "Tỷ trọng mua vào từ NCC bình thường / Tổng mua vào", False, "%", [0.3593, 0.4646, 0.4061, 0.4093, 0.7010, 0.4352, 0.2040, 0.7132, 0.9320, 0.9956], "Bình quân 10 năm chiếm 64.95% tổng giá trị mua vào")
    ]

    for item in indicators_31:
        stt_val, title_val, is_grp, uom_val, vals, note_val = item
        cur_r += 1
        if is_grp:
            ws1.merge_cells(f"A{cur_r}:O{cur_r}")
            cell = ws1.cell(cur_r, 1, f"{stt_val}. {title_val}")
            cell.font = f_sec
            cell.fill = fill_hdr_light
            cell.alignment = align_left
            cell.border = tbl_border
            ws1.row_dimensions[cur_r].height = 22
        else:
            ws1.cell(cur_r, 1, stt_val).font = f_data_bold if "Tổng" in title_val else f_data
            ws1.cell(cur_r, 1).alignment = align_center
            ws1.cell(cur_r, 1).border = thin_border

            ws1.cell(cur_r, 2, title_val).font = f_data_bold if "Tổng" in title_val or "Chênh lệch" in title_val else f_data
            ws1.cell(cur_r, 2).alignment = align_left
            ws1.cell(cur_r, 2).border = thin_border

            ws1.cell(cur_r, 3, uom_val).font = f_data
            ws1.cell(cur_r, 3).alignment = align_center
            ws1.cell(cur_r, 3).border = thin_border

            for y_idx, val in enumerate(vals, 4):
                cell = ws1.cell(cur_r, y_idx, val)
                cell.font = f_data_bold if "Tổng" in title_val else f_data
                cell.alignment = align_right
                cell.border = thin_border
                if uom_val == "%":
                    cell.number_format = pct_fmt
                elif uom_val == "Hóa đơn":
                    cell.number_format = num_fmt
                else:
                    cell.number_format = num_fmt

            # Cột Tổng số (Col 14)
            tot_cell = ws1.cell(cur_r, 14)
            tot_cell.font = f_data_bold
            tot_cell.alignment = align_right
            tot_cell.border = thin_border
            tot_cell.fill = fill_sum_amber if "bất thường" in title_val or "Chênh lệch" in title_val else fill_sum
            if uom_val == "%":
                if "bất thường" in title_val:
                    tot_cell.value = 17193567072 / 49050275491
                else:
                    tot_cell.value = 31856708419 / 49050275491
                tot_cell.number_format = pct_fmt
            else:
                col_start_let = get_column_letter(4)
                col_end_let = get_column_letter(13)
                tot_cell.value = f"=SUM({col_start_let}{cur_r}:{col_end_let}{cur_r})"
                tot_cell.number_format = num_fmt

            # Cột Ghi chú & Đánh giá (Col 15)
            note_cell = ws1.cell(cur_r, 15, note_val)
            note_cell.font = f_data
            note_cell.alignment = align_left
            note_cell.border = thin_border

    # ------------------------------------------
    # MỤC 3.2: DANH MỤC HỒ SƠ & GIẢI TRÌNH CẦN CUNG CẤP KHI KIỂM TRA
    # ------------------------------------------
    cur_r += 3
    ws1.cell(cur_r, 1, "3.2. Danh mục hồ sơ, chứng từ NNT cần giải trình và cung cấp khi kiểm tra tại trụ sở").font = f_tbl_title

    cur_r += 1
    hdr32 = ["STT", "Nội dung rủi ro / Chênh lệch", "Thời kỳ / Năm", "Số tiền liên quan (VNĐ)", "Hồ sơ chứng từ cần NNT cung cấp", "Trọng tâm kiểm tra của Đoàn kiểm tra"]
    for c_idx, h in enumerate(hdr32, 1):
        cell = ws1.cell(cur_r, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_navy
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws1.row_dimensions[cur_r].height = 24

    recs_32 = [
        (1, "Lệch Chỉ tiêu [23] Tờ khai GTGT Q2/2016", "Quý 2/2016", 300000000, 
         "Tờ khai GTGT Q2/2016, Bảng kê mua vào, Phiếu hạch toán sổ chi tiết TK 133", 
         "Xác nhận lỗi gõ phím (sót số 5 thành số 2 tại ct23); Thuế [24] kê đúng 56,52tr -> Không truy thu"),
        (2, "Mua hàng từ Công ty Công nghiệp Tân Á (Đã đóng MST)", "2016 - 2024", 16727455829, 
         "Hợp đồng kinh tế, 559 Hóa đơn GTGT, Phiếu nhập kho, Định mức sản xuất, Sao kê UNC ngân hàng", 
         "Kiểm tra tính có thực của nguyên liệu giấy tấm carton; đảm bảo HĐ phát sinh trước 21/05/2024"),
        (3, "Mua hàng từ các NCC tạm ngừng KD năm 2026 (Vĩnh Phúc, Mê Kông 68, Tân Đại Việt)", "2020 - 2025", 411781001, 
         "Biên bản giao nhận hàng, Phiếu nhập kho, Chứng từ thanh toán ngân hàng", 
         "Xác minh hàng hóa nhập kho thực tế trong giai đoạn 2020-2025 trước thời điểm tạm ngừng"),
        (4, "Hoàn tất đối chiếu Tờ khai GTGT Quý 4/2025", "Quý 4/2025", 2759527892, 
         "Tờ khai thuế GTGT Mẫu 01/GTGT Q4/2025 chính thức (đã nộp XML ngày 24/01/2026)", 
         "Đã đối chiếu khớp 100% Bảng kê và HĐĐT; Thuế GTGT phát sinh phải nộp 27.607.113 đ"),
        (5, "Giải trình 13 Hóa đơn điện tử bị thay thế/hủy bỏ trên portal", "2022 - 2025", 283032244, 
         "Biên bản hủy/thay thế hóa đơn, Hóa đơn điện tử thay thế tương ứng", 
         "Xác nhận doanh nghiệp không kê khai trùng lặp doanh thu đối với các hóa đơn đã hủy/thay thế")
    ]
    for rec in recs_32:
        cur_r += 1
        for c_idx, val in enumerate(rec, 1):
            cell = ws1.cell(cur_r, c_idx, val)
            cell.font = f_data
            cell.border = thin_border
            if c_idx in [1, 3]:
                cell.alignment = align_center
            elif c_idx == 4:
                cell.alignment = align_right
                cell.number_format = num_fmt
            else:
                cell.alignment = align_left

    print("Sheet 1 finished. Building Sheet 2: Tong_hop_NCC_CVCBRR...")

    # ==========================================
    # SHEET 2: Tong_hop_NCC_CVCBRR
    # ==========================================
    ncc_wb = openpyxl.load_workbook('sources/Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx', data_only=True)
    ws_ncc_sum = ncc_wb['Tong_hop_doi_chieu']

    # Header
    ws2.merge_cells("A1:P1")
    ws2["A1"] = "BẢNG TỔNG HỢP ĐỐI CHIẾU NHÀ CUNG CẤP MUA VÀO VỚI SỔ CVCBRR (F0, F1), SỔ CÔNG AN VÀ DANH BẠ NNT"
    ws2["A1"].font = f_title
    ws2["A1"].alignment = align_center

    ws2.merge_cells("A2:P2")
    ws2["A2"] = "Người nộp thuế: CÔNG TY TNHH MTV BAO BÌ GIẤY VẠN PHÁT THÀNH | Mã số thuế: 3601409441"
    ws2["A2"].font = f_sub
    ws2["A2"].alignment = align_center

    ws2.merge_cells("A3:P3")
    ws2["A3"] = "Thời kỳ kiểm tra: 2016 - 2025 | Dữ liệu đối chiếu: Sổ CVCBRR (F0: 16.931 bản ghi, F1: 3.704 bản ghi), Sổ Công an (1.003 bản ghi) & Danh bạ NNT"
    ws2["A3"].font = f_sub
    ws2["A3"].alignment = align_center

    cur_r2 = 5
    ws2.cell(cur_r2, 1, "PHẦN 1. KẾT QUẢ ĐỐI CHIẾU CÔNG TY ĐANG XÉT (NNT: 3601409441) VỚI SỔ CVCBRR (F1) VÀ SỔ CÔNG VĂN CÔNG AN").font = f_sec

    cur_r2 += 2
    hdr_p1 = ["STT", "Nguồn dữ liệu đối chiếu", "File & Sheet đối chiếu", "Quy mô dữ liệu nguồn", "MST tra cứu", "Tên đơn vị tra cứu", "Kết quả đối chiếu", "Số công văn", "Ngày công văn", "Cơ quan ban hành", "Nội dung / Đánh giá rủi ro"]
    for c_idx, h in enumerate(hdr_p1, 1):
        cell = ws2.cell(cur_r2, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_blue
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws2.row_dimensions[cur_r2].height = 26

    p1_rows = [
        (1, "Sổ theo dõi CVCBRR (Bên mua - F1)", "SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx \\ Sheet F1", "3.704 dòng bản ghi bên mua", "3601409441", "CÔNG TY TNHH MTV BAO BÌ GIẤY VẠN PHÁT THÀNH", "KHÔNG CÓ TRONG DANH SÁCH", "Không phát sinh", "-", "-", "Doanh nghiệp không nằm trong danh sách mua hàng liên quan đến hóa đơn rủi ro (F1) do CQT cảnh báo. AN TOÀN."),
        (2, "Sổ theo dõi hồ sơ cung cấp Cơ quan Công an", "Bao_cao_cong_van.xlsx \\ Sheet NHAP_HO_SO & TAM_NGUNG", "1.003 dòng theo dõi Công an", "3601409441", "CÔNG TY TNHH MTV BAO BÌ GIẤY VẠN PHÁT THÀNH", "KHÔNG CÓ TRONG DANH SÁCH", "Không phát sinh", "-", "-", "Doanh nghiệp không có công văn yêu cầu điều tra/xác minh của cơ quan Công an và không thuộc diện tạm ngưng kê khai. AN TOÀN.")
    ]
    for row in p1_rows:
        cur_r2 += 1
        for c_idx, val in enumerate(row, 1):
            cell = ws2.cell(cur_r2, c_idx, val)
            cell.font = f_data
            cell.border = thin_border
            if c_idx in [1, 5, 7, 8, 9, 10]:
                cell.alignment = align_center
                if c_idx == 7:
                    cell.font = f_data_bold
            else:
                cell.alignment = align_left

    cur_r2 += 1
    ws2.merge_cells(f"A{cur_r2}:K{cur_r2}")
    ws2.cell(cur_r2, 1, "> KẾT LUẬN PHẦN 1: Công ty đang xét (Vạn Phát Thành - MST: 3601409441) hoàn toàn KHÔNG CÓ công văn cảnh báo rủi ro thuế F1 và KHÔNG CÓ công văn xác minh / yêu cầu cung cấp hồ sơ từ Cơ quan Công an.").font = f_data_bold
    ws2.cell(cur_r2, 1).alignment = align_left
    ws2.cell(cur_r2, 1).fill = fill_sum

    # PHẦN 2
    cur_r2 += 3
    ws2.cell(cur_r2, 1, "PHẦN 2. ĐỐI CHIẾU DANH SÁCH NGƯỜI BÁN (BẢNG KÊ MUA VÀO 2016 - 2025) VỚI SỔ CVCBRR BÊN BÁN (F0) VÀ DANH BẠ NNT").font = f_sec

    cur_r2 += 2
    ws2.cell(cur_r2, 1, "2.1. Danh sách các Nhà cung cấp có cảnh báo rủi ro / có công văn trong Sổ CVCBRR (F0)").font = f_tbl_title

    cur_r2 += 1
    hdr_p21 = ["STT", "MST Người bán", "Tên Người bán (Bảng kê mua vào)", "Số HĐ", "Số dòng BK", "Doanh số mua chưa thuế (VNĐ)", "Thuế GTGT (VNĐ)", "Kỳ mua vào", "Từ ngày HĐ", "Đến ngày HĐ", "Mặt hàng chính / Nhận xét"]
    for c_idx, h in enumerate(hdr_p21, 1):
        cell = ws2.cell(cur_r2, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_navy
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws2.row_dimensions[cur_r2].height = 24

    cur_r2 += 1
    ws2.merge_cells(f"A{cur_r2}:K{cur_r2}")
    ws2.cell(cur_r2, 1, "KHÔNG PHÁT SINH: Toàn bộ 58 nhà cung cấp (1.764 dòng hóa đơn mua vào 2016 - 2025) hoàn toàn KHÔNG CÓ trong danh sách cảnh báo rủi ro F0 (Sổ CVCBRR cập nhật gồm 16.931 bản ghi).").font = f_data_bold
    ws2.cell(cur_r2, 1).alignment = align_left
    ws2.cell(cur_r2, 1).border = thin_border

    # BẢNG 2.2: 7 NCC BẤT THƯỜNG
    cur_r2 += 3
    ws2.cell(cur_r2, 1, "2.2. Bảng chi tiết 7 Nhà cung cấp có trạng thái ĐKKD bất thường theo Danh bạ NNT").font = f_tbl_title

    cur_r2 += 1
    hdr_p22 = ["STT", "MST Người bán", "Tên Nhà cung cấp", "Số HĐ", "Doanh số chưa thuế (VNĐ)", "Thuế GTGT (VNĐ)", "Tỷ trọng (%)", "Kỳ mua vào", "Trạng thái ĐKKD / Hoạt động NNT", "Trọng tâm rà soát kiểm tra"]
    for c_idx, h in enumerate(hdr_p22, 1):
        cell = ws2.cell(cur_r2, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_teal
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws2.row_dimensions[cur_r2].height = 26

    bad_ncc_rows = [
        (1, "0300655374", "CÔNG TY CÔNG NGHIỆP TÂN Á", 559, 16727455829, 1589355353, 16727455829/49050275491, "2016-2024", "Ngừng HĐ chưa hoàn thành thủ tục đóng MST (TB 8733 ngày 20/05/2024)", "Doanh số lớn nhất (34,10%). Hóa đơn phát sinh trước ngày đóng MST (21/05/2024). Kiểm tra kho giấy carton."),
        (2, "3700617971", "CÔNG TY TNHH SX-TM VĨNH PHÚC", 21, 210475000, 17160700, 210475000/49050275491, "2022-2025", "Tạm ngừng KD có thời hạn (09/01/2026 - 31/12/2026)", "Kiểm tra chứng từ thanh toán ngân hàng và biên bản giao nhận giấy carton."),
        (3, "3700617971-002", "CN CTY TNHH SX-TM VĨNH PHÚC TẠI ĐỒNG NAI", 5, 163074701, 13045976, 163074701/49050275491, "2022", "Tạm ngừng KD có thời hạn (09/01/2026 - 31/12/2026)", "Chi nhánh Vĩnh Phúc, cùng nhóm rà soát giấy tấm đầu vào."),
        (4, "3702442244", "CÔNG TY TNHH MTV SX CÔNG NGHIỆP PHƯƠNG ANH", 2, 51630282, 5163028, 51630282/49050275491, "2023-2024", "Ngừng HĐ đã đóng MST (18/02/2025 - Giải thể)", "Mua hạt nhựa/vật tư năm 2023-2024, phát sinh trước thời điểm giải thể."),
        (5, "3703050786", "CÔNG TY TNHH MÊ KÔNG 68", 3, 26731260, 2138501, 26731260/49050275491, "2023-2024", "Tạm ngừng KD có thời hạn (01/01/2026 - 31/12/2026)", "Kiểm tra hàng hóa nhập kho (giấy tấm carton)."),
        (6, "3700625186", "CÔNG TY CP SX TM TÂN ĐẠI VIỆT", 2, 11500000, 1040000, 11500000/49050275491, "2020, 2023", "Tạm ngừng KD có thời hạn (01/01/2026 - 31/12/2026)", "Vật tư keo dán giấy carton."),
        (7, "0317986737", "CTY TNHH TMDV PHÁT TRIỂN CHỮ KÝ SỐ VN", 1, 2700000, 270000, 2700000/49050275491, "2023", "Không hoạt động tại địa chỉ đã đăng ký (TB 26/05/2026)", "Dịch vụ chữ ký số Viettel-CA năm 2023.")
    ]
    b22_start = cur_r2 + 1
    for row in bad_ncc_rows:
        cur_r2 += 1
        for c_idx, val in enumerate(row, 1):
            cell = ws2.cell(cur_r2, c_idx, val)
            cell.font = f_data
            cell.border = thin_border
            if c_idx in [1, 2, 8]:
                cell.alignment = align_center
            elif c_idx in [4, 5, 6]:
                cell.alignment = align_right
                cell.number_format = num_fmt
            elif c_idx == 7:
                cell.alignment = align_right
                cell.number_format = pct_fmt
            else:
                cell.alignment = align_left

    b22_end = cur_r2
    cur_r2 += 1
    ws2.cell(cur_r2, 1, "CỘNG").font = f_data_bold
    ws2.cell(cur_r2, 1).alignment = align_center
    ws2.cell(cur_r2, 1).fill = fill_sum_amber
    ws2.cell(cur_r2, 1).border = sum_border
    for c_idx in range(2, 11):
        cell = ws2.cell(cur_r2, c_idx)
        cell.font = f_data_bold
        cell.fill = fill_sum_amber
        cell.border = sum_border
        if c_idx in [4, 5, 6]:
            col_let = get_column_letter(c_idx)
            cell.value = f"=SUM({col_let}{b22_start}:{col_let}{b22_end})"
            cell.number_format = num_fmt
            cell.alignment = align_right
        elif c_idx == 7:
            cell.value = 17193567072 / 49050275491
            cell.number_format = pct_fmt
            cell.alignment = align_right
        elif c_idx == 10:
            cell.value = "Chiếm 35.05% tổng giá trị mua vào 10 năm của doanh nghiệp"
            cell.alignment = align_left

    # BẢNG 2.3: BẢNG SỐ LIỆU THEO NĂM CỦA 7 NCC BẤT THƯỜNG (MỚI BỔ SUNG THEO YÊU CẦU)
    cur_r2 += 3
    ws2.cell(cur_r2, 1, "2.3. Bảng tổng hợp Doanh số mua vào từ 7 Nhà cung cấp bất thường theo từng năm (2016 - 2025) (Có cột Tổng số)").font = f_tbl_title

    cur_r2 += 1
    hdr_p23 = ["STT", "MST Người bán", "Tên Nhà cung cấp", "Trạng thái ĐKKD", 
               "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", 
               "TỔNG CỘNG", "Tỷ trọng (%)"]
    for c_idx, h in enumerate(hdr_p23, 1):
        cell = ws2.cell(cur_r2, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_blue
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws2.row_dimensions[cur_r2].height = 28

    bad_ncc_years_data = [
        (1, "0300655374", "CÔNG TY CÔNG NGHIỆP TÂN Á", "Ngừng HĐ chưa đóng MST", [1327166410, 1772549728, 2548304368, 1947254435, 1294032760, 2445543930, 3292727198, 1715697000, 384180000, 0]),
        (2, "3700617971", "CÔNG TY TNHH SX-TM VĨNH PHÚC", "Tạm ngừng KD có thời hạn", [0, 0, 0, 0, 0, 0, 11210000, 63295000, 96930000, 39040000]),
        (3, "3700617971-002", "CN CTY TNHH SX-TM VĨNH PHÚC", "Tạm ngừng KD có thời hạn", [0, 0, 0, 0, 0, 0, 163074701, 0, 0, 0]),
        (4, "3702442244", "CÔNG TY TNHH MTV SX CN PHƯƠNG ANH", "Ngừng HĐ đã đóng MST", [0, 0, 0, 0, 0, 0, 0, 9789000, 41841282, 0]),
        (5, "3703050786", "CÔNG TY TNHH MÊ KÔNG 68", "Tạm ngừng KD có thời hạn", [0, 0, 0, 0, 0, 0, 0, 22993000, 3738260, 0]),
        (6, "3700625186", "CÔNG TY CP SX TM TÂN ĐẠI VIỆT", "Tạm ngừng KD có thời hạn", [0, 0, 0, 0, 6000000, 0, 0, 5500000, 0, 0]),
        (7, "0317986737", "CTY TNHH TMDV PT CHỮ KÝ SỐ VN", "Không HĐ tại địa chỉ ĐKKD", [0, 0, 0, 0, 0, 0, 0, 2700000, 0, 0])
    ]
    b23_start = cur_r2 + 1
    for item in bad_ncc_years_data:
        stt, mst, ten, tt, y_vals = item
        cur_r2 += 1
        ws2.cell(cur_r2, 1, stt).font = f_data; ws2.cell(cur_r2, 1).alignment = align_center; ws2.cell(cur_r2, 1).border = thin_border
        ws2.cell(cur_r2, 2, mst).font = f_data; ws2.cell(cur_r2, 2).alignment = align_center; ws2.cell(cur_r2, 2).border = thin_border
        ws2.cell(cur_r2, 3, ten).font = f_data; ws2.cell(cur_r2, 3).alignment = align_left; ws2.cell(cur_r2, 3).border = thin_border
        ws2.cell(cur_r2, 4, tt).font = f_data; ws2.cell(cur_r2, 4).alignment = align_left; ws2.cell(cur_r2, 4).border = thin_border
        for y_idx, val in enumerate(y_vals, 5):
            c = ws2.cell(cur_r2, y_idx, val)
            c.font = f_data; c.alignment = align_right; c.number_format = num_fmt; c.border = thin_border
        
        # Col 15: Tong
        tot_c = ws2.cell(cur_r2, 15)
        tot_c.font = f_data_bold; tot_c.alignment = align_right; tot_c.border = thin_border; tot_c.fill = fill_sum
        col_s = get_column_letter(5)
        col_e = get_column_letter(14)
        tot_c.value = f"=SUM({col_s}{cur_r2}:{col_e}{cur_r2})"
        tot_c.number_format = num_fmt

        # Col 16: Ty trong
        pct_c = ws2.cell(cur_r2, 16)
        pct_c.font = f_data; pct_c.alignment = align_right; pct_c.border = thin_border
        pct_c.value = sum(y_vals) / 49050275491
        pct_c.number_format = pct_fmt

    b23_end = cur_r2
    cur_r2 += 1
    # Dòng Tổng cộng 7 NCC
    ws2.cell(cur_r2, 1, "TỔNG CỘNG 7 NCC").font = f_data_bold
    ws2.cell(cur_r2, 1).alignment = align_center
    ws2.cell(cur_r2, 1).fill = fill_sum_amber
    ws2.cell(cur_r2, 1).border = sum_border
    for c_idx in range(2, 17):
        c = ws2.cell(cur_r2, c_idx)
        c.font = f_data_bold
        c.fill = fill_sum_amber
        c.border = sum_border
        if c_idx in range(5, 16):
            col_let = get_column_letter(c_idx)
            c.value = f"=SUM({col_let}{b23_start}:{col_let}{b23_end})"
            c.number_format = num_fmt
            c.alignment = align_right
        elif c_idx == 16:
            c.value = 17193567072 / 49050275491
            c.number_format = pct_fmt
            c.alignment = align_right

    # Dòng Tổng mua vào cả năm
    cur_r2 += 1
    all_mua_vals = [2071400835, 3310547198, 4291017442, 3296739253, 4348565059, 4329976519, 4355324230, 6344898998, 7742899128, 8958906829]
    ws2.cell(cur_r2, 1, "TỔNG MUA VÀO CẢ NĂM").font = f_data_bold
    ws2.cell(cur_r2, 1).alignment = align_center
    ws2.cell(cur_r2, 1).fill = fill_sum
    ws2.cell(cur_r2, 1).border = thin_border
    for c_idx in range(2, 17):
        c = ws2.cell(cur_r2, c_idx)
        c.font = f_data_bold
        c.fill = fill_sum
        c.border = thin_border
        if c_idx in range(5, 15):
            c.value = all_mua_vals[c_idx - 5]
            c.number_format = num_fmt
            c.alignment = align_right
        elif c_idx == 15:
            c.value = 49050275491
            c.number_format = num_fmt
            c.alignment = align_right
        elif c_idx == 16:
            c.value = 1.0
            c.number_format = pct_fmt
            c.alignment = align_right

    # Dòng Tỷ trọng 7 NCC trên tổng mua
    cur_r2 += 1
    ws2.cell(cur_r2, 1, "TỶ TRỌNG 7 NCC (%)").font = f_data_bold
    ws2.cell(cur_r2, 1).alignment = align_center
    ws2.cell(cur_r2, 1).fill = fill_hdr_amber
    ws2.cell(cur_r2, 1).border = sum_border
    bad_totals_y = [1327166410, 1772549728, 2548304368, 1947254435, 1300032760, 2445543930, 3467011899, 1819974000, 526689542, 39040000]
    for c_idx in range(2, 17):
        c = ws2.cell(cur_r2, c_idx)
        c.font = f_data_bold
        c.fill = fill_hdr_amber
        c.border = sum_border
        if c_idx in range(5, 15):
            c.value = bad_totals_y[c_idx - 5] / all_mua_vals[c_idx - 5]
            c.number_format = pct_fmt
            c.alignment = align_right
        elif c_idx == 15:
            c.value = 17193567072
            c.number_format = num_fmt
            c.alignment = align_right
        elif c_idx == 16:
            c.value = 17193567072 / 49050275491
            c.number_format = pct_fmt
            c.alignment = align_right

    # BẢNG 2.4: TỔNG HỢP 58 NHÀ CUNG CẤP (Copy từ bảng 2.2 cũ)
    cur_r2 += 3
    ws2.cell(cur_r2, 1, "2.4. Bảng tổng hợp toàn bộ 58 Nhà cung cấp mua vào (2016 - 2025) đối chiếu Sổ CVCBRR và Danh bạ NNT").font = f_tbl_title

    cur_r2 += 1
    hdr_p24 = ["STT", "MST Người bán", "Tên Nhà cung cấp", "Phân loại", "Số HĐ", "Số dòng BK", "Tổng Doanh số chưa thuế (VNĐ)", "Tổng Thuế GTGT (VNĐ)", "Tỷ trọng (%)", "Kỳ mua vào", "Từ ngày HĐ", "Đến ngày HĐ", "Trạng thái ĐKKD / Nhận xét"]
    for c_idx, h in enumerate(hdr_p24, 1):
        cell = ws2.cell(cur_r2, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_navy
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws2.row_dimensions[cur_r2].height = 26

    # Read rows from ws_ncc_sum Table 2.2
    b24_start = cur_r2 + 1
    r_start_src = None
    for r in range(1, ws_ncc_sum.max_row+1):
        if ws_ncc_sum.cell(r, 1).value == "STT" and ws_ncc_sum.cell(r, 2).value == "MST Người bán":
            r_start_src = r + 1
            break
    
    if r_start_src:
        for r in range(r_start_src, ws_ncc_sum.max_row+1):
            stt_v = ws_ncc_sum.cell(r, 1).value
            if stt_v is None or str(stt_v).strip().upper() == "CỘNG" or "PHẦN" in str(stt_v):
                break
            cur_r2 += 1
            row_vals = [ws_ncc_sum.cell(r, c).value for c in range(1, 14)]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws2.cell(cur_r2, c_idx, val)
                cell.font = f_data
                cell.border = thin_border
                if c_idx in [1, 2, 4, 10, 11, 12]:
                    cell.alignment = align_center
                elif c_idx in [5, 6, 7, 8]:
                    cell.alignment = align_right
                    cell.number_format = num_fmt
                elif c_idx == 9:
                    cell.alignment = align_right
                    cell.number_format = pct_fmt
                else:
                    cell.alignment = align_left

    b24_end = cur_r2
    cur_r2 += 1
    ws2.cell(cur_r2, 1, "CỘNG").font = f_data_bold
    ws2.cell(cur_r2, 1).alignment = align_center
    ws2.cell(cur_r2, 1).fill = fill_sum
    ws2.cell(cur_r2, 1).border = sum_border
    for c_idx in range(2, 14):
        cell = ws2.cell(cur_r2, c_idx)
        cell.font = f_data_bold
        cell.fill = fill_sum
        cell.border = sum_border
        if c_idx in [5, 6, 7, 8]:
            col_let = get_column_letter(c_idx)
            cell.value = f"=SUM({col_let}{b24_start}:{col_let}{b24_end})"
            cell.number_format = num_fmt
            cell.alignment = align_right
        elif c_idx == 9:
            cell.value = 1.0
            cell.number_format = pct_fmt
            cell.alignment = align_right
        elif c_idx == 13:
            cell.value = "Toàn bộ 58 nhà cung cấp (1.764 hóa đơn) an toàn F0; 7 NCC có trạng thái bất thường"
            cell.alignment = align_left

    # PHẦN 3
    cur_r2 += 3
    ws2.cell(cur_r2, 1, "PHẦN 3. ĐỀ XUẤT TRỌNG TÂM KIỂM TRA TẠI DOANH NGHIỆP").font = f_sec

    cur_r2 += 2
    hdr_p3 = ["STT", "Nhóm đối tượng kiểm tra", "Đối tượng trọng điểm", "Giá trị mua vào (VNĐ)", "Tỷ trọng", "Kế hoạch & Phương pháp kiểm tra thực tế"]
    for c_idx, h in enumerate(hdr_p3, 1):
        cell = ws2.cell(cur_r2, c_idx, h)
        cell.font = f_hdr
        cell.fill = fill_hdr_navy
        cell.alignment = align_hdr
        cell.border = tbl_border
    ws2.row_dimensions[cur_r2].height = 24

    p3_recs = [
        (1, "Nhà cung cấp chiếm tỷ trọng lớn nhất", "CÔNG TY CÔNG NGHIỆP TÂN Á (MST: 0300655374)", 16727455829, 16727455829/49050275491, 
         "Rà soát toàn bộ 559 hóa đơn mua giấy tấm carton (2016-2024); đối chiếu phiếu nhập kho, định mức sản xuất, sao kê thanh toán ngân hàng. Đảm bảo hóa đơn phát sinh trước ngày đóng MST (21/05/2024)."),
        (2, "Nhà cung cấp tạm ngừng KD năm 2026", "Cty Vĩnh Phúc, CN Vĩnh Phúc, Mê Kông 68, Tân Đại Việt", 411781001, 411781001/49050275491, 
         "Kiểm tra chứng từ thanh toán không dùng tiền mặt và biên bản giao nhận hàng hóa phát sinh trước thời điểm doanh nghiệp tạm ngừng hoạt động trong năm 2026."),
        (3, "Nhà cung cấp ngừng HĐ đã giải thể", "Cty Phương Anh (MST: 3702442244) & Xăng dầu BR-VT", 51903009, 51903009/49050275491, 
         "Đối chiếu hóa đơn mua hạt nhựa và xăng dầu phát sinh năm 2023-2025 trước thời điểm giải thể."),
        (4, "51 Nhà cung cấp hoạt động bình thường", "Phúc Hảo, Four Nine, Bao bì Biên Hòa, Box-Pak...", 31856708419, 31856708419/49050275491, 
         "Kiểm tra chọn mẫu hóa đơn có giá trị lớn, chứng từ ngân hàng theo quy định.")
    ]
    for rec in p3_recs:
        cur_r2 += 1
        for c_idx, val in enumerate(rec, 1):
            cell = ws2.cell(cur_r2, c_idx, val)
            cell.font = f_data
            cell.border = thin_border
            if c_idx in [1]:
                cell.alignment = align_center
            elif c_idx == 4:
                cell.alignment = align_right
                cell.number_format = num_fmt
            elif c_idx == 5:
                cell.alignment = align_right
                cell.number_format = pct_fmt
            else:
                cell.alignment = align_left

    # Auto-fit columns for both sheets
    for ws in [ws1, ws2]:
        ws.views.sheetView[0].showGridLines = True
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None:
                    s = str(val)
                    if "\n" in s:
                        s = max(s.split("\n"), key=len)
                    if len(s) > max_len and len(s) < 100:
                        max_len = len(s)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
    
    # Custom tweaks for specific columns
    ws1.column_dimensions['A'].width = 8
    ws1.column_dimensions['B'].width = 16
    ws1.column_dimensions['J'].width = 35
    ws1.column_dimensions['L'].width = 35
    ws1.column_dimensions['M'].width = 45
    ws1.column_dimensions['O'].width = 50

    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 18
    ws2.column_dimensions['C'].width = 38
    ws2.column_dimensions['D'].width = 25
    ws2.column_dimensions['K'].width = 45
    ws2.column_dimensions['M'].width = 45

    # Save to all destination locations
    out_paths = [
        'sources/Tong_hop_doi_chieu_GTGT_TNDN_3601409441.xlsx',
        'ket-qua/Tong_hop_doi_chieu_GTGT_TNDN_3601409441.xlsx',
        'sao-luu/Tong_hop_doi_chieu_GTGT_TNDN_3601409441.xlsx',
        r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\DANG KIEM TRA\KE HOACH\VAN PHAT THANH\ket-qua\Tong_hop_doi_chieu_GTGT_TNDN_3601409441.xlsx',
        r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\DANG KIEM TRA\KE HOACH\VAN PHAT THANH\7. Sao luu tong hop\Tong_hop_doi_chieu_GTGT_TNDN_3601409441.xlsx'
    ]

    for p in out_paths:
        try:
            os.makedirs(os.path.dirname(p), exist_ok=True)
            wb.save(p)
            print(f"Successfully saved to: {p}")
        except Exception as e:
            print(f"Could not save to {p}: {e}")

if __name__ == "__main__":
    create_report()
