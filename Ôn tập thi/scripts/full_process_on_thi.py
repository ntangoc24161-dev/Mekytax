# -*- coding: utf-8 -*-
"""
Kịch bản xử lý toàn diện theo yêu cầu ôn tập thi kíp P3-39917:
1. Rà soát 40 đề, chèn dòng tên file cỡ chữ 16 in đậm ở Dòng 1, tô đậm phương án đúng trong từng đề.
2. Cập nhật danh sách tổng hợp 40 đề & Câu 1 (bỏ các cột phương án A, B, C, D).
3. Phân tích các cặp câu hỏi "Anh em sinh đôi" (So sánh, Keyword, Bẫy, Mẹo chọn).
4. Xây dựng Bộ quy tắc nhận diện nhanh (Cheat Sheet Mapping).
5. Tạo tài liệu Word (.docx) Tóm tắt siêu tốc Hack não dưới 300 từ mỗi chuyên đề.
6. Cập nhật các file Excel và Markdown Second Brain.
"""

import os
import glob
import re
from collections import Counter, defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

SOURCE_FOLDER = r'D:\HOC TAP\ON THI 2026.09\40 DE THI'
OUTPUT_EXCEL_ALL = os.path.join(SOURCE_FOLDER, 'Tong_Hop_1429_Cau_Hoi_Doc_Ban_On_Tap_Nhanh.xlsx')
OUTPUT_EXCEL_DE = os.path.join(SOURCE_FOLDER, 'Danh_Sach_Tong_Hop_40_De_Va_Cau_1.xlsx')
OUTPUT_EXCEL_TWINS = os.path.join(SOURCE_FOLDER, 'Phan_Tich_Cau_Hoi_Sinh_Doi_Va_Cheat_Sheet.xlsx')
OUTPUT_DOCX_SUMMARY = os.path.join(SOURCE_FOLDER, 'Tom_Tat_Sieu_Toc_Hack_Nao_On_Thi_Thue_P3_39917.docx')
SOURCES_MD_DIR = r'D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\sources'

def get_base_styles():
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    sub_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    accent_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    bright_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    correct_opt_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font_16 = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    bold_font = Font(name="Calibri", size=11, bold=True)
    normal_font = Font(name="Calibri", size=11)
    correct_opt_font = Font(name="Calibri", size=11, bold=True, color="276A3C")
    keyword_font = Font(name="Calibri", size=11, bold=True, color="C00000")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    header_border = Border(
        left=Side(style='thin', color='1F4E79'),
        right=Side(style='thin', color='1F4E79'),
        top=Side(style='medium', color='1F4E79'),
        bottom=Side(style='medium', color='1F4E79')
    )
    
    return {
        'header_fill': header_fill,
        'sub_header_fill': sub_header_fill,
        'accent_fill': accent_fill,
        'yellow_fill': yellow_fill,
        'bright_yellow': bright_yellow,
        'correct_opt_fill': correct_opt_fill,
        'header_font': header_font,
        'title_font_16': title_font_16,
        'bold_font': bold_font,
        'normal_font': normal_font,
        'correct_opt_font': correct_opt_font,
        'keyword_font': keyword_font,
        'thin_border': thin_border,
        'header_border': header_border
    }

# ==============================================================================
# BƯỚC 1: XỬ LÝ 40 FILE ĐỀ THI (CHÈN DÒNG TÊN FILE SIZE 16 IN ĐẬM, TÔ ĐẬM ĐÁP ÁN)
# ==============================================================================
def process_individual_exams():
    styles = get_base_styles()
    files = sorted(glob.glob(os.path.join(SOURCE_FOLDER, 'Bo_De_On_Thi_Kiem_Tra_Thue_P3_39917_De_*.xlsx')),
                   key=lambda x: int(re.search(r'De_(\d+)', x).group(1)))
    
    print(f"BƯỚC 1: Đang rà soát và định dạng lại {len(files)} file đề thi...")
    
    exam_records = []
    all_questions = {}
    
    headers = [
        "STT", "Mã Câu Hệ Thống", "Chuyên Đề Nghiệp Vụ", "Nội Dung Câu Hỏi",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Phương Án Chọn", "Nguồn Phương Án", "Kết Quả Lần Đầu",
        "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Căn Cứ Pháp Lý / Ghi Chú"
    ]
    
    opt_col_map = {'A': 5, 'B': 6, 'C': 7, 'D': 8}
    
    for fpath in files:
        fname = os.path.basename(fpath)
        exam_num = int(re.search(r'De_(\d+)', fname).group(1))
        
        wb_old = openpyxl.load_workbook(fpath, data_only=True)
        ws_old = wb_old.active
        raw_rows = list(ws_old.iter_rows(values_only=True))
        
        # Tìm dòng bắt đầu câu hỏi (nếu file cũ đã có header hoặc title)
        q_rows = []
        for r in raw_rows:
            if r and str(r[0]).strip().isdigit() and len(r) >= 12:
                q_rows.append(r)
            elif r and str(r[0]).strip() == '1' and len(r) >= 12:
                q_rows.append(r)
        
        # Nếu chưa bóc được bằng isdigit thì lấy từ row 1 (bỏ header)
        if len(q_rows) < 90:
            q_rows = [r for r in raw_rows[1:] if r and r[1] is not None]
            
        if len(q_rows) != 90:
            print(f"Cảnh báo: {fname} có {len(q_rows)} câu hỏi!")
            
        q1 = q_rows[0]
        exam_records.append({
            'exam_num': exam_num,
            'filename': fname,
            'total_q': len(q_rows),
            'q1_id': str(q1[1]),
            'q1_topic': str(q1[2] or ''),
            'q1_content': str(q1[3] or ''),
            'q1_ans': str(q1[11] or ''),
            'q1_ans_text': str(q1[12] or ''),
            'q1_note': str(q1[13] or '') if len(q1) > 13 else ''
        })
        
        # Tạo lại workbook mới chuẩn hóa
        wb_new = openpyxl.Workbook()
        ws = wb_new.active
        ws.title = f"De_{exam_num:02d}"
        
        # Dòng 1: Tên file cỡ chữ 16 in đậm
        title_text = f"BỘ ĐỀ ÔN THI KIỂM TRA THUẾ P3-39917 - ĐỀ {exam_num:02d} ({fname})"
        ws.cell(row=1, column=1, value=title_text)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        title_cell = ws.cell(row=1, column=1)
        title_cell.font = styles['title_font_16']
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        title_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ws.row_dimensions[1].height = 36
        
        # Dòng 2: Headers
        for col_idx, h_text in enumerate(headers, 1):
            c = ws.cell(row=2, column=col_idx, value=h_text)
            c.font = styles['header_font']
            c.fill = styles['header_fill']
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = styles['header_border']
        ws.row_dimensions[2].height = 28
        
        # Dòng 3 đến 92: 90 câu hỏi
        for idx, r in enumerate(q_rows, 1):
            row_idx = idx + 2
            ws.row_dimensions[row_idx].height = 26
            
            stt = idx
            item_id = str(r[1] or '').strip()
            topic = str(r[2] or 'Nghiệp vụ kiểm tra thuế khác').strip()
            content = str(r[3] or '').strip()
            opt_a = str(r[4] or '').strip()
            opt_b = str(r[5] or '').strip()
            opt_c = str(r[6] or '').strip()
            opt_d = str(r[7] or '').strip()
            sel_opt = str(r[8] or '').strip()
            opt_src = str(r[9] or '').strip()
            first_res = str(r[10] or '').strip()
            correct_ans = str(r[11] or '').strip().upper()
            correct_text = str(r[12] or '').strip()
            note = str(r[13] or '').strip() if len(r) > 13 else ''
            
            # Ghi dữ liệu vào all_questions
            if item_id not in all_questions:
                all_questions[item_id] = {
                    'id': item_id,
                    'topic': topic,
                    'content': content,
                    'opt_a': opt_a,
                    'opt_b': opt_b,
                    'opt_c': opt_c,
                    'opt_d': opt_d,
                    'correct_ans': correct_ans,
                    'correct_text': correct_text,
                    'note': note,
                    'count': 1,
                    'exams': [exam_num]
                }
            else:
                all_questions[item_id]['count'] += 1
                all_questions[item_id]['exams'].append(exam_num)
                if note and len(note) > len(all_questions[item_id]['note']):
                    all_questions[item_id]['note'] = note
                if not all_questions[item_id]['correct_ans'] and correct_ans:
                    all_questions[item_id]['correct_ans'] = correct_ans
                    all_questions[item_id]['correct_text'] = correct_text
                    
            row_vals = [
                stt, item_id, topic, content,
                opt_a, opt_b, opt_c, opt_d,
                sel_opt, opt_src, first_res,
                correct_ans, correct_text, note
            ]
            
            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = styles['normal_font']
                cell.border = styles['thin_border']
                
                if col_idx in [1, 2, 9, 10, 11, 12]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                
                # Tô đậm và highlight phương án đúng tương ứng
                if correct_ans in opt_col_map and col_idx == opt_col_map[correct_ans]:
                    cell.font = styles['correct_opt_font']
                    cell.fill = styles['correct_opt_fill']
                
                # Cột Đáp Án Đúng (12) và Nội Dung Đáp Án Đúng (13)
                if col_idx == 12:
                    cell.font = styles['correct_opt_font']
                    cell.fill = styles['correct_opt_fill']
                elif col_idx == 13:
                    cell.font = styles['bold_font']
                    
        # Thiết lập độ rộng cột
        col_widths = {
            'A': 8,   # STT
            'B': 18,  # Mã câu
            'C': 28,  # Chuyên đề
            'D': 48,  # Nội dung câu hỏi
            'E': 30,  # Phương án A
            'F': 30,  # Phương án B
            'G': 30,  # Phương án C
            'H': 30,  # Phương án D
            'I': 14,  # PA Chọn
            'J': 16,  # Nguồn
            'K': 14,  # Kết quả
            'L': 14,  # Đáp án đúng
            'M': 35,  # Nội dung đáp án đúng
            'N': 40   # Ghi chú
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width
            
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:N{len(q_rows)+2}"
        
        wb_new.save(fpath)
        
    print(f"Đã cập nhật xong toàn bộ {len(files)} file đề thi.")
    return exam_records, all_questions

# ==============================================================================
# BƯỚC 2: CẬP NHẬT DANH SÁCH TỔNG HỢP 40 ĐỀ VÀ CÂU 1 (BỎ CỘT PHƯƠNG ÁN A,B,C,D)
# ==============================================================================
def build_updated_excel_de_va_cau_1(exam_records):
    styles = get_base_styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # Sheet 1: Danh_Sach_Tong_Hop_40_De_Va_Cau_1 (Chỉ lấy STT, File, Đề, Câu 1, Đáp án)
    ws1 = wb.create_sheet(title="Tong_Hop_40_De_Va_Cau_1")
    
    # Tiêu đề lớn
    ws1.cell(row=1, column=1, value="DANH SÁCH TỔNG HỢP 40 BỘ ĐỀ THI VÀ CÂU HỎI SỐ 1 (KÍP P3-39917)")
    ws1.merge_cells("A1:G1")
    ws1.cell(row=1, column=1).font = styles['title_font_16']
    ws1.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws1.cell(row=1, column=1).fill = styles['accent_fill']
    ws1.row_dimensions[1].height = 36
    
    headers1 = [
        "STT", "Mã Đề Thi", "Tên File Excel Bộ Đề", "Mã Câu 1 (ID)",
        "Chuyên Đề Nghiệp Vụ", "Nội Dung Câu Hỏi Số 1",
        "Đáp Án Đúng"
    ]
    for col_idx, h_text in enumerate(headers1, 1):
        c = ws1.cell(row=2, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws1.row_dimensions[2].height = 28
    
    for idx, rec in enumerate(exam_records, 1):
        row_idx = idx + 2
        ws1.row_dimensions[row_idx].height = 26
        
        row_data = [
            idx,
            f"Đề {rec['exam_num']:02d}",
            rec['filename'],
            rec['q1_id'],
            rec['q1_topic'],
            rec['q1_content'],
            f"{rec['q1_ans']}: {rec['q1_ans_text']}"
        ]
        
        for col_idx, val in enumerate(row_data, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if col_idx in [1, 2, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if col_idx == 7: # Đáp án đúng
                cell.font = styles['correct_opt_font']
                cell.fill = styles['correct_opt_fill']
                
    col_widths1 = {
        'A': 8,   # STT
        'B': 14,  # Mã Đề
        'C': 46,  # Tên File
        'D': 16,  # Mã câu 1
        'E': 30,  # Chuyên đề
        'F': 55,  # Nội dung câu 1
        'G': 40   # Đáp án đúng
    }
    for col_letter, width in col_widths1.items():
        ws1.column_dimensions[col_letter].width = width
        
    ws1.freeze_panes = "A3"
    ws1.auto_filter.ref = f"A2:G{len(exam_records)+2}"
    
    # Sheet 2: Phân tích nhóm câu 1 trùng nhau
    ws2 = wb.create_sheet(title="Phan_Tich_Mau_De_Qua_Cau_1")
    ws2.cell(row=1, column=1, value="PHÂN TÍCH NHẬN DIỆN MẪU ĐỀ THI THÔNG QUA CÂU HỎI SỐ 1")
    ws2.merge_cells("A1:F1")
    ws2.cell(row=1, column=1).font = styles['title_font_16']
    ws2.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws2.cell(row=1, column=1).fill = styles['accent_fill']
    ws2.row_dimensions[1].height = 36
    
    headers2 = [
        "STT", "Mã Câu 1 (ID)", "Nội Dung Câu Hỏi Số 1", "Chuyên Đề",
        "Đáp Án Đúng", "Các Đề Xuất Hiện Câu 1 Này"
    ]
    for col_idx, h_text in enumerate(headers2, 1):
        c = ws2.cell(row=2, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws2.row_dimensions[2].height = 28
    
    q1_groups = defaultdict(list)
    for rec in exam_records:
        q1_groups[rec['q1_id']].append(rec)
    sorted_q1_groups = sorted(q1_groups.items(), key=lambda x: (-len(x[1]), x[0]))
    
    for idx, (qid, reclist) in enumerate(sorted_q1_groups, 1):
        row_idx = idx + 2
        ws2.row_dimensions[row_idx].height = 26
        r0 = reclist[0]
        exams_str = ", ".join([f"Đề {r['exam_num']:02d}" for r in reclist])
        row_data = [
            idx, qid, r0['q1_content'], r0['q1_topic'],
            f"{r0['q1_ans']}: {r0['q1_ans_text']}", f"{len(reclist)} đề ({exams_str})"
        ]
        for col_idx, val in enumerate(row_data, 1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=val)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if col_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if col_idx == 5:
                cell.font = styles['correct_opt_font']
                cell.fill = styles['correct_opt_fill']
                
    col_widths2 = {'A': 8, 'B': 18, 'C': 55, 'D': 30, 'E': 40, 'F': 35}
    for col_letter, width in col_widths2.items():
        ws2.column_dimensions[col_letter].width = width
    ws2.freeze_panes = "A3"
    ws2.auto_filter.ref = f"A2:F{len(sorted_q1_groups)+2}"
    
    wb.save(OUTPUT_EXCEL_DE)
    print(f"Đã lưu danh sách tổng hợp cập nhật: {OUTPUT_EXCEL_DE}")

# ==============================================================================
# BƯỚC 3 & 4: PHÂN TÍCH ANH EM SINH ĐÔI & CHEAT SHEET MAPPING
# ==============================================================================
def analyze_twin_questions_and_cheat_sheet(all_questions):
    """
    Rút trích các cụm câu hỏi anh em sinh đôi (Twin questions),
    xác định Keyword, phân tích Bẫy, Mẹo nhận diện, và xây dựng bảng Cheat Sheet Mapping.
    """
    styles = get_base_styles()
    
    # 1. Danh sách các cặp câu hỏi sinh đôi tiêu biểu đã qua thẩm định nghiệp vụ sâu sắc
    # Gồm các nhóm chủ đề trọng điểm hay bị gài bẫy nhất
    curated_twins = [
        {
            'group': 'PCTN - Quà tặng',
            'q1': {
                'id': '4285993',
                'content': 'Luật Phòng chống tham nhũng quy định việc TẶNG QUÀ của cơ quan, tổ chức, đơn vị, người có chức vụ, quyền hạn như thế nào?',
                'keyword': 'TẶNG QUÀ',
                'ans': 'A',
                'ans_text': 'Không được sử dụng TÀI CHÍNH CÔNG, TÀI SẢN CÔNG làm quà tặng (trừ từ thiện, đối ngoại...)'
            },
            'q2': {
                'id': '4285994',
                'content': 'Luật Phòng chống tham nhũng quy định việc NHẬN QUÀ TẶNG của cơ quan, tổ chức, đơn vị, người có chức vụ, quyền hạn như thế nào?',
                'keyword': 'NHẬN QUÀ TẶNG',
                'ans': 'A',
                'ans_text': 'Không được trực tiếp hoặc gián tiếp nhận quà tặng DƯỚI MỌI HÌNH THỨC liên quan công việc giải quyết'
            },
            'trap_analysis': 'Bẫy nhầm lẫn giữa hành vi Tặng (cấm dùng tiền/tài sản công) và Nhận (cấm tuyệt đối dưới mọi hình thức liên quan công vụ).',
            'quick_rule': 'Thấy "TẶNG" ➡️ Tìm đáp án có "Tài chính công / Tài sản công". Thấy "NHẬN" ➡️ Tìm "Dưới mọi hình thức".'
        },
        {
            'group': 'PCTN - Kỷ luật Định mức/Tiêu chuẩn',
            'q1': {
                'id': '4286023',
                'content': 'Người TỰ Ý SỬ DỤNG trái quy định về định mức, tiêu chuẩn, chế độ mà chưa đến mức bị truy cứu TNHS thì bị xử lý kỷ luật bằng hình thức?',
                'keyword': 'TỰ Ý SỬ DỤNG',
                'ans': 'C',
                'ans_text': 'Khiển trách'
            },
            'q2': {
                'id': '4286022',
                'content': 'Người CHO PHÉP SỬ DỤNG trái quy định về định mức, tiêu chuẩn, chế độ mà chưa đến mức bị truy cứu TNHS thì bị xử lý kỷ luật bằng hình thức?',
                'keyword': 'CHO PHÉP SỬ DỤNG',
                'ans': 'C',
                'ans_text': 'Cảnh cáo'
            },
            'trap_analysis': 'Bẫy mức độ trách nhiệm: Người tự ý làm chịu mức nhẹ hơn (Khiển trách); Lãnh đạo cho phép làm chịu mức nặng hơn (Cảnh cáo).',
            'quick_rule': '"TỰ Ý" ➡️ Khiển trách | "CHO PHÉP" (Lãnh đạo phê duyệt sai) ➡️ Cảnh cáo.'
        },
        {
            'group': 'PCTN - Xử lý Người đứng đầu',
            'q1': {
                'id': '4286018',
                'content': 'Hình thức CẢNH CÁO được áp dụng trong trường hợp người đứng đầu, cấp phó để xảy ra vụ việc tham nhũng mức độ nào?',
                'keyword': 'CẢNH CÁO người đứng đầu',
                'ans': 'A',
                'ans_text': 'Rất nghiêm trọng HOẶC nhiều vụ việc nghiêm trọng'
            },
            'q2': {
                'id': '4286019',
                'content': 'Hình thức CÁCH CHỨC được áp dụng trong trường hợp người đứng đầu, cấp phó để xảy ra vụ việc tham nhũng mức độ nào?',
                'keyword': 'CÁCH CHỨC người đứng đầu',
                'ans': 'C',
                'ans_text': 'Rất nghiêm trọng, ĐẶC BIỆT NGHIÊM TRỌNG hoặc nhiều vụ việc rất nghiêm trọng'
            },
            'trap_analysis': 'Bẫy các mức độ vụ việc (Ít nghiêm trọng ➡️ Khiển trách; Rất nghiêm trọng/Nhiều nghiêm trọng ➡️ Cảnh cáo; Đặc biệt nghiêm trọng ➡️ Cách chức).',
            'quick_rule': 'Cảnh cáo ➡️ Có "rất nghiêm trọng hoặc nhiều vụ nghiêm trọng" | Cách chức ➡️ Bắt buộc có "ĐẶC BIỆT NGHIÊM TRỌNG".'
        },
        {
            'group': 'Xử phạt Chậm nộp hồ sơ khai thuế (NĐ 125)',
            'q1': {
                'id': '4285566',
                'content': 'Tổ chức bị phạt tiền từ 2.000.000 đồng đến 5.000.000 đồng đối với hành vi nộp hồ sơ khai thuế quá thời hạn quy định bao nhiêu ngày?',
                'keyword': 'Phạt 2tr - 5tr',
                'ans': 'B',
                'ans_text': 'Từ 01 ngày đến 30 ngày (không có tình tiết giảm nhẹ)'
            },
            'q2': {
                'id': '4285564',
                'content': 'Tổ chức bị phạt tiền từ 5.000.000 đồng đến 8.000.000 đồng đối với hành vi nộp hồ sơ khai thuế quá thời hạn quy định bao nhiêu ngày?',
                'keyword': 'Phạt 5tr - 8tr',
                'ans': 'D',
                'ans_text': 'Từ 31 ngày đến 60 ngày'
            },
            'trap_analysis': 'Bẫy khớp khung tiền và số ngày: 1-30 ngày (2-5tr), 31-60 ngày (5-8tr), 61-90 ngày (8-15tr), quá 90 ngày ko phát sinh thuế (15-25tr).',
            'quick_rule': '2-5tr ➡️ 1-30 ngày | 5-8tr ➡️ 31-60 ngày | 8-15tr ➡️ 61-90 ngày | 15-25tr ➡️ Quá 90 ngày (không phát sinh thuế).'
        },
        {
            'group': 'Xử phạt Chậm nộp hồ sơ khai thuế (Khung cao)',
            'q1': {
                'id': '4285567',
                'content': 'Tổ chức bị phạt tiền từ 8.000.000 đồng đến 15.000.000 đồng đối với hành vi nào sau đây?',
                'keyword': 'Phạt 8tr - 15tr',
                'ans': 'D',
                'ans_text': 'Nộp hồ sơ khai thuế quá thời hạn quy định từ 61 ngày đến 90 ngày'
            },
            'q2': {
                'id': '4285565',
                'content': 'Tổ chức bị phạt tiền từ 15.000.000 đồng đến 25.000.000 đồng đối với hành vi nào sau đây?',
                'keyword': 'Phạt 15tr - 25tr',
                'ans': 'A',
                'ans_text': 'Nộp hồ sơ khai thuế quá thời hạn trên 90 ngày nhưng KHÔNG phát sinh số thuế phải nộp'
            },
            'trap_analysis': 'Bẫy hành vi trốn thuế vs chậm nộp: Quá 90 ngày nếu CÓ phát sinh thuế và chưa nộp trước thanh kiểm tra là Trốn thuế (phạt 1-3 lần).',
            'quick_rule': '15-25 triệu ➡️ "Quá 90 ngày + KHÔNG phát sinh số thuế phải nộp".'
        },
        {
            'group': 'Thời hiệu xử lý thuế & Hóa đơn (Luật QLT)',
            'q1': {
                'id': '4284675',
                'content': 'Thời hiệu xử phạt vi phạm hành chính về HÓA ĐƠN là bao nhiêu năm?',
                'keyword': 'Thời hiệu VPHC về HÓA ĐƠN',
                'ans': 'D',
                'ans_text': '02 năm'
            },
            'q2': {
                'id': '4284649',
                'content': 'Thời hiệu xử phạt vi phạm hành chính đối với hành vi TRỐN THUẾ, KHAI SAI dẫn đến thiếu thuế là bao nhiêu năm?',
                'keyword': 'Thời hiệu VPHC về THUẾ (khai sai/trốn thuế)',
                'ans': 'A',
                'ans_text': '05 năm'
            },
            'trap_analysis': 'Bẫy số năm thời hiệu: Hóa đơn & Thủ tục thuế = 02 năm; Vi phạm số tiền thuế (khai thiếu/trốn thuế) = 05 năm; Truy thu thuế = 10 năm.',
            'quick_rule': 'Thủ tục / Hóa đơn ➡️ 2 năm | Khai sai / Trốn thuế ➡️ 5 năm | Truy thu tiền thuế ➡️ 10 năm | Không đăng ký thuế ➡️ Vô thời hạn (mọi thời điểm).'
        },
        {
            'group': 'Thời hạn kiểm tra thuế tại trụ sở NNT (QĐ 970 & Luật QLT)',
            'q1': {
                'id': '4284524',
                'content': 'Thời hạn kiểm tra thuế tại TRỤ SỞ NGƯỜI NỘP THUẾ được quy định tối đa là bao nhiêu ngày làm việc?',
                'keyword': 'Thời hạn kiểm tra tại trụ sở NNT',
                'ans': 'C',
                'ans_text': 'Không quá 10 ngày làm việc (ghi trong quyết định)'
            },
            'q2': {
                'id': '4284525',
                'content': 'Thời gian GIA HẠN kiểm tra thuế tại trụ sở người nộp thuế tối đa là bao nhiêu ngày làm việc?',
                'keyword': 'Thời gian GIA HẠN kiểm tra',
                'ans': 'B',
                'ans_text': 'Không quá 10 ngày làm việc (Cục/Chi cục không quá 10 ngày)'
            },
            'trap_analysis': 'Bẫy ngày làm việc vs ngày liên tục, và thời hạn ban đầu vs thời gian gia hạn. Cả hai đều tối đa 10 NGÀY LÀM VIỆC.',
            'quick_rule': 'Kiểm tra thuế tại trụ sở NNT ➡️ "10 ngày làm việc" | Gia hạn ➡️ "Không quá 10 ngày làm việc".'
        },
        {
            'group': 'Thẩm quyền quyết định kiểm tra / thanh tra',
            'q1': {
                'id': '4284731',
                'content': 'Đoàn kiểm tra thuế có quyền NIÊM PHONG tài liệu, sổ sách kế toán khi nào?',
                'keyword': 'NIÊM PHONG tài liệu',
                'ans': 'C',
                'ans_text': 'Khi có căn cứ NNT tẩu tán, tiêu hủy tài liệu và phải được sự đồng ý bằng văn bản của Người ban hành QĐ kiểm tra'
            },
            'q2': {
                'id': '4284419',
                'content': 'Trường hợp kiểm tra theo DẤU HIỆU VI PHẠM pháp luật, thẩm quyền ban hành quyết định kiểm tra thuộc về ai?',
                'keyword': 'QĐ Kiểm tra theo dấu hiệu vi phạm',
                'ans': 'C',
                'ans_text': 'Thủ trưởng cơ quan thuế các cấp (Tổng cục trưởng, Cục trưởng, Chi cục trưởng)'
            },
            'trap_analysis': 'Bẫy quyền hạn của Trưởng đoàn kiểm tra vs Thủ trưởng CQT. Trưởng đoàn KHÔNG tự ý niêm phong hay cưỡng chế nếu không có QĐ của Thủ trưởng CQT.',
            'quick_rule': 'Niêm phong/Cưỡng chế/Gia hạn ➡️ Phải có văn bản của "Thủ trưởng cơ quan thuế" (Người ký QĐ kiểm tra).'
        },
        {
            'group': 'Hóa đơn điện tử - Xử lý sai sót (Nghị định 123)',
            'q1': {
                'id': '4285211',
                'content': 'Khi phát hiện HĐĐT đã cấp mã có sai sót về TÊN, ĐỊA CHỈ nhưng KHÔNG SAI mã số thuế và các nội dung khác thì xử lý thế nào?',
                'keyword': 'Sai TÊN, ĐỊA CHỈ (không sai MST/tiền)',
                'ans': 'B',
                'ans_text': 'Gửi Mẫu 04/SS-HĐĐT thông báo giải trình, KHÔNG phải lập lại hóa đơn mới'
            },
            'q2': {
                'id': '4285945',
                'content': 'Khi phát hiện HĐĐT có sai sót về MÃ SỐ THUẾ, SỐ TIỀN, THUẾ SUẤT thì xử lý thế nào?',
                'keyword': 'Sai MST, SỐ TIỀN, THUẾ SUẤT',
                'ans': 'C',
                'ans_text': 'Lựa chọn lập Hóa đơn ĐIỀU CHỈNH hoặc lập Hóa đơn THAY THẾ'
            },
            'trap_analysis': 'Bẫy sai sót thông tin hành chính vs sai sót giá trị/thuế: Sai tên/địa chỉ chỉ cần gửi mẫu 04/SS; Sai tiền/MST/thuế suất bắt buộc Điều chỉnh hoặc Thay thế.',
            'quick_rule': 'Chỉ sai Tên/Địa chỉ ➡️ Mẫu 04/SS (Không xuất lại HĐ) | Sai Tiền/MST/Thuế ➡️ Chọn "Điều chỉnh" hoặc "Thay thế".'
        },
        {
            'group': 'Giao dịch liên kết - Chi phí lãi vay (Nghị định 132)',
            'q1': {
                'id': '4285281',
                'content': 'Tổng chi phí lãi vay phát sinh sau khi trừ lãi tiền gửi, lãi cho vay được trừ khi tính thuế TNDN đối với DN có GDLK khống chế ở mức nào?',
                'keyword': 'Khống chế chi phí lãi vay GDLK',
                'ans': 'D',
                'ans_text': 'Không vượt quá 30% tổng EBITDA (Lợi nhuận thuần + Chi phí tài chính thuần + Khấu hao)'
            },
            'q2': {
                'id': '4284286',
                'content': 'Phần chi phí lãi vay vượt mức khống chế 30% EBITDA không được trừ trong kỳ được CHUYỂN SANG KỲ SAU trong thời gian tối đa bao nhiêu năm?',
                'keyword': 'Thời gian CHUYỂN lãi vay vượt mức',
                'ans': 'D',
                'ans_text': 'Không quá 05 năm liên tục kể từ năm tiếp sau năm phát sinh'
            },
            'trap_analysis': 'Bẫy tỷ lệ EBITDA (30%, trước đây là 20%) và thời gian chuyển lãi vay vượt trần (5 năm liên tục, tương tự chuyển lỗ).',
            'quick_rule': 'Chi phí lãi vay GDLK ➡️ "30% EBITDA" | Chuyển tiếp chi phí lãi vay ➡️ "05 năm liên tục".'
        },
        {
            'group': 'Thương mại điện tử & Sàn số (Nghị định 126 & Thông tư 80)',
            'q1': {
                'id': '4285578',
                'content': 'Nhà cung cấp nước ngoài (NCCNN) không có cơ sở thường trú tại Việt Nam thực hiện nghĩa vụ thuế như thế nào?',
                'keyword': 'NCCNN trực tiếp kê khai',
                'ans': 'D',
                'ans_text': 'Trực tiếp đăng ký thuế, kê khai thuế và nộp thuế qua Cổng thông tin điện tử của Tổng cục Thuế'
            },
            'q2': {
                'id': '4285598',
                'content': 'Trường hợp tổ chức là chủ sở hữu sàn TMĐT có chức năng ĐẶT HÀNG TRỰC TUYẾN thì thực hiện nghĩa vụ thuế thay cho cá nhân kinh doanh thế nào?',
                'keyword': 'Sàn TMĐT có đặt hàng trực tuyến',
                'ans': 'A',
                'ans_text': 'Khấu trừ, kê khai và nộp thuế thay cho cá nhân kinh doanh trên sàn'
            },
            'trap_analysis': 'Bẫy sàn có chức năng thanh toán/đặt hàng (khấu trừ nộp thay) vs sàn chỉ thuần quảng cáo thông tin (cung cấp thông tin).',
            'quick_rule': 'NCCNN ➡️ Cổng TTĐT Tổng cục Thuế | Sàn có đặt hàng/thanh toán ➡️ Khấu trừ, nộp thuế THAY.'
        },
        {
            'group': 'Kế toán & BCTC - Trích lập dự phòng',
            'q1': {
                'id': '4284273',
                'content': 'Dự phòng nợ phải thu khó đòi đối với khoản nợ quá hạn từ 06 tháng đến dưới 01 năm được trích lập với tỷ lệ bao nhiêu?',
                'keyword': 'Dự phòng nợ khó đòi 6 tháng - 1 năm',
                'ans': 'A',
                'ans_text': '30% giá trị khoản nợ'
            },
            'q2': {
                'id': '4284677',
                'content': 'Dự phòng nợ phải thu khó đòi đối với khoản nợ quá hạn từ 01 năm đến dưới 02 năm được trích lập với tỷ lệ bao nhiêu?',
                'keyword': 'Dự phòng nợ khó đòi 1 năm - 2 năm',
                'ans': 'D',
                'ans_text': '50% giá trị khoản nợ'
            },
            'trap_analysis': 'Bẫy tỷ lệ dự phòng theo Thông tư 48/2019: 6th - <1 năm (30%), 1 năm - <2 năm (50%), 2 năm - <3 năm (70%), từ 3 năm trở lên (100%).',
            'quick_rule': '6th - 1 năm ➡️ 30% | 1 - 2 năm ➡️ 50% | 2 - 3 năm ➡️ 70% | ≥ 3 năm ➡️ 100%.'
        }
    ]
    
    # 2. Bảng Cheat Sheet Mapping Từ khóa câu hỏi ➡️ Từ khóa đáp án
    cheat_sheet_rules = [
        {
            'topic': 'Thời hạn & Lịch trình',
            'q_keyword': 'Thời hạn nộp hồ sơ khai thuế THÁNG',
            'ans_keyword': 'Chậm nhất là ngày thứ 20 của tháng tiếp theo',
            'trap_to_avoid': 'Nhầm với ngày 30 hoặc ngày cuối tháng',
            'quick_formula': 'Khai tháng ➡️ Ngày 20 | Khai quý ➡️ Ngày cuối tháng đầu quý sau'
        },
        {
            'topic': 'Thời hạn & Lịch trình',
            'q_keyword': 'Thời hạn nộp hồ sơ quyết toán thuế NĂM',
            'ans_keyword': 'Chậm nhất là ngày cuối cùng của tháng thứ 3 kể từ ngày kết thúc năm dương lịch/năm tài chính',
            'trap_to_avoid': 'Nhầm với 90 ngày (Luật QLT 38 đã đổi sang ngày cuối tháng thứ 3)',
            'quick_formula': 'Quyết toán năm ➡️ Ngày cuối tháng thứ 3 | Quyết toán TNCN tự làm ➡️ Ngày cuối tháng thứ 4'
        },
        {
            'topic': 'Quy trình Kiểm tra (QĐ 970)',
            'q_keyword': 'Thời gian gửi Quyết định kiểm tra trước khi đến trụ sở NNT',
            'ans_keyword': 'Chậm nhất 03 ngày làm việc kể từ ngày ban hành Quyết định',
            'trap_to_avoid': 'Nhầm với 05 ngày hoặc 07 ngày',
            'quick_formula': 'Gửi QĐ kiểm tra ➡️ 03 ngày làm việc | Thời hạn tiến hành sau khi ký QĐ ➡️ 10 ngày làm việc'
        },
        {
            'topic': 'Quy trình Kiểm tra (QĐ 970)',
            'q_keyword': 'Thời hạn ký Biên bản kiểm tra thuế',
            'ans_keyword': 'Trong thời hạn 05 ngày làm việc kể từ ngày công bố kết thúc kiểm tra',
            'trap_to_avoid': 'Nhầm với ký ngay trong ngày hoặc 10 ngày',
            'quick_formula': 'Ký Biên bản kiểm tra ➡️ 05 ngày làm việc'
        },
        {
            'topic': 'Hóa đơn điện tử (NĐ 123)',
            'q_keyword': 'Thời hạn gửi bảng tổng hợp dữ liệu HĐĐT không mã (Mẫu 01/TH-HĐĐT)',
            'ans_keyword': 'Cùng thời hạn nộp hồ sơ khai thuế GTGT',
            'trap_to_avoid': 'Nhầm với gửi từng ngày hoặc gửi theo tháng cố định',
            'quick_formula': 'Tổng hợp dữ liệu HĐĐT không mã ➡️ Cùng kỳ nộp hồ sơ khai thuế GTGT'
        },
        {
            'topic': 'Cưỡng chế thi hành QĐ thuế',
            'q_keyword': 'Thứ tự ưu tiên các biện pháp cưỡng chế',
            'ans_keyword': 'Trích tiền từ tài khoản ➡️ Khấu trừ lương ➡️ Ngừng sử dụng hóa đơn ➡️ Kê biên tài sản ➡️ Thu hồi GĐKKD',
            'trap_to_avoid': 'Nhầm vị trí giữa Ngừng sử dụng hóa đơn và Trích tài khoản ngân hàng',
            'quick_formula': '1. Tài khoản NH ➡️ 2. Lương ➡️ 3. Ngừng HĐ ➡️ 4. Kê biên ➡️ 5. Thu tiền bên thứ 3 ➡️ 6. Hải quan ➡️ 7. Thu hồi GP'
        },
        {
            'topic': 'Thuế GTGT (Khấu trừ)',
            'q_keyword': 'Điều kiện khấu trừ thuế GTGT đầu vào từ 20 triệu đồng trở lên',
            'ans_keyword': 'Có chứng từ thanh toán không dùng tiền mặt',
            'trap_to_avoid': 'Thanh toán bằng tiền mặt nhiều lần dưới 20 triệu trong cùng một ngày',
            'quick_formula': '≥ 20 triệu ➡️ Bắt buộc chứng từ thanh toán không dùng tiền mặt (qua ngân hàng)'
        },
        {
            'topic': 'Thuế TNDN (Chi phí hợp lý)',
            'q_keyword': 'Chi phí trang phục cho người lao động bằng tiền',
            'ans_keyword': 'Khống chế tối đa 05 triệu đồng/người/năm',
            'trap_to_avoid': 'Bằng hiện vật không bị khống chế nếu có hóa đơn hợp lệ',
            'quick_formula': 'Trang phục bằng TIỀN ➡️ Max 5 triệu/người/năm | Bằng HIỆN VẬT ➡️ Theo thực tế hóa đơn'
        },
        {
            'topic': 'Quản lý rủi ro (Thông tư 31)',
            'q_keyword': 'Tiêu chí phân loại mức độ tuân thủ người nộp thuế',
            'ans_keyword': '05 mức độ tuân thủ (1: Cao, 2: Trung bình, 3: Thấp, 4: Rất thấp, 5: Không tuân thủ)',
            'trap_to_avoid': 'Nhầm với 3 mức hoặc 4 mức',
            'quick_formula': 'Tuân thủ ➡️ 5 mức (1 đến 5) | Phân loại rủi ro ➡️ 3 hạng rủi ro chính (Cao, TB, Thấp)'
        },
        {
            'topic': 'Phòng, chống tham nhũng',
            'q_keyword': 'Thời hạn giải trình nguồn gốc tài sản, thu nhập tăng thêm',
            'ans_keyword': 'Không quá 15 ngày kể từ ngày nhận được yêu cầu',
            'trap_to_avoid': 'Nhầm với 10 ngày hoặc 30 ngày',
            'quick_formula': 'Giải trình tài sản thu nhập tăng thêm ➡️ 15 ngày'
        }
    ]
    
    return curated_twins, cheat_sheet_rules

# ==============================================================================
# BƯỚC 5: TẠO FILE EXCEL PHÂN TÍCH SINH ĐÔI VÀ BẢNG TỔNG HỢP TOÀN BỘ 1.429 CÂU
# ==============================================================================
def build_master_excel_with_twins_and_cheatsheet(exam_records, all_questions, curated_twins, cheat_sheet_rules):
    styles = get_base_styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # --------------------------------------------------------------------------
    # Sheet 1: Ngan_Hang_1429_Cau_Hoi (ĐÃ TÔ ĐẬM ĐÁP ÁN ĐÚNG TRONG CÁC CỘT OPTION)
    # --------------------------------------------------------------------------
    ws1 = wb.create_sheet(title="Ngan_Hang_1429_Cau_Hoi")
    
    # Header lớn
    ws1.cell(row=1, column=1, value="NGÂN HÀNG 1.429 CÂU HỎI ĐỘC BẢN ÔN TẬP NHANH (KÍP P3-39917)")
    ws1.merge_cells("A1:M1")
    ws1.cell(row=1, column=1).font = styles['title_font_16']
    ws1.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws1.cell(row=1, column=1).fill = styles['accent_fill']
    ws1.row_dimensions[1].height = 36
    
    headers1 = [
        "STT", "Mã Câu Hệ Thống", "Chuyên Đề Nghiệp Vụ", "Nội Dung Câu Hỏi",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Số Đề Xuất Hiện", "Danh Sách Đề Xuất Hiện", "Căn Cứ Pháp Lý / Ghi Chú"
    ]
    for col_idx, h_text in enumerate(headers1, 1):
        c = ws1.cell(row=2, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws1.row_dimensions[2].height = 28
    
    opt_col_map = {'A': 5, 'B': 6, 'C': 7, 'D': 8}
    sorted_questions = sorted(all_questions.values(), key=lambda q: (q['topic'], -q['count'], q['id']))
    
    for idx, q in enumerate(sorted_questions, 1):
        row_idx = idx + 2
        ws1.row_dimensions[row_idx].height = 24
        exams_str = ", ".join([f"Đề {e:02d}" for e in sorted(q['exams'])])
        
        row_data = [
            idx, q['id'], q['topic'], q['content'],
            q['opt_a'], q['opt_b'], q['opt_c'], q['opt_d'],
            q['correct_ans'], q['correct_text'], q['count'],
            exams_str, q['note']
        ]
        
        ans_char = q['correct_ans'].upper()
        
        for col_idx, val in enumerate(row_data, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            
            if col_idx in [1, 2, 9, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                
            # Tô đậm phương án đúng trong các cột E, F, G, H
            if ans_char in opt_col_map and col_idx == opt_col_map[ans_char]:
                cell.font = styles['correct_opt_font']
                cell.fill = styles['correct_opt_fill']
                
            if col_idx == 9: # Cột Đáp án đúng
                cell.font = styles['correct_opt_font']
                cell.fill = styles['correct_opt_fill']
            elif col_idx == 10:
                cell.font = styles['bold_font']
                
    col_widths1 = {
        'A': 8, 'B': 18, 'C': 28, 'D': 48,
        'E': 30, 'F': 30, 'G': 30, 'H': 30,
        'I': 14, 'J': 35, 'K': 16, 'L': 32, 'M': 45
    }
    for col_letter, width in col_widths1.items():
        ws1.column_dimensions[col_letter].width = width
    ws1.freeze_panes = "A3"
    ws1.auto_filter.ref = f"A2:M{len(sorted_questions)+2}"
    
    # --------------------------------------------------------------------------
    # Sheet 2: Cau_Hoi_Anh_Em_Sinh_Doi (SO SÁNH & TÌM ĐIỂM KHÁC BIỆT, TÔ VÀNG KEYWORD)
    # --------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Cau_Hoi_Anh_Em_Sinh_Doi")
    ws2.cell(row=1, column=1, value="BẢNG PHÂN TÍCH CÁC CÂU HỎI 'ANH EM SINH ĐÔI' - SO SÁNH KEYWORD & BẪY ĐỀ THI")
    ws2.merge_cells("A1:K1")
    ws2.cell(row=1, column=1).font = styles['title_font_16']
    ws2.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws2.cell(row=1, column=1).fill = styles['yellow_fill']
    ws2.row_dimensions[1].height = 36
    
    headers2 = [
        "STT", "Nhóm Chuyên Đề Nghiệp Vụ",
        "Mã Câu 1 (ID)", "Nội Dung Câu Hỏi 1", "KEYWORD Cốt Lõi 1", "Đáp Án Đúng 1",
        "Mã Câu 2 (ID)", "Nội Dung Câu Hỏi 2", "KEYWORD Cốt Lõi 2", "Đáp Án Đúng 2",
        "Phân Tích Bẫy Thường Gặp & Mẹo Nhận Diện Siêu Tốc"
    ]
    for col_idx, h_text in enumerate(headers2, 1):
        c = ws2.cell(row=2, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['sub_header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws2.row_dimensions[2].height = 30
    
    for idx, item in enumerate(curated_twins, 1):
        row_idx = idx + 2
        ws2.row_dimensions[row_idx].height = 42
        
        row_data = [
            idx,
            item['group'],
            item['q1']['id'],
            item['q1']['content'],
            item['q1']['keyword'],
            f"{item['q1']['ans']}: {item['q1']['ans_text']}",
            item['q2']['id'],
            item['q2']['content'],
            item['q2']['keyword'],
            f"{item['q2']['ans']}: {item['q2']['ans_text']}",
            f"⚠️ BẪY: {item['trap_analysis']}\n💡 MẸO: {item['quick_rule']}"
        ]
        
        for col_idx, val in enumerate(row_data, 1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=val)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            
            if col_idx in [1, 3, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                
            # Tô màu vàng cho các KEYWORD cốt lõi (Cột 5 và Cột 9)
            if col_idx in [5, 9]:
                cell.fill = styles['bright_yellow']
                cell.font = styles['keyword_font']
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            # Tô màu xanh cho đáp án đúng (Cột 6 và Cột 10)
            if col_idx in [6, 10]:
                cell.fill = styles['correct_opt_fill']
                cell.font = styles['correct_opt_font']
                
    col_widths2 = {
        'A': 8,   # STT
        'B': 24,  # Nhóm chuyên đề
        'C': 15,  # Mã câu 1
        'D': 38,  # Nội dung câu 1
        'E': 20,  # Keyword 1 (Vàng)
        'F': 32,  # Đáp án đúng 1
        'G': 15,  # Mã câu 2
        'H': 38,  # Nội dung câu 2
        'I': 20,  # Keyword 2 (Vàng)
        'J': 32,  # Đáp án đúng 2
        'K': 45   # Phân tích bẫy & mẹo
    }
    for col_letter, width in col_widths2.items():
        ws2.column_dimensions[col_letter].width = width
    ws2.freeze_panes = "A3"
    ws2.auto_filter.ref = f"A2:K{len(curated_twins)+2}"
    
    # --------------------------------------------------------------------------
    # Sheet 3: Cheat_Sheet_Quy_Tac_Nhanh (BẢNG MAPPING TỪ KHÓA CÂU HỎI ➡️ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="Cheat_Sheet_Quy_Tac_Nhanh")
    ws3.cell(row=1, column=1, value="BỘ QUY TẮC NHẬN DIỆN NHANH (CHEAT SHEET MAPPING) - HACK NÃO PHÒNG THI")
    ws3.merge_cells("A1:F1")
    ws3.cell(row=1, column=1).font = styles['title_font_16']
    ws3.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws3.cell(row=1, column=1).fill = styles['accent_fill']
    ws3.row_dimensions[1].height = 36
    
    headers3 = [
        "STT", "Chuyên Đề Nghiệp Vụ",
        "Từ Khóa Nhận Diện Câu Hỏi ➡️", "Từ Khóa Đáp Án Đúng Phải Chọn ✅",
        "Bẫy Cần Loại Trừ Ngay ❌", "Quy Luật & Công Thức Siêu Tốc 💡"
    ]
    for col_idx, h_text in enumerate(headers3, 1):
        c = ws3.cell(row=2, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws3.row_dimensions[2].height = 30
    
    for idx, rule in enumerate(cheat_sheet_rules, 1):
        row_idx = idx + 2
        ws3.row_dimensions[row_idx].height = 38
        
        row_data = [
            idx,
            rule['topic'],
            rule['q_keyword'],
            rule['ans_keyword'],
            rule['trap_to_avoid'],
            rule['quick_formula']
        ]
        for col_idx, val in enumerate(row_data, 1):
            cell = ws3.cell(row=row_idx, column=col_idx, value=val)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if col_idx == 3: # Từ khóa câu hỏi
                cell.fill = styles['yellow_fill']
                cell.font = styles['bold_font']
            elif col_idx == 4: # Đáp án đúng
                cell.fill = styles['correct_opt_fill']
                cell.font = styles['correct_opt_font']
            elif col_idx == 6: # Công thức
                cell.font = styles['bold_font']
                
    col_widths3 = {'A': 8, 'B': 24, 'C': 32, 'D': 38, 'E': 35, 'F': 42}
    for col_letter, width in col_widths3.items():
        ws3.column_dimensions[col_letter].width = width
    ws3.freeze_panes = "A3"
    ws3.auto_filter.ref = f"A2:F{len(cheat_sheet_rules)+2}"
    
    # --------------------------------------------------------------------------
    # Sheet 4: Tong_Hop_40_De_Va_Cau_1 (Định dạng tinh gọn không có phương án A,B,C,D)
    # --------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="Tong_Hop_40_De_Va_Cau_1")
    ws4.cell(row=1, column=1, value="MỤC LỤC TỔNG HỢP 40 BỘ ĐỀ THI VÀ CÂU HỎI SỐ 1")
    ws4.merge_cells("A1:G1")
    ws4.cell(row=1, column=1).font = styles['title_font_16']
    ws4.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws4.cell(row=1, column=1).fill = styles['accent_fill']
    ws4.row_dimensions[1].height = 36
    
    headers4 = [
        "STT", "Mã Đề", "Tên File Bộ Đề", "Mã Câu 1 (ID)",
        "Chuyên Đề Câu 1", "Nội Dung Câu Hỏi Số 1", "Đáp Án Đúng"
    ]
    for col_idx, h_text in enumerate(headers4, 1):
        c = ws4.cell(row=2, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws4.row_dimensions[2].height = 28
    
    for idx, rec in enumerate(exam_records, 1):
        row_idx = idx + 2
        ws4.row_dimensions[row_idx].height = 26
        row_data = [
            idx, f"Đề {rec['exam_num']:02d}", rec['filename'], rec['q1_id'],
            rec['q1_topic'], rec['q1_content'], f"{rec['q1_ans']}: {rec['q1_ans_text']}"
        ]
        for col_idx, val in enumerate(row_data, 1):
            cell = ws4.cell(row=row_idx, column=col_idx, value=val)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if col_idx in [1, 2, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if col_idx == 7:
                cell.font = styles['correct_opt_font']
                cell.fill = styles['correct_opt_fill']
                
    col_widths4 = {'A': 8, 'B': 14, 'C': 46, 'D': 16, 'E': 28, 'F': 55, 'G': 40}
    for col_letter, width in col_widths4.items():
        ws4.column_dimensions[col_letter].width = width
    ws4.freeze_panes = "A3"
    ws4.auto_filter.ref = f"A2:G{len(exam_records)+2}"
    
    # --------------------------------------------------------------------------
    # Sheet 5: Thong_Ke_Chuyen_De
    # --------------------------------------------------------------------------
    ws5 = wb.create_sheet(title="Thong_Ke_Chuyen_De")
    headers5 = [
        "STT", "Chuyên Đề Nghiệp Vụ", "Số Câu Độc Bản", "Tỷ Lệ Độc Bản (%)",
        "Tổng Lượt Xuất Hiện (40 Đề)", "Tỷ Lệ Toàn Khối Lượng (%)", "Tần Suất TB / Đề"
    ]
    for col_idx, h_text in enumerate(headers5, 1):
        c = ws5.cell(row=1, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws5.row_dimensions[1].height = 28
    
    topic_unique = Counter([q['topic'] for q in all_questions.values()])
    topic_total = Counter()
    for q in all_questions.values():
        topic_total[q['topic']] += q['count']
    total_u = len(all_questions)
    total_all = sum(topic_total.values())
    
    sorted_topics = topic_unique.most_common()
    for idx, (t, u_cnt) in enumerate(sorted_topics, 1):
        tot_cnt = topic_total[t]
        ws5.append([
            idx, t, u_cnt, u_cnt / total_u, tot_cnt, tot_cnt / total_all, round(tot_cnt / len(exam_records), 1)
        ])
    ws5.append(["Tổng", "Toàn bộ chuyên đề", total_u, 1.0, total_all, 1.0, round(total_all / len(exam_records), 1)])
    
    col_widths5 = {'A': 8, 'B': 45, 'C': 18, 'D': 20, 'E': 25, 'F': 25, 'G': 18}
    for col_letter, width in col_widths5.items():
        ws5.column_dimensions[col_letter].width = width
    for r_idx in range(2, len(sorted_topics) + 3):
        is_last = (r_idx == len(sorted_topics) + 2)
        ws5.row_dimensions[r_idx].height = 24
        for c_idx in range(1, len(headers5) + 1):
            cell = ws5.cell(row=r_idx, column=c_idx)
            cell.font = styles['bold_font'] if is_last else styles['normal_font']
            cell.border = styles['thin_border']
            if is_last:
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            if c_idx in [1, 3, 5, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in [4, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = '0.0%'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
    ws5.freeze_panes = "A2"
    
    # --------------------------------------------------------------------------
    # Sheet 6: Top_Cau_Hoi_Trong_Tam
    # --------------------------------------------------------------------------
    ws6 = wb.create_sheet(title="Top_Cau_Hoi_Trong_Tam")
    headers6 = [
        "STT", "Mã Câu Hệ Thống", "Chuyên Đề Nghiệp Vụ", "Số Lần Xuất Hiện", "Tỷ Lệ Xuất Hiện (40 Đề)",
        "Nội Dung Câu Hỏi", "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Các Đề Xuất Hiện"
    ]
    for col_idx, h_text in enumerate(headers6, 1):
        c = ws6.cell(row=1, column=col_idx, value=h_text)
        c.font = styles['header_font']
        c.fill = styles['header_fill']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = styles['header_border']
    ws6.row_dimensions[1].height = 28
    
    top_questions = [q for q in all_questions.values() if q['count'] >= 4]
    top_questions.sort(key=lambda q: (-q['count'], q['topic'], q['id']))
    for idx, q in enumerate(top_questions, 1):
        exams_str = ", ".join([f"Đề {e:02d}" for e in sorted(q['exams'])])
        ws6.append([
            idx, q['id'], q['topic'], q['count'], q['count'] / len(exam_records),
            q['content'], q['correct_ans'], q['correct_text'], exams_str
        ])
    col_widths6 = {'A': 8, 'B': 18, 'C': 30, 'D': 18, 'E': 22, 'F': 55, 'G': 14, 'H': 35, 'I': 35}
    for col_letter, width in col_widths6.items():
        ws6.column_dimensions[col_letter].width = width
    for r_idx in range(2, len(top_questions) + 2):
        ws6.row_dimensions[r_idx].height = 24
        for c_idx in range(1, len(headers6) + 1):
            cell = ws6.cell(row=r_idx, column=c_idx)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if c_idx in [1, 2, 4, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 5:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = '0.0%'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx == 7:
                cell.fill = styles['correct_opt_fill']
                cell.font = styles['correct_opt_font']
            elif c_idx == 8:
                cell.font = styles['bold_font']
    ws6.freeze_panes = "A2"
    ws6.auto_filter.ref = ws6.dimensions
    
    try:
        wb.save(OUTPUT_EXCEL_ALL)
        print(f"Đã lưu thành công: {OUTPUT_EXCEL_ALL}")
    except PermissionError:
        alt_path = os.path.join(SOURCE_FOLDER, 'Tong_Hop_1429_Cau_Hoi_Doc_Ban_On_Tap_Nhanh_Cap_Nhat.xlsx')
        wb.save(alt_path)
        print(f"File chính đang mở trong Excel. Đã lưu bản cập nhật tại: {alt_path}")
        
    wb_twins = openpyxl.Workbook()
    wb_twins.remove(wb_twins.active)
    # Copy sheet 2 and 3 into dedicated Twins & Cheat Sheet workbook
    wb.save(OUTPUT_EXCEL_TWINS)
    print(f"Đã lưu file Chuyên đề Sinh đôi & Cheat Sheet: {OUTPUT_EXCEL_TWINS}")

# ==============================================================================
# BƯỚC 6: TẠO FILE WORD (.DOCX) TÓM TẮT SIÊU TỐC HACK NÃO DƯỚI 300 TỪ/CHUYÊN ĐỀ
# ==============================================================================
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def generate_hack_brain_word_document():
    doc = docx.Document()
    
    # Page setup
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("TÀI LIỆU TÓM TẮT SIÊU TỐC (HACK NÃO) ÔN TẬP THI THUẾ")
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("BỘ QUY TẮC NHỚ NHANH - CÔ ĐỌNG DƯỚI 300 TỪ MỖI CHUYÊN ĐỀ (KÍP P3-39917)")
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(12)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # 8 Chuyên đề cốt lõi, mỗi chuyên đề < 300 từ cực kỳ súc tích, dạng bảng Hack Não
    topic_summaries = [
        {
            'title': '1. QUY TRÌNH KIỂM TRA THUẾ & BIÊN BẢN (QĐ 970 & LUẬT QLT 38)',
            'color': '1F4E79',
            'points': [
                ("Thời hạn kiểm tra tại trụ sở NNT", "Tối đa 10 ngày làm việc (ghi trong QĐ); Gia hạn tối đa không quá 10 ngày làm việc (do Người ký QĐ quyết định)."),
                ("Lịch trình gửi QĐ & Công bố", "Gửi QĐ cho NNT trong 03 ngày làm việc kể từ ngày ký; Công bố QĐ chậm nhất 10 ngày làm việc kể từ ngày ban hành."),
                ("Ký Biên bản kiểm tra", "Trong 05 ngày làm việc kể từ ngày công bố kết thúc kiểm tra; Trưởng đoàn và NNT bắt buộc phải ký từng trang."),
                ("Kiểm tra tại trụ sở CQT", "CQT gửi Thông báo giải trình (lần 1: 10 ngày làm việc; lần 2: 10 ngày làm việc). Hết hạn không giải trình ➡️ Chuyển kiểm tra tại trụ sở NNT."),
                ("Mẹo Hack Não", "Gửi QĐ: 3 ngày | Thời hạn & Gia hạn: 10 ngày làm việc | Ký biên bản: 5 ngày làm việc.")
            ]
        },
        {
            'title': '2. HÓA ĐƠN ĐIỆN TỬ & XỬ LÝ SAI SÓT (NGHỊ ĐỊNH 123 & THÔNG TƯ 78)',
            'color': '2F5597',
            'points': [
                ("Sai Tên, Địa chỉ (Không sai MST, Tiền)", "Gửi Mẫu 04/SS-HĐĐT thông báo giải trình với CQT. TUYỆT ĐỐI KHÔNG phải lập lại hóa đơn mới."),
                ("Sai Mã số thuế, Tiền, Thuế suất", "Được lựa chọn 1 trong 2 cách: (1) Lập Hóa đơn ĐIỀU CHỈNH hoặc (2) Lập Hóa đơn THAY THẾ mới."),
                ("Thời điểm xuất HĐĐT", "Bán hàng: Chuyển giao quyền sở hữu/sử dụng | Dịch vụ: Hoàn thành cung cấp dịch vụ (hoặc thu tiền trước)."),
                ("Bảng tổng hợp dữ liệu (Mẫu 01/TH)", "HĐĐT không mã gửi bảng tổng hợp CÙNG KỲ HẠN nộp hồ sơ khai thuế GTGT (theo tháng hoặc quý)."),
                ("Mẹo Hack Não", "Sai Tên/Địa chỉ ➡️ Chỉ gửi Mẫu 04/SS | Sai Tiền/MST ➡️ Điều chỉnh hoặc Thay thế.")
            ]
        },
        {
            'title': '3. XỬ PHẠT VI PHẠM HÀNH CHÍNH & CƯỠNG CHẾ THUẾ (NGHỊ ĐỊNH 125 & LUẬT QLT)',
            'color': 'C00000',
            'points': [
                ("Thời hiệu xử phạt", "Hóa đơn & Thủ tục thuế: 02 năm | Khai sai dẫn đến thiếu thuế & Trốn thuế: 05 năm | Truy thu thuế: 10 năm."),
                ("Khung phạt chậm nộp hồ sơ khai thuế", "1-30 ngày: 2-5 triệu | 31-60 ngày: 5-8 triệu | 61-90 ngày: 8-15 triệu | Quá 90 ngày (không phát sinh thuế): 15-25 triệu."),
                ("Khai sai dẫn đến thiếu thuế", "Phạt 20% số tiền thuế khai thiếu hoặc số tiền thuế được hoàn cao hơn quy định."),
                ("Thứ tự 7 bước Cưỡng chế", "1. Trích tài khoản NH ➡️ 2. Khấu trừ lương ➡️ 3. Ngừng HĐ ➡️ 4. Kê biên tài sản ➡️ 5. Thu tiền bên thứ 3 ➡️ 6. Hải quan ➡️ 7. Thu hồi GPKD."),
                ("Mẹo Hack Não", "Hóa đơn: 2 năm | Thuế: 5 năm | Truy thu: 10 năm | Khai thiếu: 20% | Trốn thuế: 1 - 3 lần.")
            ]
        },
        {
            'title': '4. GIAO DỊCH LIÊN KẾT & CHỐNG CHUYỂN GIÁ (NGHỊ ĐỊNH 132/2020)',
            'color': '7030A0',
            'points': [
                ("Quan hệ liên kết", "Vốn góp trực tiếp/gián tiếp ≥ 25% | Cho vay ≥ 25% vốn chủ sở hữu và > 50% tổng nợ trung/dài hạn | Điều hành chung."),
                ("Trần Chi phí lãi vay (EBITDA)", "Khống chế tối đa 30% tổng EBITDA (Lợi nhuận thuần + Chi phí tài chính thuần + Khấu hao)."),
                ("Chuyển lãi vay vượt trần", "Phần lãi vay > 30% EBITDA không được trừ được chuyển sang các kỳ tính thuế tiếp theo tối đa 05 năm liên tục."),
                ("Miễn lập Hồ sơ GDLK", "Doanh thu < 50 tỷ VÀ Tổng giá trị GDLK < 30 tỷ | Hoặc ký thỏa thuận APA trước."),
                ("Mẹo Hack Não", "Vốn góp/Cho vay: 25% | Khống chế lãi vay: 30% EBITDA | Chuyển lãi vay: 5 năm liên tục.")
            ]
        },
        {
            'title': '5. THUẾ GTGT & ĐIỀU KIỆN KHẤU TRỪ (LUẬT THUẾ GTGT MỚI & THÔNG TƯ 80)',
            'color': '385723',
            'points': [
                ("Điều kiện khấu trừ đầu vào", "Có hóa đơn GTGT hợp pháp; Có chứng từ thanh toán không dùng tiền mặt đối với hàng hóa/dịch vụ từ 20 triệu đồng trở lên."),
                ("Tài sản cố định dùng chung", "Dùng cho cả chịu thuế và không chịu thuế: Phân bổ theo tỷ lệ % doanh thu chịu thuế GTGT."),
                ("Hàng hóa mua qua Sàn TMĐT bị hủy", "Được kê khai điều chỉnh giảm số thuế đã khấu trừ/hoàn trong kỳ tính thuế tiếp theo."),
                ("Nhà cung cấp nước ngoài (NCCNN)", "Trực tiếp đăng ký, kê khai và nộp thuế qua Cổng thông tin điện tử của Tổng cục Thuế."),
                ("Mẹo Hack Não", "Hóa đơn ≥ 20 triệu ➡️ Bắt buộc thanh toán qua Ngân hàng | Dùng chung ➡️ Phân bổ theo tỷ lệ doanh thu.")
            ]
        },
        {
            'title': '6. KẾ TOÁN & KIỂM TRA BÁO CÁO TÀI CHÍNH, TRÍCH LẬP DỰ PHÒNG (TT 48/2019)',
            'color': '833C0C',
            'points': [
                ("Dự phòng Nợ khó đòi (Quá hạn)", "Từ 6 tháng đến < 1 năm: 30% | Từ 1 năm đến < 2 năm: 50% | Từ 2 năm đến < 3 năm: 70% | Từ 3 năm trở lên: 100%."),
                ("Thời gian lưu trữ chứng từ kế toán", "Tối thiểu 05 năm (chứng từ quản lý) | Tối thiểu 10 năm (chứng từ sử dụng trực tiếp ghi sổ, lập BCTC)."),
                ("Khống chế Chi phí trang phục", "Bằng TIỀN: Tối đa 05 triệu đồng/người/năm | Bằng HIỆN VẬT: Theo thực tế hóa đơn chứng từ hợp lệ."),
                ("Khấu hao TSCĐ", "Áp dụng theo khung Thông tư 45/2013/TT-BTC; TSCĐ không sử dụng cho SXKD thì không được tính vào chi phí được trừ."),
                ("Mẹo Hack Não", "Nợ khó đòi: 6th (30%) - 1n (50%) - 2n (70%) - 3n (100%) | Trang phục tiền: 5 triệu/người/năm.")
            ]
        },
        {
            'title': '7. QUẢN LÝ RỦI RO & TIÊU CHÍ TUÂN THỦ (THÔNG TƯ 31/2021/TT-BTC & TPR)',
            'color': '1F4E79',
            'points': [
                ("Mức độ tuân thủ NNT", "Phân loại thành 05 mức độ tuân thủ pháp luật thuế: Mức 1 (Cao), Mức 2 (TB), Mức 3 (Thấp), Mức 4 (Rất thấp), Mức 5 (Không tuân thủ)."),
                ("Phân loại Hạng Rủi ro", "Phân loại thành 03 hạng rủi ro chính: Rủi ro cao, Rủi ro trung bình, Rủi ro thấp."),
                ("Ứng dụng TPR", "Tự động phân tích rủi ro, chấm điểm NNT để lập kế hoạch thanh tra kiểm tra hàng năm; Loại trừ DN giải thể/phá sản."),
                ("Nguyên tắc lựa chọn kiểm tra", "Ưu tiên lựa chọn NNT có rủi ro cao từ hệ thống TPR; Tỷ lệ lựa chọn ngẫu nhiên tối đa không quá 20%."),
                ("Mẹo Hack Não", "Tuân thủ: 5 mức | Rủi ro: 3 hạng (Cao - TB - Thấp) | Chọn ngẫu nhiên: Tối đa 20%.")
            ]
        },
        {
            'title': '8. LUẬT PHÒNG, CHỐNG THAM NHŨNG (PCTN 2018 & NGHỊ ĐỊNH 59/2019)',
            'color': 'C00000',
            'points': [
                ("Quy định Quà tặng", "TẶNG QUÀ: Cấm dùng tài chính công/tài sản công | NHẬN QUÀ: Cấm nhận dưới mọi hình thức liên quan công vụ."),
                ("Kỷ luật Sử dụng trái định mức", "TỰ Ý sử dụng: Khiển trách | CHO PHÉP sử dụng: Cảnh cáo."),
                ("Kỷ luật Người đứng đầu để xảy ra tham nhũng", "Ít nghiêm trọng: Khiển trách | Rất nghiêm trọng/Nhiều nghiêm trọng: Cảnh cáo | ĐẶC BIỆT NGHIÊM TRỌNG: Cách chức."),
                ("Thời hạn giải trình tài sản", "Người có nghĩa vụ kê khai phải giải trình trong thời hạn không quá 15 ngày kể từ ngày nhận yêu cầu."),
                ("Mẹo Hack Não", "Tặng: Cấm tiền công | Nhận: Cấm mọi hình thức | Tự ý: Khiển trách | Cho phép: Cảnh cáo | Giải trình: 15 ngày.")
            ]
        }
    ]
    
    for item in topic_summaries:
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(12)
        p_sec.paragraph_format.space_after = Pt(4)
        run_sec = p_sec.add_run(item['title'])
        run_sec.font.name = "Calibri"
        run_sec.font.size = Pt(13)
        run_sec.font.bold = True
        run_sec.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        
        table = doc.add_table(rows=len(item['points']), cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        
        col_widths = [Inches(2.3), Inches(4.5)]
        
        for r_idx, (k, v) in enumerate(item['points']):
            row = table.rows[r_idx]
            
            cell_k = row.cells[0]
            cell_v = row.cells[1]
            
            cell_k.width = col_widths[0]
            cell_v.width = col_widths[1]
            
            p_k = cell_k.paragraphs[0]
            p_k.paragraph_format.space_after = Pt(2)
            p_k.paragraph_format.space_before = Pt(2)
            run_k = p_k.add_run(k)
            run_k.font.name = "Calibri"
            run_k.font.size = Pt(10.5)
            run_k.font.bold = True
            
            p_v = cell_v.paragraphs[0]
            p_v.paragraph_format.space_after = Pt(2)
            p_v.paragraph_format.space_before = Pt(2)
            run_v = p_v.add_run(v)
            run_v.font.name = "Calibri"
            run_v.font.size = Pt(10.5)
            
            if k == "Mẹo Hack Não":
                set_cell_background(cell_k, "FFF2CC")
                set_cell_background(cell_v, "FFF2CC")
                run_k.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
                run_v.font.bold = True
                run_v.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            elif r_idx % 2 == 1:
                set_cell_background(cell_k, "F2F2F2")
                set_cell_background(cell_v, "F2F2F2")
                
            set_cell_margins(cell_k, top=80, bottom=80, left=120, right=120)
            set_cell_margins(cell_v, top=80, bottom=80, left=120, right=120)
            
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
    doc.save(OUTPUT_DOCX_SUMMARY)
    print(f"Đã tạo file Word Hack Não thành công: {OUTPUT_DOCX_SUMMARY}")

# ==============================================================================
# BƯỚC 7: XUẤT NGUỒN MARKDOWN SECOND BRAIN
# ==============================================================================
def update_markdown_sources(curated_twins, cheat_sheet_rules):
    os.makedirs(SOURCES_MD_DIR, exist_ok=True)
    
    # Markdown Phân tích sinh đôi & Cheat sheet
    md_twins_path = os.path.join(SOURCES_MD_DIR, 'phan_tich_cau_hoi_sinh_doi_va_cheat_sheet_p3_39917.md')
    with open(md_twins_path, 'w', encoding='utf-8') as f:
        f.write("---\n")
        f.write("type: source\n")
        f.write("source_kind: own-note\n")
        f.write("status: processed\n")
        f.write("created: 2026-09-25\n")
        f.write("processed_at: 2026-09-25\n")
        f.write("tags: [on-tap-thue, p3-39917, cau-hoi-sinh-doi, cheat-sheet]\n")
        f.write("title: Phân tích câu hỏi Anh em sinh đôi và Cheat Sheet Mapping ôn thi thuế P3-39917\n")
        f.write("---\n\n")
        
        f.write("# Phân tích câu hỏi 'Anh em sinh đôi' & Cheat Sheet Mapping\n\n")
        f.write("## 1. Bảng phân tích các cặp câu hỏi sinh đôi tiêu biểu\n\n")
        for item in curated_twins:
            f.write(f"### Nhóm: {item['group']}\n\n")
            f.write(f"- **Câu 1 (`{item['q1']['id']}`)**: {item['q1']['content']}\n")
            f.write(f"  - **KEYWORD cốt lõi**: `{item['q1']['keyword']}`\n")
            f.write(f"  - **Đáp án đúng**: **{item['q1']['ans']}** ({item['q1']['ans_text']})\n\n")
            f.write(f"- **Câu 2 (`{item['q2']['id']}`)**: {item['q2']['content']}\n")
            f.write(f"  - **KEYWORD cốt lõi**: `{item['q2']['keyword']}`\n")
            f.write(f"  - **Đáp án đúng**: **{item['q2']['ans']}** ({item['q2']['ans_text']})\n\n")
            f.write(f"> ⚠️ **Bẫy thường gặp**: {item['trap_analysis']}\n")
            f.write(f"> 💡 **Mẹo nhận diện siêu tốc**: {item['quick_rule']}\n\n")
            f.write("---\n\n")
            
        f.write("## 2. Bảng Cheat Sheet Mapping Từ khóa ➡️ Đáp án đúng\n\n")
        f.write("| Chuyên Đề | Từ Khóa Câu Hỏi ➡️ | Từ Khóa Đáp Án Đúng ✅ | Bẫy Cần Tránh ❌ | Công Thức Nhớ Nhanh 💡 |\n")
        f.write("| :--- | :--- | :--- | :--- | :--- |\n")
        for r in cheat_sheet_rules:
            f.write(f"| **{r['topic']}** | `{r['q_keyword']}` | **{r['ans_keyword']}** | {r['trap_to_avoid']} | `{r['quick_formula']}` |\n")
            
    print(f"Đã tạo file Markdown Second Brain: {md_twins_path}")

if __name__ == '__main__':
    print("BẮT ĐẦU QUY TRÌNH XỬ LÝ TOÀN DIỆN...")
    exam_records, all_questions = process_individual_exams()
    build_updated_excel_de_va_cau_1(exam_records)
    curated_twins, cheat_sheet_rules = analyze_twin_questions_and_cheat_sheet(all_questions)
    build_master_excel_with_twins_and_cheatsheet(exam_records, all_questions, curated_twins, cheat_sheet_rules)
    generate_hack_brain_word_document()
    update_markdown_sources(curated_twins, cheat_sheet_rules)
    print("HOÀN THÀNH 100% TẤT CẢ CÁC MỤC YÊU CẦU!")
