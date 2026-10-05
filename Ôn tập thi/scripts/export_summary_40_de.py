# -*- coding: utf-8 -*-
"""
Kịch bản tổng hợp ngân hàng 1.429 câu hỏi độc bản và danh sách 40 đề thi kèm Câu 1 kíp P3-39917.
Lưu tại D:\HOC TAP\ON THI 2026.09\40 DE THI\
"""

import os
import glob
import re
from collections import Counter, defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SOURCE_FOLDER = r'D:\HOC TAP\ON THI 2026.09\40 DE THI'
OUTPUT_EXCEL_ALL = os.path.join(SOURCE_FOLDER, 'Tong_Hop_1429_Cau_Hoi_Doc_Ban_On_Tap_Nhanh.xlsx')
OUTPUT_EXCEL_DE = os.path.join(SOURCE_FOLDER, 'Danh_Sach_Tong_Hop_40_De_Va_Cau_1.xlsx')
SOURCES_MD_DIR = r'D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\sources'

def create_styles():
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    sub_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    ans_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    ans_font = Font(name="Calibri", size=11, bold=True, color="276A3C")
    
    bold_font = Font(name="Calibri", size=11, bold=True)
    normal_font = Font(name="Calibri", size=11)
    
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
        'header_font': header_font,
        'ans_fill': ans_fill,
        'ans_font': ans_font,
        'bold_font': bold_font,
        'normal_font': normal_font,
        'thin_border': thin_border,
        'header_border': header_border
    }

def process_data():
    files = sorted(glob.glob(os.path.join(SOURCE_FOLDER, 'Bo_De_On_Thi_Kiem_Tra_Thue_P3_39917_De_*.xlsx')),
                   key=lambda x: int(re.search(r'De_(\d+)', x).group(1)))
    
    print(f"Đang xử lý {len(files)} file đề thi...")
    
    all_questions = {}
    exam_records = []
    
    for fpath in files:
        fname = os.path.basename(fpath)
        exam_num = int(re.search(r'De_(\d+)', fname).group(1))
        wb = openpyxl.load_workbook(fpath, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))[1:] # bỏ header
        
        q1 = rows[0]
        exam_records.append({
            'exam_num': exam_num,
            'filename': fname,
            'total_q': len(rows),
            'q1_id': str(q1[1]),
            'q1_topic': str(q1[2] or ''),
            'q1_content': str(q1[3] or ''),
            'q1_a': str(q1[4] or ''),
            'q1_b': str(q1[5] or ''),
            'q1_c': str(q1[6] or ''),
            'q1_d': str(q1[7] or ''),
            'q1_ans': str(q1[11] or ''),
            'q1_ans_text': str(q1[12] or ''),
            'q1_note': str(q1[13] or '')
        })
        
        for r in rows:
            stt, item_id, topic, content, opt_a, opt_b, opt_c, opt_d, sel_opt, opt_src, first_res, correct_ans, correct_text, note = r[:14]
            q_id = str(item_id).strip()
            if q_id not in all_questions:
                all_questions[q_id] = {
                    'id': q_id,
                    'topic': str(topic or 'Nghiệp vụ kiểm tra thuế khác').strip(),
                    'content': str(content or '').strip(),
                    'opt_a': str(opt_a or '').strip(),
                    'opt_b': str(opt_b or '').strip(),
                    'opt_c': str(opt_c or '').strip(),
                    'opt_d': str(opt_d or '').strip(),
                    'correct_ans': str(correct_ans or '').strip(),
                    'correct_text': str(correct_text or '').strip(),
                    'note': str(note or '').strip(),
                    'count': 1,
                    'exams': [exam_num]
                }
            else:
                all_questions[q_id]['count'] += 1
                all_questions[q_id]['exams'].append(exam_num)
                if note and len(str(note)) > len(str(all_questions[q_id]['note'] or '')):
                    all_questions[q_id]['note'] = str(note).strip()
                if not all_questions[q_id]['correct_ans'] and correct_ans:
                    all_questions[q_id]['correct_ans'] = str(correct_ans).strip()
                    all_questions[q_id]['correct_text'] = str(correct_text).strip()

    return exam_records, all_questions

def build_excel_full(exam_records, all_questions):
    styles = create_styles()
    wb = openpyxl.Workbook()
    # Xóa default sheet
    wb.remove(wb.active)
    
    # -------------------------------------------------------------
    # Sheet 1: Ngan_Hang_1429_Cau_Hoi
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="Ngan_Hang_1429_Cau_Hoi")
    headers1 = [
        "STT", "Mã Câu Hệ Thống", "Chuyên Đề Nghiệp Vụ", "Nội Dung Câu Hỏi",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Số Đề Xuất Hiện", "Danh Sách Đề Xuất Hiện", "Căn Cứ Pháp Lý / Ghi Chú"
    ]
    ws1.append(headers1)
    
    # Sort questions: Chuyên đề, rồi đến số lần xuất hiện giảm dần, rồi đến Mã câu
    sorted_questions = sorted(all_questions.values(), key=lambda q: (q['topic'], -q['count'], q['id']))
    
    for idx, q in enumerate(sorted_questions, 1):
        exams_str = ", ".join([f"Đề {e:02d}" for e in sorted(q['exams'])])
        row_data = [
            idx,
            q['id'],
            q['topic'],
            q['content'],
            q['opt_a'],
            q['opt_b'],
            q['opt_c'],
            q['opt_d'],
            q['correct_ans'],
            q['correct_text'],
            q['count'],
            exams_str,
            q['note']
        ]
        ws1.append(row_data)
        
    # Format Sheet 1
    col_widths1 = {
        'A': 8,   # STT
        'B': 18,  # Mã câu
        'C': 30,  # Chuyên đề
        'D': 50,  # Nội dung câu hỏi
        'E': 30,  # Phương án A
        'F': 30,  # Phương án B
        'G': 30,  # Phương án C
        'H': 30,  # Phương án D
        'I': 14,  # Đáp án đúng
        'J': 35,  # Nội dung đáp án đúng
        'K': 16,  # Số đề xuất hiện
        'L': 32,  # Danh sách đề
        'M': 45   # Căn cứ pháp lý
    }
    
    for col_letter, width in col_widths1.items():
        ws1.column_dimensions[col_letter].width = width
        
    for col_num in range(1, len(headers1) + 1):
        cell = ws1.cell(row=1, column=col_num)
        cell.fill = styles['header_fill']
        cell.font = styles['header_font']
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles['header_border']
    ws1.row_dimensions[1].height = 28
    
    for r_idx in range(2, len(sorted_questions) + 2):
        ws1.row_dimensions[r_idx].height = 24
        for c_idx in range(1, len(headers1) + 1):
            cell = ws1.cell(row=r_idx, column=c_idx)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if c_idx in [1, 2, 9, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx == 9: # Đáp án đúng
                cell.fill = styles['ans_fill']
                cell.font = styles['ans_font']
                
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions

    # -------------------------------------------------------------
    # Sheet 2: Tong_Hop_40_De_Va_Cau_1
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Tong_Hop_40_De_Va_Cau_1")
    headers2 = [
        "STT Đề", "Tên File Bộ Đề", "Tổng Số Câu", "Mã Câu 1 (ID)",
        "Chuyên Đề Câu 1", "Nội Dung Câu Hỏi Số 1",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Căn Cứ Pháp Lý / Ghi Chú"
    ]
    ws2.append(headers2)
    
    for rec in exam_records:
        row_data = [
            f"Đề {rec['exam_num']:02d}",
            rec['filename'],
            rec['total_q'],
            rec['q1_id'],
            rec['q1_topic'],
            rec['q1_content'],
            rec['q1_a'],
            rec['q1_b'],
            rec['q1_c'],
            rec['q1_d'],
            rec['q1_ans'],
            rec['q1_ans_text'],
            rec['q1_note']
        ]
        ws2.append(row_data)
        
    col_widths2 = {
        'A': 10,  # STT Đề
        'B': 45,  # Tên File
        'C': 14,  # Tổng số câu
        'D': 16,  # Mã câu 1
        'E': 28,  # Chuyên đề
        'F': 50,  # Nội dung câu 1
        'G': 28,  # Phương án A
        'H': 28,  # Phương án B
        'I': 28,  # Phương án C
        'J': 28,  # Phương án D
        'K': 14,  # Đáp án đúng
        'L': 35,  # Nội dung đáp án đúng
        'M': 40   # Ghi chú
    }
    for col_letter, width in col_widths2.items():
        ws2.column_dimensions[col_letter].width = width
        
    for col_num in range(1, len(headers2) + 1):
        cell = ws2.cell(row=1, column=col_num)
        cell.fill = styles['header_fill']
        cell.font = styles['header_font']
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles['header_border']
    ws2.row_dimensions[1].height = 28
    
    for r_idx in range(2, len(exam_records) + 2):
        ws2.row_dimensions[r_idx].height = 26
        for c_idx in range(1, len(headers2) + 1):
            cell = ws2.cell(row=r_idx, column=c_idx)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if c_idx in [1, 3, 4, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx == 11:
                cell.fill = styles['ans_fill']
                cell.font = styles['ans_font']
                
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = ws2.dimensions

    # -------------------------------------------------------------
    # Sheet 3: Thong_Ke_Chuyen_De
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="Thong_Ke_Chuyen_De")
    headers3 = [
        "STT", "Chuyên Đề Nghiệp Vụ", "Số Câu Độc Bản", "Tỷ Lệ Độc Bản (%)",
        "Tổng Lượt Xuất Hiện (40 Đề)", "Tỷ Lệ Toàn Khối Lượng (%)", "Tần Suất TB / Đề"
    ]
    ws3.append(headers3)
    
    topic_unique = Counter()
    topic_total = Counter()
    for q in all_questions.values():
        topic_unique[q['topic']] += 1
        topic_total[q['topic']] += q['count']
        
    total_u = len(all_questions)
    total_all = sum(topic_total.values())
    
    sorted_topics = topic_unique.most_common()
    for idx, (t, u_cnt) in enumerate(sorted_topics, 1):
        tot_cnt = topic_total[t]
        u_pct = u_cnt / total_u
        tot_pct = tot_cnt / total_all
        avg_per_exam = tot_cnt / len(exam_records)
        ws3.append([
            idx, t, u_cnt, u_pct, tot_cnt, tot_pct, round(avg_per_exam, 1)
        ])
    
    # Dòng tổng cộng
    total_row = [
        "Tổng", "Toàn bộ chuyên đề", total_u, 1.0, total_all, 1.0, round(total_all / len(exam_records), 1)
    ]
    ws3.append(total_row)
    
    col_widths3 = {'A': 8, 'B': 45, 'C': 18, 'D': 20, 'E': 25, 'F': 25, 'G': 18}
    for col_letter, width in col_widths3.items():
        ws3.column_dimensions[col_letter].width = width
        
    for col_num in range(1, len(headers3) + 1):
        cell = ws3.cell(row=1, column=col_num)
        cell.fill = styles['header_fill']
        cell.font = styles['header_font']
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles['header_border']
    ws3.row_dimensions[1].height = 28
    
    for r_idx in range(2, len(sorted_topics) + 3):
        is_last = (r_idx == len(sorted_topics) + 2)
        ws3.row_dimensions[r_idx].height = 24
        for c_idx in range(1, len(headers3) + 1):
            cell = ws3.cell(row=r_idx, column=c_idx)
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
                
    ws3.freeze_panes = "A2"

    # -------------------------------------------------------------
    # Sheet 4: Top_Cau_Hoi_Trong_Tam (Lặp lại >= 4 lần)
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="Top_Cau_Hoi_Trong_Tam")
    headers4 = [
        "STT", "Mã Câu Hệ Thống", "Chuyên Đề Nghiệp Vụ", "Số Lần Xuất Hiện", "Tỷ Lệ Xuất Hiện (40 Đề)",
        "Nội Dung Câu Hỏi", "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Các Đề Xuất Hiện"
    ]
    ws4.append(headers4)
    
    top_questions = [q for q in all_questions.values() if q['count'] >= 4]
    top_questions.sort(key=lambda q: (-q['count'], q['topic'], q['id']))
    
    for idx, q in enumerate(top_questions, 1):
        exams_str = ", ".join([f"Đề {e:02d}" for e in sorted(q['exams'])])
        pct_exam = q['count'] / len(exam_records)
        ws4.append([
            idx, q['id'], q['topic'], q['count'], pct_exam,
            q['content'], q['correct_ans'], q['correct_text'], exams_str
        ])
        
    col_widths4 = {
        'A': 8, 'B': 18, 'C': 30, 'D': 18, 'E': 22,
        'F': 55, 'G': 14, 'H': 35, 'I': 35
    }
    for col_letter, width in col_widths4.items():
        ws4.column_dimensions[col_letter].width = width
        
    for col_num in range(1, len(headers4) + 1):
        cell = ws4.cell(row=1, column=col_num)
        cell.fill = styles['header_fill']
        cell.font = styles['header_font']
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles['header_border']
    ws4.row_dimensions[1].height = 28
    
    for r_idx in range(2, len(top_questions) + 2):
        ws4.row_dimensions[r_idx].height = 24
        for c_idx in range(1, len(headers4) + 1):
            cell = ws4.cell(row=r_idx, column=c_idx)
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
                cell.fill = styles['ans_fill']
                cell.font = styles['ans_font']
                
    ws4.freeze_panes = "A2"
    ws4.auto_filter.ref = ws4.dimensions

    wb.save(OUTPUT_EXCEL_ALL)
    print(f"Đã lưu thành công: {OUTPUT_EXCEL_ALL}")

def build_excel_de_va_cau_1(exam_records):
    styles = create_styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # -------------------------------------------------------------
    # Sheet 1: Danh_Sach_40_De_Va_Cau_1
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="Danh_Sach_40_De_Va_Cau_1")
    headers1 = [
        "STT Đề", "Tên File Excel", "Số Câu", "Mã Câu 1 (ID)",
        "Chuyên Đề Nghiệp Vụ", "Nội Dung Câu Hỏi Số 1",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Căn Cứ Pháp Lý / Ghi Chú"
    ]
    ws1.append(headers1)
    
    for rec in exam_records:
        row_data = [
            f"Đề {rec['exam_num']:02d}",
            rec['filename'],
            rec['total_q'],
            rec['q1_id'],
            rec['q1_topic'],
            rec['q1_content'],
            rec['q1_a'],
            rec['q1_b'],
            rec['q1_c'],
            rec['q1_d'],
            rec['q1_ans'],
            rec['q1_ans_text'],
            rec['q1_note']
        ]
        ws1.append(row_data)
        
    col_widths1 = {
        'A': 10, 'B': 45, 'C': 12, 'D': 16, 'E': 28, 'F': 50,
        'G': 28, 'H': 28, 'I': 28, 'J': 28, 'K': 14, 'L': 35, 'M': 40
    }
    for col_letter, width in col_widths1.items():
        ws1.column_dimensions[col_letter].width = width
        
    for col_num in range(1, len(headers1) + 1):
        cell = ws1.cell(row=1, column=col_num)
        cell.fill = styles['header_fill']
        cell.font = styles['header_font']
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles['header_border']
    ws1.row_dimensions[1].height = 28
    
    for r_idx in range(2, len(exam_records) + 2):
        ws1.row_dimensions[r_idx].height = 26
        for c_idx in range(1, len(headers1) + 1):
            cell = ws1.cell(row=r_idx, column=c_idx)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if c_idx in [1, 3, 4, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx == 11:
                cell.fill = styles['ans_fill']
                cell.font = styles['ans_font']
                
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = ws1.dimensions

    # -------------------------------------------------------------
    # Sheet 2: Phan_Tich_Cau_1
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Phan_Tich_Cau_1")
    headers2 = [
        "STT", "Mã Câu 1 (ID)", "Nội Dung Câu Hỏi Số 1", "Chuyên Đề",
        "Đáp Án Đúng", "Số Đề Có Chung Câu 1 Này", "Danh Sách Các Đề"
    ]
    ws2.append(headers2)
    
    # Nhóm theo q1_id
    q1_groups = defaultdict(list)
    for rec in exam_records:
        q1_groups[rec['q1_id']].append(rec)
        
    sorted_q1_groups = sorted(q1_groups.items(), key=lambda x: (-len(x[1]), x[0]))
    for idx, (qid, reclist) in enumerate(sorted_q1_groups, 1):
        r0 = reclist[0]
        exams_str = ", ".join([f"Đề {r['exam_num']:02d}" for r in reclist])
        ws2.append([
            idx, qid, r0['q1_content'], r0['q1_topic'],
            r0['q1_ans'], len(reclist), exams_str
        ])
        
    col_widths2 = {'A': 8, 'B': 18, 'C': 55, 'D': 30, 'E': 14, 'F': 25, 'G': 35}
    for col_letter, width in col_widths2.items():
        ws2.column_dimensions[col_letter].width = width
        
    for col_num in range(1, len(headers2) + 1):
        cell = ws2.cell(row=1, column=col_num)
        cell.fill = styles['header_fill']
        cell.font = styles['header_font']
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles['header_border']
    ws2.row_dimensions[1].height = 28
    
    for r_idx in range(2, len(sorted_q1_groups) + 2):
        ws2.row_dimensions[r_idx].height = 26
        for c_idx in range(1, len(headers2) + 1):
            cell = ws2.cell(row=r_idx, column=c_idx)
            cell.font = styles['normal_font']
            cell.border = styles['thin_border']
            if c_idx in [1, 2, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx == 5:
                cell.fill = styles['ans_fill']
                cell.font = styles['ans_font']
                
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = ws2.dimensions

    wb.save(OUTPUT_EXCEL_DE)
    print(f"Đã lưu thành công: {OUTPUT_EXCEL_DE}")

def build_markdown_sources(exam_records, all_questions):
    os.makedirs(SOURCES_MD_DIR, exist_ok=True)
    
    # 1. Markdown 40 Đề và Câu 1
    md_de_path = os.path.join(SOURCES_MD_DIR, 'danh_sach_40_de_va_cau_1_p3_39917.md')
    with open(md_de_path, 'w', encoding='utf-8') as f:
        f.write("---\n")
        f.write("type: source\n")
        f.write("source_kind: own-note\n")
        f.write("status: processed\n")
        f.write("created: 2026-09-25\n")
        f.write("processed_at: 2026-09-25\n")
        f.write("tags: [on-tap-thue, p3-39917, muc-luc-40-de]\n")
        f.write("title: Danh sách tổng hợp 40 đề thi và Câu 1 kíp P3-39917\n")
        f.write("---\n\n")
        f.write("# Danh sách tổng hợp 40 bộ đề thi và Câu hỏi số 1 kíp P3-39917\n\n")
        f.write("Dữ liệu trích xuất từ 40 file đề thi chuẩn hóa tại `D:\\HOC TAP\\ON THI 2026.09\\40 DE THI\\`.\n\n")
        f.write("| STT Đề | Mã Câu 1 (ID) | Chuyên Đề | Nội Dung Câu 1 | Đáp Án Đúng | Tên File Excel |\n")
        f.write("| :--- | :--- | :--- | :--- | :---: | :--- |\n")
        for rec in exam_records:
            clean_q = rec['q1_content'].replace('|', '/').replace('\n', ' ')
            f.write(f"| **Đề {rec['exam_num']:02d}** | `{rec['q1_id']}` | {rec['q1_topic']} | {clean_q[:65]}... | **{rec['q1_ans']}** | `{rec['filename']}` |\n")
            
    print(f"Đã tạo: {md_de_path}")
    
    # 2. Markdown Tổng hợp 1.429 câu độc bản tóm tắt
    md_1429_path = os.path.join(SOURCES_MD_DIR, 'tong_hop_1429_cau_hoi_doc_ban_p3_39917.md')
    with open(md_1429_path, 'w', encoding='utf-8') as f:
        f.write("---\n")
        f.write("type: source\n")
        f.write("source_kind: own-note\n")
        f.write("status: processed\n")
        f.write("created: 2026-09-25\n")
        f.write("processed_at: 2026-09-25\n")
        f.write("tags: [on-tap-thue, p3-39917, ngan-hang-1429-cau]\n")
        f.write("title: Tổng hợp ngân hàng 1.429 câu hỏi độc bản kíp P3-39917 ôn tập nhanh\n")
        f.write("---\n\n")
        f.write("# Ngân hàng 1.429 câu hỏi độc bản kíp P3-39917\n\n")
        f.write("- **Tổng số bộ đề khảo sát**: 40 đề thi (3.600 câu lượt).\n")
        f.write("- **Số lượng câu hỏi độc bản**: **1.429 câu** (đạt 100% đáp án chuẩn).\n")
        f.write("- **File Excel ngân hàng chi tiết**: `D:\\HOC TAP\\ON THI 2026.09\\40 DE THI\\Tong_Hop_1429_Cau_Hoi_Doc_Ban_On_Tap_Nhanh.xlsx`.\n\n")
        f.write("## Thống kê chuyên đề nghiệp vụ\n\n")
        
        topic_unique = Counter([q['topic'] for q in all_questions.values()])
        f.write("| STT | Chuyên Đề Nghiệp Vụ | Số Câu Độc Bản | Tỷ Lệ (%) |\n")
        f.write("| :---: | :--- | :---: | :---: |\n")
        for idx, (t, c) in enumerate(topic_unique.most_common(), 1):
            f.write(f"| {idx} | {t} | **{c}** | {c/len(all_questions)*100:.1f}% |\n")
            
    print(f"Đã tạo: {md_1429_path}")

if __name__ == '__main__':
    exam_records, all_questions = process_data()
    build_excel_full(exam_records, all_questions)
    build_excel_de_va_cau_1(exam_records)
    build_markdown_sources(exam_records, all_questions)
    print("HOÀN THÀNH TẤT CẢ CÁC BƯỚC XUẤT DỮ LIỆU!")
