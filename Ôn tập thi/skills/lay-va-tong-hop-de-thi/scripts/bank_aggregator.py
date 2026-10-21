import sys
import os
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

DEFAULT_MASTER_EXCEL = r"D:\HOC TAP\ON THI 2026.09\Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.xlsx"
DEFAULT_BRAIN_MD = r"D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\sources\Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.md"

def normalize_text(t):
    if not t:
        return ""
    t = re.sub(r'<[^>]+>', ' ', str(t))
    t = t.lower()
    t = re.sub(r'[\r\n\t]+', ' ', t)
    t = re.sub(r'[^\w\s]', '', t)
    return re.sub(r'\s+', ' ', t).strip()

def classify_topic(q_text):
    q_lower = q_text.lower()
    if '970' in q_lower or 'quy trình kiểm tra' in q_lower or '1161' in q_lower or 'kiểm tra tại trụ sở' in q_lower:
        return "Quy trình kiểm tra thuế & QĐ 970"
    elif '108' in q_lower or 'luật quản lý thuế' in q_lower or '125' in q_lower or 'xử phạt' in q_lower or 'vi phạm hành chính' in q_lower:
        return "Luật Quản lý thuế & Xử lý vi phạm hành chính"
    elif 'giá trị gia tăng' in q_lower or 'gtgt' in q_lower or 'khấu trừ' in q_lower or 'hoàn thuế' in q_lower:
        return "Thuế Giá trị gia tăng & Khấu trừ hoàn thuế"
    elif 'thương mại điện tử' in q_lower or 'tmdt' in q_lower or '117/2025' in q_lower or 'sàn' in q_lower:
        return "Thương mại điện tử & Nghị định 117/2025/NĐ-CP"
    elif 'hóa đơn' in q_lower or '123/2020' in q_lower or '254/2026' in q_lower or 'hđđt' in q_lower:
        return "Hóa đơn điện tử & Nghị định 123/2020/NĐ-CP"
    elif 'thu nhập doanh nghiệp' in q_lower or 'tndn' in q_lower or 'chi phí được trừ' in q_lower:
        return "Thuế Thu nhập doanh nghiệp & Chính sách thuế"
    elif 'tpr' in q_lower or 'quản lý rủi ro' in q_lower or 'bộ tiêu chí' in q_lower or 'hệ số k' in q_lower:
        return "Quản lý rủi ro & Ứng dụng TPR"
    elif 'doanh nghiệp lớn' in q_lower or 'giao dịch liên kết' in q_lower or '132/2020' in q_lower:
        return "Quản lý thuế Doanh nghiệp lớn & Giao dịch liên kết"
    elif 'tài khoản' in q_lower or 'kế toán' in q_lower or 'bctc' in q_lower or 'báo cáo tài chính' in q_lower or 'sổ sách' in q_lower:
        return "Kế toán doanh nghiệp & Kiểm tra Báo cáo tài chính"
    elif 'tham nhũng' in q_lower or 'khiếu nại' in q_lower or 'tố cáo' in q_lower:
        return "Phòng, chống tham nhũng, Luật Khiếu nại & Tố cáo"
    else:
        return "Nghiệp vụ kiểm tra thuế khác"

def aggregate_new_questions(source_excel, master_excel=DEFAULT_MASTER_EXCEL, source_label="Đề thi bổ sung"):
    print(f"=== BẮT ĐẦU TỔNG HỢP VÀ ĐỐI SOÁT VÀO MASTER BANK ===")
    print(f"Nguồn: {source_excel}")
    print(f"Đích Master: {master_excel}")
    
    if not os.path.exists(master_excel):
        print(f"Lỗi: Không tìm thấy file Master Bank tại {master_excel}")
        return 0, 0
        
    wb_master = openpyxl.load_workbook(master_excel)
    ws_master = wb_master['Ngan_Hang_Cau_Hoi_Chuan'] if 'Ngan_Hang_Cau_Hoi_Chuan' in wb_master.sheetnames else wb_master.active
    
    # 1. Index existing questions
    existing_map = {}
    row_count = 0
    for row_idx in range(3, ws_master.max_row + 1):
        q_val = ws_master.cell(row=row_idx, column=3).value
        if q_val:
            norm_q = normalize_text(q_val)
            if norm_q:
                existing_map[norm_q] = row_idx
                row_count += 1
    print(f"-> Master Bank hiện có: {row_count} câu hỏi đã lập chỉ mục.")
    
    # 2. Read source excel
    wb_src = openpyxl.load_workbook(source_excel, data_only=True)
    ws_src = wb_src.active
    
    # Determine column indexes in source
    # Expecting: STT, Mã Câu/Chuyên đề, ..., Nội Dung, A, B, C, D, ..., Đáp Án
    src_rows = list(ws_src.iter_rows(values_only=True))
    if not src_rows:
        print("Lỗi: File nguồn không có dữ liệu.")
        return 0, 0
        
    header = [str(c).lower() if c else "" for c in src_rows[0]]
    q_col = -1
    a_col, b_col, c_col, d_col = -1, -1, -1, -1
    ans_col = -1
    exp_col = -1
    
    for idx, name in enumerate(header):
        if 'nội dung' in name or 'câu hỏi' in name:
            q_col = idx
        elif name in ['phương án a', 'a']:
            a_col = idx
        elif name in ['phương án b', 'b']:
            b_col = idx
        elif name in ['phương án c', 'c']:
            c_col = idx
        elif name in ['phương án d', 'd']:
            d_col = idx
        elif 'đáp án chuẩn' in name or 'đáp án đúng' in name or name == 'đáp án':
            ans_col = idx
        elif 'căn cứ' in name or 'giải thích' in name or 'ghi chú' in name:
            exp_col = idx
            
    if q_col == -1 or ans_col == -1:
        print("Lỗi: Không tự nhận diện được cột Nội Dung hoặc cột Đáp Án trong file nguồn.")
        return 0, 0
        
    added_count = 0
    duplicate_count = 0
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(name="Calibri", size=11, bold=True, color="006100")
    
    for row in src_rows[1:]:
        if not row or len(row) <= max(q_col, ans_col):
            continue
        q_text = str(row[q_col]).strip() if row[q_col] else ""
        if not q_text or len(q_text) < 5:
            continue
            
        norm_q = normalize_text(q_text)
        if norm_q in existing_map:
            duplicate_count += 1
            continue
            
        ans_val = str(row[ans_col]).strip().upper() if row[ans_col] else ""
        if not ans_val or ans_val not in ['A', 'B', 'C', 'D']:
            continue
            
        choice_a = str(row[a_col]).strip() if a_col != -1 and len(row) > a_col and row[a_col] else ""
        choice_b = str(row[b_col]).strip() if b_col != -1 and len(row) > b_col and row[b_col] else ""
        choice_c = str(row[c_col]).strip() if c_col != -1 and len(row) > c_col and row[c_col] else ""
        choice_d = str(row[d_col]).strip() if d_col != -1 and len(row) > d_col and row[d_col] else ""
        exp_text = str(row[exp_col]).strip() if exp_col != -1 and len(row) > exp_col and row[exp_col] else ""
        
        topic = classify_topic(q_text)
        new_row_idx = ws_master.max_row + 1
        stt = new_row_idx - 2
        
        row_data = [
            stt, topic, q_text, choice_a, choice_b, choice_c, choice_d,
            ans_val, exp_text, source_label
        ]
        ws_master.append(row_data)
        
        # Style row
        cell_ans = ws_master.cell(row=new_row_idx, column=8)
        cell_ans.fill = green_fill
        cell_ans.font = green_font
        cell_ans.alignment = Alignment(horizontal="center", vertical="center")
        
        existing_map[norm_q] = new_row_idx
        added_count += 1

    # 3. Recalculate Statistics sheet
    if 'Thong_Ke_Chuyen_De' in wb_master.sheetnames:
        ws_stat = wb_master['Thong_Ke_Chuyen_De']
        # Recalculate topic counts
        topic_counts = {}
        for r_idx in range(3, ws_master.max_row + 1):
            t_val = ws_master.cell(row=r_idx, column=2).value
            if t_val:
                t_str = str(t_val).strip()
                topic_counts[t_str] = topic_counts.get(t_str, 0) + 1
                
        total_questions = sum(topic_counts.values())
        for s_idx in range(3, ws_stat.max_row):
            cat_name = ws_stat.cell(row=s_idx, column=1).value
            if cat_name:
                c_clean = str(cat_name).strip()
                cnt = topic_counts.get(c_clean, 0)
                ws_stat.cell(row=s_idx, column=2, value=cnt)
                ws_stat.cell(row=s_idx, column=3, value=round(cnt * 100.0 / total_questions, 1) if total_questions > 0 else 0)
        # Total row
        last_r = ws_stat.max_row
        ws_stat.cell(row=last_r, column=2, value=total_questions)
        ws_stat.cell(row=last_r, column=3, value=100.0)

    wb_master.save(master_excel)
    print(f"-> Hoàn tất tổng hợp!")
    print(f"   + Câu hỏi mới bổ sung: {added_count}")
    print(f"   + Câu hỏi trùng lặp được lọc bỏ: {duplicate_count}")
    print(f"   + Tổng quy mô Master Bank hiện tại: {len(existing_map)} câu.")
    return added_count, duplicate_count

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tổng hợp câu hỏi mới vào Master Bank")
    parser.add_argument("--source", required=True, help="Đường dẫn file Excel đề thi cần tổng hợp")
    parser.add_argument("--master", default=DEFAULT_MASTER_EXCEL, help="Đường dẫn file Master Bank")
    parser.add_argument("--label", default="Đề thi bổ sung", help="Nhãn nguồn gốc đề thi")
    args = parser.parse_args()
    
    aggregate_new_questions(args.source, master_excel=args.master, source_label=args.label)
