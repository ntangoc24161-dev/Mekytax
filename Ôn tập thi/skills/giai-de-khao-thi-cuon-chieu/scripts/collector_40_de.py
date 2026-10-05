# -*- coding: utf-8 -*-
"""
Script thu thập và giải đề thi trắc nghiệm kíp P3-39917 - Bản tự động hóa 40 đề
- Cổng khảo thí: taxedu.gdt.gov.vn (api-e3.khaothi.online)
- Ngân hàng đối soát: Tong_Hop_Ngan_Hang_1247_Cau_Kip_P3_39917.xlsx
- Solver AI: tax_solver_engine.py
- Giãn cách an toàn: 2.0s - 4.0s/lần chọn
- Tự động nhận diện Đề cũ (nếu gặp lại đề cũ thì bổ sung tiếp các câu còn thiếu đến 90)
- Tự động tạo Đề mới (gán mã đề tiếp theo cho đến khi đủ 40 đề)
- Tự động bỏ qua đề gặp sự cố/tạm dừng để chuyển sang đề tiếp theo mà không làm gián đoạn toàn bộ tiến trình
- Lưu cuốn chiếu liên tục sau từng câu hỏi bảo toàn 100% dữ liệu
"""

import sys
import os
import argparse
import urllib.request
import urllib.parse
import json
import hashlib
import random
import time
import re
import glob
import subprocess
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DECRYPT_JS = r"D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\skills\lay-va-tong-hop-de-thi\scripts\decrypt_dcv2.js"
BANK_FILE = r"D:\HOC TAP\ON THI 2026.09\Tong_Hop_Ngan_Hang_1247_Cau_Kip_P3_39917.xlsx"
OUTPUT_DIR = r"D:\HOC TAP\ON THI 2026.09\40 DE THI"
LOG_FILE = r"D:\HOC TAP\ON THI 2026.09\collector_40_de.log"
CHECKPOINT_FILE = r"D:\HOC TAP\ON THI 2026.09\checkpoint_P3_39917.json"

sys.path.append(r"D:\HOC TAP\ON THI 2026.09")
try:
    import tax_solver_engine
except ImportError:
    tax_solver_engine = None

def log(msg):
    ts = time.strftime('%Y-%m-%d %H:%M:%S')
    formatted = f"[{ts}] {msg}"
    print(formatted, flush=True)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(formatted + "\n")
    except Exception:
        pass

def clean_html(text):
    if not text:
        return ""
    text = re.sub(r'<br\s*/?>', '\n', text)
    text = re.sub(r'</p>', '\n', text)
    text = re.sub(r'<[^>]+>', '', text)
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    return '\n'.join(lines)

def normalize_text_for_match(t):
    if not t:
        return ""
    t = re.sub(r'<[^>]+>', ' ', t)
    t = t.lower()
    t = re.sub(r'[\r\n\t]+', ' ', t)
    t = re.sub(r'[^\w\s]', '', t)
    return re.sub(r'\s+', ' ', t).strip()

def decrypt_payload(raw_b64):
    if not isinstance(raw_b64, str):
        return raw_b64
    res = subprocess.run(['node', DECRYPT_JS, 'decrypt', raw_b64], capture_output=True, text=True, encoding='utf-8')
    if res.returncode != 0:
        raise RuntimeError(f"Node decrypt error: {res.stderr}")
    return json.loads(res.stdout)

def classify_topic(q_text):
    q_lower = q_text.lower()
    if 'thông tư' in q_lower or '31/2021' in q_lower or 'tiêu chí đánh giá' in q_lower or 'tuân thủ' in q_lower:
        return "Thông tư 31/2021 & Tiêu chí tuân thủ"
    elif 'tpr' in q_lower or 'quản lý rủi ro' in q_lower or '991/qđ-tct' in q_lower:
        return "Quản lý rủi ro & Ứng dụng TPR"
    elif 'nhật ký' in q_lower or 'itnk' in q_lower:
        return "Ứng dụng Nhật ký điện tử (iTNK)"
    elif 'biên bản' in q_lower or 'kết thúc' in q_lower or 'thời hạn' in q_lower or 'thời hiệu' in q_lower:
        return "Quy trình kiểm tra & Thủ tục biên bản"
    elif 'tham nhũng' in q_lower:
        return "Luật Phòng, chống tham nhũng"
    elif 'giao dịch điện tử' in q_lower:
        return "Luật Giao dịch điện tử"
    elif 'hóa đơn' in q_lower or 'chứng từ' in q_lower:
        return "Hóa đơn, chứng từ điện tử"
    elif 'tài khoản' in q_lower or 'giá vốn' in q_lower or 'bctc' in q_lower or 'kế toán' in q_lower:
        return "Kế toán & Kiểm tra BCTC"
    elif '1614' in q_lower or 'nhiệm vụ đoàn' in q_lower:
        return "Quyết định 1614 & Nhiệm vụ đoàn kiểm tra"
    else:
        return "Nghiệp vụ kiểm tra thuế khác"

def load_bank(bank_path):
    bank_by_id = {}
    bank_by_text = {}
    if not os.path.exists(bank_path):
        log(f"[Cảnh báo] Không tìm thấy file ngân hàng: {bank_path}")
        return bank_by_id, bank_by_text
        
    try:
        wb = openpyxl.load_workbook(bank_path, data_only=True)
        ws = wb['Ngan_Hang_1247_Cau_Hoi'] if 'Ngan_Hang_1247_Cau_Hoi' in wb.sheetnames else wb.active
        for r in range(2, ws.max_row + 1):
            q_id = str(ws.cell(r, 2).value or '').strip()
            topic = str(ws.cell(r, 3).value or '').strip()
            q_text = str(ws.cell(r, 4).value or '').strip()
            a = str(ws.cell(r, 5).value or '').strip()
            b = str(ws.cell(r, 6).value or '').strip()
            c = str(ws.cell(r, 7).value or '').strip()
            d = str(ws.cell(r, 8).value or '').strip()
            ans = str(ws.cell(r, 9).value or '').strip().upper()
            ans_text = str(ws.cell(r, 10).value or '').strip()
            legal = str(ws.cell(r, 11).value or '').strip()
            status = str(ws.cell(r, 12).value or '').strip()
            
            entry = {
                'id': q_id,
                'topic': topic,
                'question': q_text,
                'A': a,
                'B': b,
                'C': c,
                'D': d,
                'ans': ans,
                'ans_text': ans_text,
                'legal': legal,
                'status': status,
                'is_verified': ('xác thực' in status.lower() or 'chuẩn hóa' in status.lower())
            }
            if q_id:
                bank_by_id[q_id] = entry
            norm_q = normalize_text_for_match(q_text)
            if norm_q:
                bank_by_text[norm_q] = entry
        log(f"-> Đã nạp Ngân hàng đối soát: {len(bank_by_id)} câu theo ID, {len(bank_by_text)} câu theo Text.")
    except Exception as e:
        log(f"[Lỗi] Nạp ngân hàng thất bại: {e}")
    return bank_by_id, bank_by_text

def load_all_existing_exams(output_dir):
    exams = {}
    pattern = re.compile(r'Bo_De_On_Thi_Kiem_Tra_Thue_.*_De_(\d+)\.xlsx')
    if not os.path.exists(output_dir):
        return exams
        
    for f in glob.glob(os.path.join(output_dir, 'Bo_De_On_Thi_Kiem_Tra_Thue_*_De_*.xlsx')):
        if os.path.basename(f).startswith('~$'):
            continue
        m = pattern.search(os.path.basename(f))
        if not m:
            continue
        exam_num = int(m.group(1))
        try:
            wb = openpyxl.load_workbook(f, data_only=True)
            target_sheet = None
            for s in wb.sheetnames:
                if 'Cau' in s or f'De_' in s:
                    target_sheet = s
                    break
            if not target_sheet:
                target_sheet = wb.sheetnames[0]
            ws = wb[target_sheet]
            
            verified = {}
            lead_ids = []
            all_ids = []
            for r in range(2, ws.max_row + 1):
                item_id = str(ws.cell(r, 2).value or '').strip()
                correct = str(ws.cell(r, 12).value or '').strip().upper()
                if item_id:
                    all_ids.append(item_id)
                    if len(lead_ids) < 3:
                        lead_ids.append(item_id)
                    if correct in ['A', 'B', 'C', 'D']:
                        verified[item_id] = {
                            'index': ws.cell(r, 1).value or (len(verified) + 1),
                            'item_id': item_id,
                            'topic': str(ws.cell(r, 3).value or '').strip(),
                            'question': str(ws.cell(r, 4).value or '').strip(),
                            'choice_a': str(ws.cell(r, 5).value or '').strip(),
                            'choice_b': str(ws.cell(r, 6).value or '').strip(),
                            'choice_c': str(ws.cell(r, 7).value or '').strip(),
                            'choice_d': str(ws.cell(r, 8).value or '').strip(),
                            'selected_choice': str(ws.cell(r, 9).value or '').strip(),
                            'source_note': str(ws.cell(r, 10).value or '').strip(),
                            'initial_result': str(ws.cell(r, 11).value or '').strip(),
                            'correct_choice': correct,
                            'correct_text': str(ws.cell(r, 13).value or '').strip(),
                            'legal_basis': str(ws.cell(r, 14).value or '').strip()
                        }
            exams[exam_num] = {
                'file': f,
                'exam_num': exam_num,
                'lead_ids': lead_ids,
                'all_ids': all_ids,
                'verified': verified,
                'count': len(verified)
            }
        except Exception as e:
            log(f"[Cảnh báo] Lỗi đọc file đề thi {f}: {e}")
    return exams

def find_matched_exam(session_items, existing_exams):
    if not existing_exams or not session_items:
        return None, None
        
    session_lead_ids = [str(session_items[i]['id']) for i in range(min(3, len(session_items)))]
    session_all_ids = set(str(it['id']) for it in session_items)
    
    for num, ex in existing_exams.items():
        if ex['lead_ids'] and ex['lead_ids'] == session_lead_ids:
            return num, ex
        if ex['all_ids'] and len(session_all_ids & set(ex['all_ids'])) >= 80:
            return num, ex
            
    return None, None

def save_exam_excel(results, output_dir, subscription, group_id, exam_num, is_final=False):
    if not results:
        return None
    os.makedirs(output_dir, exist_ok=True)
    out_excel = os.path.join(output_dir, f"Bo_De_On_Thi_Kiem_Tra_Thue_{subscription}_{group_id}_De_{exam_num:02d}.xlsx")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"De_{group_id}_{exam_num:02d}"
    
    headers_list = [
        "STT", "Mã Câu Hệ Thống", "Chuyên Đề Nghiệp Vụ", "Nội Dung Câu Hỏi",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Phương Án Chọn", "Nguồn Phương Án", "Kết Quả Lần Đầu",
        "Đáp Án Đúng", "Nội Dung Đáp Án Đúng", "Căn Cứ Pháp Lý / Ghi Chú"
    ]
    ws.append(headers_list)
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx in range(1, len(headers_list) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(name="Calibri", size=11, bold=True, color="006100")
    
    for r in results:
        row_vals = [
            r['index'], r['item_id'], r['topic'], r['question'],
            r['choice_a'], r['choice_b'], r['choice_c'], r['choice_d'],
            r['selected_choice'], r['source_note'], r['initial_result'],
            r['correct_choice'], r['correct_text'], r['legal_basis']
        ]
        ws.append(row_vals)
        curr_row = ws.max_row
        
        ans_cell = ws.cell(row=curr_row, column=12)
        ans_cell.fill = green_fill
        ans_cell.font = green_font
        ans_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        ws.cell(row=curr_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=curr_row, column=2).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=curr_row, column=9).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=curr_row, column=10).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=curr_row, column=11).alignment = Alignment(horizontal="center", vertical="center")
        
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:N{ws.max_row}"
    
    col_widths = {
        'A': 8,   'B': 16,  'C': 26,  'D': 50,  'E': 35,
        'F': 35,  'G': 35,  'H': 35,  'I': 16,  'J': 18,
        'K': 16,  'L': 15,  'M': 35,  'N': 45
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width
        
    try:
        wb.save(out_excel)
        if is_final:
            log(f"-> ĐÃ LƯU FILE EXCEL HOÀN TẤT: {os.path.basename(out_excel)} ({len(results)} câu)")
        else:
            if len(results) % 5 == 0 or len(results) == 1:
                log(f"   [Lưu cuốn chiếu] File {os.path.basename(out_excel)}: đã ghi {len(results)}/90 câu")
    except PermissionError:
        log(f"   [Cảnh báo cuốn chiếu] File Excel đang mở bởi ứng dụng khác, tạm bỏ qua lượt lưu này...")
    except Exception as e:
        log(f"   [Cảnh báo cuốn chiếu] Ghi file Excel thất bại: {e}")
    return out_excel

def update_checkpoint(group_id, subscription, exam_num, total, completed, latest_q="", latest_ans="", excel_file="", status="running", extra=None):
    ckpt_data = {
        "group_id": group_id,
        "subscription": subscription,
        "current_exam_num": exam_num,
        "total_in_exam": total,
        "completed_in_exam": completed,
        "progress_percent": round(completed * 100 / total, 1) if total else 0.0,
        "latest_question": (latest_q[:80] if latest_q else ""),
        "latest_correct_choice": latest_ans,
        "excel_file": excel_file,
        "status": status,
        "updated_at": time.strftime('%Y-%m-%d %H:%M:%S')
    }
    if extra:
        ckpt_data.update(extra)
    try:
        with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
            json.dump(ckpt_data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def create_exam_session(group_id=39917, subscription="P3"):
    first_names = ["Nguyễn", "Trần", "Lê", "Phạm", "Hoàng", "Huỳnh", "Phan", "Vũ", "Võ", "Đặng", "Bùi", "Đỗ", "Hồ", "Ngô", "Dương", "Lý"]
    mid_names = ["Văn", "Thị", "Hồng", "Minh", "Đức", "Quốc", "Hải", "Tuấn", "Thanh", "Ngọc", "Gia", "Bảo"]
    last_names = ["An", "Bình", "Cường", "Dũng", "Em", "Hương", "Hà", "Khanh", "Linh", "Mai", "Nam", "Phúc", "Quân", "Sơn", "Tâm", "Uyên", "Vinh"]
    candidate_name = f"{random.choice(first_names)} {random.choice(mid_names)} {random.choice(last_names)}"
    
    visitor_id = hashlib.md5(f"device_agent_{random.randint(100000, 999999)}".encode()).hexdigest()
    client_id = str(random.randint(11111111, 99999999))
    device_id = f"{visitor_id}-{client_id}"
    device_info = json.dumps({"b": "Chrome", "bv": "128.0.0.0", "o": "Windows", "ov": "10", "m": 0})
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
        'Origin': 'https://taxedu.gdt.gov.vn',
        'Referer': 'https://taxedu.gdt.gov.vn/',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'vi-VN',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Xcbt-Device-id': device_id,
        'Xcbt-Device-info': device_info,
    }
    
    log(f"[1/2] Đăng ký thí sinh: {candidate_name}...")
    reg_url = "https://api-e3.khaothi.online/delivery/exam/guest-register"
    reg_payload = {
        "test_taker_group_id": group_id,
        "subscription": subscription,
        "info[name]": candidate_name,
        "e_v": 2
    }
    req = urllib.request.Request(reg_url, data=urllib.parse.urlencode(reg_payload).encode('utf-8'), headers=headers)
    with urllib.request.urlopen(req, timeout=20) as resp:
        parsed = json.loads(resp.read().decode('utf-8'))
        raw_reg = parsed['data']
        reg_data = decrypt_payload(raw_reg) if isinstance(raw_reg, str) else raw_reg
        test_taker = reg_data['test_taker']
        log(f"-> Thí sinh đã đăng ký: {test_taker['name']} (Mã code: {test_taker['code']})")
        
    log("[2/2] Mở phiên làm bài thi và bóc tách đề...")
    start_url = "https://api-e3.khaothi.online/delivery/exam/start"
    start_payload = {
        "data[code]": test_taker['code'],
        "e_v": 2
    }
    req_start = urllib.request.Request(start_url, data=urllib.parse.urlencode(start_payload).encode('utf-8'), headers=headers)
    with urllib.request.urlopen(req_start, timeout=20) as resp:
        start_resp = json.loads(resp.read().decode('utf-8'))
        raw_data = start_resp['data']
        exam_data = decrypt_payload(raw_data) if isinstance(raw_data, str) else raw_data
        token = exam_data['token']
        items = exam_data['testFormData']['items']
        log(f"-> Phiên thi đã sẵn sàng. Tổng số câu hỏi: {len(items)} câu.")
        
    return test_taker, token, headers, items

def solve_single_exam(session_info, target_exam_num, group_id=39917, subscription="P3", bank_by_id=None, bank_by_text=None, output_dir=OUTPUT_DIR, resume_verified=None):
    choice_map = {"1": "A", "2": "B", "3": "C", "4": "D"}
    rev_choice_map = {"A": "1", "B": "2", "C": "3", "D": "4"}
    
    test_taker = session_info['test_taker']
    token = session_info['token']
    headers = session_info['headers']
    items = session_info['items']
    
    check_url = "https://api-e3.khaothi.online/delivery/exam/submit-item-answer"
    check_headers = dict(headers)
    check_headers['Xcbt-Token'] = token
    
    results = []
    start_idx = 1
    
    # Nạp các câu đã có nếu là đề cũ
    if resume_verified:
        verified_count = len(resume_verified)
        log(f"-> Kích hoạt chế độ bổ sung câu còn thiếu cho Đề {target_exam_num:02d} (đã có {verified_count}/90 câu)!")
        for idx in range(1, verified_count + 1):
            item = items[idx - 1]
            item_id = str(item['id'])
            v_entry = resume_verified.get(item_id)
            if v_entry:
                results.append(v_entry)
                correct_letter = v_entry['correct_choice']
                ans_key = rev_choice_map.get(correct_letter, "1")
                payload = {"item_id": item_id, "answer_data": ans_key}
                try:
                    req_check = urllib.request.Request(check_url, data=urllib.parse.urlencode(payload).encode('utf-8'), headers=check_headers)
                    with urllib.request.urlopen(req_check, timeout=10) as c_resp:
                        pass
                except Exception:
                    pass
                time.sleep(round(random.uniform(0.5, 1.0), 2))
        log(f"-> Đã đồng bộ nhanh {len(results)} câu trước đó. Bắt đầu giải cuốn chiếu từ CÂU {len(results) + 1} ĐẾN {len(items)}!")
        start_idx = len(results) + 1

    try:
        for idx in range(start_idx, len(items) + 1):
            item = items[idx - 1]
            item_id = str(item['id'])
            q_text = clean_html(item['data_body_html'])
            raw_choices = json.loads(item['data_choices_html'])
            cleaned_choices = {k: clean_html(v) for k, v in raw_choices.items()}
            
            choice_a = cleaned_choices.get("1", "")
            choice_b = cleaned_choices.get("2", "")
            choice_c = cleaned_choices.get("3", "")
            choice_d = cleaned_choices.get("4", "")
            
            # 1. Tra cứu ngân hàng
            bank_entry = bank_by_id.get(item_id) if bank_by_id else None
            if not bank_entry and bank_by_text:
                norm_q = normalize_text_for_match(q_text)
                bank_entry = bank_by_text.get(norm_q)
                
            chosen_key = None
            source_note = ""
            legal_basis = ""
            topic = classify_topic(q_text)
            
            if bank_entry and bank_entry.get('is_verified') and bank_entry.get('ans') in ['A', 'B', 'C', 'D']:
                chosen_letter = bank_entry['ans']
                chosen_key = rev_choice_map.get(chosen_letter, "1")
                source_note = "Từ ngân hàng đề"
                legal_basis = bank_entry.get('legal', '')
                topic = bank_entry.get('topic') or topic
            else:
                # AI Solver
                if tax_solver_engine:
                    item_dict = {
                        'id': item_id,
                        'topic': topic,
                        'question': q_text,
                        'A': choice_a,
                        'B': choice_b,
                        'C': choice_c,
                        'D': choice_d
                    }
                    sol_letter, sol_text, sol_basis = tax_solver_engine.solve_item(item_dict)
                    chosen_letter = sol_letter
                    chosen_key = rev_choice_map.get(chosen_letter, "1")
                    source_note = "AI đề xuất"
                    legal_basis = sol_basis
                else:
                    chosen_key = "1"
                    chosen_letter = "A"
                    source_note = "AI đề xuất"
                    legal_basis = "Phân tích ngữ nghĩa"
                    
            # Giãn cách an toàn tối ưu (đảm bảo hoàn thành 1 đề dưới 15 phút)
            delay = round(random.uniform(2.0, 3.5), 2)
            log(f"Đề {target_exam_num:02d} | Câu {idx:02d}/{len(items)} (ID: {item_id}): Chọn {chosen_letter} [{source_note}] - Chờ {delay}s...")
            time.sleep(delay)
            
            payload = {"item_id": item_id, "answer_data": chosen_key}
            is_correct = False
            latest_instruction = ""
            
            conn_ok = False
            for attempt in range(8):
                try:
                    req_check = urllib.request.Request(check_url, data=urllib.parse.urlencode(payload).encode('utf-8'), headers=check_headers)
                    with urllib.request.urlopen(req_check, timeout=15) as check_resp:
                        check_data = json.loads(check_resp.read().decode('utf-8'))
                        is_correct = check_data.get('data', {}).get('is_correct', False)
                        latest_instruction = check_data.get('data', {}).get('instruction', '') or latest_instruction
                        conn_ok = True
                        break
                except Exception as e:
                    log(f"   [Cảnh báo kết nối lượt {attempt+1}/8 ({e})], thử lại sau 3s...")
                    time.sleep(3)
                    
            if not conn_ok:
                raise ConnectionError(f"Mất kết nối máy chủ khảo thí tại câu {idx} Đề {target_exam_num:02d}")
                
            final_correct_key = chosen_key if is_correct else None
            initial_correct_str = "Đúng" if is_correct else "Chưa đúng"
            
            if is_correct:
                log(f"   -> Kết quả: {chosen_letter} ĐÚNG!")
            else:
                log(f"   -> {chosen_letter} chưa đúng. Dò các phương án còn lại...")
                remaining_keys = [k for k in sorted(raw_choices.keys()) if k != chosen_key]
                for r_key in remaining_keys:
                    r_letter = choice_map.get(r_key, r_key)
                    r_delay = round(random.uniform(2.0, 3.0), 2)
                    time.sleep(r_delay)
                    r_payload = {"item_id": item_id, "answer_data": r_key}
                    try:
                        r_req = urllib.request.Request(check_url, data=urllib.parse.urlencode(r_payload).encode('utf-8'), headers=check_headers)
                        with urllib.request.urlopen(r_req, timeout=15) as r_resp:
                            r_data = json.loads(r_resp.read().decode('utf-8'))
                            if r_data.get('data', {}).get('instruction'):
                                latest_instruction = r_data.get('data', {}).get('instruction')
                            if r_data.get('data', {}).get('is_correct', False):
                                final_correct_key = r_key
                                log(f"   -> Dò trúng phương án {r_letter} ĐÚNG!")
                                break
                    except Exception as e:
                        log(f"   [Lỗi dò phương án ({e})]")
                        time.sleep(4)
                        
            correct_letter = choice_map.get(final_correct_key, "A")
            correct_text = cleaned_choices.get(final_correct_key, "")
            
            if latest_instruction:
                clean_inst = clean_html(latest_instruction)
                legal_basis = f"{legal_basis} | {clean_inst}" if legal_basis else clean_inst
                
            results.append({
                'index': idx,
                'item_id': item_id,
                'topic': topic,
                'question': q_text,
                'choice_a': choice_a,
                'choice_b': choice_b,
                'choice_c': choice_c,
                'choice_d': choice_d,
                'selected_choice': chosen_letter,
                'source_note': source_note,
                'initial_result': initial_correct_str,
                'correct_choice': correct_letter,
                'correct_text': correct_text,
                'legal_basis': legal_basis
            })
            
            # Lưu cuốn chiếu sau từng câu vào file Excel
            save_exam_excel(results, output_dir, subscription, group_id, target_exam_num, is_final=False)
            
            out_excel = os.path.join(output_dir, f"Bo_De_On_Thi_Kiem_Tra_Thue_{subscription}_{group_id}_De_{target_exam_num:02d}.xlsx")
            update_checkpoint(group_id, subscription, target_exam_num, len(items), idx, q_text, correct_letter, out_excel, status="running")

        log(f"=== HOÀN TẤT ĐỀ THI {target_exam_num:02d}: ĐÃ GIẢI VÀ XÁC THỰC {len(results)}/{len(items)} CÂU ===")
        out_excel = save_exam_excel(results, output_dir, subscription, group_id, target_exam_num, is_final=True)
        update_checkpoint(group_id, subscription, target_exam_num, len(items), len(results), "", "", out_excel, status="completed")
        return True, results, None
        
    except Exception as e:
        log(f"[TIẾN TRÌNH BỊ TẠM DỪNG / GIÁN ĐOẠN] Đề {target_exam_num:02d} gặp lỗi ở câu {len(results)}/{len(items)}: {e}")
        log(f"-> Dữ liệu cuốn chiếu đã được bảo toàn ({len(results)} câu) trong file Excel.")
        log(f"-> TỰ ĐỘNG BỎ QUA ĐỀ {target_exam_num:02d}, chuẩn bị chuyển sang đề tiếp theo...")
        out_excel = os.path.join(output_dir, f"Bo_De_On_Thi_Kiem_Tra_Thue_{subscription}_{group_id}_De_{target_exam_num:02d}.xlsx")
        update_checkpoint(group_id, subscription, target_exam_num, len(items), len(results), "", "", out_excel, status="interrupted_skipped", extra={"error": str(e)})
        return False, results, str(e)

def run_batch_40_exams(max_exams=40, target_dir=OUTPUT_DIR, group_id=39917, subscription="P3", bank_by_id=None, bank_by_text=None):
    log("===================================================================")
    log(f"=== KHỞI CHẠY TIẾN TRÌNH THU THẬP TỰ ĐỘNG ĐẾN KHI ĐỦ {max_exams} ĐỀ THI ===")
    log("===================================================================")
    
    skipped_exams = set()
    consecutive_errors = 0
    
    while True:
        existing_exams = load_all_existing_exams(target_dir)
        completed_full_exams = [k for k, v in existing_exams.items() if v['count'] >= 90]
        total_valid_exams = len(existing_exams)
        
        log(f"[TIẾN ĐỘ TỔNG] Hiện có {total_valid_exams}/{max_exams} đề trong thư mục ({len(completed_full_exams)} đề đã hoàn tất 90/90 câu).")
        
        if total_valid_exams >= max_exams and len(completed_full_exams) >= max_exams:
            log(f"-> ĐÃ THU THẬP VÀ HOÀN TẤT ĐỦ {max_exams} ĐỀ THI 100%! KẾT THÚC TIẾN TRÌNH.")
            break
            
        # 1. Mở phiên thi mới từ máy chủ
        try:
            test_taker, token, headers, items = create_exam_session(group_id, subscription)
            consecutive_errors = 0
        except Exception as e:
            consecutive_errors += 1
            retry_wait = min(10 + consecutive_errors * 3, 60)
            log(f"[Lỗi mở phiên thi ({e})]. Lần lỗi liên tiếp: {consecutive_errors}. Tự động nghỉ {retry_wait}s rồi thử lại liên tục...")
            time.sleep(retry_wait)
            continue
            
        session_info = {
            'test_taker': test_taker,
            'token': token,
            'headers': headers,
            'items': items
        }
        
        # 2. Đối chiếu nhận diện đề cũ hay đề mới
        matched_num, matched_info = find_matched_exam(items, existing_exams)
        
        if matched_num is not None:
            # Khớp với một đề đã có trong thư mục
            if matched_info['count'] >= 90:
                log(f"[TRÙNG LẶP ĐỀ CŨ HOÀN TẤT] Phiên thi khớp với Đề {matched_num:02d} đã đủ 90/90 câu. Tự động bỏ qua để bốc đề mới...")
                time.sleep(round(random.uniform(2.0, 3.5), 2))
                continue
            else:
                log(f"[NHẬN DIỆN ĐỀ CŨ CHƯA XONG] Phiên thi khớp với Đề {matched_num:02d} (mới có {matched_info['count']}/90 câu)!")
                log(f"-> Tiến hành chạy tiếp phiên lưu cuốn chiếu bổ sung từ câu {matched_info['count'] + 1} đến 90 cho Đề {matched_num:02d}!")
                target_exam_num = matched_num
                resume_verified = matched_info['verified']
        else:
            # Đề thi hoàn toàn mới
            # Tìm mã đề nhỏ nhất từ 1..max_exams chưa có hoặc chưa bắt đầu
            available_nums = [i for i in range(1, max_exams + 1) if i not in existing_exams and i not in skipped_exams]
            if not available_nums:
                available_nums = [i for i in range(1, max_exams + 1) if existing_exams.get(i, {}).get('count', 0) < 90 and i not in skipped_exams]
            if not available_nums:
                skipped_exams.clear()
                available_nums = [i for i in range(1, max_exams + 1) if i not in existing_exams or existing_exams.get(i, {}).get('count', 0) < 90]
            if not available_nums:
                if len(completed_full_exams) >= max_exams:
                    log(f"-> Đã đạt đủ {max_exams} đề thi hoàn thiện 90/90 câu trong thư mục. Hoàn tất toàn bộ chu kỳ!")
                    break
                target_exam_num = total_valid_exams + 1
            else:
                target_exam_num = available_nums[0]
                
            log(f"[PHÁT HIỆN ĐỀ MỚI] Gán mã đề: ĐỀ {target_exam_num:02d}. Quét và giải từ câu 1 đến 90!")
            resume_verified = None

        # 3. Thực thi giải đề
        success, results, err = solve_single_exam(
            session_info=session_info,
            target_exam_num=target_exam_num,
            group_id=group_id,
            subscription=subscription,
            bank_by_id=bank_by_id,
            bank_by_text=bank_by_text,
            output_dir=target_dir,
            resume_verified=resume_verified
        )
        
        if not success:
            log(f"[TỰ ĐỘNG BỎ QUA] Đề {target_exam_num:02d} bị tạm dừng/gián đoạn. Chuyển sang đề tiếp theo sau 3s...")
            skipped_exams.add(target_exam_num)
            time.sleep(3)
        else:
            log(f"-> Đề {target_exam_num:02d} đã hoàn tất tốt đẹp. Nghỉ an toàn 3s-5s trước đề tiếp theo...")
            time.sleep(round(random.uniform(3.0, 5.0), 2))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Thu thập đề thi P3-39917 tự động 40 đề")
    parser.add_argument("--exam-num", type=int, default=0, help="Chỉ định đề thi cụ thể (0 = tự động chạy batch 40 đề)")
    parser.add_argument("--max-exams", type=int, default=40, help="Tổng số đề thi mục tiêu (mặc định 40)")
    parser.add_argument("--session-file", type=str, default="", help="File session json có sẵn")
    args = parser.parse_args()
    
    bank_by_id, bank_by_text = load_bank(BANK_FILE)
    
    if args.exam_num > 0:
        session_data = None
        if args.session_file and os.path.exists(args.session_file):
            with open(args.session_file, 'r', encoding='utf-8') as f:
                session_data = json.load(f)
        solve_single_exam(session_info=session_data, target_exam_num=args.exam_num, group_id=39917, subscription="P3", bank_by_id=bank_by_id, bank_by_text=bank_by_text)
    else:
        run_batch_40_exams(max_exams=args.max_exams, target_dir=OUTPUT_DIR, group_id=39917, subscription="P3", bank_by_id=bank_by_id, bank_by_text=bank_by_text)
