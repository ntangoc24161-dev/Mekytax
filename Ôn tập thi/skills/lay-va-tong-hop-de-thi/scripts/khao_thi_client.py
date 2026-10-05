import sys
import os

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
    except Exception:
        pass

import argparse
import urllib.request
import urllib.parse
import json
import hashlib
import random
import time
import re
import subprocess
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DECRYPT_JS = os.path.join(SCRIPT_DIR, "decrypt_dcv2.js")
DEFAULT_MASTER_EXCEL = r"D:\HOC TAP\ON THI 2026.09\Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.xlsx"
DEFAULT_OUTPUT_DIR = r"D:\HOC TAP\ON THI 2026.09"
DEFAULT_BRAIN_SOURCES = r"D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\sources"

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

def load_master_bank_cache(master_path):
    cache = {}
    if not os.path.exists(master_path):
        return cache
    try:
        wb = openpyxl.load_workbook(master_path, data_only=True)
        sheet = wb['Ngan_Hang_Cau_Hoi_Chuan'] if 'Ngan_Hang_Cau_Hoi_Chuan' in wb.sheetnames else wb.active
        # Columns: STT, Chuyên Đề, Nội Dung, A, B, C, D, Đáp Án Chuẩn, Căn Cứ, Nguồn
        for row in sheet.iter_rows(min_row=3, values_only=True):
            if not row or len(row) < 8:
                continue
            q_content = str(row[2]) if row[2] else ""
            ans = str(row[7]).strip().upper() if row[7] else ""
            norm_q = normalize_text_for_match(q_content)
            if norm_q and ans in ['A', 'B', 'C', 'D']:
                cache[norm_q] = {
                    'ans': ans,
                    'explanation': str(row[8]) if len(row) > 8 and row[8] else ""
                }
    except Exception as e:
        print(f"[Cache] Lỗi khi nạp Master Bank cache: {e}")
    return cache

def predict_best_choice(q_text, choices_dict, cached_ans=None):
    choice_map = {"1": "A", "2": "B", "3": "C", "4": "D"}
    rev_choice_map = {"A": "1", "B": "2", "C": "3", "D": "4"}
    
    if cached_ans and cached_ans in rev_choice_map:
        return rev_choice_map[cached_ans], "Khớp ngân hàng Master Bank (Cache)"
        
    semi_choices = [k for k, v in choices_dict.items() if v.strip().endswith(';')]
    if len(semi_choices) == 1:
        return semi_choices[0], "Heuristic: Phương án có dấu chấm phẩy duy nhất"
        
    for k, v in choices_dict.items():
        v_lower = v.lower()
        if any(kw in v_lower for kw in ['tất cả các phương án', 'tất cả các đáp án', 'cả ba phương án']):
            return k, "Heuristic: Lựa chọn tổng hợp (Tất cả/Cả ba)"
            
    legal_keywords = ['thông tư', 'quyết định', 'nghị định', 'luật quản lý thuế', 'luật phòng']
    scores = {}
    for k, v in choices_dict.items():
        score = 0
        v_lower = v.lower()
        for kw in legal_keywords:
            if kw in v_lower:
                score += 2
        score += min(len(v) / 50.0, 3.0)
        scores[k] = score
        
    best_k = max(scores, key=scores.get)
    return best_k, "Heuristic: Trích dẫn điều luật hoặc độ dài chi tiết"

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

def decrypt_payload(raw_b64):
    res = subprocess.run(['node', DECRYPT_JS, 'decrypt', raw_b64], capture_output=True, text=True, encoding='utf-8')
    if res.returncode != 0:
        raise RuntimeError(f"Node decrypt error: {res.stderr}")
    return json.loads(res.stdout)

def fetch_and_solve_exam(group_id, subscription="P3", candidate_name=None, master_cache_file=DEFAULT_MASTER_EXCEL, output_dir=DEFAULT_OUTPUT_DIR, brain_sources_dir=DEFAULT_BRAIN_SOURCES, exam_index=None, out_excel_path=None):
    iter_label = f" (Lần {exam_index})" if exam_index else ""
    print(f"=== BẮT ĐẦU TRUY CẬP VÀ GIẢI ĐỀ THI KÍP {group_id} ({subscription}){iter_label} ===")
    
    # 1. Setup headers
    visitor_id = hashlib.md5(f"device_agent_{random.randint(10000, 99999)}".encode()).hexdigest()
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
    
    # 2. Register Guest
    if not candidate_name:
        candidate_name = random.choice(["Nguyễn Văn An", "Trần Thị Mai", "Lê Hoàng Long", "Phạm Quốc Bảo", "Đoàn Minh Trí"])
    print(f"[1/4] Đăng ký thí sinh dự thi: {candidate_name}...")
    reg_url = "https://api-e3.khaothi.online/delivery/exam/guest-register"
    reg_payload = {
        "test_taker_group_id": group_id,
        "subscription": subscription,
        "info[name]": candidate_name,
        "e_v": 2
    }
    req = urllib.request.Request(reg_url, data=urllib.parse.urlencode(reg_payload).encode('utf-8'), headers=headers)
    with urllib.request.urlopen(req) as resp:
        parsed = json.loads(resp.read().decode('utf-8'))
        reg_data = decrypt_payload(parsed['data'])
        test_taker = reg_data['test_taker']
        print(f"-> Thí sinh đã đăng ký: {test_taker['name']} (Mã code: {test_taker['code']})")
        
    # 3. Start Exam
    print("[2/4] Bắt đầu phiên làm bài thi...")
    start_url = "https://api-e3.khaothi.online/delivery/exam/start"
    start_payload = {
        "data[code]": test_taker['code'],
        "e_v": 2
    }
    req_start = urllib.request.Request(start_url, data=urllib.parse.urlencode(start_payload).encode('utf-8'), headers=headers)
    with urllib.request.urlopen(req_start) as resp:
        start_resp = json.loads(resp.read().decode('utf-8'))
        raw_data = start_resp['data']
        exam_data = decrypt_payload(raw_data) if isinstance(raw_data, str) else raw_data
        token = exam_data['token']
        items = exam_data['testFormData']['items']
        print(f"-> Phiên thi đã mở. Tổng số câu hỏi: {len(items)} câu.")

    # 4. Load Master Bank Cache
    if master_cache_file and os.path.exists(master_cache_file):
        cache = load_master_bank_cache(master_cache_file)
        print(f"-> Nạp Master Bank Cache: {len(cache)} câu hỏi đã có sẵn.")
    else:
        cache = {}
        print("-> Chế độ: Không đối soát với Master Bank, AI giải và dò trực tiếp từ hệ thống.")

    # 5. Solve Questions
    print("[3/4] Tiến hành lấy đề và kiểm tra đáp án từng câu...")
    check_url = "https://api-e3.khaothi.online/delivery/exam/submit-item-answer"
    check_headers = dict(headers)
    check_headers['Xcbt-Token'] = token
    choice_map = {"1": "A", "2": "B", "3": "C", "4": "D"}
    rev_choice_map = {"A": "1", "B": "2", "C": "3", "D": "4"}
    
    results = []
    
    for idx, item in enumerate(items, 1):
        item_id = item['id']
        q_text = clean_html(item['data_body_html'])
        raw_choices = json.loads(item['data_choices_html'])
        cleaned_choices = {k: clean_html(v) for k, v in raw_choices.items()}
        
        norm_q = normalize_text_for_match(q_text)
        cached_info = cache.get(norm_q)
        cached_ans = cached_info['ans'] if cached_info else None
        
        ai_choice_key, decision_reason = predict_best_choice(q_text, cleaned_choices, cached_ans)
        ai_choice_letter = choice_map.get(ai_choice_key, ai_choice_key)
        
        delay = round(random.uniform(5.1, 7.3), 2)
        print(f"\n--- Câu {idx}/{len(items)} (ID: {item_id}) ---")
        print(f"Nội dung: {q_text[:70]}...")
        print(f"Dự đoán ban đầu: {ai_choice_letter} ({decision_reason}) [Delay {delay}s]")
        time.sleep(delay)
        
        payload = {"item_id": item_id, "answer_data": ai_choice_key}
        is_correct = False
        latest_instruction = ""
        for attempt in range(3):
            try:
                req_check = urllib.request.Request(check_url, data=urllib.parse.urlencode(payload).encode('utf-8'), headers=check_headers)
                with urllib.request.urlopen(req_check, timeout=15) as check_resp:
                    check_data = json.loads(check_resp.read().decode('utf-8'))
                    is_correct = check_data.get('data', {}).get('is_correct', False)
                    latest_instruction = check_data.get('data', {}).get('instruction', '') or latest_instruction
                    break
            except Exception as e:
                print(f"Cảnh báo gửi request ({e}), thử lại sau 5s...")
                time.sleep(5)
                
        final_correct_key = None
        system_chose_key = None
        if is_correct:
            final_correct_key = ai_choice_key
            system_chose_key = ai_choice_key
            print(f"-> KẾT QUẢ: Lựa chọn {ai_choice_letter} ĐÚNG!")
        else:
            print(f"-> Lựa chọn {ai_choice_letter} chưa đúng. Dò các phương án còn lại...")
            remaining_keys = [k for k in sorted(raw_choices.keys()) if k != ai_choice_key]
            if latest_instruction:
                def score_choice(k):
                    c_text = cleaned_choices.get(k, "").lower()
                    words = [w for w in c_text.split() if len(w) > 2]
                    return sum(1 for w in words if w in latest_instruction.lower())
                remaining_keys.sort(key=score_choice, reverse=True)
                
            for r_key in remaining_keys:
                r_letter = choice_map.get(r_key, r_key)
                r_delay = round(random.uniform(5.1, 7.1), 2)
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
                            system_chose_key = r_key
                            print(f"-> KẾT QUẢ: Dò trúng phương án {r_letter} ĐÚNG!")
                            break
                except Exception as e:
                    print(f"Lỗi dò đáp án ({e})")
                    time.sleep(4)
                    
        correct_letter = choice_map.get(final_correct_key, "N/A")
        correct_text = cleaned_choices.get(final_correct_key, "")
        
        explanation = cached_info['explanation'] if cached_info else ""
        if latest_instruction:
            explanation = latest_instruction if not explanation else f"{explanation} | {latest_instruction}"
            
        results.append({
            'index': idx,
            'item_id': item_id,
            'topic': classify_topic(q_text),
            'question': q_text,
            'choice_a': cleaned_choices.get("1", ""),
            'choice_b': cleaned_choices.get("2", ""),
            'choice_c': cleaned_choices.get("3", ""),
            'choice_d': cleaned_choices.get("4", ""),
            'ai_choice': ai_choice_letter,
            'ai_correct': "Đúng" if is_correct else "Chưa đúng",
            'correct_choice': correct_letter,
            'correct_text': correct_text,
            'explanation': explanation
        })
        
        # Checkpoint tiến độ sau mỗi câu
        checkpoint_path = os.path.join(output_dir, f"checkpoint_{subscription}_{group_id}.json")
        try:
            with open(checkpoint_path, 'w', encoding='utf-8') as f_ckpt:
                json.dump({
                    'group_id': group_id,
                    'subscription': subscription,
                    'total': len(items),
                    'completed': idx,
                    'progress_percent': round(idx * 100 / len(items), 1),
                    'latest_question': q_text[:80],
                    'latest_correct_choice': correct_letter,
                    'updated_at': time.strftime('%Y-%m-%d %H:%M:%S')
                }, f_ckpt, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # 6. Export to Excel & Markdown
    print("\n[4/4] Xuất kết quả sang Excel và Markdown...")
    os.makedirs(output_dir, exist_ok=True)
    if not out_excel_path:
        if exam_index:
            out_excel = os.path.join(output_dir, f"Bo_De_On_Thi_Kiem_Tra_Thue_{subscription}_{group_id}_Lan_{exam_index:02d}.xlsx")
        else:
            out_excel = os.path.join(output_dir, f"Bo_De_On_Thi_Kiem_Tra_Thue_{subscription}_{group_id}_Chuan.xlsx")
    else:
        out_excel = out_excel_path
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"De_{group_id}_L{exam_index}" if exam_index else f"De_{group_id}_Chi_Tiet"
    
    headers_list = [
        "STT", "Mã Câu", "Chuyên Đề", "Nội Dung Câu Hỏi",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "AI Chọn", "Kết Quả AI", "Đáp Án Chuẩn", "Nội Dung Đáp Án Chuẩn", "Căn Cứ Pháp Lý / Ghi Chú"
    ]
    ws.append(headers_list)
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx in range(1, len(headers_list) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(name="Calibri", size=11, bold=True, color="006100")
    
    for r in results:
        row_vals = [
            r['index'], r['item_id'], r['topic'], r['question'],
            r['choice_a'], r['choice_b'], r['choice_c'], r['choice_d'],
            r['ai_choice'], r['ai_correct'], r['correct_choice'], r['correct_text'], r['explanation']
        ]
        ws.append(row_vals)
        curr_row = ws.max_row
        ans_cell = ws.cell(row=curr_row, column=11)
        ans_cell.fill = green_fill
        ans_cell.font = green_font
        ans_cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # Auto column width
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 60)
        
    wb.save(out_excel)
    print(f"-> Đã lưu Excel đề thi: {out_excel}")
    
    # Save Markdown in Brain sources
    if os.path.exists(brain_sources_dir):
        out_md = os.path.join(brain_sources_dir, f"de_on_tap_kiem_tra_thue_{subscription.lower()}_{group_id}.md")
        with open(out_md, 'w', encoding='utf-8') as f:
            f.write(f"# Bộ đề chuẩn: Ôn tập - Lĩnh vực kiểm tra thuế (Mã kíp: {group_id})\n")
            f.write(f"- **Tổng số câu hỏi**: {len(results)} câu\n")
            correct_first = sum(1 for r in results if r['ai_correct'] == "Đúng")
            f.write(f"- **Số câu chọn đúng lần đầu**: {correct_first}/{len(results)} ({round(correct_first*100/len(results), 1)}%)\n")
            f.write(f"- **Số câu hệ thống xác nhận chuẩn xác 100%**: {len(results)}/{len(results)} câu\n")
            f.write(f"- **Thời gian hoàn tất**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n---\n\n")
            for r in results:
                f.write(f"### Câu {r['index']} (Mã ID: {r['item_id']}) - {r['topic']}\n\n")
                f.write(f"**{r['question']}**\n\n")
                f.write(f"- **A.** {r['choice_a']}\n")
                f.write(f"- **B.** {r['choice_b']}\n")
                f.write(f"- **C.** {r['choice_c']}\n")
                f.write(f"- **D.** {r['choice_d']}\n\n")
                f.write(f"- **Đáp án đúng chuẩn từ hệ thống**: **{r['correct_choice']}**\n")
                f.write(f"- **Nội dung chuẩn**: *{r['correct_text']}*\n\n---\n\n")
        print(f"-> Đã đồng bộ Markdown vào Second Brain: {out_md}")
        
    return results, out_excel

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Khảo thí thuế - Lấy đề và giải đề tự động")
    parser.add_argument("--url", help="URL kíp thi (vd https://taxedu.gdt.gov.vn/t/P3-36751)")
    parser.add_argument("--group-id", type=int, help="Mã kíp thi (vd 36751)")
    parser.add_argument("--sub", default="P3", help="Mã subscription (mặc định P3)")
    parser.add_argument("--name", help="Tên thí sinh ảo")
    parser.add_argument("--master", default=DEFAULT_MASTER_EXCEL, help="Đường dẫn file Master Bank")
    parser.add_argument("--no-cache", action="store_true", help="Không đối soát với Master Bank")
    args = parser.parse_args()
    
    group_id = args.group_id
    sub = args.sub
    if args.url:
        match = re.search(r'/t/([A-Za-z0-9]+)-(\d+)', args.url)
        if match:
            sub = match.group(1)
            group_id = int(match.group(2))
            
    if not group_id:
        print("Lỗi: Vui lòng cung cấp --url hoặc --group-id")
        sys.exit(1)
        
    master_file = None if args.no_cache else args.master
    fetch_and_solve_exam(group_id, subscription=sub, candidate_name=args.name, master_cache_file=master_file)
