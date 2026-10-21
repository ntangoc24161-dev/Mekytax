import os
import sys
import json
import time
import random
import re
import hashlib
import urllib.request
import urllib.parse
import subprocess
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = r"D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\skills\lay-va-tong-hop-de-thi\scripts"
DECRYPT_JS = os.path.join(SCRIPT_DIR, "decrypt_dcv2.js")
OUTPUT_DIR = r"D:\HOC TAP\ON THI 2026.09"
BRAIN_SOURCES = r"D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\sources"

def decrypt_payload(raw_b64):
    res = subprocess.run(['node', DECRYPT_JS, 'decrypt', raw_b64], capture_output=True, text=True, encoding='utf-8')
    if res.returncode != 0:
        raise RuntimeError(f"Node decrypt error: {res.stderr}")
    return json.loads(res.stdout)

def clean_html_preserve_text(text):
    if not text:
        return ""
    # Convert <br> and </p> to newline
    t = re.sub(r'<br\s*/?>', '\n', text)
    t = re.sub(r'</p>', '\n', t)
    # Remove remaining HTML tags
    t = re.sub(r'<[^>]+>', '', t)
    # Replace html entities
    t = t.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&#39;', "'")
    lines = [l.strip() for l in t.split('\n') if l.strip()]
    return '\n'.join(lines)

def run_extraction():
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

    # 1. Guest register
    candidate_name = f"Thí Sinh {random.randint(100, 999)}"
    print(f"1. Đăng nhập với tên ngẫu nhiên: {candidate_name}")
    reg_url = "https://api-e3.khaothi.online/delivery/exam/guest-register"
    reg_payload = {
        "test_taker_group_id": 39917,
        "subscription": "P3",
        "info[name]": candidate_name,
        "e_v": 2
    }
    req = urllib.request.Request(reg_url, data=urllib.parse.urlencode(reg_payload).encode('utf-8'), headers=headers)
    with urllib.request.urlopen(req) as resp:
        parsed = json.loads(resp.read().decode('utf-8'))
        reg_data = decrypt_payload(parsed['data'])
        test_taker = reg_data['test_taker']
        print(f"-> Mã thí sinh: {test_taker['code']}")

    # 2. Start exam
    print("2. Bắt đầu phiên thi...")
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
        print(f"-> Đã tải toàn bộ {len(items)} câu hỏi từ hệ thống.")

    # 3. Submit random answers and collect data
    print("3. Chọn ngẫu nhiên đáp án để chuyển câu và trích xuất toàn bộ 90 câu...")
    check_url = "https://api-e3.khaothi.online/delivery/exam/submit-item-answer"
    check_headers = dict(headers)
    check_headers['Xcbt-Token'] = token

    choice_map = {"1": "A", "2": "B", "3": "C", "4": "D"}
    extracted_questions = []

    for idx, item in enumerate(items, 1):
        item_id = item['id']
        raw_body = item['data_body_html']
        clean_body = clean_html_preserve_text(raw_body)
        
        raw_choices = json.loads(item['data_choices_html'])
        choice_a = clean_html_preserve_text(raw_choices.get("1", ""))
        choice_b = clean_html_preserve_text(raw_choices.get("2", ""))
        choice_c = clean_html_preserve_text(raw_choices.get("3", ""))
        choice_d = clean_html_preserve_text(raw_choices.get("4", ""))
        
        # Chọn ngẫu nhiên đáp án để qua câu mới trên hệ thống
        available_keys = list(raw_choices.keys())
        chosen_key = random.choice(available_keys) if available_keys else "1"
        chosen_letter = choice_map.get(chosen_key, chosen_key)
        
        payload = {"item_id": item_id, "answer_data": chosen_key}
        is_correct = False
        instruction = ""
        try:
            req_check = urllib.request.Request(check_url, data=urllib.parse.urlencode(payload).encode('utf-8'), headers=check_headers)
            with urllib.request.urlopen(req_check, timeout=10) as check_resp:
                res_data = json.loads(check_resp.read().decode('utf-8'))
                is_correct = res_data.get('data', {}).get('is_correct', False)
                instruction = res_data.get('data', {}).get('instruction', '')
        except Exception as e:
            print(f"Lỗi gửi câu {idx}: {e}")

        extracted_questions.append({
            'stt': idx,
            'item_id': item_id,
            'question': clean_body,
            'choice_a': choice_a,
            'choice_b': choice_b,
            'choice_c': choice_c,
            'choice_d': choice_d,
            'chosen_answer': chosen_letter,
            'is_correct': "Đúng" if is_correct else "Chưa đúng",
            'instruction': instruction
        })
        
        if idx % 10 == 0 or idx == len(items):
            print(f"-> Tiến độ trích xuất: {idx}/{len(items)} câu...")
        time.sleep(0.1)

    # 4. Generate Excel
    excel_path = os.path.join(OUTPUT_DIR, "Bo_De_On_Thi_Kiem_Tra_Thue_P3_39917_90_Cau.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "De_P3_39917_90_Cau"

    headers_list = [
        "STT", "Mã Câu Hỏi", "Nội Dung Câu Hỏi",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Đáp Án Chọn Hệ Thống", "Kết Quả Hệ Thống", "Căn Cứ / Hướng Dẫn"
    ]
    ws.append(headers_list)

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx in range(1, len(headers_list) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for q in extracted_questions:
        row = [
            q['stt'], q['item_id'], q['question'],
            q['choice_a'], q['choice_b'], q['choice_c'], q['choice_d'],
            q['chosen_answer'], q['is_correct'], q['instruction']
        ]
        ws.append(row)
        curr_row = ws.max_row
        for col_idx in range(1, len(headers_list) + 1):
            c = ws.cell(row=curr_row, column=col_idx)
            c.border = thin_border
            if col_idx in [1, 2, 8, 9]:
                c.alignment = Alignment(horizontal="center", vertical="top")
            else:
                c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Set column widths
    col_widths = {
        'A': 8,   # STT
        'B': 14,  # Mã Câu Hỏi
        'C': 45,  # Nội Dung Câu Hỏi
        'D': 35,  # Phương Án A
        'E': 35,  # Phương Án B
        'F': 35,  # Phương Án C
        'G': 35,  # Phương Án D
        'H': 22,  # Đáp Án Chọn Hệ Thống
        'I': 18,  # Kết Quả Hệ Thống
        'J': 45   # Căn Cứ / Hướng Dẫn
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(excel_path)
    print(f"-> Đã xuất thành công file Excel 90 câu: {excel_path}")

    # Also save Markdown backup
    md_path = os.path.join(BRAIN_SOURCES, "de_on_tap_kiem_tra_thue_p3_39917_90_cau.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("# Bộ Đề Ôn Tập Kiểm Tra Thuế - Kíp P3-39917 (90 Câu Gốc)\n\n")
        f.write(f"- **Tổng số câu**: {len(extracted_questions)} câu\n")
        f.write(f"- **File Excel xuất**: `{excel_path}`\n\n---\n\n")
        for q in extracted_questions:
            f.write(f"### Câu {q['stt']} (Mã ID: {q['item_id']})\n\n")
            f.write(f"**{q['question']}**\n\n")
            f.write(f"- **A.** {q['choice_a']}\n")
            f.write(f"- **B.** {q['choice_b']}\n")
            f.write(f"- **C.** {q['choice_c']}\n")
            f.write(f"- **D.** {q['choice_d']}\n\n")
            if q['instruction']:
                f.write(f"> **Căn cứ/Hướng dẫn**: {q['instruction']}\n\n")
            f.write("---\n\n")
    print(f"-> Đã lưu Markdown: {md_path}")

if __name__ == "__main__":
    run_extraction()
