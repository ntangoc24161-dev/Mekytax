"""
Tien trinh xu ly hang loat 30 bo de kiem tra thue Dot 2
- Trich xuat anh tu 30 file Word (.docx) tai 'D:\\HOC TAP\\ON THI 2026.09\\On tap thi nghiep vu\\Dot 2\\KIEM TRA THUE\\NEW'
- Nhan dien nut radio mau xanh (dap an he thong da chon) va cac nut xam bang OpenCV HSV
- Nhan dien toan bo van ban tieng Viet chuan xac bang PaddleOCR PP-OCRv5 Latin trong mot luot (single-pass)
- Phan bo chinh xac noi dung cau hoi va cac phuong an A, B, C, D dua tren toa do hinh hoc
- Luu checkpoint lien tuc vao 'dot2_checkpoint.json' (ho tro tiep tuc bat ky luc nao khong mat du lieu)
- Tu dong tong hop va khau tru trung lap vao 'Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.xlsx'
"""

import os
import sys
import glob
import json
import time
import zipfile
import re
import datetime
import argparse
from pathlib import Path

# Ensure UTF-8 output on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Disable paddle MKLDNN / PIR warnings
os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["FLAGS_enable_pir_api"] = "0"

import cv2
import numpy as np
import openpyxl
from paddleocr import PaddleOCR

# Cau hinh duong dan
SOURCE_DIR = r"D:\HOC TAP\ON THI 2026.09\Ôn tập thi nghiệp vụ\Đợt 2\KIEM TRA THUE\NEW"
OUTPUT_DIR = r"D:\HOC TAP\ON THI 2026.09"
CACHE_IMG_DIR = os.path.join(OUTPUT_DIR, ".cache_dot2_images")
CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "dot2_checkpoint.json")
PROGRESS_FILE = os.path.join(OUTPUT_DIR, "tien_do_boc_tach_dot2.json")
EXCEL_MASTER = os.path.join(OUTPUT_DIR, "Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.xlsx")
BRAIN_MD = r"D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\sources\Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.md"
BRAIN2_MD = r"D:\SECOND BRAIN\javis-os\brains\tổng hợp ôn tập\sources\Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.md"

os.makedirs(CACHE_IMG_DIR, exist_ok=True)


def log(msg):
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now_str}] {msg}", flush=True)


def load_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            log(f"Loi doc checkpoint: {e}")
    return {"processed_tests": {}, "questions": {}, "total_processed": 0}


def save_checkpoint(ckpt):
    try:
        temp_file = CHECKPOINT_FILE + ".tmp"
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(ckpt, f, ensure_ascii=False, indent=2)
        if os.path.exists(CHECKPOINT_FILE):
            os.replace(temp_file, CHECKPOINT_FILE)
        else:
            os.rename(temp_file, CHECKPOINT_FILE)
    except Exception as e:
        log(f"Loi ghi checkpoint: {e}")


def update_progress(current_test, current_q, total_tests, total_q, status_msg="Đang xử lý"):
    progress_data = {
        "status": status_msg,
        "current_test": current_test,
        "current_question": current_q,
        "total_tests": total_tests,
        "total_questions_processed": total_q,
        "target_total_questions": 1800,
        "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(progress_data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def extract_images_from_docx(docx_path, out_subfolder):
    """Boc tach 60 anh theo dung thu tu xuat hien trong document.xml"""
    os.makedirs(out_subfolder, exist_ok=True)
    images_in_order = []
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        rels_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")
        rel_map = {}
        for m in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml):
            rel_map[m.group(1)] = m.group(2)
            
        doc_xml = z.read("word/document.xml").decode("utf-8")
        embeds = re.findall(r'r:embed="([^"]+)"', doc_xml)
        
        for idx, rId in enumerate(embeds):
            target = rel_map.get(rId, "")
            if not target:
                continue
            if target.startswith("media/"):
                zip_path = "word/" + target
            elif not target.startswith("word/"):
                zip_path = "word/media/" + target.split("/")[-1]
            else:
                zip_path = target
                
            out_img_name = f"q_{idx+1:02d}.png"
            out_img_path = os.path.join(out_subfolder, out_img_name)
            if not os.path.exists(out_img_path) or os.path.getsize(out_img_path) == 0:
                with open(out_img_path, "wb") as f_out:
                    f_out.write(z.read(zip_path))
            images_in_order.append((idx + 1, out_img_path))
            
    return images_in_order


def extract_all_tests():
    """Trich xuat anh cho toan bo 30 de thi"""
    docx_files = glob.glob(os.path.join(SOURCE_DIR, "*.docx"))
    docx_files.sort()
    log(f"Bat dau trich xuat anh cho toan bo {len(docx_files)} file docx...")
    total_imgs = 0
    for idx, docx_path in enumerate(docx_files, 1):
        filename = os.path.basename(docx_path)
        m = re.search(r'Đề\s*(\d+)', filename, re.IGNORECASE)
        test_num = int(m.group(1)) if m else idx
        subfolder = os.path.join(CACHE_IMG_DIR, f"De_{test_num:02d}")
        imgs = extract_images_from_docx(docx_path, subfolder)
        total_imgs += len(imgs)
        log(f"  [{idx:02d}/{len(docx_files)}] Đề {test_num:02d}: {len(imgs)} ảnh -> {os.path.basename(subfolder)}")
    log(f"Hoan tat trich xuat tong cong {total_imgs} anh tu 30 bo de.")
    return total_imgs


def parse_single_image(img_path, ocr_engine):
    """Nhan dien cau hoi, cac dap an va nut radio duoc chon tu 1 anh bang Single-Pass OCR"""
    img = cv2.imread(img_path)
    if img is None:
        return None
    h, w, _ = img.shape
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # 1. Tim nut mau xanh (dap an da chon)
    lower_blue = np.array([90, 80, 80])
    upper_blue = np.array([135, 255, 255])
    mask_blue = cv2.inRange(hsv, lower_blue, upper_blue)
    cnts_blue, _ = cv2.findContours(mask_blue, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    blue_pt = None
    for c in cnts_blue:
        x, y, bw, bh = cv2.boundingRect(c)
        if 15 <= bw <= 80 and 15 <= bh <= 80 and abs(bw - bh) <= 20:
            blue_pt = (x + bw/2, y + bh/2)
            break
            
    # 2. Tim cac vong tron radio bang HoughCircles va contour xam
    circles = cv2.HoughCircles(gray, cv2.HOUGH_GRADIENT, 1, 20, param1=50, param2=30, minRadius=15, maxRadius=40)
    c_list = []
    if circles is not None:
        for c in circles[0]:
            cx, cy, r = float(c[0]), float(c[1]), float(c[2])
            if cx < 100 and cy < 80:  # Bo qua huy hieu so thu tu cau o goc trai tren
                continue
            c_list.append((cx, cy, r))
            
    # Neu Hough khong du hoac khong co, thu bo sung tu contour xam
    if len(c_list) < 3:
        lower_grey = np.array([0, 0, 180])
        upper_grey = np.array([180, 50, 250])
        mask_grey = cv2.inRange(hsv, lower_grey, upper_grey)
        cnts_grey, _ = cv2.findContours(mask_grey, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for c in cnts_grey:
            x, y, bw, bh = cv2.boundingRect(c)
            if 15 <= bw <= 75 and 15 <= bh <= 75 and abs(bw - bh) <= 20:
                cx, cy = x + bw/2, y + bh/2
                if cx < 100 and cy < 80:
                    continue
                if not any(abs(cx - ex[0]) < 15 and abs(cy - ex[1]) < 15 for ex in c_list):
                    c_list.append((cx, cy, bw/2))
        if blue_pt and not any(abs(blue_pt[0] - ex[0]) < 15 and abs(blue_pt[1] - ex[1]) < 15 for ex in c_list):
            c_list.append((blue_pt[0], blue_pt[1], 25))

    is_horizontal = False
    chosen_idx = -1
    labels = ['A', 'B', 'C', 'D']
    
    if len(c_list) >= 3:
        ys = [c[1] for c in c_list]
        is_horizontal = (max(ys) - min(ys) < 40)
        if is_horizontal:
            c_list.sort(key=lambda c: c[0])
        else:
            c_list.sort(key=lambda c: c[1])
        c_list = c_list[:4]
        
        # Xac dinh radio duoc chon
        if blue_pt:
            for idx, c in enumerate(c_list):
                if abs(c[0] - blue_pt[0]) < 35 and abs(c[1] - blue_pt[1]) < 35:
                    chosen_idx = idx
                    break
        if chosen_idx == -1:
            # Thu kiem tra tam vong tron tren kenh xam (cho truong hop anh don sac/xam)
            inners = []
            for c in c_list:
                cx, cy = int(c[0]), int(c[1])
                inner = gray[max(0, cy-5):min(h, cy+5), max(0, cx-5):min(w, cx+5)]
                inners.append(np.mean(inner) if inner.size > 0 else 255)
            min_idx = int(np.argmin(inners))
            if inners[min_idx] < 200:
                chosen_idx = min_idx
                
    chosen_answer = labels[chosen_idx] if 0 <= chosen_idx < len(labels) else ''

    # 3. Single-Pass OCR toan bo buc anh
    res = ocr_engine.ocr(img)
    lines = []
    if res and res[0]:
        rec_texts = res[0].get('rec_texts', [])
        dt_polys = res[0].get('dt_polys', [])
        for text, poly in zip(rec_texts, dt_polys):
            poly_np = np.array(poly)
            cx = float(np.mean(poly_np[:, 0]))
            cy = float(np.mean(poly_np[:, 1]))
            lines.append({'text': text.strip(), 'cx': cx, 'cy': cy})

    options = {"A": "", "B": "", "C": "", "D": ""}
    
    if is_horizontal and len(c_list) >= 2:
        first_y = min(c[1] for c in c_list)
        # Cau hoi o phia tren hang radio ngang, loai bo badge goc trai tren (cx < 100, cy < 80)
        q_lines = [l['text'] for l in lines if (l['cy'] < (first_y - 12)) and not (l['cx'] < 100 and l['cy'] < 80)]
        q_text = " ".join(q_lines).strip()
        
        for idx in range(len(c_list)):
            lbl = labels[idx]
            curr_x = c_list[idx][0]
            prev_x = c_list[idx-1][0] if idx > 0 else 0
            next_x = c_list[idx+1][0] if idx < len(c_list)-1 else w
            x_min = (prev_x + curr_x)/2 if idx > 0 else curr_x - 10
            x_max = (curr_x + next_x)/2 if idx < len(c_list)-1 else w
            opt_lines = [l['text'] for l in lines if (l['cy'] >= first_y - 30) and (x_min <= l['cx'] < x_max)]
            opt_str = " ".join(opt_lines).strip()
            opt_str = re.sub(r'^[A-D][\.\:\/\-\)]\s*', '', opt_str)
            options[lbl] = opt_str
    else:
        first_radio_y = c_list[0][1] if c_list else h * 0.35
        # Cau hoi o phia tren radio dau tien, loai bo badge goc trai tren
        q_lines = [l['text'] for l in lines if (l['cy'] < (first_radio_y - 12)) and not (l['cx'] < 100 and l['cy'] < 80)]
        q_text = " ".join(q_lines).strip()
        
        for idx in range(min(len(c_list), 4)):
            lbl = labels[idx]
            curr_y = c_list[idx][1]
            radio_cx = c_list[idx][0]
            
            prev_y = c_list[idx-1][1] if idx > 0 else (first_radio_y - 20)
            next_y = c_list[idx+1][1] if idx < len(c_list) - 1 else h
            
            y_min = (prev_y + curr_y) / 2 if idx > 0 else (first_radio_y - 12)
            y_max = (curr_y + next_y) / 2 if idx < len(c_list) - 1 else h
            
            opt_lines = [l['text'] for l in lines if (y_min <= l['cy'] < y_max) and (l['cx'] > radio_cx - 10)]
            opt_str = " ".join(opt_lines).strip()
            opt_str = re.sub(r'^[A-D][\.\:\/\-\)]\s*', '', opt_str)
            options[lbl] = opt_str

    # Loai bo so thu tu o dau cau neu con sot (vi du "1.", "12.", "15 ")
    q_text = re.sub(r'^\d+[\.\:\s\-]+', '', q_text).strip()

    return {
        "question": q_text,
        "options": options,
        "chosen_answer": chosen_answer
    }


def update_master_excel(ckpt):
    """Cap nhat sheet Dot_2_30_De_Chi_Tiet va bo sung Ngan_Hang_Cau_Hoi_Chuan"""
    if not os.path.exists(EXCEL_MASTER):
        log("Chua tim thay file Excel Master de cap nhat.")
        return 0
        
    log("Dang cap nhat du lieu vao file Excel Tong Hop Master...")
    wb = openpyxl.load_workbook(EXCEL_MASTER)
    
    # 1. Sheet Dot_2_30_De_Chi_Tiet
    sheet_name = "Dot_2_30_De_Chi_Tiet"
    if sheet_name in wb.sheetnames:
        ws_dot2 = wb[sheet_name]
        wb.remove(ws_dot2)
    ws_dot2 = wb.create_sheet(title=sheet_name)
    
    headers = [
        "STT Toàn Đợt", "Mã Đề", "STT Câu Trong Đề", "Nội Dung Câu Hỏi",
        "Phương Án A", "Phương Án B", "Phương Án C", "Phương Án D",
        "Đáp Án Đã Chọn (Radio Xanh)", "Nội Dung Đáp Án Đúng", "File Ảnh Gốc"
    ]
    ws_dot2.append(headers)
    
    all_qs = ckpt.get("questions", {})
    sorted_keys = sorted(all_qs.keys(), key=lambda k: (all_qs[k].get("test_num", 999), all_qs[k].get("q_idx", 999)))
    
    for row_idx, k in enumerate(sorted_keys, 1):
        item = all_qs[k]
        opts = item.get("options", {})
        ans = item.get("chosen_answer", "")
        ans_text = opts.get(ans, "") if ans else ""
        
        ws_dot2.append([
            row_idx,
            item.get("test_name", ""),
            item.get("q_idx", ""),
            item.get("question", ""),
            opts.get("A", ""),
            opts.get("B", ""),
            opts.get("C", ""),
            opts.get("D", ""),
            ans,
            ans_text,
            item.get("img_path", "")
        ])

    # 2. Khử trùng lặp và bổ sung vào Ngan_Hang_Cau_Hoi_Chuan
    ws_master = wb["Ngan_Hang_Cau_Hoi_Chuan"]
    existing_questions = set()
    for row in ws_master.iter_rows(min_row=2, values_only=True):
        if row and len(row) > 2 and row[2]:
            q_norm = re.sub(r'[\s\W]+', '', str(row[2]).lower())
            existing_questions.add(q_norm)
            
    added_count = 0
    next_stt = ws_master.max_row
    
    for k in sorted_keys:
        item = all_qs[k]
        q_text = item.get("question", "").strip()
        if not q_text or len(q_text) < 15:
            continue
        q_norm = re.sub(r'[\s\W]+', '', q_text.lower())
        if q_norm not in existing_questions:
            existing_questions.add(q_norm)
            opts = item.get("options", {})
            ans = item.get("chosen_answer", "")
            
            ws_master.append([
                next_stt,
                f"Đợt 2 - {item.get('test_name', '')}",
                q_text,
                opts.get("A", ""),
                opts.get("B", ""),
                opts.get("C", ""),
                opts.get("D", ""),
                ans,
                "Bóc tách từ ảnh chụp màn hình bài thi Đợt 2 (Nút chọn màu xanh)",
                item.get("source_docx", "")
            ])
            next_stt += 1
            added_count += 1
            
    try:
        wb.save(EXCEL_MASTER)
        log(f"Da ghi thanh cong vao Excel Master: bo sung them {added_count} cau hoi doc nhat moi (tong {len(sorted_keys)} cau Dot 2).")
    except Exception as e_save:
        log(f"[Canh bao] Khong the ghi de Excel (co the file dang mo): {e_save}")
        backup_xlsx = os.path.join(OUTPUT_DIR, "Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026_Dot2.xlsx")
        try:
            wb.save(backup_xlsx)
            log(f"Da luu ban sao luu vao: {backup_xlsx}")
        except Exception:
            pass

    # Cap nhat Markdown cho Brain sources
    try:
        md_content = f"""# Ngân Hàng Câu Hỏi Ôn Thi Kiểm Tra Thuế 2026 (Master)

- **Ngày cập nhật**: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M")}
- **Tổng số câu hỏi Đợt 2 đã bóc tách**: **{len(sorted_keys)} câu trắc nghiệm**
- **Tổng số đề Đợt 2 đã hoàn thành**: **{len(ckpt.get('processed_tests', {}))}/30 đề**
- **Trạng thái tiến trình**: Tự động bóc tách và đối soát liên tục bằng PaddleOCR PP-OCRv5

---

## Danh sách tiến độ bóc tách 30 đề Đợt 2

| Đề thi | Số câu đã bóc | Trạng thái | Thời gian hoàn thành |
| :--- | :---: | :---: | :--- |
"""
        for t_idx in range(1, 31):
            t_key = f"de_{t_idx:02d}"
            info = ckpt.get("processed_tests", {}).get(t_key, {})
            st = " Hoàn thành" if info.get("status") == "completed" else " Đang chờ / Đang xử lý"
            c_cnt = info.get("questions_count", 0)
            c_time = info.get("completed_at", "-")
            md_content += f"| Đề {t_idx:02d} | {c_cnt}/60 | {st} | {c_time} |\n"

        for md_path in [BRAIN_MD, BRAIN2_MD]:
            os.makedirs(os.path.dirname(md_path), exist_ok=True)
            with open(md_path, "w", encoding="utf-8") as f_md:
                f_md.write(md_content)
    except Exception as e_md:
        log(f"Loi cap nhat Markdown brain: {e_md}")

    return added_count


def process_tests(max_tests=None, specific_test=None):
    docx_files = glob.glob(os.path.join(SOURCE_DIR, "*.docx"))
    docx_files.sort()
    
    ckpt = load_checkpoint()
    log(f"Checkpoint hien tai: {len(ckpt.get('questions', {}))} cau hoi da luu.")
    
    log("Khoi tao PaddleOCR PP-OCRv5 Latin (chuan tieng Viet co dau, Single-Pass)...")
    ocr = PaddleOCR(
        ocr_version='PP-OCRv5',
        lang='vi',
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False
    )
    
    tests_processed_this_run = 0
    
    for file_idx, docx_path in enumerate(docx_files, 1):
        test_filename = os.path.basename(docx_path)
        m = re.search(r'Đề\s*(\d+)', test_filename, re.IGNORECASE)
        test_num = int(m.group(1)) if m else file_idx
        test_key = f"de_{test_num:02d}"
        
        if specific_test is not None and test_num != specific_test:
            continue
            
        if test_key in ckpt.get("processed_tests", {}) and ckpt["processed_tests"][test_key].get("status") == "completed":
            log(f"[{file_idx}/{len(docx_files)}] Đề {test_num:02d} ({test_filename}) da hoan tat truoc do. Bo qua.")
            continue
            
        log(f"\n>>> Dang xu ly [{file_idx}/{len(docx_files)}]: Đề {test_num:02d} ({test_filename})...")
        subfolder = os.path.join(CACHE_IMG_DIR, f"De_{test_num:02d}")
        
        # 1. Trích xuất ảnh
        try:
            images = extract_images_from_docx(docx_path, subfolder)
        except Exception as e:
            log(f"Loi trich xuat anh tu docx: {e}")
            continue
            
        log(f"  Co {len(images)} anh cau hoi trong de {test_num:02d}.")
        
        # 2. Xu ly tung anh
        test_questions_done = 0
        t0_test = time.time()
        for q_idx, img_path in images:
            q_id = f"de_{test_num:02d}_q_{q_idx:02d}"
            if q_id in ckpt.get("questions", {}):
                test_questions_done += 1
                continue
                
            t0_q = time.time()
            try:
                res = parse_single_image(img_path, ocr)
                if res:
                    ckpt.setdefault("questions", {})[q_id] = {
                        "test_num": test_num,
                        "test_name": f"Đề {test_num:02d}",
                        "q_idx": q_idx,
                        "question": res["question"],
                        "options": res["options"],
                        "chosen_answer": res["chosen_answer"],
                        "img_path": img_path,
                        "source_docx": test_filename,
                        "timestamp": datetime.datetime.now().isoformat()
                    }
                    test_questions_done += 1
                    total_q_now = len(ckpt.get("questions", {}))
                    ckpt["total_processed"] = total_q_now
                    dt_q = time.time() - t0_q
                    
                    if q_idx % 5 == 0 or q_idx == len(images):
                        log(f"  - Đề {test_num:02d}: cau {q_idx:02d}/{len(images)} (Đã chọn: {res['chosen_answer']}) [{dt_q:.1f}s/cau] -> Tong da bóc: {total_q_now}/1800")
                        save_checkpoint(ckpt)
                        update_progress(f"Đề {test_num:02d}", q_idx, len(docx_files), total_q_now)
            except Exception as e:
                log(f"  [!] Loi khi xu ly anh cau {q_idx:02d}: {e}")
                
        # Danh dau de hoan thanh
        dt_test = time.time() - t0_test
        ckpt.setdefault("processed_tests", {})[test_key] = {
            "status": "completed",
            "file": test_filename,
            "questions_count": test_questions_done,
            "time_seconds": dt_test,
            "completed_at": datetime.datetime.now().isoformat()
        }
        save_checkpoint(ckpt)
        update_progress(f"Đề {test_num:02d}", len(images), len(docx_files), len(ckpt.get("questions", {})), "Đã hoàn thành 1 đề")
        log(f">>> HOAN TAT Đề {test_num:02d}: {test_questions_done} cau trong {dt_test:.1f}s ({dt_test/60:.1f} phut).")
        
        # Cap nhat vao Excel sau moi de
        try:
            update_master_excel(ckpt)
        except Exception as e:
            log(f"Loi cap nhat Excel sau de {test_num:02d}: {e}")
            
        tests_processed_this_run += 1
        if max_tests is not None and tests_processed_this_run >= max_tests:
            log(f"Dat gioi han {max_tests} de thi cho lan chay nay. Tam dung.")
            break
            
    total_final = len(ckpt.get("questions", {}))
    update_progress("Hoàn tất phiên", 1800, len(docx_files), total_final, f"Đã bóc tách {total_final} câu")
    update_master_excel(ckpt)
    log(f"================ KET THUC PHIEN XU LY: TONG {total_final}/1800 CAU HOI ================")


def main():
    parser = argparse.ArgumentParser(description="Tien trinh xu ly hang loat Dot 2")
    parser.add_argument("--extract-only", action="store_true", help="Chi trich xuat anh tu docx")
    parser.add_argument("--max-tests", type=int, default=None, help="So de thi toi da can xu ly trong dot nay")
    parser.add_argument("--test", type=int, default=None, help="Xu ly duy nhat 1 de thi chi dinh")
    parser.add_argument("--update-excel", action="store_true", help="Cap nhat lai file Excel tu checkpoint")
    args = parser.parse_args()
    
    if args.extract_only:
        extract_all_tests()
    elif args.update_excel:
        ckpt = load_checkpoint()
        update_master_excel(ckpt)
    else:
        process_tests(max_tests=args.max_tests, specific_test=args.test)


if __name__ == "__main__":
    main()
