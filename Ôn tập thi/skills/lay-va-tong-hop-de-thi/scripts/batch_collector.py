import sys
import os
import time
import random
import argparse
import json

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from khao_thi_client import fetch_and_solve_exam, DEFAULT_MASTER_EXCEL, DEFAULT_OUTPUT_DIR, DEFAULT_BRAIN_SOURCES
from bank_aggregator import aggregate_new_questions

CANDIDATE_NAMES = [
    "Nguyễn Hải Đăng", "Vũ Minh Tuấn", "Trần Thu Hà", "Lê Thanh Bình",
    "Đặng Thùy Dung", "Bùi Anh Khoa", "Hoàng Ngọc Mai", "Phan Gia Huy",
    "Ngô Bảo Châu", "Đỗ Quỳnh Anh", "Lương Đức Trọng", "Trịnh Thảo Vy",
    "Cao Tiến Đạt", "Dương Phương Anh", "Mai Văn Hùng", "Tạ Đình Phong",
    "Phùng Khánh Linh", "Đinh Tuấn Kiệt", "Hồ Bích Ngọc", "Võ Hữu Nghĩa",
    "Nguyễn Thị Bích", "Trần Đình Trọng", "Lê Phương Thảo", "Phạm Văn Nam"
]

def run_batch_collection(group_id, subscription="P3", start_index=1, target_exams=20, master_excel=DEFAULT_MASTER_EXCEL, output_dir=DEFAULT_OUTPUT_DIR, brain_sources_dir=DEFAULT_BRAIN_SOURCES):
    print("=" * 70)
    print(f"KHỞI CHẠY TIẾN TRÌNH THU THẬP ĐỀ THI TỰ ĐỘNG - KÍP {subscription}-{group_id}")
    print(f"Lộ trình: Thu thập từ Đề {start_index:02d} đến Đề {target_exams:02d} (Mục tiêu: {target_exams} đề)")
    print(f"Bảng tính Master Bank: {master_excel}")
    print(f"Thư mục lưu trữ: {output_dir}")
    print("=" * 70)
    
    start_time = time.time()
    total_added_all = 0
    total_dup_all = 0
    collected_files = []
    
    checkpoint_file = os.path.join(output_dir, f"tien_do_thu_thap_{subscription}_{group_id}_20_de.json")
    
    for i in range(start_index, target_exams + 1):
        name = random.choice(CANDIDATE_NAMES)
        print(f"\n>>>>>>>> BẮT ĐẦU THU THẬP ĐỀ THỨ {i:02d}/{target_exams:02d} (Thí sinh: {name}) <<<<<<<<")
        
        try:
            results, out_excel = fetch_and_solve_exam(
                group_id=group_id,
                subscription=subscription,
                candidate_name=name,
                master_cache_file=master_excel,
                output_dir=output_dir,
                brain_sources_dir=brain_sources_dir,
                exam_index=i
            )
            
            collected_files.append(out_excel)
            
            # Aggregate to Master Bank
            print(f"\n[Gộp & Khử Trùng] Nạp đề thứ {i:02d} vào Master Bank...")
            added, dup = aggregate_new_questions(
                source_excel=out_excel,
                master_excel=master_excel,
                source_label=f"Kíp {subscription}-{group_id} (Lần {i:02d})"
            )
            total_added_all += added
            total_dup_all += dup
            
            # Save checkpoint
            elapsed = round(time.time() - start_time, 1)
            checkpoint_data = {
                "group_id": group_id,
                "subscription": subscription,
                "target_exams": target_exams,
                "completed_exams": i,
                "current_progress_percent": round(i * 100 / target_exams, 1),
                "total_new_questions_added": total_added_all,
                "total_duplicates_filtered": total_dup_all,
                "elapsed_seconds": elapsed,
                "latest_excel_file": out_excel,
                "last_updated": time.strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(checkpoint_file, 'w', encoding='utf-8') as f_ckpt:
                json.dump(checkpoint_data, f_ckpt, ensure_ascii=False, indent=2)
                
            print(f"-> Đã hoàn tất Đề {i:02d}/{target_exams:02d}: Thêm mới {added} câu, trùng {dup} câu.")
            
        except Exception as e:
            print(f"-> CẢNH BÁO LỖI tại Đề {i:02d}: {e}. Tiếp tục sau 15 giây...")
            time.sleep(15)
            continue
            
        if i < target_exams:
            cooldown = random.uniform(10.0, 15.0)
            print(f"-> Giãn cách an toàn giữa các đề: nghỉ {round(cooldown, 1)}s trước khi mở đề tiếp theo...")
            time.sleep(cooldown)
            
    total_time = round((time.time() - start_time) / 60, 2)
    print("\n" + "=" * 70)
    print(f"HOÀN TẤT TOÀN BỘ {target_exams} ĐỀ THI!")
    print(f"- Tổng thời gian thực thi: {total_time} phút")
    print(f"- Tổng câu hỏi mới nạp vào Master Bank: {total_added_all} câu")
    print(f"- Tổng câu hỏi trùng lặp đã nhận diện: {total_dup_all} câu")
    print(f"- Bảng tính Master Bank cập nhật tại: {master_excel}")
    print("=" * 70)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Thu thập hàng loạt đề thi và gộp Master Bank")
    parser.add_argument("--group-id", type=int, default=39917, help="Mã kíp thi (vd 39917)")
    parser.add_argument("--sub", default="P3", help="Mã subscription (mặc định P3)")
    parser.add_argument("--start", type=int, default=1, help="Số thứ tự đề bắt đầu (mặc định 1)")
    parser.add_argument("--target", type=int, default=20, help="Số lượng đề cần thu thập (mặc định 20)")
    parser.add_argument("--master", default=DEFAULT_MASTER_EXCEL, help="Đường dẫn file Master Bank")
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, help="Thư mục lưu kết quả")
    args = parser.parse_args()
    
    run_batch_collection(
        group_id=args.group_id,
        subscription=args.sub,
        start_index=args.start,
        target_exams=args.target,
        master_excel=args.master,
        output_dir=args.output_dir
    )
