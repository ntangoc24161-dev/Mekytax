"""
run_continuous_p3_39917.py - Supervisor tự động kích hoạt và chạy liên tục tiến trình thu thập & giải đề kíp P3-39917.
- Tự động chạy lại ngay khi tiến trình con bị tắt hoặc gặp sự cố mạng / API.
- Tự động đối soát và lưu cuốn chiếu từng câu vào file Excel và checkpoint.
- Chạy liên tục cho đến khi thu thập đủ toàn bộ 40 đề thi (mỗi đề 90 câu).
"""
import os
import sys
import time
import subprocess
import json
import glob
import re
import datetime
from pathlib import Path

# Ensure UTF-8 output on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

PYTHON_EXE = sys.executable
TARGET_DIR = r"D:\HOC TAP\ON THI 2026.09"
EXAMS_DIR = os.path.join(TARGET_DIR, "40 DE THI")
TARGET_SCRIPT = os.path.join(TARGET_DIR, "collector_40_de.py")
CHECKPOINT_FILE = os.path.join(TARGET_DIR, "checkpoint_P3_39917.json")
SUPERVISOR_LOG = os.path.join(TARGET_DIR, "supervisor_p3_39917.log")

def log(msg):
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{now_str}] [SUPERVISOR P3-39917] {msg}"
    print(formatted, flush=True)
    try:
        with open(SUPERVISOR_LOG, "a", encoding="utf-8") as f:
            f.write(formatted + "\n")
    except Exception:
        pass

def count_completed_exams():
    if not os.path.exists(EXAMS_DIR):
        return 0, 0, {}
    
    pattern = re.compile(r'Bo_De_On_Thi_Kiem_Tra_Thue_.*_De_(\d+)\.xlsx')
    completed_90 = 0
    total_exams = 0
    details = {}
    
    for f in glob.glob(os.path.join(EXAMS_DIR, "Bo_De_On_Thi_Kiem_Tra_Thue_*_De_*.xlsx")):
        if os.path.basename(f).startswith("~$"):
            continue
        m = pattern.search(os.path.basename(f))
        if not m:
            continue
        num = int(m.group(1))
        total_exams += 1
        # Check rows without loading full openpyxl to be fast
        try:
            import openpyxl
            wb = openpyxl.load_workbook(f, data_only=True)
            ws = wb.active
            cnt = ws.max_row - 1
            details[num] = cnt
            if cnt >= 90:
                completed_90 += 1
        except Exception:
            pass
            
    return completed_90, total_exams, details

def run_loop():
    log("=" * 75)
    log("KHỞI ĐỘNG BỘ GIÁM SÁT TỰ ĐỘNG CHẠY LIÊN TỤC KÍP P3-39917 (MỤC TIÊU: 40 ĐỀ)")
    log(f"Script thực thi: {TARGET_SCRIPT}")
    log(f"Thư mục lưu trữ: {EXAMS_DIR}")
    log("=" * 75)

    iteration = 0

    while True:
        completed_90, total_exams, details = count_completed_exams()
        log(f"--- CHU KỲ GIÁM SÁT #{iteration + 1} ---")
        log(f"Tiến độ hiện tại: {completed_90}/40 đề đã đủ 90 câu. Tổng file: {total_exams} đề.")
        
        if completed_90 >= 40:
            log(" CHÚC MỪNG! ĐÃ THU THẬP VÀ GIẢI ĐỦ TOÀN BỘ 40/40 ĐỀ THI KÍP P3-39917!")
            break

        log(f"Bắt đầu khởi chạy tiến trình collector_40_de.py...")
        t0 = time.time()
        
        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            proc = subprocess.run(
                [PYTHON_EXE, TARGET_SCRIPT, "--max-exams", "40"],
                cwd=TARGET_DIR,
                env=env
            )
            elapsed = time.time() - t0
            log(f"Tiến trình con kết thúc (Exit Code: {proc.returncode}) sau {elapsed:.1f}s ({elapsed/60:.1f} phút).")
        except Exception as e:
            log(f"[Lỗi thực thi supervisor]: {e}")

        completed_90, total_exams, details = count_completed_exams()
        if completed_90 >= 40:
            log(" ĐÃ HOÀN TẤT ĐỦ 40/40 ĐỀ THI KÍP P3-39917!")
            break

        iteration += 1
        log(f"Tiến trình vừa tạm dừng. Tự động khôi phục và tiếp tục ngay sau 3 giây...")
        time.sleep(3)

    log("Supervisor P3-39917 kết thúc phiên.")

if __name__ == "__main__":
    run_loop()
