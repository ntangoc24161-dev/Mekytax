"""
watchdog_priority_process.py
Giám sát và tự động kích hoạt lại tiến trình bóc tách công văn cảnh báo (Ưu tiên 1, Ưu tiên 2, nạp Excel)
khi bị tạm dừng đột ngột, crash, hoặc bị tắt tiến trình cho đến khi hoàn thành toàn bộ 3 giai đoạn.
"""

import os
import sys
import time
import subprocess
import datetime
from pathlib import Path

os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

PYTHON_EXE = r"D:\SECOND BRAIN\javis-os\.venv-paddleocr\Scripts\python.exe"
SCRIPT_DIR = Path(__file__).resolve().parent
TARGET_SCRIPT = SCRIPT_DIR / "run_priority_1_and_2.py"
LOG_FILE = SCRIPT_DIR / "watchdog.log"
COMPLETION_MARKER = SCRIPT_DIR / "COMPLETED_PRIORITY_1_AND_2.txt"

def safe_print(msg):
    try:
        print(msg, flush=True)
    except Exception:
        try:
            print(msg.encode('ascii', errors='replace').decode('ascii'), flush=True)
        except Exception:
            pass

def log_msg(msg):
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{now_str}] {msg}"
    safe_print(formatted)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(formatted + "\n")
    except Exception:
        pass

def is_fully_completed():
    return COMPLETION_MARKER.exists()

def run_worker_and_monitor():
    log_msg(f"Khởi chạy tiến trình chính: {TARGET_SCRIPT}")
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    
    with open(SCRIPT_DIR / "run.log", "a", encoding="utf-8") as run_out:
        proc = subprocess.Popen(
            [PYTHON_EXE, "-u", str(TARGET_SCRIPT)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            cwd=str(SCRIPT_DIR.parents[2]),
            env=env,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        log_msg(f"Tiến trình đang chạy với PID: {proc.pid}")
        
        for line in iter(proc.stdout.readline, ''):
            if line:
                clean_line = line.strip()
                if clean_line:
                    run_out.write(line)
                    run_out.flush()
                    if "[OCR OK]" in clean_line or "[Cache hit]" in clean_line or "[Tái sử dụng SCV]" in clean_line or "[Digital OK]" in clean_line or ">>>" in clean_line or "[OK]" in clean_line or "[HOÀN TẤT]" in clean_line or "KẾT QUẢ" in clean_line:
                        log_msg(f"  -> {clean_line}")
        
        proc.stdout.close()
        ret_code = proc.wait()
        return ret_code

def main():
    log_msg("=" * 60)
    log_msg("KHỞI ĐỘNG HỆ THỐNG GIÁM SÁT TỰ ĐỘNG (WATCHDOG SUPERVISOR)")
    log_msg("Tự động khôi phục & tiếp tục tiến trình đến khi hoàn thành trọn vẹn 3 giai đoạn")
    log_msg("=" * 60)

    retry_count = 0
    max_retries = 50

    while retry_count < max_retries:
        if is_fully_completed():
            log_msg(">>> TOÀN BỘ 3 GIAI ĐOẠN ĐÃ HOÀN TẤT TRỌN VẸN. WATCHDOG DỪNG GIÁM SÁT. <<<")
            break

        retry_count += 1
        log_msg(f"Bắt đầu phiên xử lý #{retry_count}...")
        try:
            ret = run_worker_and_monitor()
            log_msg(f"Tiến trình worker kết thúc với mã trả về: {ret}")
        except Exception as e:
            log_msg(f"Ngoại lệ khi thực thi worker: {e}")

        if is_fully_completed():
            log_msg(">>> XÁC NHẬN HOÀN TẤT THÀNH CÔNG TẤT CẢ 3 GIAI ĐOẠN! <<<")
            break
        else:
            log_msg("[CẢNH BÁO] Tiến trình chưa hoàn tất 3 giai đoạn. Tự động phục hồi lại sau 5 giây...")
            time.sleep(5)

if __name__ == "__main__":
    main()
