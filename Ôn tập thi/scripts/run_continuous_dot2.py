"""
run_continuous_dot2.py - Supervisor tự động chạy liên tục tiến trình bóc tách đề thi Đợt 2 cho đến khi hoàn thành 100% (đủ 30 đề, 1800 câu hỏi).
Tự động phục hồi khi gặp sự cố, tự động lưu checkpoint và cập nhật Excel Master + Brain.
"""
import os
import sys
import time
import subprocess
import json
import datetime
from pathlib import Path

# Ensure UTF-8 output
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parents[2]
PYTHON_EXE = PROJECT_ROOT / ".venv-paddleocr" / "Scripts" / "python.exe"
TARGET_SCRIPT = SCRIPT_DIR / "batch_processor_dot2.py"
CHECKPOINT_FILE = Path(r"D:\HOC TAP\ON THI 2026.09\dot2_checkpoint.json")

def log(msg):
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now_str}] [SUPERVISOR] {msg}", flush=True)

def is_fully_completed():
    if not CHECKPOINT_FILE.exists():
        return False, 0
    try:
        with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
            ckpt = json.load(f)
        done_tests = len([k for k, v in ckpt.get("processed_tests", {}).items() if v.get("status") == "completed"])
        total_q = len(ckpt.get("questions", {}))
        return done_tests >= 30, total_q
    except Exception:
        return False, 0

def run_loop():
    log("Khởi động Supervisor chạy liên tục bóc tách 30 đề thi...")
    log(f"Python interpreter: {PYTHON_EXE}")
    log(f"Target script: {TARGET_SCRIPT}")

    retry_count = 0
    max_retries = 100

    while retry_count < max_retries:
        completed, total_q = is_fully_completed()
        if completed:
            log(f" Tuyệt vời! Đã thu thập đủ toàn bộ 30/30 đề thi ({total_q} câu hỏi).")
            break

        log(f"Bắt đầu / Tiếp tục chu kỳ xử lý (Hiện tại: {total_q} câu hỏi)...")
        t0 = time.time()

        try:
            # Chạy batch_processor_dot2.py
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            proc = subprocess.run(
                [str(PYTHON_EXE), str(TARGET_SCRIPT)],
                cwd=str(SCRIPT_DIR),
                env=env
            )
            elapsed = time.time() - t0
            log(f"Tiến trình con kết thúc với mã trả về: {proc.returncode} sau {elapsed:.1f}s")
        except Exception as e:
            log(f"Lỗi khi thực thi tiến trình: {e}")

        completed, total_q = is_fully_completed()
        if completed:
            log(f" Hoàn thành toàn bộ 30/30 đề thi ({total_q} câu hỏi).")
            break

        retry_count += 1
        log(f"Chưa hoàn tất đủ 30 đề. Tự động khởi động lại sau 3 giây (Lần thử {retry_count}/{max_retries})...")
        time.sleep(3)

    log("Supervisor kết thúc phiên làm việc.")

if __name__ == "__main__":
    run_loop()
