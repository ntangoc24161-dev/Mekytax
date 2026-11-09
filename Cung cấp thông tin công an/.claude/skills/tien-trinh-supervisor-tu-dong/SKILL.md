---
name: Tiến trình supervisor tự động
description: "Supervisor tự động điều phối, giám sát chu trình giải đề khảo thí cuốn chiếu, tự phục hồi khi gián đoạn mạng và kiểm kê đến khi đủ mục tiêu."
group: Năng suất
---

# Tiến trình supervisor tự động - Điều phối và tự phục hồi giải đề cuốn chiếu

Kỹ năng điều phối kiến trúc **Bộ giám sát (Supervisor) kết hợp Bộ thực thi (Worker)**, vận hành hoàn toàn offline nhằm tự động hóa thu thập, bóc tách và giải đề khảo thí cuốn chiếu. Supervisor liên tục kiểm kê tiến độ thực tế từ file bảng tính, quản lý vòng lặp thực thi, tự động bắt lỗi và tái kích hoạt worker khi gặp sự cố mạng hoặc timeout cho đến khi hoàn thành 100% mục tiêu đề thi.

## Khi nào sử dụng skill

- Người dùng yêu cầu: "chạy supervisor tự động", "chạy liên tục đến khi đủ 40 đề", "tự động khôi phục khi gián đoạn", "kiểm kê tiến độ và giải tiếp", "supervisor giải đề khảo thí", "tự động hóa thu thập đề thi chạy nền".
- Cần thu thập và giải một khối lượng lớn đề thi trắc nghiệm (ví dụ 40 đề, 3.600 câu) trên các cổng khảo thí trực tuyến (`taxedu.gdt.gov.vn`, `api-e3.khaothi.online`).
- Môi trường mạng hoặc máy chủ khảo thí thiếu ổn định, thường xuyên bị ngắt kết nối giữa chừng, cần cơ chế tự phục hồi (Self-Healing) mà không làm mất mát các câu hỏi đã thu thập trước đó.
- Cần cơ chế kiểm kê dữ liệu thực tế: đọc trực tiếp số dòng từ file Excel (`ws.max_row - 1`) để xác định chính xác số đề đã đủ 90/90 câu, không phụ thuộc vào bộ nhớ tạm.

## Chuẩn bị

1. Thư mục làm việc và dữ liệu học tập: `D:\HOC TAP\ON THI 2026.09`.
2. Script Supervisor điều phối chính: `D:\HOC TAP\ON THI 2026.09\run_continuous_p3_39917.py`.
3. Script Worker thực thi giải đề: `D:\HOC TAP\ON THI 2026.09\collector_40_de.py`.
4. Thư mục lưu trữ bộ đề thi: `D:\HOC TAP\ON THI 2026.09\40 DE THI`.
5. Tệp nhật ký Supervisor: `D:\HOC TAP\ON THI 2026.09\supervisor_p3_39917.log`.
6. Tệp checkpoint đồng bộ tiến độ: `D:\HOC TAP\ON THI 2026.09\checkpoint_P3_39917.json`.
7. Môi trường thực thi: Python 3 (thư viện `openpyxl`, `subprocess`, `glob`, `re`, `json`).

## Cách chạy

### 1. Khởi chạy Supervisor giám sát liên tục (Khuyên dùng)

Mở terminal hoặc PowerShell và chạy:

```powershell
python "D:\HOC TAP\ON THI 2026.09\run_continuous_p3_39917.py"
```

Supervisor sẽ tự động:
- Kiểm tra toàn bộ thư mục `40 DE THI` để đếm số đề đã hoàn tất 90/90 câu.
- Khởi động worker `collector_40_de.py` chạy ngầm.
- Theo dõi vòng đời của worker; nếu worker bị ngắt do mạng hoặc lỗi máy chủ, Supervisor ghi log và tự động bật lại sau 3 giây.
- Tiếp tục giải cuốn chiếu các câu còn thiếu của đề dang dở hoặc bốc đề mới cho đến khi hoàn tất đủ 40/40 đề.

### 2. Tùy biến mục tiêu số lượng đề

Có thể điều chỉnh biến `TARGET_TOTAL` trong file supervisor hoặc truyền tham số tương ứng để đặt mục tiêu 10, 20, 30 hoặc 40 đề tùy theo từng đợt ôn thi.

## Quy trình 5 bước thực thi chi tiết

### Bước 1: Quét và kiểm kê tiến độ thực tế (Inventory & Count)
- Truy cập thư mục lưu trữ `40 DE THI`.
- Sử dụng biểu thức chính quy `Bo_De_On_Thi_Kiem_Tra_Thue_.*_De_(\d+)\.xlsx` để quét tất cả các file Excel hiện hữu.
- Bỏ qua các file tạm của hệ thống bắt đầu bằng `~$`.
- Mở từng file bằng `openpyxl` với chế độ `data_only=True`, đếm số dòng thực tế (`ws.max_row - 1`).
- Nếu số dòng `>= 90`, tính là 01 đề hoàn tất (`completed_90 += 1`).

### Bước 2: Kích hoạt và quản lý tiến trình con Worker
- Nếu `completed_90 < TARGET_TOTAL`, Supervisor thiết lập môi trường `PYTHONIOENCODING=utf-8`.
- Khởi chạy subprocess gọi `collector_40_de.py --max-exams TARGET_TOTAL`.
- Ghi nhận thời điểm bắt đầu (`t0`) và mở luồng ghi log song song ra console và file `supervisor_p3_39917.log`.

### Bước 3: Điều phối chu trình Worker giải đề cuốn chiếu
- Worker thực hiện đăng ký danh tính ảo, giải mã payload và đối chiếu nhận diện mẫu đề.
- Nếu trùng đề cũ chưa hoàn tất: tải lại dữ liệu đã có và tiếp tục giải từ câu còn thiếu đến câu 90.
- Nếu trùng đề cũ đã đủ 90/90 câu: tự động bỏ qua (auto-skip) để bốc đề mới.
- Giải đề qua 2 tầng (Ngân hàng Master + AI Solver) và dò đáp án chuẩn 100% từ máy chủ với giãn cách an toàn 2.0s - 4.0s.
- Ghi cuốn chiếu liên tục sau từng câu hỏi vào file Excel và cập nhật checkpoint.

### Bước 4: Bắt lỗi và tự phục hồi sự cố (Self-Healing)
- Khi Worker kết thúc hoặc bị gián đoạn (timeout, rớt mạng, máy chủ ngắt socket), Supervisor thu nhận mã thoát `Exit Code` và thời gian chạy.
- Toàn bộ dữ liệu trước thời điểm gián đoạn đã được lưu cuốn chiếu an toàn trong file Excel nên không bị mất mát câu hỏi nào.

### Bước 5: Tự động lặp lại chu kỳ đến khi hoàn thành 100%
- Supervisor gọi lại hàm `count_completed_exams()` để cập nhật số lượng đề hoàn tất mới nhất.
- Nếu chưa đủ mục tiêu: nghỉ chờ **3 giây**, tăng số chu kỳ (`iteration += 1`), sau đó tự động tái khởi động Worker để tiếp tục làm bài.
- Nếu đã đủ 100% mục tiêu (`completed_90 >= TARGET_TOTAL`): Supervisor ghi nhận thông điệp hoàn thành thắng lợi và kết thúc tiến trình an toàn.

## Mẫu mã nguồn Supervisor chuẩn (Python)

```python
import os
import sys
import time
import subprocess
import glob
import re
import datetime

TARGET_DIR = r"D:\HOC TAP\ON THI 2026.09"
EXAMS_DIR = os.path.join(TARGET_DIR, "40 DE THI")
TARGET_SCRIPT = os.path.join(TARGET_DIR, "collector_40_de.py")
SUPERVISOR_LOG = os.path.join(TARGET_DIR, "supervisor_p3_39917.log")
TARGET_TOTAL = 40

def log(msg):
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{now_str}] [SUPERVISOR] {msg}"
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

def run_supervisor():
    iteration = 0
    while True:
        completed_90, total_exams, details = count_completed_exams()
        log(f"--- CHU KỲ #{iteration + 1} | Tiến độ: {completed_90}/{TARGET_TOTAL} đề đủ 90 câu ---")
        if completed_90 >= TARGET_TOTAL:
            log("ĐÃ HOÀN TẤT 100% MỤC TIÊU!")
            break
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        try:
            subprocess.run([sys.executable, TARGET_SCRIPT, "--max-exams", str(TARGET_TOTAL)], cwd=TARGET_DIR, env=env)
        except Exception as e:
            log(f"Lỗi thực thi: {e}")
        completed_90, total_exams, details = count_completed_exams()
        if completed_90 >= TARGET_TOTAL:
            break
        iteration += 1
        time.sleep(3)

if __name__ == "__main__":
    run_supervisor()
```

## Bẫy và Lưu ý khi vận hành

- **Bẫy file tạm Excel (`~$...`)**: Khi mở file Excel, ứng dụng tự tạo file lock bắt đầu bằng `~$`. Hàm kiểm kê bắt buộc phải lọc bỏ tiền tố này để tránh lỗi crash script.
- **Bẫy mã hóa Unicode Console trên Windows**: Luôn gán biến môi trường `PYTHONIOENCODING=utf-8` khi gọi subprocess để tránh lỗi ký tự tiếng Việt khi ghi log.
- **Bẫy xung đột khóa file**: Khi người dùng mở file Excel xem kết quả, lệnh ghi của Worker có thể bị `PermissionError`. Phải bọc khối try/except để bỏ qua lượt ghi đó mà không ngắt tiến trình.
- **Bảo toàn dữ liệu tuyệt đối**: Nhờ cơ chế cuốn chiếu từng câu, khi sự cố xảy ra, toàn bộ câu hỏi trước đó vẫn được lưu vẹn toàn.
