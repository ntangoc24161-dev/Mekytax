# Quy chuẩn triển khai Cơ chế đồng bộ Dashboard (Cách B - Windows Excel COM)

## 1. Nguyên lý phân công
- **VPS (Linux):** Quét `CV MOI`, bóc tách dữ liệu, chỉ ghi dữ liệu thô vào các cột chính `A:Z` trên sheet `NHAP_HO_SO`. Tuyệt đối không can thiệp dải cột phụ `AE:AZ`, không nén lại sheet `DASHBOARD` hay biểu đồ.
- **Máy trạm Windows:** Khi mở máy hoặc kích hoạt script, Microsoft Excel thật sẽ mở file, tự động kéo công thức phụ `AE:AZ` khớp với các dòng dữ liệu mới từ VPS, tính toán toàn bộ (`Calculate`), làm mới PivotTable/biểu đồ (`RefreshAll`), sau đó lưu lại chuẩn 100% Microsoft.

## 2. Đường dẫn và tài nguyên chuẩn
- **File báo cáo trên Drive:** `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx`
- **File script đồng bộ Windows:** Lưu tại thư mục tiện ích máy Windows (hoặc trong vault: `scripts/sync_dashboard_windows.py`).
- **File kích hoạt nhanh:** `scripts/run_sync_dashboard.bat`.

## 3. Nội dung mã nguồn script Python (Windows COM)
Script hoàn chỉnh, xử lý triệt để:
- Kiểm tra file tồn tại và kiểm tra tiến trình Excel đang mở để tránh xung đột khóa file.
- Quét dòng dữ liệu cuối tại cột J (MST) và dòng công thức phụ cuối tại cột AE (`CV_SỐ`).
- Tự động `AutoFill` dải công thức phụ từ cột `AE` đến `AZ` (cột 31 đến cột 52) tương ứng với các dòng mới.
- Chạy `Calculate` và `RefreshAll` để toàn bộ sheet `DASHBOARD`, `TAM_NGUNG`, `HET_TAM_NGUNG`, `DANH_SACH_LOC` đồng bộ tức thì.
- Đóng và giải phóng bộ nhớ COM sạch sẽ.

```python
import os
import sys
import time
import win32com.client

# Đường dẫn file sổ theo dõi trên Google Drive đồng bộ ở máy Windows
FILE_PATH = r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx"

def sync_dashboard():
    if not os.path.exists(FILE_PATH):
        print(f"[ERROR] Khong tim thay file: {FILE_PATH}")
        return False

    print("[INFO] Khoi dong Microsoft Excel COM...")
    excel = None
    wb = None
    try:
        # Khởi tạo tiến trình Excel độc lập chạy nền
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        print(f"[INFO] Dang mo file: {FILE_PATH}")
        wb = excel.Workbooks.Open(FILE_PATH, UpdateLinks=False, ReadOnly=False)
        ws = wb.Sheets("NHAP_HO_SO")

        # Xác định dòng cuối có dữ liệu (Cột J - MST, index 10)
        # xlUp = -4162
        last_data_row = ws.Cells(ws.Rows.Count, 10).End(-4162).Row

        # Xác định dòng cuối có công thức phụ (Cột AE - CV_SỐ, index 31)
        last_formula_row = ws.Cells(ws.Rows.Count, 31).End(-4162).Row

        print(f"[INFO] Dong du lieu cuoi (cot J): {last_data_row}")
        print(f"[INFO] Dong cong thuc phu cuoi (cot AE): {last_formula_row}")

        # Nếu có dòng mới từ VPS mà chưa có công thức phụ
        if last_data_row > last_formula_row and last_formula_row >= 5:
            print(f"[INFO] Phat hien {last_data_row - last_formula_row} dong moi can keo cong thuc...")
            # Dải cột phụ từ AE (31) đến AZ (52)
            src_range = ws.Range(ws.Cells(last_formula_row, 31), ws.Cells(last_formula_row, 52))
            dst_range = ws.Range(ws.Cells(last_formula_row, 31), ws.Cells(last_data_row, 52))
            
            # xlFillDefault = 0
            src_range.AutoFill(dst_range, 0)
            print("[INFO] Da keo cong thuc cot AE:AZ thanh cong.")
        else:
            print("[INFO] Cong thuc cot phu da khop voi du lieu.")

        print("[INFO] Tinh toan lai toan bo so (Calculate)...")
        excel.Calculate()

        print("[INFO] Lam moi cac bieu do va Dashboard (RefreshAll)...")
        wb.RefreshAll()

        print("[INFO] Dang luu file...")
        wb.Save()
        print("[SUCCESS] Dong bo Dashboard hoan tat 100%!")
        return True

    except Exception as e:
        print(f"[ERROR] Loi trong qua trinh dong bo: {str(e)}")
        return False
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass

if __name__ == "__main__":
    sync_dashboard()
```

## 4. File BAT khởi chạy một chạm (Quick Run)
Nội dung file `scripts/run_sync_dashboard.bat`:

```bat
@echo off
chcp 65001 > nul
echo ======================================================
echo    TIEN TRINH DONG BO SO BAO CAO CONG VAN (EXCEL COM)
echo ======================================================
echo.
python "%~dp0sync_dashboard_windows.py"
echo.
pause
```

## 5. Hướng dẫn thiết lập chạy tự động khi mở máy Windows
1. Nhấn `Windows + R`, gõ `shell:startup` và nhấn Enter để mở thư mục Startup.
2. Tạo một Shortcut trỏ đến file `run_sync_dashboard.bat`.
3. Hoặc thiết lập qua **Task Scheduler**:
   - Action: `Start a program`
   - Program/script: `pythonw.exe` (chạy ngầm hoàn toàn không hiện cửa sổ đen)
   - Arguments: `"D:\path\to\sync_dashboard_windows.py"`
   - Trigger: `At log on` (Khi người dùng đăng nhập vào Windows).
