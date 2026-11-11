---
type: business
provenance: source
origin: javis-learned
created: 2026-11-09
updated: 2026-11-09
source: chat 2026-11-09 (5001691f)
---
Công ty TNHH Môi trường Đô thị Thành Sen (MST 3603403068), thư mục `Ke hoach nam\Bo sung\THANH_SEN_3603403068\0. BIEN BAN\Bang tinh.xlsx`. Kỳ header giữ **2024-2025**. Sổ `Danh sach kiem tra quyet toan pro.xlsm` chưa có MST này.

Đã đổ `KQKDSXKD` (backup `Bang tinh_backup_before_nhap_20261109_035039.xlsx`):
- 2024 từ `PTDN02_20261106152005.xlsx`. Doanh thu thuần mã 10 = 9.888.903.035 (khớp BCTC `NamTruoc`). Giá vốn 8.371.486.703, giữ PTDN (BCTC `NamTruoc` là 8.360.368.554). CP QLDN 1.210.134.467. LN trước thuế 18.841.600.
- 2025 PTDN toàn 0 nên bù `3603403068000-200_BCTC-01012025_31122025_VND-L00.xml` khối `NamNay`, không lấy `NamTruoc`. Doanh thu thuần 21.161.811.599. Giá vốn 20.115.352.946. CP QLDN 680.425.850. LN trước thuế -65.046.352.

Đã đổ CDKT (backup `Bang tinh_backup_before_cdkt_20261109_035442.xlsx`):
- 2024 từ `PTDN01_20261106151913.xlsx` (TT200 map TT133). Tổng tài sản = NV = 25.512.302.119.
- 2025 PTDN01 toàn 0 nên bù BCTC `SoCuoiNam`. Tổng tài sản = NV = 37.260.675.631. Sheet `CDKT` cột O sẵn có 2025.
- Sheet `CĐKT ` (có dấu, thừa dấu cách) vốn không có cột 2025. User bảo tạo cột: đã chèn cột P = 2025 sau O = 2024, copy merge 2 dòng, đổ cùng số CDKT/BCTC. Backup `Bang tinh_backup_before_cdkt_col2025_20261109_035838.xlsx`. Cân đối lệch 0. Định dạng `#,##0`, không `#VALUE!`.

Ngày 09/11/2026 NNT bổ sung hồ sơ trong `GỬI THUẾ` (2024-2025). Đã chuyển 32 file vào `2.1` (24), `2.2` (4), `2.4` (4), xóa `GỬI THUẾ` sau đối soát size. Checklist cập nhật tại `4. KET QUA PHAN TICH\CHECKLIST_TONG_HOP_HO_SO.xlsx`. Còn thiếu: XNT/giá thành chi tiết, sao kê NH+UNC, bảng lương/chấm công/HĐLĐ, bảng kê NNT, CVCBRR, danh bạ ĐKKD. HĐĐT dump lớn còn ở `HDDT TAI CHUNG`.

Chưa có `KQ2.1_*_FINAL` nên chưa đổ GTGT/EBITDA. Chưa bóc tách DI-01.
