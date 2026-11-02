---
type: user
provenance: user
origin: javis-learned
created: 2026-10-22
updated: 2026-10-22
---
IN02 là sheet `IN02_BangKe` (24 cột, header dòng 2) trong bộ biểu mẫu Drive `Bo_bieu_mau_v2_TT_BIEN_HOA_truocIN01IN02IN07.xlsx`. Dữ liệu Phát Hoàng Long (MST 3600974507) nằm ở Drive tài khoản ntangoc2416.1@gmail.com (connector googledrive_amala-ensign), thư mục `PHAT_HOANG_LONG_3600974507-chua so/2. HO SO NNT CUNG CAP/2.5. Bang ke mua vao ban ra/BANG KE MUA VAO BAN RA` gồm 5 file (bảng kê 2022 vào/ra, VAT 2023-2025).

Đã tạo `PHAT_HOANG_LONG_3600974507/4. KET QUA PHAN TICH/cau-thanh/IN02_BangKe_PHAT_HOANG_LONG_3600974507.xlsx` (5.263 dòng, 667 KB).

**Thực tế tính đến 22/10/2026:**
- User yêu cầu đưa IN02 lên Drive; Javis không upload được, user chọn **tự tải** từ trang Tệp tin rồi kéo vào thư mục đích `4. KET QUA PHAN TICH` của `PHAT_HOANG_LONG_3600974507-chua so` (https://drive.google.com/drive/folders/1jEn7zqL3Wafpjl-FD2IWPE4WF9vybGuE, tài khoản ntangoc2416.1).
- Cột `Trang_thai_DC` vẫn `CHUA_DOI_CHIEU`; bước tiếp theo đề xuất là đối chiếu với HĐĐT thư mục 1.2 (user chưa chốt).

**Áp dụng:** tải Drive từng file một (gọi song song dễ lỗi 403); khi lọc dòng Tổng cộng chỉ so khớp ô đúng "TỔNG CỘNG", không so chuỗi con (sẽ xóa nhầm HĐ "TỔNG CÔNG TY"); ngày có dạng dd/mm/yy; file 2023-2025 là danh sách HĐ cổng thuế nên quý/tháng tính theo ngày HĐ, thuế suất suy từ thuế/giá trị.
