---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
---
Sau khi xóa sheet đệm `DuLieuMerge` và dùng Dòng 1 ẩn trên `01_DA_BAN_HANH`, cần lưu ý 4 điểm vận hành:
1. **Bộ lọc trong Word:** Word quét toàn bộ hàng trên sheet gồm cả các dòng tiêu đề cũ (dòng 2-4) và dòng trống, do đó khi in/duyệt cần dùng bộ lọc Mailings chọn theo MST/Tên hoặc bỏ tích các dòng tiêu đề.
2. **Kéo công thức chữ:** Khi thêm doanh nghiệp mới, phải sao chép công thức hàm `=DocSo(...)` ở 6 cột cuối (BA đến BF, cột 53 đến 58) để không trống trường tiền bằng chữ.
3. **Phụ thuộc VBA:** Hàm `DocSo` nằm trong module VBA nên sổ theo dõi bắt buộc lưu `.xlsm` và cho phép macro chạy.
4. **Giới hạn trường đơn (1-1):** Mail Merge chỉ thế chỗ các trường đơn lẻ, các bảng số liệu chi tiết phân bổ nhiều năm từ `Bang tinh.xlsx` vẫn dựa vào khung dựng sẵn hoặc script điền tự động.
