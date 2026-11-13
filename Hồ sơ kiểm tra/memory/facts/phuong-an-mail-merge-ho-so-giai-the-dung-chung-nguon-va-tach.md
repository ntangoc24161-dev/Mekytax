---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
---
Thống nhất phương án xử lý nhóm kiểm tra giải thể:
- Nguồn dữ liệu: Dùng chung sheet `01_DA_BAN_HANH$` của sổ `Danh sach kiem tra quyet toan pro.xlsm` qua Dòng 1 ẩn.
- Phân biệt nghiệp vụ: Cột `Loai_hinh` (cột 3) phân loại `KH` (Kế hoạch) và `GT` (Giải thể).
- Phôi Word: Tách 2 phôi riêng biệt thay vì lồng IF/THEN/ELSE phức tạp. Nhóm `KH` dùng `templates/MAUTONG HOP.docx`, nhóm `GT` dùng `templates/MAUTONG HOP GIAI THE.docx`.
- Điều phối: Tự động bắt mã `Loai_hinh` để chọn đúng phôi và gọi skill tương ứng (`ho-so-kiem-tra-ke-hoach` hoặc `ho-so-kiem-tra-giai-the`).
