---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
source: chat 2026-11-13 (3bfe9e24)
---
Quy tắc vận hành sổ `Danh sach kiem tra quyet toan pro.xlsm`:
- Ngay khi hoàn tất tạo/xuất bộ hồ sơ đi (trình ký), vụ việc được tự động chuyển từ sheet `02_DANG_THUC_HIEN` sang sheet `01_DA_BAN_HANH` (gán trạng thái `1. Đã ban hành`).
- Đồng thời tự động di chuyển thư mục của công ty sang thư mục `1. TD` (đặt STT nối tiếp, ví dụ `1. TD\<STT>. <TÊN CÔNG TY>\`).
- Không cần chờ đến khi có Số QĐKT chính thức mới chuyển sheet và chuyển thư mục.
- Khi có số QĐKT thật sau này, chỉ cần cập nhật số vào dòng đã nằm sẵn ở sheet `01_DA_BAN_HANH`.
