---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
source: chat 2026-11-13 (3bfe9e24)
---
**Quy tắc:** Khi hồ sơ kiểm tra chuyển sang trạng thái `1. Đã ban hành` (hoặc ngay khi hoàn tất tạo bộ hồ sơ đi trình ký), hệ thống đồng thời tự động di chuyển thư mục của công ty sang thư mục `1. TD` (đặt theo số thứ tự tiếp theo, ví dụ `1. TD\<STT>. <TÊN CÔNG TY>\`).

**Vì sao:** Thư mục `1. TD` là nơi lưu trữ các vụ việc chính thức đang theo dõi và kiểm tra tại trụ sở NNT, tách biệt khỏi thư mục chuẩn bị/kế hoạch ban đầu.

**Áp dụng:** Tự động thực hiện song song với việc chuyển dòng sang sheet `01_DA_BAN_HANH` trên sổ `Danh sach kiem tra quyet toan pro.xlsm`, không hỏi lại người dùng. Đã cập nhật vào các skill `lap-ho-so-kiem-tra` và `ho-so-kiem-tra-ke-hoach`.
