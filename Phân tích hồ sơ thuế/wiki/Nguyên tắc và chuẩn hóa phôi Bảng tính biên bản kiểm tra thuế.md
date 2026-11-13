---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
source: [[conversations/2026-11-13]]
---
Phôi `Bang tinh.xlsx` là bảng tính biên bản của cơ quan thuế phục vụ thanh kiểm tra thuế và hồ sơ giải thể tại bàn, thuộc nhóm tài liệu nghiệp vụ biên bản/giải trình `3.4` hoặc dữ liệu CQT `1.3`, không thuộc danh mục sổ sách nội bộ doanh nghiệp `IN02–IN13` [[conversations/2026-11-13]].

### 1. Cấu trúc các sheet chuẩn
- `BB`: Biên bản tổng hợp kết quả thanh kiểm tra [[conversations/2026-11-13]].
- `KQKDSXKD` và `CĐKT`: Bảng kết quả kinh doanh và cân đối kế toán dùng để đối chiếu số kê khai so với số kiểm tra qua các năm thực tế [[conversations/2026-11-13]].
- `GTGT`: Bảng đối chiếu số thuế GTGT đầu ra, đầu vào và số phải nộp [[conversations/2026-11-13]].
- `BANG KE KO HOP LE` (1 & 2): Danh mục chi tiết hóa đơn không hợp lệ cần loại trừ khỏi chi phí hoặc thuế GTGT [[conversations/2026-11-13]].
- `TINH PHAT + CHAM NOP`: Công thức tự động tính tiền phạt VPHC và tiền chậm nộp theo thời điểm phát sinh [[conversations/2026-11-13]].

### 2. Nguyên tắc bóc tách và nạp số liệu
- Chỉ lập và bóc tách các năm phát sinh thực tế; tuyệt đối không tạo số liệu hoặc điền năm trước ngày doanh nghiệp thành lập [[conversations/2026-11-13]].
- Khi phôi ở trạng thái khung mẫu chưa có số liệu, bắt buộc bóc tách từ XML tờ khai (GTGT, QTT TNDN, BCTC), không điền số 0 giả [[conversations/2026-11-13]].
- Khung mẫu chuẩn trước khi khởi tạo vụ mới phải đảm bảo 0 lỗi `#REF!`, xóa sạch các cột tính nháp ngoài rìa và xóa bỏ toàn bộ dữ liệu mẫu về lao động, số liệu hóa đơn đối tác cũ [[conversations/2026-11-13]].
