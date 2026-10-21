---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-09-30
updated: 2026-09-30
source: [[conversations/2026-09-30]]
---
Hiện tượng file Excel bị lỗi 'Removed Records: Formula from /xl/worksheets/sheet1.xml part' khi mở thường xảy ra do một ô văn bản bắt đầu bằng ký tự dấu bằng (`=`), ví dụ như ký hiệu suy ra `=> KẾT LUẬN...` [[conversations/2026-09-30]]. Khi đó trình tạo file tự gán kiểu dữ liệu của ô là Formula và lưu vào thẻ `<f>` trong XML thay vì thẻ chuỗi `<t>` [[conversations/2026-09-30]]. Khi Excel mở, parser không thấy hàm toán học hợp lệ sẽ xóa bỏ công thức lỗi và hiện hộp thoại phục hồi file (thực tế tính đến 2026-09-30) [[conversations/2026-09-30]]. Cách khắc phục là thay thế ký tự đầu tiên (ví dụ đổi `=>` thành `>`) hoặc ép kiểu chuỗi tường minh cho ô dữ liệu [[conversations/2026-09-30]].
