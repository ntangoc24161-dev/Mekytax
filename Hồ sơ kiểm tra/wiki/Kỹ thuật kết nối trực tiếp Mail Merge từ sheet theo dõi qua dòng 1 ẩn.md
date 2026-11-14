---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
source: [[conversations/2026-11-13]]
---
Để Word Mail Merge kết nối trực tiếp với bảng tính theo dõi có tiêu đề nhiều dòng tiếng Việt mà không cần sheet đệm trung gian, giải pháp chuẩn là chèn Dòng 1 ẩn chứa mã trường không dấu tiếng Anh/ASCII [[conversations/2026-11-13]].

Word Mail Merge luôn mặc định lấy đúng hàng đầu tiên của sheet làm tên trường dữ liệu (Field Names) [[conversations/2026-11-13]]. Khi ẩn Dòng 1, giao diện hiển thị cho người dùng từ Dòng 2 trở đi vẫn giữ nguyên bố cục nghiệp vụ, màu sắc và tiêu đề tiếng Việt gốc (thực tế tính đến 13/11/2026) [[conversations/2026-11-13]].

Cơ chế điền cột chữ tiền tự tạo qua VBA `=DocSo(...)` cần được script tự động gán công thức tương ứng theo hàng mỗi khi thêm mới hoặc cập nhật dòng dữ liệu để tránh thiếu sót dữ liệu trộn sang phôi Word [[conversations/2026-11-13]].
