---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-04
updated: 2026-10-04
source: [[conversations/2026-10-04]]
---
Trong công tác soạn thảo hồ sơ kiểm tra thuế bằng Word/python-docx, các Quyết định trình ký độc lập (Quyết định kiểm tra, Quyết định giám sát) phải được bố trí gọn trong đúng 1 trang A4 [[conversations/2026-10-04]]. Để tránh tràn chữ ký sang trang 2, căn cứ pháp lý được đặt cỡ chữ 11.5 pt, điều khoản 12 pt và khoảng cách dòng 1.0 line spacing [[conversations/2026-10-04]]. Khung ký hiệu mẫu biểu ở góc trên bên phải cần dùng bảng căn phải dạng inline (thay vì floating table) để không bị đẩy sang trang riêng lẻ [[conversations/2026-10-04]]. Đối với file tổng hợp nhiều mẫu, cần loại bỏ các đoạn văn rỗng trước và sau ngắt trang (Page Break) và thiết lập khóa chống vỡ dòng bảng (`cantSplit`) để triệt tiêu các trang trắng thừa [[conversations/2026-10-04]].
