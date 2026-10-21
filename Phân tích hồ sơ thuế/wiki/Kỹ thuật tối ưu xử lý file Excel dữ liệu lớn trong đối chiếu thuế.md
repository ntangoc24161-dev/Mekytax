---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-08
updated: 2026-10-08
source: [[conversations/2026-10-08]]
---
Khi xử lý dữ liệu kiểm tra thuế quy mô trên 10.000 dòng (HĐĐT, bảng kê), đọc từng ô qua `cell(row, col)` trong thư viện openpyxl gây suy giảm hiệu năng nghiêm trọng và dễ treo tiến trình [[conversations/2026-10-08]]. Sử dụng `iter_rows(values_only=True)` để đọc hàng loạt giúp giảm thời gian xử lý xuống còn khoảng 22 giây đối với tập dữ liệu 12.000 dòng [[conversations/2026-10-08]]. Đồng thời, ghi trực tiếp ra ổ đĩa cục bộ thay vì đồng bộ sang ổ mạng/ổ chia sẻ giúp tránh lỗi nghẽn I/O và treo file [[conversations/2026-10-08]].
