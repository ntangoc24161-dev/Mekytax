---
type: reference
provenance: assistant
origin: javis-learned
created: 2026-10-08
updated: 2026-10-08
---
Khi xử lý tập dữ liệu lớn (~12.000 dòng từ Bảng kê, HĐĐT), việc đọc từng ô bằng `ws.cell()` trong Python openpyxl làm chậm và treo script; chuyển sang dùng `iter_rows(values_only=True)` và tránh copy trực tiếp sang ổ mạng G: giúp hoàn thành trong khoảng 22 giây.
