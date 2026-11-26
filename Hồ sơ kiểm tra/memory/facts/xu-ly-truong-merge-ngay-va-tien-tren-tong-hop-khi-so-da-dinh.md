---
type: preference
provenance: assistant
origin: javis-learned
created: 2026-11-25
updated: 2026-11-25
---
Khi nguồn Excel đã trả về chuỗi ngày `dd/MM/yyyy` và tiền có phân cách hàng nghìn `#.##0`:
1. Bỏ switch `\@ "dd/MM/yyyy"` ở các trường ngày đầy đủ để tránh Word diễn giải nhầm tháng/ngày.
2. Bỏ switch `\# "#,##0"` ở các trường tiền.
3. Đối với các trường ngày tách lẻ (`dd`, `MM`, `yyyy`), kết nối Mail Merge dạng chuỗi không tự bóc tách được các phần tử nên cần xử lý cố định hoặc bổ sung cột ngày tách riêng từ nguồn.
