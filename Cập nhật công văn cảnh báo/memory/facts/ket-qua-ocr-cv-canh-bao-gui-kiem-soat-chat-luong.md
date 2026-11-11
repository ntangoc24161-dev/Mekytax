---
type: preference
provenance: user
origin: javis-learned
created: 2026-11-11
updated: 2026-11-11
---
Ngày 11/11/2026, khi hỏi agent `so-hoa-ocr-cong-van` làm xong gửi ai, user chỉ đạo: **gửi Kiểm soát chất lượng**. Kết quả gửi gồm danh sách đường dẫn `.md`, cờ chất lượng, trích dòng đầu mỗi file.

**Vì sao:** user muốn đầu ra OCR được kiểm soát độc lập trước khi đi tiếp.
**Áp dụng:** sau bước 2 (OCR), bàn giao cho agent Kiểm soát chất lượng dữ liệu trong brain, không chỉ báo cáo lại cho user; không hỏi lại tên/Zalo người QC vì QC là agent nội bộ.
