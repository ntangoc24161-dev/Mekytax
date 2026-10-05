---
type: decision
provenance: user
origin: javis-learned
created: 2026-09-24
updated: 2026-09-24
---
Người dùng quyết định chọn Phương án 2 để bóc tách 30 bộ đề thi Đợt 2: xử lý cuốn chiếu theo từng cụm 5 đề (khoảng 300 câu) mỗi lần thay vì chờ chạy trọn bộ 30 đề.

**Vì sao:** Quá trình OCR chạy CPU cục bộ cho 30 đề (1.800 câu) mất 5-7 tiếng; chạy cụm 5 đề giúp có ngay dữ liệu mới nạp vào ngân hàng ôn tập.
**Áp dụng:** Triển khai bóc tách ưu tiên từ Đề 01 đến Đề 05 và cập nhật ngay vào bảng tính Master.
