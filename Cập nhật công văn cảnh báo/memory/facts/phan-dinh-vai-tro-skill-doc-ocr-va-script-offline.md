---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-23
updated: 2026-11-23
---
Phân công kỹ thuật trong luồng xử lý hồ sơ công văn cảnh báo:
- **Script offline (PaddleOCR/Regex):** Giữ vai trò xử lý bóc tách hàng loạt danh sách lớn trực tiếp thành bảng dữ liệu nạp vào `Bang_Tam_CVCBRR.xlsx`, tiết kiệm thời gian và 0 token.
- **Skill doc-ocr:** Đóng vai trò đọc, tra cứu trích đoạn và đối chiếu kiểm tra chi tiết từng công văn cụ thể (Số CV, ngày, đơn vị, MST nghi ngờ) hoặc fallback khi gặp văn bản dị biệt/chất lượng kém.
