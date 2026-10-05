---
type: preference
provenance: assistant
origin: javis-learned
created: 2026-10-05
updated: 2026-10-05
---
- Trong các đợt chạy hàng loạt hồ sơ phân tích thuế, một số doanh nghiệp bị lỗi ghi file do vượt quá giới hạn độ dài đường dẫn Windows (MAX_PATH).
- Khắc phục triệt để bằng cách chuẩn hóa đường dẫn tuyệt đối với prefix `\\?\` trước khi thực hiện ghi file.
