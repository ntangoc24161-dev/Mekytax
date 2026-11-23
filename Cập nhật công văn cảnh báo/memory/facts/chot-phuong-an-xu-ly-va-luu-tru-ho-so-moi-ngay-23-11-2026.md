---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-23
updated: 2026-11-23
---
- Bóc tách chi tiết bảng kê từ CV 9955 (30 DN F0, 40 DN F1) và CV 5645 (193 DN F1).
- Sheet F0 tăng lên 17.167 dòng (thêm 31 dòng F0 mới gồm 30 DN CV 9955 và 1 DN cty Nguyễn Tài).
- Sheet F1 tăng lên 14.284 dòng (thêm 254 dòng F1 mới gồm 61 dòng CV 9955 và 193 dòng CV 5645).
- Đã phân loại và di chuyển toàn bộ tệp PDF đã nạp sang `01_HO SO CONG AN CHUYEN DEN/2026` và `02_NGUON RUI RO NGANH THUE/2026`.
- Đã dọn sạch 48 tệp trung gian OCR, đưa `HO SO MOI` về trạng thái trống (0 tệp).
- Thống nhất quy trình xử lý: phân luồng tệp số/scan OCR offline không tốn token, kiểm soát qua sổ tạm `Bang_Tam_CVCBRR.xlsx`, backup `.bak` trước khi ghi, và chỉ dọn thư mục nguồn khi nạp thành công.
