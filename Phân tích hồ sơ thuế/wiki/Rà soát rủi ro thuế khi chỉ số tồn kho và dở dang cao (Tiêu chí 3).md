---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-01
updated: 2026-10-01
source: [[conversations/2026-10-01]]
---
Khi hồ sơ thuế có tỷ lệ (Hàng tồn kho + Chi phí SXKD dở dang) / Doanh thu thuần vượt ngưỡng rủi ro, cần tập trung kiểm tra 3 nhóm chi phí: (1) Chi phí giá vốn hàng bán TK 632 để phát hiện tồn kho ảo bán hàng không xuất hóa đơn hoặc ghìm giá vốn tránh lỗ; (2) Chi phí tập hợp giá thành TK 621, TK 622, TK 627 (đặc biệt phân bổ chi phí SXC cố định theo công suất thực tế theo VAS 02 và tính có thật của nguyên vật liệu mua vào); (3) Chi phí trích lập dự phòng giảm giá hàng tồn kho TK 2294 [[conversations/2026-10-01]].


## Mâu thuẫn
- Quan điểm mới ([[conversations/2026-10-05]]): Trên Báo cáo tài chính mẫu DNN dùng cho phân tích hồ sơ kiểm tra thuế, chỉ tiêu hàng tồn kho thường nằm ở mã **ct140**. (thực tế tính đến 2026-10-05) [[conversations/2026-10-05]]

Mã **ct141** có thể bằng 0 dù doanh nghiệp vẫn có tồn kho trên sổ XNT/TK 156; nếu lấy ct141 sẽ làm tỷ lệ HTK/Doanh thu sai và TC3 lệch. (thực tế tính đến 2026-10-05) [[conversations/2026-10-05]]

Khi nhãn `kyKKhai` trên XML BCTC lệch 1 năm so với niên độ doanh thu, phải **map năm theo doanh thu** khớp GTGT/TNDN trước k
- Cần xác minh (append _open-questions).
