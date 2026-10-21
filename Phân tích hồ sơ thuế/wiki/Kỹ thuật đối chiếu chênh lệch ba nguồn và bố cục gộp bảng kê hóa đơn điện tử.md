---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-14
updated: 2026-10-14
source: [[conversations/2026-10-14]]
---
Trong quy trình phân tích và làm sạch dữ liệu kiểm tra thuế theo chuyên đề BT5-BT9, việc đối chiếu giữa 3 nguồn dữ liệu độc lập (Tờ khai, Bảng kê, HĐĐT) được chuẩn hóa theo các nguyên tắc [[conversations/2026-10-14]]:
- Bảng đối chiếu doanh thu bán ra (Bảng 1.2): Đối chiếu đồng thời cả 2 cặp Tờ khai vs Bảng kê và Tờ khai vs HĐĐT, tính chênh lệch tuyệt đối để phát hiện doanh thu chưa kê khai hoặc sai lệch thời điểm [[conversations/2026-10-14]].
- Bảng đối chiếu mua vào (Bảng 1.3): Tính chênh lệch giá trị và tiền thuế giữa Tờ khai và HĐĐT điện tử cơ quan thuế; lược bỏ các cột đếm số dòng cơ học để tập trung vào giá trị tiền [[conversations/2026-10-14]].
- Đánh dấu rủi ro: Áp dụng định dạng có điều kiện (conditional formatting) tô màu nổi bật đối với toàn bộ các chỉ tiêu có giá trị chênh lệch khác 0 [[conversations/2026-10-14]].
- Tinh gọn bảng tính: Gộp các danh mục mua vào và bán ra trên cùng một sheet theo bố cục 2 phần xếp dọc (`Chuan MST BK`, `DC_HDDT_vs_BangKe`) giúp giảm tải số lượng sheet trong workbook nhưng vẫn đảm bảo tách bạch chiều mua - bán [[conversations/2026-10-14]].
