---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-09-30
updated: 2026-09-30
source: [[conversations/2026-09-30]]
---
Quy trình TTKT: lấy MST người bán trên bảng kê **mua vào** của NNT đang kiểm tra, đối chiếu với sổ theo dõi công văn cảnh báo rủi ro thuế (CVCBRR, gồm cả cảnh báo công an) và danh bạ NCC theo trạng thái hoạt động, rồi xuất một file Excel tổng hợp trong `ket-qua`. [[conversations/2026-09-30]]

### Nguồn đầu vào
- Bảng kê mua vào (sheet `mua vao` hoặc `4.MV` đã gộp). [[conversations/2026-09-30]]
- Sổ CVCBRR dạng Excel theo dõi công văn (đường dẫn lồng `...xlsx\...xlsx` thường **không tồn tại**; dùng file sổ thật một tầng). [[conversations/2026-09-30]]
- Danh bạ NCC (`.xls`/`.xlsx`) có trạng thái: không hoạt động tại địa chỉ đăng ký, ngừng hoạt động chưa hoàn thành thủ tục chấm dứt hiệu lực, tạm ngưng kinh doanh, kèm ngày bắt đầu KD / tạm nghỉ / đóng MST. [[conversations/2026-09-30]]

### Khớp khóa
Ghép theo **MST gốc** (phần trước dấu `-`). (thực tế tính đến 2026-09-30) Khớp MST gốc có thể dính nhánh khác cùng 10 số đầu; cần đối chiếu lại tên trên công văn trước khi kết luận. [[conversations/2026-09-30]]

### Đầu ra chuẩn (mục tiêu theo yêu cầu user)
- Sheet **chi tiết**: HĐ mua + thông tin CV + trạng thái; sắp theo MST NCC, theo CV, và theo thời gian. [[conversations/2026-09-30]]
- Phần **tổng hợp**: theo từng NCC, chia nhóm **công văn cảnh báo** (rủi ro thuế/HĐ + công an) và nhóm **tình trạng hoạt động**. [[conversations/2026-09-30]]

### Đọc kết quả
Ưu tiên NCC vừa có doanh số mua đáng kể vừa dính CV công an/rủi ro thuế hoặc trạng thái ngừng/tạm ngừng. [[conversations/2026-09-30]]
Không kết luận gian lận chỉ vì có tên trong sổ CVCBRR nếu thiếu số/ngày CV hoặc lệch tên/nhánh. (cần xác minh) [[conversations/2026-09-30]]


## Mâu thuẫn
- Quan điểm mới ([[conversations/2026-10-05]]): File `Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx` phải liệt kê **100% NCC** trên bảng kê mua vào (kể cả NCC hoạt động bình thường), kèm dòng Tổng cộng toàn bộ; không chỉ lọc nhóm trạng thái bất thường. (mục tiêu) [[conversations/2026-10-05]]

Sheet tổng hợp bắt buộc có **Phần 3** nhận xét đánh giá rủi ro và đề xuất trọng tâm kiểm tra. (mục tiêu) [[conversations/2026-10-05]]

Sheet chi tiết hóa đơn giữ đủ nhóm cột chứng từ kế toán và cột rủi ro F0/ĐKKD tách riêng theo chuẩn bản Vạn Phát Thành. (mục 
- Cần xác minh (append _open-questions).
