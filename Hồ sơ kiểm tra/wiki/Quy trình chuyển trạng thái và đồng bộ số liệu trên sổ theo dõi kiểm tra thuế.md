---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
source: [[conversations/2026-11-12]]
---
Quy trình vận hành sổ theo dõi kiểm tra quyết toán thuế (`Danh sach kiem tra quyet toan pro.xlsm`) quy định rõ các mốc chuyển trạng thái và đồng bộ số liệu giữa hồ sơ và sổ theo dõi:

1. **Khâu tạo hồ sơ đi**: Khi hoàn tất khởi tạo bộ hồ sơ đi trình ký, hồ sơ tự động chuyển từ sheet `02_DANG_THUC_HIEN` sang sheet `01_DA_BAN_HANH` (trạng thái `1. Đã ban hành`) mà không cần chờ cấp Số QĐKT chính thức [[conversations/2026-11-12]].
2. **Khâu hồ sơ về và chốt số liệu**: Định kỳ hoặc trước khi chốt hồ sơ về, tự động đồng bộ số tiền truy thu, phạt 20% và tiền chậm nộp từ sheet `TINH PHAT + CHAM NOP (C T)` của `Bang tinh.xlsx` sang các cột 40-49 trên sổ `pro.xlsm` [[conversations/2026-11-12]]. Mốc thời gian tính số ngày chậm nộp là Ngày ký BBKT tại cột 36 của sổ [[conversations/2026-11-12]].
3. **Cảnh báo phạt VPHC**: Riêng tiền phạt vi phạm thủ tục hành chính thuế (cột 47), hệ thống đưa ra cảnh báo riêng khi có số liệu phát sinh lớn hơn 0 để rà soát quyết định xử phạt [[conversations/2026-11-12]].
4. **Kỷ luật lưu trữ**: Thư mục `00. Theo doi chung\` duy trì file làm việc chính `Danh sach kiem tra quyet toan pro.xlsm` và giữ file gốc `Danh sach kiem tra quyet toan.xlsx` để tham chiếu đối chiếu; tuyệt đối không lưu file `.bak` tại đây [[conversations/2026-11-12]].
