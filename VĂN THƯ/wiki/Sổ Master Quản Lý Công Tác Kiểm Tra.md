---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-17
updated: 2026-11-17
source: [[conversations/2026-11-17]]
---
Sổ Master Quản lý công tác kiểm tra là mô hình cơ sở dữ liệu bảng tính tập trung nhằm theo dõi đa dạng các đầu việc kiểm tra thuế (kiểm tra bổ sung KHBS, kiểm tra quyết toán tại trụ sở NNT, thông báo mời giải trình, đóng mở hóa đơn...) [[conversations/2026-11-17]].

### Cấu trúc chuẩn 5 thành phần
1. **DASHBOARD**: Tổng hợp tự động các chỉ số theo từng nhóm việc: tổng số người nộp thuế (NNT), số lượng chưa gửi lần 1, quá hạn lần 1, quá hạn lần 2, cấp đỏ cảnh báo [[conversations/2026-11-17]].
2. **KT_TONG_HOP**: Theo dõi tiến độ chi tiết cho các vụ kiểm tra (KHBS, kiểm tra quyết toán) theo từng niên độ và mốc thời gian thông báo [[conversations/2026-11-17]].
3. **DANH_SACH_THEO_DOI_CHUNG**: Tổng hợp toàn bộ NNT trên địa bàn cùng các loại công việc phát sinh khác nhau nhằm tránh sót việc [[conversations/2026-11-17]].
4. **CONFIG**: Bảng tham số quy định biểu mẫu, thời hạn xử lý giữa các lần gửi thông báo (thông thường 7 ngày làm việc), hành động xử lý cuối cùng (phiếu chuyển xác minh hoặc đóng mã số thuế) và cán bộ phụ trách [[conversations/2026-11-17]].
5. **LEGEND**: Hướng dẫn mã trạng thái và quy ước màu sắc cảnh báo [[conversations/2026-11-17]].
