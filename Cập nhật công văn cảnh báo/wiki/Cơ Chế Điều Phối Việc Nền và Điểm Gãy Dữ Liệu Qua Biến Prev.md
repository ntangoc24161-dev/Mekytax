---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-09
updated: 2026-11-09
source: [[conversations/2026-11-09]]
---
Trong kiến trúc thực thi tác vụ đa tác tử (multi-agent workflow) của Javis OS, điều phối việc nền (background dispatcher) truyền dữ liệu tuần tự giữa các agent thông qua biến ngữ cảnh `{{prev}}` [[conversations/2026-11-09]].

### Cơ chế hoạt động
- Giao việc trực tiếp trong chat: Trợ lý nắm toàn bộ ngữ cảnh hội thoại, có thể hỏi lại người dùng để làm rõ thông tin thiếu trước khi gọi công cụ [[conversations/2026-11-09]].
- Giao việc qua workflow nền: Hệ thống dispatcher gọi tuần tự từng agent theo cấu hình. Tác tử phía sau chỉ nhận duy nhất kết quả dạng văn bản do tác tử liền trước sinh ra thông qua biến `{{prev}}` [[conversations/2026-11-09]].

### Điểm gãy (Failure Mode)
- Tác tử nền không có cơ chế tương tác ngược với người dùng khi thiếu tham số đầu vào. Khi gặp thông tin mơ hồ, nếu tác tử viết câu hỏi hoặc chuỗi thông báo rỗng/sẵn sàng (như `Đã sẵn sàng`) ra đầu ra, tác tử kế tiếp sẽ nhận nhầm nội dung đó làm dữ liệu nghiệp vụ đầu vào, dẫn đến gãy đổ toàn bộ dây chuyền xử lý [[conversations/2026-11-09]].

### Nguyên tắc thiết kế và vận hành an toàn
- Luôn cung cấp đường dẫn tuyệt đối trực tiếp vào trường `intent` thay vì để tác tử tự dò tìm thư mục [[conversations/2026-11-09]].
- Ưu tiên chia nhỏ quy trình thành từng khâu đơn lẻ hoặc xử lý trực tiếp trong phiên chat đối với các tác vụ gọn nhẹ [[conversations/2026-11-09]].
