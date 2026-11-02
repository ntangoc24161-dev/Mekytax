---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-27
updated: 2026-10-27
source: chat 2026-10-27 (65cbc484)
---
Ngày 27/10/2026 người dùng hỏi việc cập nhật sheet GTGT trong `Bang tinh.xlsx` của Sơn Tiền Hải do ai làm. Họ chốt: phải **phân việc cho trợ lý tại hồ sơ kiểm tra thuế** (không phải Văn thư làm, cũng không để Javis làm chung chung), đồng thời **thực hiện chạy thẳng, không chạy nền**. [[conversations/2026-10-27]]

**Vì sao:** lượt trước Javis báo "đang chạy nền" nhưng thực ra không gọi tool nào, việc không được làm; người dùng muốn có kết quả thật ngay trong lượt và việc đi đúng trợ lý chuyên môn hồ sơ kiểm tra. [[conversations/2026-10-27]]

**Áp dụng:** với yêu cầu xử lý Bảng tính / số liệu của một vụ kiểm tra thuế, ghi rõ trợ lý hồ sơ kiểm tra thuế phụ trách (dùng đúng agent/skill chuyên môn, ví dụ nhóm Phân tích hồ sơ thuế) và làm trực tiếp trong lượt chat, trả kết quả thật; không xếp Kanban hay loop nền trừ khi người dùng yêu cầu. Văn thư chỉ phụ trách công văn đến/đi. [[conversations/2026-10-27]]
