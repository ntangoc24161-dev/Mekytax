---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-25
updated: 2026-11-25
source: [[conversations/2026-11-25]]
---
Quy trình lập hồ sơ kiểm tra thuế bằng mô hình Hybrid Mail Merge qua thực tế triển khai bộc lộ các hạn chế kỹ thuật và định hướng khắc phục cụ thể:

### 1. Hạn chế về Mail Merge và nguồn dữ liệu
- **Kiểm chứng Word qua mạng:** Word COM thao tác trên ổ mạng (ổ `G:`) có tỷ lệ treo cao, việc kiểm chứng kết quả trộn chủ yếu dựa trên phân tích cấu trúc file XML hoặc OLEDB [[conversations/2026-11-25]].
- **Trường ngày tách (ngày, tháng, năm riêng lẻ):** Các vị trí cần điền riêng lẻ `ngày... tháng... năm...` (khoảng 12 vị trí trong hồ sơ) vẫn phải để tĩnh nếu sổ chưa có các cột phụ chiết tách `dd`, `MM`, `yyyy` từ ngày gốc [[conversations/2026-11-25]].
- **Khóa tệp và cảnh báo SQL:** Excel khi mở sẽ khóa ghi tệp nguồn, đồng thời Word khi mở file merge luôn hiện hộp thoại SQL xác nhận kết nối dữ liệu từ phía người dùng [[conversations/2026-11-25]].
- **Độ nhạy của vị trí cột:** Đọc dữ liệu theo chỉ số cột cố định (khoảng 58 trường) dễ bị xô lệch khi cấu trúc sheet thay đổi hoặc chèn cột mới [[conversations/2026-11-25]].

### 2. Hạn chế về phôi văn bản và logic kiểm tra
- **Tính linh hoạt của chuyên đề chi phí:** Phôi `MAUTONG HOP` mặc định tách hai dòng QLDN và Chi phí bán hàng, trong khi một số vụ việc yêu cầu gộp chung thành một dòng 'Chi phí quản lý kinh doanh' [[conversations/2026-11-25]].
- **Tính nhất quán giữa dự thảo và ban hành:** Quy định điền số quyết định cần thống nhất rõ: nếu đã có số quyết định chính thức trên sổ thì điền đầy đủ, nếu đang ở bước dự thảo thì để trống [[conversations/2026-11-25]].
- **Kiểm soát sai sót tổng thể:** Thiếu công cụ tự động (script lint) để rà soát sót số liệu mẫu cũ, sai lệch font chữ/cỡ chữ, dòng trắng thừa và lỗi định dạng ngày `..../....` trước khi bàn giao hồ sơ [[conversations/2026-11-25]].
