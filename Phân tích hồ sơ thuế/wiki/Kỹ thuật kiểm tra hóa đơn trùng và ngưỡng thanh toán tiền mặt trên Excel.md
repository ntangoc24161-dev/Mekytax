---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
source: [[conversations/2026-10-07]]
---
Trong kiểm tra hồ sơ thuế qua Excel, rà soát trùng lặp hóa đơn và kiểm soát điều kiện thanh toán không dùng tiền mặt là hai thủ tục kỹ thuật trọng tâm [[Kỹ năng sử dụng Excel trong xử lý dữ liệu kế toán phục vụ kiểm tra thuế]].

### 1. Kỹ thuật đếm hóa đơn trùng (BT5)
- Khóa kiểm tra trùng lặp chuẩn gồm 3 trường ghép: `Số HĐ & MST người bán & Giá trị chưa thuế` [[Kỹ năng sử dụng Excel trong xử lý dữ liệu kế toán phục vụ kiểm tra thuế]].
- Sử dụng hàm `COUNTIFS` trên dải dữ liệu bảng kê mua vào (ví dụ `F9:F1247`) quét đồng thời cột Số HĐ, MST và Doanh thu chưa thuế; nếu kết quả trả về `> 1` thì cảnh báo nghi vấn kê khai trùng lặp hóa đơn đầu vào [[Kỹ năng sử dụng Excel trong xử lý dữ liệu kế toán phục vụ kiểm tra thuế]].

### 2. Chuẩn hóa mã số thuế (BT6)
- Bảng kê mua vào thường gặp lỗi MST 10 số bị mất số `0` ở đầu do định dạng số (Number) hoặc MST chi nhánh 13 số bị thiếu dấu gạch ngang `-` [[Kỹ năng sử dụng Excel trong xử lý dữ liệu kế toán phục vụ kiểm tra thuế]].
- Sử dụng hàm `LEN()` kết hợp hàm điều kiện `IF()` lồng nhau để phân loại độ dài ký tự và chèn lại số `0` hoặc dấu `-` trước khi dùng làm khóa đối chiếu dữ liệu [[Kỹ năng sử dụng Excel trong xử lý dữ liệu kế toán phục vụ kiểm tra thuế]].

### 3. Rà soát hóa đơn trên ngưỡng tiền mặt trong cùng ngày (BT7A/BT7B)
- Đối với hóa đơn mua cùng một nhà cung cấp trong cùng một ngày có tổng thanh toán từ 20 triệu đồng trở lên (hoặc từ 5 triệu đồng trở lên theo quy định mới), kế toán/công chức kiểm tra tính cột `Tổng giá trị HĐ = Tiền hàng + Tiền thuế` [[Kỹ năng sử dụng Excel trong xử lý dữ liệu kế toán phục vụ kiểm tra thuế]].
- Sử dụng hàm `SUMIFS()` với 2 điều kiện khóa: `MST người bán` và `Ngày hóa đơn` để cộng gộp toàn bộ giá trị mua của NCC trong ngày; sau đó dùng `IF()` so sánh với ngưỡng luật định để bóc tách các khoản thanh toán bằng tiền mặt vi phạm điều kiện khấu trừ [[Kỹ năng sử dụng Excel trong xử lý dữ liệu kế toán phục vụ kiểm tra thuế]].
