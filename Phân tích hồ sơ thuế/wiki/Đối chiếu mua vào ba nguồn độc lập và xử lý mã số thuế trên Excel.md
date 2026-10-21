---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
source: [[conversations/2026-10-07]]
---
Trong công tác kiểm tra thuế qua Excel, kiểm tra mua vào được thực hiện theo chuỗi đối chiếu ba nguồn độc lập: HĐĐT mua vào ↔ Bảng kê mua vào ↔ Sổ Nhật ký chung (TK 133) [[sources/2026_CD1.pdf]].

### 1. Chuẩn hóa khóa đối chiếu Mã số thuế (MST)
Khi kết xuất dữ liệu sang Excel, MST thường bị lỗi mất số 0 ở đầu hoặc thiếu dấu gạch nối chi nhánh (`-`), làm sai lệch các phép tìm kiếm và gom nhóm [[sources/2026_CD1.pdf]].
- Sử dụng kết hợp hàm `LEN` và `IF` lồng để xử lý các trường hợp: MST 9 ký tự (bổ sung số `0` thành 10 ký tự: `"0"&MST`), MST 12 hoặc 13 ký tự liền (chèn dấu `-` sau 10 số đầu để chuẩn hóa thành `XXXXXXXXXX-XXX`) [[sources/2026_CD1.pdf]].
- Sau khi chuẩn hóa MST, tạo khóa đối chiếu liên kết giữa Số hóa đơn và MST bằng phép ghép chuỗi `[Số HĐ] & "_" & [MST]` để phục vụ hàm dò tìm `VLOOKUP` hoặc `COUNTIFS` [[sources/2026_CD1.pdf]].

### 2. Quy trình đối chiếu mua vào 3 nguồn
- Tách sổ Nhật ký chung về dạng một Nợ - một Có, lọc các định khoản phát sinh Thuế GTGT đầu vào (TK Nợ 133 đối ứng TK Có liên quan như 331, 111, 112) [[sources/2026_CD1.pdf]].
- Dùng Pivot Table hoặc hàm `SUMIFS` tổng hợp giá trị mua vào chưa thuế và tiền thuế GTGT theo từng tháng/quý của 3 nguồn độc lập [[sources/2026_CD1.pdf]].
- Thiết lập bảng so sánh chênh lệch giữa các nguồn, lọc các dòng lỗi `#N/A` hoặc phát sinh lệch để xác định hóa đơn bỏ ngoài sổ sách, hóa đơn kê khai sai kỳ, hoặc chênh lệch giữa số liệu khai thuế và hạch toán kế toán [[sources/2026_CD1.pdf]].
