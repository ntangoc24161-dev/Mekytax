---
name: Đối chiếu CUP GDLK
description: "Đối chiếu giao dịch liên kết và độc lập theo phương pháp CUP, kiểm tra giá mua bán và lệch xuất nhập tồn kho."
group: Tài chính
---

# Đối chiếu CUP GDLK - Quy tắc và quy trình rà soát giao dịch liên kết

## Khi nào dùng
Kích hoạt khi người dùng yêu cầu:
- Làm CUP, tính CUP, đối chiếu GDLK hoặc giao dịch liên kết.
- Rà soát chênh lệch giá mua bán giữa bên liên kết và bên độc lập theo phương pháp so sánh giá giao dịch độc lập (CUP).
- Kiểm tra, đối chiếu số liệu mua bán M02-M05 với xuất nhập tồn kho (sheet M8-yyyy).
- Kiểm tra tính hợp lệ và cấu trúc file dữ liệu nguồn `Tinh_Cup_Chuan.xlsx` hoặc `Tinh Cup.xlsx`.

## Chuẩn bị
1. **File dữ liệu nguồn:**
   - File nguồn chuẩn: `Tinh_Cup_Chuan.xlsx` hoặc file dữ liệu thực tế `Tinh Cup.xlsx`.
   - Các sheet bắt buộc và cấu trúc:
     - `M01`: Danh sách bên liên kết (dữ liệu bắt đầu từ hàng 7).
     - `M02`: Mua vào liên kết (tiêu đề hàng 8, dữ liệu bắt đầu từ hàng 9).
     - `M03`: Mua vào độc lập (tiêu đề hàng 8, dữ liệu bắt đầu từ hàng 9).
     - `M04`: Bán ra liên kết (tiêu đề hàng 8, cột B = Số chứng từ, cột C = Ngày, dữ liệu bắt đầu từ hàng 9).
     - `M05`: Bán ra độc lập (tiêu đề hàng 8, cột B = Số chứng từ, cột C = Ngày, dữ liệu bắt đầu từ hàng 9).
     - `M8-yyyy`: Xuất nhập tồn kho từng năm (tiêu đề hàng 1, dữ liệu bắt đầu từ hàng 2; cột B = Mã, cột G = SL nhập, cột I = SL xuất). Số lượng năm linh hoạt theo hồ sơ (ví dụ `M8-2020` đến `M8-2024`).
     - `M8-MAU`: Sheet mẫu để nhân bản sang năm mới khi cần bổ sung năm tài chính.
   - Sheet không dùng cho CUP: `M06`, `M07` (định mức và tổng hợp GDLK). Bỏ qua các sheet mẫu trống như `M08`, `M08 (2)` nếu không có dữ liệu phát sinh.
2. **Công cụ thực thi sẵn có:**
   - Ứng dụng HTML độc lập chạy ngay: `CUP_GDLK_ChayNgay.html` (chạy trên trình duyệt Chrome/Edge, không cần cài đặt môi trường).
   - Script Python xử lý tự động qua `openpyxl`: `python scripts/run_cup.py --input "Tinh Cup.xlsx" --output "Tinh_Cup_DoiChieu_CUP.xlsx"`.

## Quy tắc đối chiếu (CUP & NXT)
Các nguyên tắc nghiệp vụ đã chốt, áp dụng nhất quán và không hỏi lại:
1. **Xác định năm tài chính:**
   - Niên độ tài chính doanh nghiệp: từ **Tháng 4 năm N đến Tháng 3 năm N+1**.
   - Ưu tiên số năm tại cột **Kỳ tính thuế** (ví dụ: `2024`).
   - Nếu thiếu cột kỳ tính thuế: Tháng > 3 giữ nguyên năm dương lịch; Tháng <= 3 tính bằng năm dương lịch trừ 1.
2. **Khóa khớp dữ liệu NXT:**
   - Khóa khớp duy nhất: Cùng **(Năm TC + Mã hàng hóa)**.
   - Chuẩn hóa mã hàng: Chuyển toàn bộ thành chữ in hoa (UPPERCASE) và cắt bỏ khoảng trắng thừa đầu cuối (TRIM).
   - Không gom theo nhóm mã hoặc cắt đuôi mã giá, trừ khi có chỉ đạo cụ thể.
   - Xử lý lệch mã: Khi mã chứng từ khác mã kho (ví dụ `FOA401811` so với `TPFOA401811`), hệ thống sẽ phát sinh cảnh báo lệch NXT hoặc dòng phát sinh riêng.
3. **Quy tắc tính toán CUP:**
   - Mua vào: Tổng hợp M02 (LK) và M03 (ĐL) vào `Tổng mua vào`, gom theo cặp (Năm TC + Mã) sang `Cup mua vào`.
   - Bán ra: Tổng hợp M04 (LK) và M05 (ĐL) vào `Tổng bán ra`, gom sang `Cup bán ra`. Cột Lãi/Lỗ: `= Thành tiền (cột I) - Giá vốn hàng bán (cột P)`.
   - Giá trị chênh lệch phía mua: Khi có cả hai phía LK và ĐL, tính theo công thức: `MAX(Đơn giá LK - Đơn giá ĐL, 0) * Số lượng LK`.
   - Chênh lệch NXT mua vào = `SL nhập trên M8 - (SL LK + SL ĐL)`.
   - Chênh lệch NXT bán ra = `SL xuất trên M8 - (SL LK + SL ĐL)`.
4. **Dòng phát sinh bổ sung cuối sheet CUP:**
   - Mã có SL nhập trên M8 > 0 nhưng không có trên Tổng mua vào: Bổ sung dòng cuối `Cup mua vào` với ghi chú `Có trên M8 (nhập), không có trên tổng mua vào`.
   - Mã có SL xuất trên M8 > 0 nhưng không có trên Tổng bán ra: Bổ sung dòng cuối `Cup bán ra` với ghi chú `Có trên M8 (xuất), không có trên tổng bán ra`.

## Cấu trúc file kết quả
File kết quả chuẩn `Tinh_Cup_DoiChieu_CUP.xlsx` gồm 6 sheet theo đúng thứ tự:
1. `Thong ke`: Thống kê tổng hợp số dòng, số mã và hiện trạng chung của hồ sơ.
2. `KQ M2-M5`: Tổng hợp 1 dòng duy nhất cho mỗi cặp (Năm TC + Mã hàng), chia thành 4 khối cột M2, M3, M4, M5 (SL, ĐG, TT). Đơn giá = TT / SL khi SL > 0. Freeze cố định từ hàng A6.
3. `Tổng mua vào`: Dữ liệu chi tiết kết hợp M02 + M03, có cột phân loại LK/ĐL, tách SL/GT ở cột K-N. Hàng 1 tiêu đề, hàng 2 là dòng `SUBTOTAL(9, ...)` cho cột G, I, K-N, hàng 3 là tiêu đề cột, hàng 4 là chữ giải thích nền vàng. Dữ liệu từ hàng 5. Freeze cố định A5, bật bộ lọc filter A3:O.
4. `Tổng bán ra`: Dữ liệu chi tiết kết hợp M04 + M05. Cột Q là Lãi/Lỗ (`= I - P`), cột R là Năm TC. Freeze cố định A5, bật bộ lọc filter A3:R.
5. `Cup mua vào`: Gom theo (Năm TC + Mã). Nhóm cột M02 Mua vào LK và M03 Mua vào ĐL, cột chênh lệch ĐG, giá trị chênh lệch, NXT nhập M8, chênh lệch NXT, Năm TC và Ghi chú. Freeze cố định A6.
6. `Cup bán ra`: Gom theo (Năm TC + Mã). Cột A là Tên hàng, cột B là Mã hàng. Gồm các cột: Xuất vãng lai (để trống), Tổng xuất HĐ, NXT xuất M8, Chênh lệch, Ghi chú, Năm TC. Freeze cố định A6.

## Bẫy & Điều cấm
- **Không đặt công thức tại hàng nền vàng:** Hàng nền vàng dưới tiêu đề là văn bản giải thích thuần túy, tuyệt đối không bắt đầu bằng dấu `=` để tránh phát sinh lỗi `#NAME?`.
- **Dòng SUBTOTAL hàng 2:** Chỉ tính `SUBTOTAL(9, ...)` theo vùng lọc dữ liệu thực tế, không được bao gồm hàng giải thích nền vàng.
- **Bảo toàn cấu trúc file nguồn:**
  - Không đổi tên các sheet chuẩn: `M02`, `M03`, `M04`, `M05`.
  - Không sửa hoặc xóa hàng tiêu đề màu xanh.
  - Không chèn thêm cột trước cột A, không merge cell từ hàng dữ liệu.
  - Không chèn dòng Cộng trung gian giữa các dòng chi tiết.
  - Không nhập số lượng hoặc giá trị dạng chuỗi kèm đơn vị tính (ví dụ: `1.234 đ`).
  - Không gộp dữ liệu nhiều năm vào cùng một sheet M8.
- **Loại bỏ dòng thừa:** Tự động bỏ qua các dòng tổng cộng cuối nguồn không có mã hàng hoặc số chứng từ.

## Kiểm chứng
Khi thực hiện đối chiếu hoặc xuất kết quả, rà soát theo checklist sau:
1. Đủ 6 sheet kết quả theo thứ tự chuẩn, định dạng freeze pane và autofilter đúng quy định.
2. Kiểm tra dòng SUBTOTAL ở hàng 2 của `Tổng mua vào` và `Tổng bán ra` khớp chính xác với chi tiết.
3. Rà soát cột Ghi chú tại `Cup mua vào` và `Cup bán ra`:
   - Các dòng có đủ hai phía: `Có CUP (LK+ĐL)`.
   - Các dòng có chênh lệch kho: `Lệch SL với NXT`.
   - Các dòng bổ sung từ kho: `Có trên M8...`.
4. Đảm bảo toàn bộ bảng tính không xuất hiện lỗi công thức Excel (`#NAME?`, `#VALUE!`, `#REF!`).
