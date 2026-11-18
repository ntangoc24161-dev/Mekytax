---
name: "Lập Bảng kê hóa đơn không hợp lệ vào Bảng tính"
description: "Lập sheet BANG KE KO HOP LE theo năm/quý từ BK LOẠI HOÁ ĐƠN, link thuế GTGT sang TINH PHAT cột C và chênh lệch DTCHEP, tách HĐ thuế 0."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-11-17
---
## Khi nào dùng
- Người dùng yêu cầu bóc tách hóa đơn mua vào bị loại từ file bảng kê GTGT NNT cung cấp sang Bảng tính của vụ kiểm tra.
- Cần tính truy thu thuế GTGT theo từng năm, từng quý và đưa vào sheet TINH PHAT + CHAM NOP và DTCHEP.
- Ví dụ: "bóc tách BK LOẠI HOÁ ĐƠN để tạo sheet BANG KE KO HOP LE, tính theo năm, chia theo quý".

## Chuẩn bị
- File nguồn: bảng kê GTGT NNT trong `2. HO SO NNT CUNG CAP\2.5. Bang ke mua vao ban ra\`, sheet `BK LOẠI HOÁ ĐƠN`.
- File đích: `0. BIEN BAN\Bang tinh.xlsx` của vụ.
- Thời kỳ kiểm tra của vụ (kiểm memory trước, vd Kiến Thành Tân bắt đầu từ 2015).
- Sao lưu Bảng tính trước khi sửa: `Bang tinh.bak-<mô tả>.xlsx` cùng thư mục.

## Quy trình
1. Đọc sheet nguồn. Liệt kê MỌI sheet có tên gần giống (vd `BK LOẠI HOÁ ĐƠN (2)`): nếu số hóa đơn khác nhau thì báo người dùng, mặc định bóc đúng sheet được chỉ định.
2. Lọc bỏ hóa đơn ngoài thời kỳ kiểm tra.
3. Trên sheet `BANG KE KO HOP LE ` (tên có dấu cách cuối): xóa sạch số liệu phôi mẫu, giữ công thức tên công ty/MST ở đầu bảng, bổ sung tên mặt hàng từ sheet `MV` nếu có.
4. Nhóm hóa đơn theo năm, trong mỗi năm theo quý; mỗi quý có dòng tổng (tiền hàng, thuế GTGT), cuối bảng có dòng TỔNG CỘNG.
5. Sheet `TINH PHAT + CHAM NOP (C T)`, cột C - Thuế GTGT: mỗi dòng quý phát sinh đặt công thức `='BANG KE KO HOP LE '!H<dòng tổng quý>`; dòng tổng năm để `=SUM` 4 quý. Xóa tham chiếu rác của công ty mẫu ở các năm không phát sinh.
6. Hóa đơn có thuế GTGT = 0: tách sang sheet `1-xem xét không loại`, cùng cấu trúc năm/quý, có tổng.
7. Sheet `DTCHEP` phần GTGT: cột Số báo cáo link dòng tổng từng năm của sheet `GTGT`; cột Chênh lệch chỉ tiêu Thuế GTGT phải nộp link tổng thuế GTGT từng năm trên `BANG KE KO HOP LE `. Xóa số mẫu cũ còn sót.
8. Phần TNDN trên DTCHEP: giữ nguyên số báo cáo theo KQKDSXKD, không trừ phần hóa đơn loại, trừ khi người dùng chỉ đạo khác.

## Bẫy
- Sheet nguồn có bản sao `(2)` chứa thêm hóa đơn: không tự gộp, hỏi người dùng.
- Khi loại một quý (vd ngoài thời kỳ), các dòng bên dưới dịch lên: phải cập nhật lại TOÀN BỘ công thức cột C trên TINH PHAT và các link trên DTCHEP.
- Tên sheet `BANG KE KO HOP LE ` có dấu cách cuối, viết công thức phải giữ đúng.
- Số mẫu cũ còn sót ở các năm không phát sinh trên DTCHEP làm lệch tổng.
- Không dùng sheet BB nếu người dùng đã bảo bỏ qua.

## Kiểm chứng
- Tổng thuế GTGT dòng TỔNG CỘNG trên `BANG KE KO HOP LE ` = tổng thuế cột C trên TINH PHAT = tổng chênh lệch GTGT trên DTCHEP (trừ phần thuế 0 không ảnh hưởng).
- Số hóa đơn sau lọc khớp số hóa đơn nguồn trừ phần ngoài thời kỳ.
- Báo lại người dùng bảng tóm tắt theo năm/quý, tên file sao lưu và các điểm chưa rõ trong dữ liệu nguồn.
