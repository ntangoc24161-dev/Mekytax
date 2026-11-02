## Bài học (tự học)
- Script build_ho_so_di_*.py đọc cột theo index cứng (v(27), v(29)...) dễ lệch khi sổ Excel JAVIS PRO chèn thêm cột mới; trước khi chạy lại một script build cũ, nên đối chiếu index cột với header row 4 thực tế của sheet, đừng tin console log đầu tiên.
- Ghi đè ô vào file .xlsm bằng openpyxl (không data_only) sẽ xoá cache hiển thị của các ô công thức khác trong cùng hàng/sheet; cần báo người dùng mở lại bằng Excel một lần để recalc, không phải lỗi dữ liệu.
- Khi đọc năm trên sheet KQKDSXKD, chỉ lấy khối số tuyệt đối trước cột "Dọc - Ngang"; nếu quét hết hàng header thì cột phụ bên phải (ví dụ cột 38 cũng ghi 2022) sẽ đè mất cột năm thật.
- Khi ghi mẫu `Số:              /QĐ-TCS3`, phải set `xml:space=preserve` trên run; nếu không Word/python-docx dễ nuốt khoảng trống còn `Số: /QĐ-TCS3`.
- `clean_doc_flow` cắt đoạn trống liên tiếp ở body, nên bước căn ≥5 dòng chữ ký phải chạy sau `clean_doc_flow`, nếu không khoảng cách bị co lại.
