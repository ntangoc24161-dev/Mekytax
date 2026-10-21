---
name: Gộp hóa đơn mua vào bán ra
description: Gộp hóa đơn mua vào/bán ra thành Excel 2 sheet, mỗi công ty 1 file; mini app HTML hoặc exe không cần admin.
group: Tài chính
updated: '2026-10-06'
learned_updated: '2026-10-06'
---
## Khi nào dùng

- Gộp danh sách HĐĐT portal hoặc bảng kê quý thành Excel 2 sheet Mua vào / Bán ra, mỗi công ty một file.
- User muốn công cụ mang USB/máy khác, không cài Python hay Excel, không quyền admin.
- Câu lệnh kiểu: mini app hóa đơn, mỗi công ty 1 file 2 sheet, xuất exe gộp hóa đơn, thêm file/folder rồi chạy.

## Chuẩn bị

- Thư mục app: `apps/gop-hoa-don-mua-ban/` (kèm `index.html`, `Chay.bat`, `vendor/`, `mau/`, `HUONG-DAN.md`).
- Bản exe: `apps/gop-hoa-don-mua-ban/GopHoaDon.exe` (Windows 64-bit). Bản nén: `apps/gop-hoa-don-mua-ban/GopHoaDon-portable.zip`.
- Đầu vào: file `.xlsx` portal `DANH SACH HOA DON*` hoặc bảng kê quý/năm. File `.xls` cũ phải lưu thành `.xlsx` trước.
- Mang đi: copy một file `GopHoaDon.exe`, hoặc copy cả thư mục HTML kèm `vendor/` và `mau/`. Không cần mạng sau khi copy.
- Máy đích cần Edge hoặc Chrome. App không ghi `Program Files`, không hỏi UAC.

## Cách chạy

1. Windows portable: bấm đúp `GopHoaDon.exe`. Cửa sổ mở bằng Edge hoặc Chrome sẵn trên máy.
2. Không có exe: Windows bấm `Chay.bat`; máy khác mở `index.html` bằng Chrome hoặc Edge.
3. Chọn file, chọn thư mục, hoặc kéo thả `.xlsx`.
4. Bấm Đọc và phân loại. Để trống MST thì app tự tách từng công ty; điền MST khi cả lô thuộc một doanh nghiệp.
5. Tải từng file hoặc zip. Đóng cửa sổ hoặc bấm Thoát app.
6. Nếu user đưa đường dẫn thư mục máy Javis đọc được: được phép gộp bằng Python/ExcelJS theo đúng quy tắc dưới, rồi mới đưa app mang đi.

## Quy trình

Mỗi công ty ra một Excel `HDDT_<MST>_<Ten>.xlsx`:

- Sheet `Mua vào` và `Bán ra`.
- Tiêu đề: danh sách hóa đơn, tên đơn vị, MST, khoảng thời gian, số HĐ.
- Sắp Ngày lập rồi Ký hiệu HĐ rồi Số HĐ, đánh lại STT.
- Font Times New Roman, in A4 ngang, FitToWidth=1, lặp hàng tiêu đề cột.
- Dòng TỔNG CỘNG bằng công thức `SUM`.

Quy tắc bắt buộc:

1. Map cột theo tên header, không theo vị trí cột. Hàm bỏ dấu phải xử lý chữ đ/Đ.
2. Lọc dòng nhãn Tổng cộng; giữ hóa đơn đối tác tên TỔNG CÔNG TY...
3. Quét nhãn Tổng cộng trên mọi cột, không chỉ 6 cột đầu.
4. File portal: phân mua/bán theo MST người mua / người bán của công ty đang xét. Bảng kê không có MST NNT trên dòng: đọc chiều từ tên sheet/file, hoặc user điền MST.
5. Không để ô chữ bắt đầu bằng `=`.
6. Nhiều công ty trong một thư mục: tách mỗi MST một file.

## Bẫy

- Dòng Tổng cộng lọt nếu chỉ soi 6 cột đầu: tên nhãn có thể nằm cột khác. Phải quét cả dòng.
- Không xóa dòng đối tác tên TỔNG CÔNG TY...; đó là hóa đơn thật.
- File `.xls` cũ app không đọc; bảo lưu thành `.xlsx` rồi kéo lại.
- Bảng kê nhiều NCC, không có MST công ty trên từng dòng: để trống MST thì tách theo đối tác; điền MST khi chắc cả lô một doanh nghiệp.
- Đã có exe: đưa `GopHoaDon.exe` hoặc zip portable. Không bảo cài Python. Không đòi quyền admin. Không nói chưa có path exe.
- SmartScreen lần đầu: Thông tin thêm -> Vẫn chạy. Đó không phải UAC.
- Thiếu Edge/Chrome thì cửa sổ exe không mở được.
- Máy đóng gói là Linux thì chưa bấm thử Windows; không khẳng định đã chạy được trên máy công vụ.

## Kiểm chứng

- Không còn dòng nhãn Tổng cộng; vẫn còn hóa đơn TỔNG CÔNG TY nếu có trong nguồn.
- Hai sheet có tiêu đề đơn vị - MST - thời gian; STT liên tục; dòng TỔNG CỘNG là công thức SUM.
- Thử `apps/gop-hoa-don-mua-ban/mau/` rồi so `mau/ket-qua-mau.xlsx`.
- Hồ sơ Vạn Phát Thành (thực tế tính đến 2026-10-06): 1.764 mua / 1.683 bán sau lọc.
- File exe là PE 64-bit, GUI, đã nhúng giao diện và ExcelJS (thực tế tính đến 2026-10-06). Bấm thử trên Windows: cần xác minh.

## Lịch sử (tự học)
- [2026-10-06] Hội thoại 2026-10-06 lộ hai chỗ skill cũ thiếu: (1) dòng Tổng cộng lọt vì tên nằm ngoài 6 cột đầu khi dựng mini app; (2) user đòi exe bấm-thêm-file/folder, không quyền admin, trong khi skill chỉ nêu HTML/Chay.bat. Bổ sung bẫy .xls và cách điền MST.
- [2026-10-06] Skill còn ghi chưa có đường dẫn .exe trong hội thoại, trong khi lượt 2026-10-06 đã đóng apps/gop-hoa-don-mua-ban/GopHoaDon.exe và GopHoaDon-portable.zip. Làm theo skill cũ sẽ đưa folder HTML và cấm nêu path, bỏ sót file mang đi. Bổ sung cách chạy exe, SmartScreen, Edge/Chrome, chưa test Windows.
