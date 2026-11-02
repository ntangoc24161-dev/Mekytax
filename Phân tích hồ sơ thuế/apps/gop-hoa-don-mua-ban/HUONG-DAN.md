---
type: note
created: 2026-10-06
---
# Mini app Gộp hóa đơn mua vào - bán ra

## Cách chạy (Windows, bản exe)

1. Copy file `GopHoaDon.exe` (hoặc giải nén `GopHoaDon-portable.zip`) sang USB / máy khác. Không cài đặt, không cần quyền admin.
2. Bấm đúp `GopHoaDon.exe`. Cửa sổ app mở bằng Edge hoặc Chrome sẵn trên máy.
3. Chọn file / chọn thư mục / kéo thả `.xlsx`. Bấm **Đọc và phân loại**, rồi tải kết quả.
4. Đóng cửa sổ hoặc bấm **Thoát app**.

Nếu Windows SmartScreen hỏi: chọn Thông tin thêm -> Vẫn chạy. Đó không phải quyền admin.

Cần Chrome hoặc Edge (máy công vụ Windows 10/11 thường đã có). App không ghi vào `C:\Program Files`, không sửa registry.

## Cách chạy (bản HTML, không exe)

1. Copy nguyên thư mục `apps/gop-hoa-don-mua-ban/` (kèm `vendor/` và `mau/`).
2. Windows: bấm `Chay.bat`. Máy khác: mở `index.html` bằng Chrome hoặc Edge.
3. Chọn file / chọn thư mục / kéo thả các file `.xlsx` hóa đơn. Thử nhanh bằng 3 file trong `mau/`.
4. Bấm **Đọc và phân loại**.
5. Tải từng công ty hoặc tải zip.

## File ra

Mỗi công ty 1 Excel:

- Sheet `Mua vào`, sheet `Bán ra`
- Tiêu đề: danh sách hóa đơn, tên đơn vị, MST, khoảng thời gian, số HĐ
- Sắp Ngày lập -> Ký hiệu HĐ -> Số HĐ, đánh lại STT
- Font Times New Roman, in A4 ngang, co vừa 1 trang rộng, lặp hàng tiêu đề cột khi in
- Dòng TỔNG CỘNG dùng công thức `SUM`
- Tên file: `HDDT_<MST>_<Ten>.xlsx`

## Nguồn đọc được

- File portal HĐĐT `DANH SACH HOA DON*.xlsx` (mua/bán lệch cột vẫn map theo tên header)
- Bảng kê quý/năm mua vào - bán ra
- Nhiều công ty trong cùng thư mục: tự tách mỗi MST một file

## Quy tắc đã gắn sẵn

- Lọc dòng nhãn `Tổng cộng`, giữ hóa đơn đối tác tên `TỔNG CÔNG TY...`
- Bỏ dấu kể cả chữ đ/Đ khi map header
- Ô chữ không bắt đầu bằng `=` (tránh lỗi Removed Records Formula)
- Phân mua/bán theo MST người mua / người bán của công ty đang xét

## Lưu ý

- File `.xls` cũ: mở Excel, Lưu thành `.xlsx`, rồi kéo lại.
- Điền MST khi chắc chắn cả lô file thuộc 1 doanh nghiệp.
- App xử lý trên máy người dùng, không gửi file ra ngoài.
