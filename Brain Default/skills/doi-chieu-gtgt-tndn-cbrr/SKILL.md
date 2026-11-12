---
name: "Đối chiếu tờ khai GTGT và TNDN chuẩn CBRR"
description: "Đối chiếu Bảng kê - Tờ khai GTGT - Quyết toán TNDN - HĐĐT gom 1 sheet chuẩn 5 bảng CBRR, bóc tách chênh lệch và xuất Excel."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-10-01
---
## Khi nào dùng

- Đối chiếu nghĩa vụ thuế **GTGT** (từng quý/tháng) và **Quyết toán TNDN** (theo năm) giữa Bảng kê mua vào/bán ra, Tờ khai thuế (Mẫu 01/GTGT, Mẫu 03/TNDN) và Hóa đơn điện tử (HĐĐT).
- Xuất file Excel tổng hợp đối chiếu gom trọn vẹn trên **1 sheet duy nhất** (`Tong_hop_doi_chieu`) theo chuẩn kỹ thuật 5 bảng số liệu công văn CBRR (`BANG_SO_LIEU_CONG_VAN.md`).
- Bóc tách nguyên nhân chênh lệch: lọc HĐĐT hợp lệ vs HĐ thay thế/hủy bỏ, mốc thời gian áp dụng HĐĐT theo NĐ 123 (từ giữa năm 2022), phát hiện lỗi gõ phím/lệch chỉ tiêu, xử lý kỳ thiếu tờ khai / kỳ ảo.

## Chuẩn bị

1. **Tờ khai thuế / BCTC**: File `TONG HOP_TO KHAI.xlsx` (sheet `GTGT 01`, `QT TNDN 03`) hoặc XML tờ khai đã giải mã.
2. **Bảng kê mua vào / bán ra**: File `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx` (sheet `ban ra`, `mua vao` 10 năm đã gộp).
3. **Dữ liệu HĐĐT portal**: Thư mục `BAN RA`, `MUA VAO` chứa file Excel HĐĐT chi tiết.
4. Xác định MST người nộp thuế đang kiểm tra và thời kỳ kiểm tra chính thức (ví dụ 2016 - 2025, loại bỏ các năm ngoài kỳ).

## Quy chuẩn 5 bảng số liệu CBRR gom 1 Sheet

Toàn bộ báo cáo nằm trên sheet **`Tong_hop_doi_chieu`** gồm 3 phần liên hoàn:

### PHẦN I. ĐỐI CHIẾU NGHĨA VỤ THUẾ GTGT THEO KỲ
- **Bảng 1 (Mẫu 3 CBRR - 7 cột)**: Kê khai thuế GTGT theo kỳ (Kỳ, GT mua vào [23], Thuế mua vào [24], Tổng DT [34], Thuế đầu ra [35], PS dương [40], Còn khấu trừ [43]/[41]). Dòng **CỘNG** chỉ tính tổng cột 2–5; cột 6–7 để trống.
- **Bảng 2 (Mẫu 5 CBRR mở rộng - 7 cột)**: Đối chiếu Bán ra theo kỳ (Kỳ, Tờ khai [34]/[35], HĐĐT hợp lệ DT/Thuế, Chênh lệch DT/Thuế = Tờ khai − HĐĐT).
- **Bảng 3 (Mẫu 6 CBRR mở rộng - 7 cột)**: Đối chiếu Mua vào theo kỳ (Kỳ, Tờ khai [23]/[24], Bảng kê mua vào DT/Thuế, Chênh lệch DT/Thuế = Tờ khai − Bảng kê).

### PHẦN II. ĐỐI CHIẾU DOANH THU & QUYẾT TOÁN TNDN THEO NĂM
- Bảng đối chiếu Doanh thu và Quyết toán TNDN (Mẫu 4 CBRR mở rộng):
  + Cột: STT, Năm, Doanh thu Bảng kê (1), HĐĐT hợp lệ (2) *(ghi chú tham khảo trước 2023)*, TNDN Tổng DT [01] (3), Giảm trừ DT [02], TNDN DT thuần [03], Chênh lệch BK - HĐĐT (1)-(2), Chênh lệch BK - TK [01] (1)-(3), Thu nhập chịu thuế, Thuế TNDN phải nộp, Đánh giá & Ghi chú.
  + Dòng CỘNG: Tính tổng các cột số liệu; lưu ý đánh giá độ khớp Bảng kê vs Chỉ tiêu [01].

### PHẦN III. TỔNG HỢP NGUYÊN NHÂN CHÊNH LỆCH & KẾT LUẬN
- Bóc tách tính nhất quán doanh thu (khớp 100% giữa Bảng kê và Tờ khai TNDN).
- Bóc tách độ phủ HĐĐT (loại trừ HĐ hủy/thay thế theo `CANCEL_KEYWORDS`).
- Nhận diện lỗi gõ phím / sai sót hình thức (ví dụ lệch ct23 nhưng ct24 khớp đúng).
- Đề xuất hồ sơ, giải trình cần NNT cung cấp khi thanh kiểm tra thực tế.

## Kỹ thuật Excel & Tránh bẫy XML

1. **Tuyệt đối tránh lỗi Formula XML**: Không để nội dung ô văn bản bắt đầu bằng dấu `=`, `+`, `-` (ví dụ `=> KẾT LUẬN...` phải đổi thành `> KẾT LUẬN...` hoặc set tường minh `cell.data_type = 's'`). Nếu bắt đầu bằng `=`, openpyxl sẽ lưu thành thẻ formula `<f>` và Excel sẽ báo lỗi *Removed Records: Formula from /xl/worksheets/sheet1.xml part*.
2. **Định dạng**: Font **Times New Roman** 100% file; tiêu đề cột in đậm; số tiền định dạng phân cách hàng nghìn chuẩn VN; dòng CỘNG dùng hàm `SUM(...)`.
3. **Hiển thị đầy đủ**: Thiết lập `wrap_text = True` cho các cột diễn giải, tự căn chỉnh chiều cao hàng và độ rộng cột để không bị khuất chữ.

## Bẫy thường gặp

- Đưa cả năm ngoài thời kỳ kiểm tra (như năm 2015) vào bảng tổng hợp khi chưa có bảng kê.
- Cộng cả hóa đơn hủy, xóa bỏ, bị thay thế vào số liệu đối chiếu HĐĐT.
- Nhầm lẫn giữa HĐĐT mua vào trên portal (thường thiếu) với Bảng kê mua vào thực tế của DN.
- Quên gán Times New Roman hoặc để ô ghi chú bắt đầu bằng dấu `=` làm hỏng file Excel.
