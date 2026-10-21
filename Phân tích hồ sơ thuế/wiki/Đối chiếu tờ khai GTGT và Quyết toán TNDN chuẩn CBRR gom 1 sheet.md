---
type: wiki
status: active
tags: [wiki, tai-chinh, ttkt]
origin: javis-learned
created: 2026-10-01
updated: 2026-10-01
source: [[conversations/2026-10-01]]
---
Quy trình đối chiếu toàn diện nghĩa vụ thuế GTGT (từng quý) và Quyết toán TNDN (theo năm) giữa Bảng kê mua vào/bán ra, Tờ khai thuế (Mẫu 01/GTGT, Mẫu 03/TNDN) và Hóa đơn điện tử (HĐĐT) chuẩn 2 sheet CBRR (`Tong_hop_doi_chieu` và `Doi_chieu_NCC_CVCBRR`) theo chuẩn kỹ thuật công văn CBRR. [[BANG_SO_LIEU_CONG_VAN.md]]

### Cấu trúc Báo cáo 2 Sheet chuẩn CBRR

#### 1. Sheet 1: `Tong_hop_doi_chieu` (GTGT, TNDN & 23 Chỉ tiêu Nguyên nhân)
- **Phần I. Đối chiếu nghĩa vụ thuế GTGT theo kỳ (2016 - 2025)**:
  + **Bảng 1 (Mẫu 3 CBRR - 7 cột)**: Kê khai thuế GTGT theo kỳ (Kỳ, Giá trị mua vào [23], Thuế mua vào [24], Tổng doanh thu [34], Thuế đầu ra [35], Phát sinh dương [40], Còn khấu trừ [43]/[41]). Dòng **CỘNG** tính tổng cột 2–5, cột 6–7 để trống. [[BANG_SO_LIEU_CONG_VAN.md]]
  + **Bảng 2 (Mẫu 5 CBRR mở rộng - 7 cột)**: Đối chiếu Bán ra theo kỳ (Kỳ, Tờ khai [34]/[35], HĐĐT hợp lệ DT/Thuế, Chênh lệch DT/Thuế = Tờ khai − HĐĐT). [[BANG_SO_LIEU_CONG_VAN.md]]
  + **Bảng 3 (Mẫu 6 CBRR mở rộng - 7 cột)**: Đối chiếu Mua vào theo kỳ (Kỳ, Tờ khai [23]/[24], Bảng kê mua vào DT/Thuế, Chênh lệch DT/Thuế = Tờ khai − Bảng kê). [[BANG_SO_LIEU_CONG_VAN.md]]
- **Phần II. Đối chiếu Doanh thu và Quyết toán TNDN theo năm (Mẫu 4 CBRR)**:
  + So sánh Doanh thu Bảng kê (1), HĐĐT hợp lệ (2), TNDN Tổng DT [01] (3), Giảm trừ DT [02], TNDN DT thuần [03], Chênh lệch BK vs TK [01], Thu nhập chịu thuế, Thuế TNDN phải nộp.
  + Bóc tách mốc triển khai HĐĐT theo NĐ 123/2020/NĐ-CP (từ tháng 7/2022); dữ liệu HĐĐT trước 2023 chỉ mang tính tham khảo.
- **Phần III. Bảng số liệu tổng hợp 23 chỉ tiêu nguyên nhân trọng yếu theo năm (10 năm)**:
  + *Nhóm A*: Doanh thu Bán ra BK, DT [01], Giảm trừ [02], DT thuần [03], Chênh lệch BK vs [01], Thu nhập chịu thuế, Thuế TNDN.
  + *Nhóm B*: DT HĐĐT hợp lệ, Lệch do chưa áp dụng NĐ 123, Giá trị HĐ thay thế/hủy.
  + *Nhóm C*: Tổng Mua vào BK, Mua vào [23], Lệch do gõ phím/hình thức, Thuế GTGT BK, Thuế GTGT [24], Lệch Thuế BK vs [24].
  + *Nhóm D*: Mua từ NCC CVCBRR F0, Mua từ NCC bất thường, Tỷ trọng NCC bất thường (%), Mua từ NCC bình thường, Tỷ trọng NCC bình thường (%).
- **Phần IV. Đề xuất trọng tâm kiểm tra tại NNT**:
  + 4 đề xuất thực tế: Trọng tâm NCC lớn/bất thường, Xác minh NCC tạm ngừng KD, Giải trình lỗi gõ phím hình thức ct23, Bổ sung tờ khai XML kỳ thiếu.

#### 2. Sheet 2: `Doi_chieu_NCC_CVCBRR` (Đối chiếu NNT & NCC với CVCBRR, Sổ CA, Danh bạ)
- **Phần 1: Đối chiếu NNT đang xét**: Đối chiếu với Sheet F1 (Bên mua cảnh báo) và Sổ Công an. [[Đối chiếu NCC mua vào với CVCBRR và danh bạ trạng thái]]
- **Phần 2: Đối chiếu Danh sách NCC mua vào**:
  + *Bảng 2.1*: Đối chiếu với Sổ CVCBRR F0 (Bên bán cảnh báo).
  + *Bảng 2.2*: Đối chiếu với Danh bạ NNT (Trạng thái bất thường: tạm ngừng KD, đóng MST, giải thể...).
  + *Bảng 2.3*: Bảng tổng hợp số liệu chi tiết các NCC bất thường theo từng năm (kèm dòng Tổng cộng và Tỷ trọng %).
- **Phần 3: Đề xuất trọng tâm kiểm tra NCC**: Kế hoạch xác minh thực tế tính có thực của giao dịch.

### Quy chuẩn Thư mục Sao lưu & Lưu trữ Hồ sơ
- Gom toàn bộ file báo cáo chính thức, file sao lưu trung gian, file thuyết minh Markdown và dữ liệu JSON vào thư mục sao lưu:
  + Thư mục trên máy: `7. Sao luu tong hop` (theo quy chuẩn phân loại `0. BIEN BAN` -> `7. Sao luu tong hop`).
  + Thư mục trong Second Brain: `sao-luu/`.

### Kỹ thuật Excel & Tránh lỗi Formula XML
- **Tránh lỗi Formula XML**: Không để ô văn bản bắt đầu bằng dấu `=`, `+`, `-` (như `=> KẾT LUẬN...` phải chuyển thành `> KẾT LUẬN...` hoặc gán kiểu chuỗi `cell.data_type = 's'`) để tránh openpyxl gán nhầm thẻ formula `<f>` gây lỗi *Removed Records: Formula from /xl/worksheets/sheet1.xml part*. [[Lỗi Removed Records Formula trong Excel khi mở file]]
- Định dạng toàn bộ file bằng font **Times New Roman**, số tiền phân cách hàng nghìn, bật `wrap_text = True` và auto fit độ cao hàng.
