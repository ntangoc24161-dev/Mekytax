---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-02
updated: 2026-10-07
---
File tổng hợp hồ sơ bị lệch định dạng so với mẫu do các nguyên nhân kỹ thuật sau:

1. Tồn tại song song file `TONG HOP.docx` (phôi thô còn trường trộn) và `TONG HOP HO SO DI.docx` (tạo từ code).
2. Script `python-docx` dựng lại từ trang trắng thay vì merge trên phôi `MAUTONG HOP.docx`, làm mất tab stop, đường kẻ trang trí và bảng chữ ký theo Nghị định 30/2020/NĐ-CP.
3. Script chuẩn hóa ép hàng loạt `line_spacing=1.15` và `space_after=1pt`, trong khi phôi `MAUTONG HOP.docx` dùng chủ yếu giãn dòng **1.0** và `space_after` để trống.
4. Script gỡ floating table (`tblpPr`) của các bảng quốc hiệu/đoàn kiểm tra, chỉ giữ floating khung Mẫu số: phôi có **7** bảng nổi, file lỗi còn **1**.

**Phương pháp chuẩn:** thao tác trực tiếp trên phôi `MAUTONG HOP.docx`; giữ nguyên spacing và floating của phôi; không normalize giãn dòng hàng loạt; chỉ xuất riêng các Quyết định, còn lại nằm trong `TONG HOP.docx`.

**Áp dụng:** mọi lần chỉnh `TONG HOP.docx` phải đối chiếu lại với phôi trước khi đóng gói `hoso.zip`.
