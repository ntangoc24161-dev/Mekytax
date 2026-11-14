---
type: decision
created: 2026-11-13
updated: 2026-11-13
source: chat 2026-11-13 (ce012a55)
---
Sáu cột tiền bằng chữ (cột 53-58 / BA-BF) trên sheet `01_DA_BAN_HANH` của sổ `Danh sach kiem tra quyet toan pro.xlsm` đang trả về chuỗi **mất hết dấu tiếng Việt** (ví dụ `Mu?i chín tri?u b?y tram...`), vì VBE lưu chuỗi ký tự trong module VBA theo code page ANSI của hệ thống, không phải Unicode.

**Vì sao:** Hàm `DocSo` được dán vào VBE có các literal tiếng Việt (`mười`, `lẻ`, `đồng`...); VBA chuyển chúng về code page ANSI nên dấu bị hỏng vĩnh viễn ngay trong mã nguồn.

**Áp dụng:** Khi merge hồ sơ, TUYỆT ĐỐI không đọc giá trị 6 cột chữ từ sổ. Luôn tính lại từ con số gốc bằng `scripts/doc_so.py` (hàm `doc_so(n)`, đã kiểm thử các mốc 10/11/15/21/25/101/105/1 tỷ và số âm). Hai module trộn thẻ MERGEFIELD cũng đã được sao vào `scripts/field_merger.py` và `scripts/field_formatter.py` để không còn phụ thuộc thư mục scratch ngoài vault. Cột chữ trên sổ chỉ còn giá trị tham khảo cho người đọc bằng mắt; muốn sửa gốc thì phải viết lại hàm VBA dùng `ChrW()` thay cho literal có dấu.
