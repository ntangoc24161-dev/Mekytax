---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-25
updated: 2026-11-25
source: [[conversations/2026-11-25]]
---
## Ý chính
- Excel lưu ngày dưới dạng số serial; Word Mail Merge in ra số thay vì `dd/MM/yyyy` khi MERGEFIELD ngày thiếu switch `\@ "dd/MM/yyyy"` (thực tế tính đến 25/11/2026). [[conversations/2026-11-25]]
- Ghi nhầm switch thành `/@` (dấu gạch chéo thay vì backslash) gây cùng lỗi in serial. [[conversations/2026-11-25]]
- Chỗ tách riêng chữ ngày, tháng, năm dùng `\@ "dd"`, `\@ "MM"`, `\@ "yyyy"`. [[conversations/2026-11-25]]

## Cách xử lý
- Vá đồng loạt trên phôi bằng `scripts/fix_mergefield_ngay_ddmmyyyy.py`; nhánh merge Python (`scripts/field_formatter.py`) cũng phải đổi serial Excel sang `dd/MM/yyyy`. [[conversations/2026-11-25]]
- Phôi `templates/MAUTONG HOP.docx` đã vá 63 chỗ và `templates/MAU TONG HOP HO SO DI.docx` đã vá 3 chỗ (thực tế tính đến 25/11/2026). [[conversations/2026-11-25]]
- File Word đã merge trước khi vá không tự hết lỗi vì số serial đã ghi cứng; phải trộn lại từ phôi mới. [[conversations/2026-11-25]]

Liên quan: [[Mô hình Hybrid Mail Merge và AI trong lập hồ sơ kiểm tra thuế]], [[Kỹ thuật kết nối trực tiếp Mail Merge từ sheet theo dõi qua dòng 1 ẩn]].
