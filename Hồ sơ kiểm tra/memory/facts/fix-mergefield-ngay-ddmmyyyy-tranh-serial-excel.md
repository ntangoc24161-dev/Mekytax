---
type: decision
created: 2026-11-25
updated: 2026-11-25
source: chat 2026-11-25 (mwvhqq6z)
---
Khi Mail Merge hồ sơ kiểm tra, mọi MERGEFIELD ngày phải có switch `\@ "dd/MM/yyyy"` (hoặc `\@ "dd"` / `\@ "MM"` / `\@ "yyyy"` khi tách ngày-tháng-năm). Thiếu switch hoặc dùng nhầm `/@` làm Word in serial Excel dạng số.

**Vì sao:** Excel lưu ngày là số serial; Word chỉ format lại khi có switch `\@` đúng.
**Áp dụng:** sửa bằng `scripts/fix_mergefield_ngay_ddmmyyyy.py` trên phôi; merge Python qua `field_formatter.py` cũng phải đổi serial sang `dd/MM/yyyy`. File đã merge trước khi vá thì merge lại từ phôi mới.
