---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
source: chat 2026-11-13 (ce012a55)
---
Đã hoàn tất chuyển đổi Mail Merge trực tiếp từ sheet `01_DA_BAN_HANH` và xóa bỏ hoàn toàn sheet `DuLieuMerge`:
1. Đã chèn Dòng 1 ẩn (Hide row) trên sheet `01_DA_BAN_HANH` chứa 58 mã trường tiếng Việt không dấu chuẩn (`STT`, `MST`, `Ten_NNT`... và 6 cột chữ `Tong_truy_thu_chu`...).
2. Phôi Word `templates/MAUTONG HOP.docx` đã được cập nhật liên kết trực tiếp vào bảng `01_DA_BAN_HANH$`.
3. Đã xóa dứt điểm sheet `DuLieuMerge` khỏi sổ `Danh sach kiem tra quyet toan pro.xlsm`.

**Áp dụng:** Khi mở `MAUTONG HOP.docx` trong Microsoft Word, hệ thống nhận diện trực tiếp dữ liệu từ `01_DA_BAN_HANH$`, có thể bấm Next/Previous hoặc lọc theo tên/MST bất kỳ công ty nào để trộn văn bản.
