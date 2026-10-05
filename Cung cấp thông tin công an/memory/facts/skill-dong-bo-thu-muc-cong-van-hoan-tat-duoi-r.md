---
type: decision
created: '2026-09-26'
updated: '2026-09-26'
---
Quy chuẩn kỹ năng rà soát và đồng bộ thư mục công văn hoàn tất:
- **Tên kỹ năng:** `dong-bo-thu-muc-ca` (Rà soát và đồng bộ thư mục Công văn hoàn tất).
- **Mục tiêu:** Quét toàn bộ các thư mục lưu trữ công văn năm 2026 trên Google Drive (`I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026`), phát hiện các thư mục đã có tệp công văn trả lời ký số (`.signed.pdf`, `.signed.signed.pdf` hoặc bản scan ký đóng dấu) nhưng tên thư mục chưa có hậu tố `-R`, và tự động đổi tên thư mục thêm `-R` ở cuối.
- **Nguyên tắc nhận diện:**
  + Phân biệt chính xác giữa công văn đến của Công an (văn bản yêu cầu) và công văn trả lời của Thuế (văn bản phúc đáp cung cấp tài liệu).
  + Tệp trả lời hợp lệ chứa tiền tố `TRA LOI...`, `CV TRẢ LỜI...`, `Cung cấp thông tin...`, `Công văn trả lời...` có chữ ký số điện tử hoặc scan dấu đỏ.
  + Kiểm tra an toàn: Thư mục đã kết thúc bằng `-R` thì không đổi tên lại để tránh lỗi `-R-R`.
- **Vì sao:** Giúp quản lý trực quan trên Google Drive và phần mềm quản lý hồ sơ, phân biệt tức thì giữa các công văn đang xử lý dở dang và các công văn đã phát hành chính thức.
- **Áp dụng:** Tự động kích hoạt khi có lệnh rà soát / đồng bộ thư mục hoàn tất hoặc sau mỗi đợt phát hành công văn trả lời ký số.
