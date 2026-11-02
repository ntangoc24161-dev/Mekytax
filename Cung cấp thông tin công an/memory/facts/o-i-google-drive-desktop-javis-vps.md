---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
---
Ổ `I:` trên máy Windows của người dùng là Google Drive for Desktop, ánh xạ `My Drive` (thư mục `HS CONG AN-2025-2026`). Chữ `I:` chỉ sống trên máy Windows đó.

Javis đang chạy trong container Linux (volume `javis-brains` → `/brains`, `javis-data` → `/data`). Container không nhìn ổ Windows, không có MCP Google Drive, không tự gắn Drive từ khung chat.

Ổ `D:\HSO CONG AN FILE MD` là đĩa máy Windows; rclone không lấy được.

Khi chưa có `/mnt/i` và chưa có `I:`: tra NNT bằng `sources/DB NNT 03.06.2026.xlsx`; lưu PDF/markdown/nháp Excel vào vault `ho-so-cv-moi-YYYY-MM/`; báo chưa ghi sổ chính `Bao_cao_cong_van.xlsx`. Không nói đã cập nhật xong khi mới có bản nháp.

**Vì sao:** 15/10/2026 không mở được dòng 761-810 vì skill chỉ đường `I:\`.

**Áp dụng:** Thử `/mnt/i/...` trước, rồi `I:\My Drive\...`, rồi mới fallback vault.
