---
type: decision
provenance: source
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
---
Google không cho đăng nhập Drive bằng mail hay mật khẩu từ khung chat. Gửi mật khẩu Gmail không dùng được.

Trên máy Javis (15/10/2026): rclone `v1.75.2` tại `/data/state/bin/rclone`, cấu hình `/data/state/rclone/rclone.conf`, remote tên `gdrive` (My Drive). Chưa có token OAuth.

Một lần trên máy Windows: `rclone authorize "drive"` (scope drive), đăng nhập đúng Gmail có thư mục `HS CONG AN-2025-2026`, dán nguyên khối JSON vào chat. Chi tiết: `ho-so-cv-moi-2026-10/huong-dan-rclone-google-drive-vps.md`.

**Vì sao:** Người dùng yêu cầu làm giúp và đăng nhập ổ Google bằng mail; OAuth phải mở trình duyệt trên máy Windows.

**Áp dụng:** Không nhận mật khẩu Gmail. Chưa có token thì dùng file Excel đính kèm hoặc vault; không gắn Drive từ trong container.
