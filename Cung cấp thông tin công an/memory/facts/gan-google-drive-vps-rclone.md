---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
---
Ngày 15/10/2026 người dùng chọn gắn Drive lâu dài bằng rclone trên VPS (không gửi Excel, không copy vào vault).

- Làm bằng SSH trên máy chủ, không làm trong khung chat; container không có quyền fuse/rclone.
- rclone mount `My Drive` tại `/mnt/gdrive`; Docker bind `/mnt/gdrive:/mnt/i` (có thể `:rshared`).
- Sổ chính lúc gắn xong: `/mnt/i/HS CONG AN-2025-2026/HS CUNG CAP CA 2026/Bao_cao_cong_van.xlsx`.
- `DB NNT.xls` nếu nằm gốc My Drive: `/mnt/i/DB NNT.xls`.
- rclone mount phải chạy trước khi start/restart Javis; sai thứ tự thì `/mnt/i` trống.
- Hướng dẫn lệnh: `ho-so-cv-moi-2026-10/huong-dan-rclone-google-drive-vps.md`.
- Gắn xong người dùng nhắn `kiểm tra /mnt/i` rồi mới rà dòng 761-810.

**Vì sao:** Cần đọc/ghi sổ live trên Drive khi Javis nằm VPS, không map được chữ `I:` từ máy Windows.

**Áp dụng:** Sau khi mount, mọi skill hồ sơ Công an ưu tiên đường `/mnt/i`. Đây là mục tiêu cho đến khi người dùng xác nhận `/mnt/i` đã có file.
