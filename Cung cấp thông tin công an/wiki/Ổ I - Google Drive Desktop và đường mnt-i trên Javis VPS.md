---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
source: [[conversations/2026-10-15]]
---
Ổ `I:` trên máy Windows là Google Drive for Desktop, ánh xạ thư mục `My Drive`. (thực tế tính đến 15/10/2026) [[conversations/2026-10-15]]

Chữ `I:` chỉ tồn tại trên máy Windows đó. Javis chạy trong container Linux (volume `/brains`, `/data`) không nhìn ổ đĩa Windows và không tự gắn Drive từ khung chat. (thực tế tính đến 15/10/2026) [[conversations/2026-10-15]]

Cách gắn lâu dài trên VPS (mục tiêu): rclone trên máy chủ mount `My Drive` tại `/mnt/gdrive`, Docker bind vào Javis tại `/mnt/i`. [[conversations/2026-10-15]]

Sau khi gắn, sổ chính là `/mnt/i/HS CONG AN-2025-2026/HS CUNG CAP CA 2026/Bao_cao_cong_van.xlsx`, tương đương `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx`. (mục tiêu) [[conversations/2026-10-15]]

Ổ `D:\HSO CONG AN FILE MD` là đĩa máy Windows; rclone không lấy được. (thực tế tính đến 15/10/2026) [[conversations/2026-10-15]]

rclone mount phải chạy trước khi start container Javis; sai thứ tự thì `/mnt/i` là thư mục trống. (cần xác minh trên từng VPS) [[conversations/2026-10-15]]
