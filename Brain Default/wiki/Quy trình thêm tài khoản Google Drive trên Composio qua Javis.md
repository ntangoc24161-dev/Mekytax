---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-02
updated: 2026-11-02
source: [[conversations/2026-11-02]]
---
Để kết nối thêm tài khoản Google Drive vào Composio qua giao diện chat của Javis [[conversations/2026-11-02]]:
1. Người dùng yêu cầu thêm tài khoản Drive kèm alias nhận diện (thực tế tính đến 02/11/2026) [[conversations/2026-11-02]].
2. Hệ thống gọi tool `COMPOSIO_MANAGE_CONNECTIONS` với tham số `action: add` cho toolkit `googledrive` để sinh OAuth URL xác thực (thực tế tính đến 02/11/2026) [[conversations/2026-11-02]].
3. Người dùng truy cập liên kết xác thực, đăng nhập tài khoản Google và cấp quyền truy cập Google Drive [[conversations/2026-11-02]].
4. Sau khi hoàn tất xác thực trên trình duyệt, gọi `COMPOSIO_WAIT_FOR_CONNECTIONS` hoặc kiểm tra danh sách kết nối để lấy alias mới [[conversations/2026-11-02]].
*Lưu ý:* Cần chỉ định rõ alias khi thực thi tool Drive để tránh rơi vào tài khoản mặc định [[conversations/2026-11-02]].
