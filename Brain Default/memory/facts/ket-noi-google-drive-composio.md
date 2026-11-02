---
type: business
provenance: assistant
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
---
Google Drive qua Composio trên Javis, chip Composio ở mức Chỉ đọc, dùng Kết nối chung.

Tài khoản (thực tế tính đến 11:00 ngày 2026-10-19):
- `ntangoc2416.1@gmail.com`: account `googledrive_amala-ensign`, đọc được.
- `cungcapthongtin.tcs3tdn@gmail.com`: account `googledrive_ralph-urbane`, đang là MẶC ĐỊNH, đọc được (chi tiết ở [[drive-mail-cungcapthongtin-tcs3tdn]]).
- `googledrive_wrack-scale`: list/about đều 403 thiếu scope, không dùng.
Khi gọi tool luôn chỉ rõ `account` theo đúng mail cần đọc.

Hai lớp phải xanh: (1) trang Kết nối của Javis: chip Composio bật; (2) Composio For You → Connect Apps: Google Drive xanh, đúng mail, tick đủ quyền Drive. Máy không đọc `G:\My Drive`, đường lấy dữ liệu là API qua Composio.

Quy mô Drive `ntangoc2416.1` (thực tế tính đến 10:16 ngày 2026-10-19): khoảng 26,9 GB; gốc My Drive 27 mục (8 thư mục gồm `00` đến `06` và `TAM_HOANG_PHAT_3603721832`, 19 file); đếm được ít nhất 328 file và 72 thư mục rồi dừng vì `rateLimitExceeded`.

Mẹo đếm: list cần cờ `allDrives`, lấy `nextPageToken` ở cấp ngoài; dính quota thì đợi hết cửa sổ rồi đếm tiếp.

Gmail toolkit chưa ACTIVE. Quyền Chỉ đọc gồm tìm, list, metadata, tải; muốn ghi thì nâng chip lên Toàn quyền.
