---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-05
updated: 2026-11-05
source: chat 2026-11-05 (524b1d04)
---
Thực tế tính đến 28/10/2026: Google Drive của `cungcapthongtin.tcs3tdn@gmail.com` ACTIVE trên Composio qua account **`googledrive_ralph-urbane`** (tên hiển thị `co so 3 Thue`). Kết nối Composio của Javis chạy mode `full` (Toàn quyền). Javis đã đọc và ghi được (CV MOI, thư mục HS CONG AN-2025-2026, Bao_cao_cong_van.xlsx). Quota khoảng 3,27 GB / 15 GB.

Lượt kiểm tra 28/10/2026 thấy tài khoản Composio mặc định lại là mail khác (`ntangoc2416@gmail.com`), không phải cungcapthongtin như đặt ngày 26/10/2026. Các kết nối khác: `ntangoc2416.1@gmail.com` (`googledrive_amala-ensign`, chỉ đọc).

**Vì sao:** Ngày 26/10/2026 người dùng yêu cầu dùng cungcapthongtin làm tài khoản cho hồ sơ Công an; ngày 28/10/2026 người dùng xác nhận đã bật Toàn quyền.

**Áp dụng:** Mọi thao tác Drive hồ sơ Công an phải truyền rõ `account=googledrive_ralph-urbane`, không dựa vào tài khoản mặc định. Không cần rclone hay /mnt/i.
