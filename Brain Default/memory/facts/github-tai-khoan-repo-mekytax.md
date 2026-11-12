---
type: reference
provenance: assistant
origin: javis-learned
created: 2026-11-11
updated: 2026-11-11
---
User dùng tài khoản GitHub `ntangoc24161-dev` với repo `https://github.com/ntangoc24161-dev/Mekytax.git` làm kho **đồng bộ 2 chiều các brain** (Brain Default, VĂN THƯ, Cung cấp thông tin công an...) giữa laptop Windows và VPS Javis.

- 2026-10-22: push bị lỗi 403 (không có quyền ghi).
- 2026-11-11: push thường (không force) đã thành công, sync báo "Đồng bộ 2 chiều 11:17 11/11", nhận 356 file.

**Bẫy đã gặp (11/11/2026):** tên file có dấu `:` (vd note `Ổ I: Google Drive Desktop và đường /mnt/i trên Javis VPS.md` trong brain Cung cấp thông tin công an) làm merge trên Windows lỗi `invalid path`. VPS (Linux) cho phép ký tự này nên từng đẩy lại tên cũ lên GitHub. Cách xử lý: commit xoá riêng file có dấu `:` trên nền bản GitHub, giữ bản đã đổi tên (`Ổ I - Google Drive Desktop và đường mnt-i trên Javis VPS.md`). Nếu lỗi tái diễn thì phải đổi tên file trực tiếp trên VPS.

Merge tự động để lại các bản `*.conflict-*-<timestamp>.*` cạnh file index/loop-state tự sinh; các bản này xoá được vì Javis tự dựng lại.
