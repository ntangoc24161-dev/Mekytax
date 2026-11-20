---
type: decision
created: 2026-11-20
updated: 2026-11-20
source: chat 2026-11-20 (mwnu8fh5tv6n99vy)
---
Sheet `NHAP_HO_SO` (`sheet2.xml`) không mở được khi `sharedStrings.xml` lệch `count` so với `uniqueCount`/`<si>` (vd count=2846, unique=2840) và thiếu khai báo `xmlns:r` / `xmlns:xr2` / `xmlns:xr3` dù `mc:Ignorable` liệt kê xr2 xr3.

**Vì sao:** Excel Windows từ chối sheet khi SST/namespace không khớp.

**Áp dụng:** Trước khi ghi đè Drive, bắt buộc `count=uniqueCount=số <si>`; header sheet2 phải đủ namespace như sheet1; giữ ZipInfo Windows (`create_system=3`, `external_attr=0x81a40000`); không serialize lại sheet bằng ElementTree.
