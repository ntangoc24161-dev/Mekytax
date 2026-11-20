---
type: source
source_kind: own-note
created: 2026-11-20
updated: 2026-11-20
---

# Sửa lỗi mở sheet NHAP_HO_SO - 20/11/2026 15:45

## Nguyên nhân còn lại sau lần sửa sáng
1. `xl/sharedStrings.xml` lệch `count="2846"` trong khi `uniqueCount` và số thẻ `<si>` = **2840**.
2. Header `sheet2.xml` (NHAP_HO_SO) thiếu `xmlns:r`, `xmlns:xr2`, `xmlns:xr3` dù `mc:Ignorable="x14ac xr xr2 xr3"`.

## Đã xử lý
- Vá SST: `count=uniqueCount=2840`.
- Bổ sung đủ namespace trên sheet2.
- Đóng gói ZipInfo Windows (`create_system=3`, `external_attr=0x81a40000`).
- Giữ O901/O902 = `Đang xử lý`, giữ 2 chart DASHBOARD.
- **Ghi đè đúng 1 file Drive** `Bao_cao_cong_van.xlsx` (id `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`), md5 `a62250c950a5527416796b4f390f2ad4`, size 762787.
- Bản vault canonical: `ho-so-cv-moi-2026-11/Bao_cao_cong_van.xlsx`.
- Đã gửi qua bot Telegram CA `@Kiem_soat_dieu_phoi_bot`.

## Quy ước
Không tạo thêm bản `Bao_cao_cong_van_*_YYYY-MM-DD.xlsx` trên Drive. Chỉ UPDATE file sống.
