---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-09
updated: 2026-11-09
source: chat 2026-11-09 (0e07ff7878014679b9f64051329a97ca)
---
Hồ sơ kiểm tra Công ty TNHH Thương mại Dịch vụ và Sản xuất Huy Thành Phát (MST 3603820015) tại `Dang kiem tra\KE HOACH\HUY_THANH_PHAT_3603820015`.

- Đã tạo `0. BIEN BAN\Bang tinh.xlsx` từ mẫu `1. TD\BIEN BAN\Bang tinh.xlsx`.
- User chốt kỳ kiểm tra: **Từ năm 2022 đến năm 2025** (mặc định KE HOACH 2026; MST chưa có trên sổ `Danh sach kiem tra quyet toan pro.xlsm`). Header CDKT!F5:F8 và KQKDSXKD khớp MST, tên NNT, địa bàn Thành phố Biên Hòa - Thuế cơ sở 3 tỉnh Đồng Nai.
- Đã đổ KQKDSXKD 2022-2024 từ PTDN02, năm 2025 từ QTT 03-1A L01 (PTDN02 năm 2025 toàn 0). Không ghi cột 2026.
- Đã đổ CDKT và `CĐKT ` năm 2022-2024 từ PTDN01 (map TT200 sang TT133). Năm 2025 để trống: PTDN01 toàn 0, không có BCTC NamNay 2025. Sheet `CĐKT ` không có cột 2025. Năm 2021 PTDN có số nhưng ngoài kỳ, đã xóa số mẫu.
- Backup: `Bang tinh_backup_before_nhap_20261109_032126.xlsx` (trước KQKD), `Bang tinh_backup_before_cdkt_20261109_033706.xlsx` (trước CDKT).
- Chưa có KQ2.1 nên chưa đổ GTGT/EBITDA. Chưa bóc tách DI-01.

**Vì sao:** User lần lượt lệnh tạo phôi, nhập KQKDSXKD chỉ đến 2025, rồi đổ sổ CDKT.
**Áp dụng:** Không bịa số 0 cho CDKT 2025; GTGT/EBITDA và DI-01 chỉ làm khi user yêu cầu hoặc có KQ2.1.
