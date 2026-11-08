---
type: business
provenance: source
origin: javis-learned
created: 2026-11-02
updated: 2026-11-02
---
Sheet `DASHBOARD` trong file `Bao_cao_cong_van.xlsx` tự động cập nhật số liệu theo dữ liệu từ sheet `NHAP_HO_SO`.

- Cơ chế: Có khoảng 235 ô công thức (`SUMIFS`, `COUNTIFS`, `SUM`, `IF`, `INDEX`...) kéo dữ liệu trực tiếp từ `NHAP_HO_SO`.
- Cột hỗ trợ: Sheet `NHAP_HO_SO` có các cột ẩn `AM`–`AU` (`CV_FIRST`, `CV_TRẠNG_THAI`, `CV_NAM`, `CV_THANG`, `LOC_KPI`...) đóng vai trò tính toán trung gian để Dashboard tổng hợp và lọc.
- Các sheet `TAM_NGUNG`, `HET_TAM_NGUNG`, `DANH_SACH_LOC`, `DANH_MUC` chỉ đóng vai trò liên kết (hyperlink), không cung cấp số liệu tổng hợp lên Dashboard.
