---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-09
updated: 2026-11-09
---
User chọn phương án gộp skill `doi-chieu-mat-hang-mua-ban` vào agent `phan-tich-vien-so-sach-dn` (Soát sổ sách công ty).

- **Phạm vi mới của vai Soát sổ sách:** 511-632, kho (XNT), mặt hàng (lọc HĐ hủy, bắt chuỗi trung gian, khớp tiền-số lượng-đơn giá), giá thành, công nợ, tiền mặt.
- **Output:** Giữ nguyên 1 file tổng hợp KQ2.2 mở rộng thay vì sinh thêm file rời.
- **Cấu trúc Lệnh 2:** Tinh gọn còn 3 vai chạy song song (Chuẩn hóa + Đối chiếu dữ liệu thuế + Soát sổ sách mở rộng), giám sát nghiệm thu 2 file KQ2.1 và KQ2.2 mở rộng trước khi sang Lệnh 3.
