---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
---
Trong quy trình phân tích hồ sơ thuế, bước Lệnh 2 (Chuẩn hóa dữ liệu đầu vào do Chuyên viên dữ liệu đầu vào thực hiện) được chốt như sau:
1. Chuẩn hóa MST theo quy tắc BT6 và quét lệch BT5-BT9 bằng skill `lam-sach-cd1-mst-hoa-don`, xuất file chuyên đề `Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx`.
2. Gọi skill `gop-hoa-don-mua-ban` 2 lần độc lập:
   - Lần 1: Input file HĐĐT → Output `HDDT_<MST>_<KY>.xlsx` (2 sheet Mua vào / Bán ra).
   - Lần 2: Input file Bảng kê → Output `BangKe_<MST>_<KY>.xlsx` (2 sheet Mua vào / Bán ra).
3. Dữ liệu chuẩn hóa gồm KQ1.1, KQ1.2, file BT5-BT9 và 2 file gộp riêng làm đầu vào chuyển tiếp cho bước Soát sổ sách doanh nghiệp (Lệnh 3). Chưa tạo skill gộp 4 sheet khi người dùng chưa yêu cầu.
