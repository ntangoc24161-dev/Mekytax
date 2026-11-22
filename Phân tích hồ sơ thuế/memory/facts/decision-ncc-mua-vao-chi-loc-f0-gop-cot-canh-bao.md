---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-20
updated: 2026-11-20
---
- **Quy tắc phân loại rủi ro:**
  - **Nhà cung cấp mua vào:** Chỉ đối soát và gắn cờ rủi ro đối với danh sách **F0** (Sổ CVCBRR sheet `F0`). Bỏ qua cờ F1 và Công văn Công an cho NCC mua vào.
  - **Người nộp thuế (công ty đang kiểm tra):** Tra cứu đầy đủ F1, CVCBRR và Sổ Công an.
  - **Tiêu chí TC2 (Giai đoạn 3):** Giữ nguyên quy tắc và logic chấm điểm cũ.
- **Bố cục cột hiển thị Excel:** Gộp chung thành 1 cột duy nhất **`Cảnh báo rủi ro`**, ghi trực tiếp *Số công văn, ngày ban hành và cơ quan ban hành* khi dính F0 (để trống nếu bình thường), không tách cột phụ thông tin công văn.
- **Đồng bộ hệ thống:** Đã áp dụng trên file kết quả `Tong_hop_HDDT_BangKe_3600840461_2020-2025.xlsx`, kỹ năng `tong-hop-hddt-bang-ke-4-sheet`, `doi-chieu-ncc-mua-vao-rui-ro`, `giai-doan-2-phan-tich-du-lieu-co-quan-thue`, `lam-sach-cd1-mst-hoa-don` và agent `phan-tich-vien-cqt`.
