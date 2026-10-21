---
type: preference
created: 2026-10-05
updated: 2026-10-05
---
**Quy trình đối chiếu tự động Doanh thu (TK 511) và Giá vốn (TK 632) cho doanh nghiệp thương mại:**
- **Đầu vào chuẩn hóa:**
  1. Trích xuất sổ chi tiết TK 511 (PS Có) và TK 632 (PS Nợ) từ Sổ Nhật ký chung (NKC) hoặc Sổ cái chi tiết TK 511, 632.
  2. Bổ sung thông tin mặt hàng từ Bảng kê bán ra (`ban ra`) để đảm bảo đầy đủ tên hàng tiếng Việt có dấu chuẩn Unicode.
  3. Chuẩn hóa số hóa đơn (bỏ tiền tố 0 đầu: `00000001` -> `1`), chuẩn hóa định dạng ngày `dd/mm/yyyy`.
- **Quy tắc ghép 3 khóa tự động:**
  - Khóa 1: Số HĐ + Ngày CT + Mặt hàng (Khớp tuyệt đối).
  - Khóa 2: Số HĐ + Mặt hàng (Khác ngày CT - xuất kho khác ngày lập HĐ).
  - Khóa 3: Số HĐ + Ngày CT (Ghép theo chứng từ khi diễn giải 632 ghi `Xuất bán theo HĐ : <so_hd>`).
  - Nhận diện bất thường: `Không khớp 632` (có DT không GV) hoặc `Có giá vốn, không có doanh thu` (có GV không DT).
- **Phân tích tài chính & Phát hiện bán dưới giá:**
  - Tính lãi/lỗ từng hóa đơn: P = Doanh thu 511 - Giá vốn 632.
  - Tự động lọc và báo cáo toàn bộ dòng bán dưới giá (P <= 0: bán lỗ hoặc hòa vốn).
  - Tính tỷ lệ lãi gộp qua từng năm và tỷ trọng doanh thu bán dưới giá / Tổng doanh thu.
- **Hệ thống file kết quả & Lưu trữ:**
  - File Tool độc lập: `4. KET QUA PHAN TICH/cau-thanh/Doi_chieu_511_632_<MST>_<ky>.xlsx` (và `.xlsm`).
  - Tích hợp Sheet `Doi_chieu_511_632` vào Master Excel (`Phan_tich_ho_so_cong_ty_<MST>.xlsx`).
  - Đồng bộ báo cáo Markdown và file kết quả vào `4. KET QUA PHAN TICH/`, `7. Sao luu tong hop/` trên máy và Second Brain (`ket-qua/`, `sao-luu/`).

**Vì sao:** Doanh nghiệp thương mại có số lượng giao dịch mua bán lớn, giá vốn hạch toán trực tiếp từ kho theo từng hóa đơn; việc đối chiếu tự động 511 - 632 giúp phát hiện nhanh chóng các giao dịch bán lỗ/bán dưới giá vốn, trích sai giá xuất kho hoặc thiếu giá vốn đối ứng.
**Áp dụng:** Áp dụng chuẩn hóa cho mọi hồ sơ doanh nghiệp thương mại trong các kỳ kiểm tra tiếp theo.
