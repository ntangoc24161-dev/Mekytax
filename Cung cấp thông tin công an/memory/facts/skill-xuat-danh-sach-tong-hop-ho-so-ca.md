---
type: decision
created: '2026-09-26'
updated: '2026-09-26'
---
Quy chuẩn kỹ năng xuất danh sách tổng hợp chi tiết hồ sơ công văn Công an:
- **Tên kỹ năng:** `xuat-danh-sach-ho-so-ca` (Xuất danh sách tổng hợp chi tiết hồ sơ Công an).
- **Mục tiêu:** Tự động trích xuất dữ liệu từ sổ theo dõi `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`), các file markdown công văn tại `D:\HSO CONG AN FILE MD` và CSDL `DB NNT.xls` để tạo tệp Excel tổng hợp chi tiết.
- **Vị trí lưu:** `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI\Danh_sach_tong_hop_ho_so_cung_cap_CV_MOI.xlsx` và đồng bộ bản ngắn `CV MOI.xlsx`.
- **Cấu trúc cột tiêu chí (24 cột):** Tách độc lập từng loại tài liệu thành từng cột kèm thời gian/kỳ thuế cụ thể:
  1. Tờ khai thuế GTGT (kỳ tháng/quý/năm yêu cầu)
  2. Tờ khai Quyết toán thuế TNDN (năm tài chính)
  3. Tờ khai Quyết toán / Khấu trừ thuế TNCN (năm quyết toán hoặc quý khấu trừ)
  4. Báo cáo tài chính (BCTC) (năm lập BCTC)
  5. Hóa đơn điện tử / Bảng kê HĐ (mua vào - bán ra, gửi email CA, bóc tách HĐ)
  6. Quyết định Xử phạt VPHC / Cưỡng chế nợ (giai đoạn rà soát)
  7. Thông báo Đóng MST / Bỏ địa chỉ kinh doanh (số/ngày TB)
  8. Thông tin Định danh & Pháp lý NNT (ĐKKD, ĐDPL, CCCD, Kế toán, CBQL, TK ngân hàng - chỉ lập khi CA yêu cầu)
  9. Nghĩa vụ nộp NSNN / Nợ thuế (số nợ hoặc số tiền đã nộp)
  10. Hồ sơ tài liệu khác / Nội dung đặc thù (hoàn thuế GTGT, biên bản thanh tra KT, phiếu chuyển Tổ...)
- **Cấu trúc Sheet:** `TONG_HOP_CHI_TIET` (chi tiết theo DN/MST), `TONG_HOP_THEO_CV` (tổng hợp theo công văn), `TOAN_BO_HO_SO_DU_THAO` (danh mục toàn bộ hồ sơ đang dự thảo), `DANH_MUC_TIEU_CHI` (hướng dẫn tra cứu từng cột).
- **Vì sao:** Giúp cán bộ kiểm tra và Lãnh đạo nắm bắt nhanh chóng, chính xác khối lượng và loại tài liệu thuế cụ thể cần xuất từ hệ thống TMS / HĐĐT mà không phải đọc lại từng bản scan PDF.
- **Áp dụng:** Tự động kích hoạt khi có lệnh xuất danh sách tổng hợp chi tiết tài liệu hoặc khi kết xuất báo cáo định kỳ.
