---
type: preference
created: '2026-10-02'
updated: '2026-10-04'
---
Người dùng yêu cầu tạo bộ thư mục mẫu chuẩn hóa theo Quy trình 4 bước (`thu-muc-mau-ho-so-doanh-nghiep/`) để áp dụng cho mọi doanh nghiệp kiểm tra mới:
- **Cấu trúc 4 thư mục chính:**
  + `1. DU LIEU THO NGANH THUE/`: Chứa các thư mục con `1.1. To khai XML/`, `1.2. Hoa don dien tu/` (file `HDDT.xlsx` tập hợp danh sách hóa đơn điện tử mua vào và bán ra), `1.3. Nghia vu nop thue & Thong bao/`, `1.4. Tong hop danh sach NLD theo QTT TNCN/` (tổng hợp danh sách người lao động theo quyết toán thuế TNCN hàng năm).
  + `2. HO SO NNT CUNG CAP/`: Chứa các thư mục con `2.1. So ke toan tong hop va chi tiet/`, `2.2. Kho - Chi phi - Gia thanh/`, `2.3. Ngan hang & Dong tien/`, `2.4. Lao dong - Tien luong - TSCD/`, `2.5. Bang ke mua vao ban ra/` (tập hợp dữ liệu bảng kê mua vào bán ra do công ty cung cấp).
  + `3. DU LIEU BO SUNG & TRA CUU QLT/`: Chứa các thư mục con `3.1. Canh bao rui ro CVCBRR & Cong an/`, `3.2. Danh ba trang thai NNT/` (gồm 01 file Excel tập hợp danh sách đối tác mua vào bán ra theo danh sách HĐĐT và Bảng kê để tra cứu trạng thái ĐKKD), `3.3. Tra cuu TNCN da nguon & Nhan su/`, `3.4. Giai trinh & Ho so bo sung/`.
  + `4. KET QUA PHAN TICH/`: Chứa file Master Excel 7 Sheet, file Master Checklist `CHECKLIST_TONG_HOP_HO_SO.xlsx`, thư mục `cau-thanh/`, `sao-luu/` và Đề cương kiểm tra.
- **Checklist và Kịch bản quét tự động:** Duy trì duy nhất 01 file Master Checklist `CHECKLIST_TONG_HOP_HO_SO.xlsx` tại thư mục kết quả `4. KET QUA PHAN TICH/` và kịch bản `kiem_tra_tinh_day_du_ho_so.py` tự động quét cây thư mục, chấm điểm thiếu/đủ và báo cáo các hồ sơ trọng yếu cần bổ sung.

**Vì sao:** Chuẩn hóa cấu trúc tiếp nhận tài liệu cho đoàn kiểm tra, giúp phát hiện ngay các khoảng trống dữ liệu trước khi tiến hành phân tích chuyên sâu.
**Áp dụng:** Khi mở hồ sơ kiểm tra cho bất kỳ doanh nghiệp nào, sao chép thư mục mẫu này sang `<TEN_DN>_<MST>/`, đổ dữ liệu vào các thư mục con và chạy kịch bản quét trước khi phân tích Bước 4.
