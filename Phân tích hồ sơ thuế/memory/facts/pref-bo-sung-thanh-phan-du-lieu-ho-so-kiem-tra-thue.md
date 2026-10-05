---
type: preference
created: '2026-10-04'
updated: '2026-10-04'
---
Người dùng yêu cầu bổ sung và chuẩn hóa các thành phần dữ liệu mới vào quy trình 4 bước và bộ thư mục hồ sơ kiểm tra thuế:
- **Bước 1 (Dữ liệu thô ngành Thuế):**
  + Bổ sung **File HDDT là tập hợp danh sách hóa đơn điện tử** (mua vào, bán ra từ hệ thống/portal HĐĐT ngành Thuế) đặt tại `1.2. Hoa don dien tu/`.
  + Bổ sung **Tổng hợp danh sách người lao động theo quyết toán thuế TNCN hàng năm** (kết xuất/tổng hợp từ tờ khai QTT TNCN 05/QTT-TNCN và phụ lục các năm) đặt tại `1.4. Tong hop danh sach NLD theo QTT TNCN/`.
- **Bước 2 (Hồ sơ NNT cung cấp):**
  + Bổ sung mục **Bảng kê** gồm tập hợp dữ liệu bảng kê mua vào bán ra do công ty cung cấp (trích xuất từ phần mềm kế toán NNT) đặt tại `2.5. Bang ke mua vao ban ra/`.
- **Bước 3 (Dữ liệu bổ sung & Tra cứu QLT):**
  + Bổ sung **01 file Excel tập hợp danh sách đối tác mua vào bán ra theo danh sách hóa đơn điện tử và bảng kê** đặt tại `3.2. Danh ba trang thai NNT/` phục vụ tra cứu trạng thái ĐKKD (đang hoạt động, bỏ trốn, ngừng hoạt động...), đối chiếu mã số thuế và rà soát rủi ro CVCBRR F0/F1, Công an.

**Vì sao:** Phân định rành mạch nguồn dữ liệu gốc của Cơ quan Thuế (HĐĐT portal, QTT TNCN) và dữ liệu do NNT tự kê khai/cung cấp (Bảng kê mua vào bán ra), đồng thời tích hợp danh sách đối tác toàn diện để tra cứu tự động.
**Áp dụng:** Áp dụng cho bộ thư mục mẫu chuẩn hóa `thu-muc-mau-ho-so-doanh-nghiep/`, cây thư mục 13 doanh nghiệp kế hoạch kiểm tra 2026 và mọi quy trình kiểm tra thuế tiếp theo.
