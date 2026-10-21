---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
---
Người dùng yêu cầu áp dụng thống nhất các chuẩn định dạng cho toàn bộ file kết quả phân tích hồ sơ thuế:
- Ngày tháng: định dạng `dd/MM/yyyy`, không hiển thị phần giờ.
- Số tiền: phân cách hàng nghìn bằng dấu chấm `.`, số thập phân bằng dấu phẩy `,`.
- Phông chữ: mặc định Times New Roman, bảng mã Unicode (nếu gặp font sai mã như VNI/TCVN3 phải tự động chuyển đổi sang Unicode).
- Trình bày: chữ đứng, tô đậm các mục/tiêu đề rõ ràng, tự động canh chỉnh độ rộng cột (`auto-fit`) để hiển thị đầy đủ nội dung ô tính.
- Dòng dữ liệu: xóa toàn bộ các dòng năm trống không có số liệu phát sinh; dòng ví dụ dưới tiêu đề mẫu phải xóa bỏ khi điền dữ liệu thực tế.
- Bố cục sheet `IN05_NLD_TNCN`: phân chia theo từng năm với dòng quyết toán (QTT) đặt trước các dòng danh sách người lao động.
