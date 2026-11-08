---
name: "Đối chiếu NCC mua vào với CVCBRR và danh bạ"
description: "Đối chiếu NCC bảng kê mua vào với sổ CVCBRR (F0, F1), sổ Công an và danh bạ NNT gom 1 sheet kèm chi tiết hóa đơn."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-09-30
updated: 2026-10-01
---
## Khi nào dùng

- Đối chiếu nhà cung cấp / bên bán trên bảng kê **mua vào** với Sổ công văn cảnh báo rủi ro thuế (CVCBRR - Sheet F0, Sheet F1), Sổ theo dõi hồ sơ Công an (`Bao_cao_cong_van.xlsx`), và Danh bạ NNT (trạng thái ĐKKD).
- Xuất file Excel đối chiếu chuẩn gom trọn vẹn trên sheet **`Tong_hop_doi_chieu`** kèm sheet phụ **`Chi_tiet_hoa_don`** (toàn bộ dòng HĐ mua vào nhiều năm sắp theo thứ tự thời gian).
- Yêu cầu kiểm tra rủi ro doanh nghiệp đang xét (F1 + Công an) và rà soát nhà cung cấp mua vào (F0 + trạng thái ĐKKD bất thường).

## Nguồn dữ liệu & Chuẩn bị

1. **Bảng kê mua vào**: `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx` (sheet `mua vao` đã gộp nhiều năm).
2. **Sổ CVCBRR**: `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` (Sheet `F0` - bên bán cảnh báo, Sheet `F1` - bên mua cảnh báo).
3. **Sổ theo dõi Công an**: `Bao_cao_cong_van.xlsx` (Sheet `NHAP_HO_SO`, `TAM_NGUNG`, `DASHBOARD`).
4. **Danh bạ NNT**: Dữ liệu danh bạ trạng thái hoạt động người nộp thuế (đang hoạt động, tạm ngừng KD, đóng MST, không hoạt động tại địa chỉ ĐK...).

## Cấu trúc Báo cáo đối chiếu gom 1 Sheet

File Excel gồm 2 sheet:

### 1. Sheet chính: `Tong_hop_doi_chieu`
- **Phần 1: Đối chiếu Người nộp thuế đang xét**:
  + Bảng đối chiếu NNT với Sheet `F1` (CVCBRR Bên mua) và Sổ Công an.
  + Đưa ra kết luận rõ ràng NNT có/không có trong danh sách cảnh báo hoặc hồ sơ Công an.
- **Phần 2: Đối chiếu Danh sách Nhà cung cấp mua vào**:
  + **Bảng 2.1: Đối chiếu với Sổ CVCBRR F0 (Bên bán)**: Liệt kê các NCC có công văn cảnh báo rủi ro thuế / hóa đơn / Công an (STT, MST, Tên NCC, Số HĐ, Doanh số, Thuế GTGT, Kỳ mua, Mặt hàng chính, Nội dung cảnh báo & Số CV, Trạng thái ĐKKD & Đánh giá rủi ro).
  + **Bảng 2.2: Đối chiếu với Danh bạ NNT (Trạng thái bất thường)**: Liệt kê các NCC đã đóng MST, tạm ngừng kinh doanh hoặc bỏ trốn / không hoạt động tại địa chỉ đăng ký.
- **Phần 3: Đề xuất trọng tâm kiểm tra tại doanh nghiệp**:
  + Trọng tâm rà soát hóa đơn, biên bản giao nhận hàng hóa, phiếu cân, định mức tiêu hao, chứng từ thanh toán ngân hàng đối với các NCC có cảnh báo hoặc đã đóng MST.

### 2. Sheet phụ: `Chi_tiet_hoa_don`
- Chi tiết toàn bộ các dòng hóa đơn mua vào (1.000 - 2.000+ dòng) trong toàn bộ thời kỳ kiểm tra.
- Sắp xếp theo thứ tự **`Ngay_HD` tăng dần (thời gian)** để phục vụ tra cứu nhanh khi làm việc thực tế.
- Gắn đầy đủ thông tin cảnh báo F0 và trạng thái ĐKKD trên từng dòng hóa đơn.

## Kỹ thuật Excel & Xử lý lỗi XML

1. **Tuyệt đối tránh lỗi Formula XML**: Các ô ghi chú kết luận không được bắt đầu bằng dấu `=`, `+`, `-` (ví dụ: đổi `=> KẾT LUẬN...` thành `> KẾT LUẬN...` hoặc set `cell.data_type = 's'`) để tránh openpyxl gán nhầm sang thẻ formula `<f>`, gây lỗi *Removed Records: Formula from /xl/worksheets/sheet1.xml part*.
2. **Định dạng**: Font **Times New Roman** 100% file; tiêu đề cột in đậm; số tiền phân cách hàng nghìn chuẩn VN; dòng CỘNG dùng hàm `SUM(...)`.
3. **Hiển thị đầy đủ**: Bật `wrap_text = True`, căn chỉnh chiều cao hàng và độ rộng cột tự động.

## Bẫy cần tránh

- Nhầm lẫn giữa tổng hợp **bên mua** (F1) với đối chiếu **bên bán** (F0).
- Khớp sót do MST có hoặc không có dấu gạch chi nhánh (phải chuẩn hóa so khớp cả MST 10 số và 13 số).
- Sort sheet chi tiết theo MST thay vì theo thời gian `Ngay_HD`.
- Để ô text bắt đầu bằng dấu `=` khiến file bị hỏng khi mở trong Microsoft Excel.
