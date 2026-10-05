---
type: preference
created: '2026-10-02'
updated: '2026-10-02'
---
**Quy chuẩn định dạng in ấn Excel Checklist 1 trang A4 cho hồ sơ kiểm tra:**
- **Kích thước & Căn lề:** Khổ giấy A4 (`PAPERSIZE_A4 = 9`), thiết lập `Page Setup: Fit to 1 page wide x 1 page tall` (`fitToWidth=1`, `fitToHeight=1`, `fitToPage=True`), lề hẹp (0.3 - 0.4 inch) và bật hiển thị đường lưới (`showGridLines=True`).
- **Hình thức văn bản hành chính:**
  + Góc trên bên trái: Tên Cục Thuế / Chi cục Thuế / Đoàn kiểm tra.
  + Góc trên bên phải: Quốc hiệu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" và Tiêu ngữ "Độc lập - Tự do - Hạnh phúc".
  + Tiêu đề in hoa, đậm, font Times New Roman, màu xanh Navy `#1F4E79`.
  + Khối thông tin doanh nghiệp: Tên DN, MST, Niên độ kiểm tra, Quyết định kiểm tra, Trưởng đoàn, Ngày rà soát.
- **Cấu trúc bảng dữ liệu:**
  + Header bảng: Nền xanh `#2F5597`, chữ trắng đậm.
  + Dòng xen kẽ: Nền `#F2F4F8` nhẹ nhàng, kẻ viền ô mỏng rõ ràng.
  + Cột trạng thái: Ô tích chọn `[ ] Đủ   [ ] Thiếu` hoặc `[  ]`.
  + Hộp kết luận đánh giá tình trạng hồ sơ: Đủ điều kiện phân tích Bước 4 / Yêu cầu bổ sung tài liệu.
  + Ô ký duyệt 2 cấp cuối trang: Cán bộ phân tích / rà soát và Trưởng đoàn kiểm tra.
- **Hệ thống file áp dụng:** Master Workbook 6 Sheet (`CHECKLIST_TONG_HOP_HO_SO.xlsx`) và 4 file checklist chuyên đề riêng lẻ trong từng thư mục con (`1.`, `2.`, `3.`, `4.`).
