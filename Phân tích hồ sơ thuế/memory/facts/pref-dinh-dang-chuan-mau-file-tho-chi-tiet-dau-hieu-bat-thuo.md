---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-23
updated: 2026-10-23
---
Quy chuẩn định dạng cho file `KQ3_NNT_Giai_trinh_Chi_tiet_<MST>.xlsx` và các báo cáo đối chiếu giải trình:
- **Font & cỡ chữ:** Times New Roman toàn bộ; dữ liệu 10pt, tiêu đề cột 10-11pt đậm, tiêu đề sheet 14pt đậm.
- **Bảng màu tiêu đề:** Tiêu đề bảng nền xanh Navy đậm (`#1F4E78`), chữ trắng in đậm, căn giữa; tiêu đề phân đoạn nền xanh nhạt (`#D9E1F2`), chữ xanh đậm.
- **Định dạng số & ngày:** Ngày `dd/mm/yyyy` căn giữa (không giờ); số tiền `#,##0` căn phải; tỷ lệ phần trăm `0.00%` căn phải.
- **Lưới & độ rộng cột:** Viền mỏng `#D9D9D9`, bật hiển thị đường lưới (`Gridlines = True`) trên toàn bộ các sheet; auto-fit độ rộng cột tối ưu theo dữ liệu để không bị lỗi `###`.
- **Màu cảnh báo rủi ro:**
  - Đỏ đào (`#FCE4D6` / chữ đỏ `#C00000`): NCC F0/Công an, hóa đơn bán dưới giá vốn.
  - Vàng cam (`#FFF2CC` / chữ cam `#B25900`): NCC có trạng thái ĐKKD bất thường.
  - Xanh lá nhạt (`#E2EFDA` / chữ xanh `#385723`): Thông tin tích cực (không phát sinh chi phí lãi vay/EBITDA an toàn).
