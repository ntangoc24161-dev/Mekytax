---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-09-30
updated: 2026-09-30
source: [[conversations/2026-09-30]]
---
Tách riêng với đối chiếu NCC bên bán trên bảng kê mua vào: lấy cột **đơn vị / MST bên mua** trên sổ theo dõi công văn cảnh báo rủi ro thuế (`SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`, sheet `THEO_DOI`) rồi xuất Excel tổng hợp. [[conversations/2026-09-30]]

### Phạm vi dòng
Chỉ lấy dòng sổ **đã có** thông tin bên mua; trên sổ, bên mua thường chỉ được nhập khi công văn nêu. (thực tế tính đến 2026-09-30) [[conversations/2026-09-30]]

### Đầu ra chuẩn (mục tiêu theo yêu cầu user)
- **`Chi_tiet`:** mỗi dòng có bên mua, kèm CV và bên bán; sắp theo ngày công văn → MST mua → số CV. [[conversations/2026-09-30]]
- **`Tong_hop`:** gom theo MST gốc bên mua; sắp theo số dòng giảm dần. [[conversations/2026-09-30]]
- Font: **Times New Roman**. [[conversations/2026-09-30]]

### Đọc kết quả
Cột doanh thu/thuế trên sổ có thể gần như trống; không suy doanh số mua từ sổ nếu hầu hết dòng trống. (thực tế tính đến 2026-09-30) [[conversations/2026-09-30]]
Tên chứa cụm từ giống NNT đang kiểm tra **không** đủ để kết luận; phải khớp đúng MST ở cột bên mua. (cần xác minh từng hồ sơ) [[conversations/2026-09-30]]
