---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
---
Người dùng chốt điều chỉnh cấu trúc các biểu mẫu phân tích hồ sơ thuế:
- `IN02_BangKe`: phải trích lọc đầy đủ cả bảng kê mua vào và bán ra, chữ đứng, in đậm rõ ràng.
- `KQ2_Doi_chieu_GTGT`: Phần I đối chiếu theo kỳ lấy dữ liệu từ tờ khai GTGT các kỳ (tháng hoặc quý tùy kỳ kê khai của doanh nghiệp).
- `KQ2_NCC_mua_vao`: chuyển thành bảng đối chiếu người bán với chi tiết toàn bộ (100%) hóa đơn mua vào hợp lệ, hiển thị thông tin công văn cảnh báo, mức độ rủi ro và tình trạng hoạt động DN từ `IN07_NCC_CanhBao`.
- `KQ2_511_632`: thay thế toàn bộ bằng nội dung sheet `Doi_chieu_511_632` từ file phân tích hồ sơ chuẩn cũ (bảng tổng hợp theo năm kèm đánh giá tuân thủ & rủi ro).
- `KQ3`: bổ sung lại sheet narrative `Danh_gia_28_tieu_chi_rui_ro` từ mẫu phân tích hồ sơ cũ bên cạnh bảng 80 quy tắc rủi ro; đưa nhận xét/đánh giá phân tích hồ sơ vào các sheet tương ứng.
