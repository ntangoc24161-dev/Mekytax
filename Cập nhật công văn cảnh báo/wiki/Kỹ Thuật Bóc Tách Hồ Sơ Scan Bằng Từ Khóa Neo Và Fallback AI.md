---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-22
updated: 2026-11-22
source: [[conversations/2026-11-22]]
---
Phương pháp tối ưu chi phí và tốc độ bóc tách hồ sơ scan thuế không quy chuẩn: kết hợp OCR offline cục bộ với thuật toán neo từ khóa (anchor-based extraction) thay vì khớp mẫu câu cứng [[conversations/2026-11-22]]. Dữ liệu nhận diện được lọc qua biểu thức chính quy (Regex) bắt các trường cố định (MST, Số CV, Ngày ban hành, Cơ quan ký, Tên DN), phân loại độ tin cậy và xuất ra bảng kiểm duyệt trung gian trước khi nạp sổ chính [[conversations/2026-11-22]]. Cơ chế dự phòng (fallback) phân tầng: chỉ chuyển các hồ sơ bị gắn cờ cảnh báo (quá mờ, cấu trúc phức tạp hoặc thiếu trường) sang mô hình AI xử lý đơn lẻ, giúp tiết kiệm trên 95% token so với bóc tách toàn bộ qua LLM (thực tế tính đến 22/11/2026) [[conversations/2026-11-22]].
