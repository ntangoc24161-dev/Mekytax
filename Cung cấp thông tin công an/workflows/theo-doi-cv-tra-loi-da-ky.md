---
type: workflow
name: Theo dõi & cập nhật CV trả lời đã ký
slug: theo-doi-cv-tra-loi-da-ky
status: active
group: Tài chính
description: Quét thư mục công văn năm 2026, phát hiện CV trả lời đã ký số, bóc tách dữ liệu, xác nhận danh sách DN, cập nhật trạng thái sổ theo dõi.
steps:
- agent: theo-doi-cv-tra-loi-da-ky
  task: Quét toàn bộ thư mục công văn năm 2026 tìm CV đã ký (-R hoặc file ký số); trích số CV, cơ quan, Tổ KT; kiểm tra sổ NHAP_HO_SO; bóc tách file .md; xác nhận danh sách DN khớp sổ; cập nhật trạng thái từ "Đang xử lý" → "Đã hoàn thành"; xuất danh sách tổng hợp; báo cáo lỗi nếu có.
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
updated: '2026-11-10'
---
