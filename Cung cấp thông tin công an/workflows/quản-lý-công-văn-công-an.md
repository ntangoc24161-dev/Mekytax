---
type: workflow
name: Quản lý công văn công an
slug: quản-lý-công-văn-công-an
status: active
group: Tài chính
description: Tiếp nhận, bóc tách công văn, kiểm soát chéo dữ liệu, phân công in ấn KT1, lập dự thảo phúc đáp và nghiệm thu đồng bộ thư mục.
steps:
- agent: khai-thac-du-lieu-ho-so
  task: Tiếp nhận công văn từ thư mục CV MOI; bóc tách file .md lưu D:\HSO CONG AN FILE MD; tra cứu CSDL DB NNT.xls và cổng thông tin; phân Tổ KT; cập nhật Bao_cao_cong_van.xlsx (sheet NHAP_HO_SO); ghi tên CBQL Tổ KT1 nếu công văn có doanh nghiệp thuộc TCS3.
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
- agent: theo-doi-phan-cong
  task: Rà soát công văn Mới; tra cứu CBQL Tổ KT1; thực hiện in phân công KT1 toàn văn kèm dòng thông tin đầu trang chẵn; lưu file tổng hợp lên Google Drive; chốt chuyển trạng thái sang Đang xử lý.
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
- agent: soan-thao-tra-loi-cong-an
  task: Lập dự thảo công văn trả lời và tờ trình 1 trang A4 đúng thể thức NĐ 30/2020 cho các DN thuộc Tổ KT1, KT2, KT3; số liệu lấy từ DB NNT; tô đỏ thông tin chưa xác định; lưu vào thư mục Drive.
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
- agent: quản-lý-hồ-sơ-công-an
  task: Lưu trữ PDF gốc vào thư mục năm 2026 theo cấu trúc STT.CV-CQ-TO KT; tự động tạo dự thảo hồ sơ trả lời (nếu KT1); xuất danh sách 24 cột tiêu chí khi có yêu cầu; rà soát ký số và cập nhật hậu tố -R cho thư mục hoàn tất; báo cáo nghiệm thu hoàn thành.
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
updated: '2026-11-06'
---
