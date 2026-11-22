---
type: workflow
name: Theo dõi & cập nhật CV trả lời đã ký
slug: theo-doi-cv-tra-loi-da-ky
status: active
group: Tài chính
description: Xử lý CV trả lời đã ký từ CV MOI (cron 12h) hoặc quét năm 2026 theo lệnh; cập nhật sổ Đang xử lý → Đã hoàn thành.
steps:
- agent: theo-doi-cv-tra-loi-da-ky
  task: "Khi cron 12h đưa file bản ký từ CV MOI: xử lý đúng file đó trước. Khi có lệnh theo dõi / quét rộng: quét thư mục công văn năm 2026 tìm CV đã ký (-R hoặc file ký số). Trích số CV, cơ quan, Tổ KT; kiểm tra sổ NHAP_HO_SO; bóc tách file .md; xác nhận danh sách DN khớp sổ; cập nhật trạng thái từ Đang xử lý → Đã hoàn thành bằng scripts/ghi-so-bao-cao-an-toan.py set-status (VPS) rồi UPLOAD_UPDATE fileId sống; xuất danh sách tổng hợp; báo cáo lỗi nếu có."
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
updated: '2026-11-18'
---

## Kích hoạt
- **Tự động (cron 12h nhánh B):** file trong `CV MOI` đạt checklist bản ký (`…/TCS3-KTr…`, `TRA LOI`/`.signed`, OCR Thuế trả lời + chữ ký/mộc). Xử lý đúng file đó trước.
- **Lệnh riêng:** quét rộng thư mục năm 2026 hoặc theo dõi một hồ sơ cụ thể.
- Checklist đầy đủ: `Javis/loops/ho-so-cung-cap-cong-an.md` và agent `theo-doi-cv-tra-loi-da-ky`.
