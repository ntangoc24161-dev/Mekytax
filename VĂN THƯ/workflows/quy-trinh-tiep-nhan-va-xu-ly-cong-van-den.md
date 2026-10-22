---
type: workflow
name: Quy trình tiếp nhận và xử lý công văn đến
slug: quy-trinh-tiep-nhan-va-xu-ly-cong-van-den
status: active
group: Văn thư
description: OCR công văn đến, vào sổ Văn thư, phân phòng thụ lý và đặt hạn theo độ khẩn
steps:
  - agent: so-hoa-ocr-cong-van
    task: "Tiếp nhận file công văn đến từ {{input}} (scan, PDF, ảnh hoặc đường dẫn iOffice). OCR bằng pdf-sang-md/paddleocr. Trích: Số công văn, Ngày ban hành, Cơ quan ban hành, Trích yếu, Độ khẩn (Hỏa tốc/Thượng khẩn/Khẩn/Thường). Bàn giao file .md và các trường đó."
    max_retries: 1
  - agent: phong-van-thu
    task: "Từ {{prev}} và {{input}}: ghi sổ `06 - Văn thư/so-van-ban-den.md` (số đến, ngày nhận, số/ký hiệu, cơ quan, trích yếu, độ khẩn). Phân phòng thụ lý theo bảng trong agent, ghi hạn (Hỏa tốc trong ngày, Thượng khẩn 1 ngày, Khẩn 2 ngày, Thường 5 ngày làm việc), trạng thái `Đã phân`. Việc phòng phải làm: một câu. Không gọi agent Yêu cầu Tài liệu."
    verify_agent: phong-van-thu
    max_retries: 1
updated: 2026-10-19
---

## Mô tả quy trình

Hai bước, một cửa Văn thư:

1. **Số hóa** (`so-hoa-ocr-cong-van`): PDF/ảnh thành .md có đủ trường.
2. **Vào sổ và phân việc** (`phong-van-thu`): ghi sổ đến, chọn phòng, đặt hạn.

Loop `van-thu-dieu-phoi-don-doc` chỉ hốt dòng `Mới` còn sót và đôn đốc quá hạn. Việc đã `Đã phân` trong bước 2 thì không phân lại.
