---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-22
updated: 2026-10-22
source: [[conversations/2026-10-22]]
---
Trong quy trình phân tích hồ sơ kiểm tra thuế, dữ liệu sao kê ngân hàng/sổ phụ nộp dưới dạng PDF được phân biệt làm hai loại chính:
1. PDF trích xuất điện tử (Internet Banking): Có sẵn lớp văn bản (text layer), cho phép trích xuất tự động thông tin tài khoản, ngân hàng, ngày giao dịch và số tiền mà không cần OCR [[conversations/2026-10-22]].
2. PDF scan ảnh thuần (0 ký tự text): Thường là bản scan giấy đóng dấu từ quầy giao dịch ngân hàng. Theo chuẩn quy trình phân tích, hệ thống ghi nhận trạng thái 'PDF - chưa kết luận' cho đối chiếu ủy nhiệm chi (UNC) và không tự ý bịa số liệu hay OCR tự động nếu chưa có cơ chế kiểm chứng độc lập [[conversations/2026-10-22]].
3. Khi thiếu sao kê số hóa chi tiết, việc phân tích công nợ và dòng tiền ưu tiên bóc tách từ sổ chi tiết kế toán Excel (TK 112, 131, 331) nếu doanh nghiệp có cung cấp [[conversations/2026-10-22]].
