---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-08
updated: 2026-10-08
---
Khi chuẩn hóa checkbox Báo cáo giám sát trong hồ sơ kiểm tra thuế từ ký tự tay (`□` Unicode) sang hộp Wingdings chuẩn, phải bảo toàn trạng thái đã tích chọn của bản cũ.

- Ô trống: Wingdings `\xa8` (`¨`)
- Ô đã chọn: Wingdings `\xfe` (`þ`)
- Nhãn chữ giữ Times New Roman; chỉ ký tự hộp dùng font Wingdings.

**Vì sao:** Gán hàng loạt mọi ô thành trống làm mất dữ liệu tích sẵn; người dùng đã yêu cầu khôi phục.

**Áp dụng:** Trước khi thay ký tự, đối chiếu bản backup/phôi để map ô đã tích; sau khi thay, kiểm chứng số ô tích khớp bản cũ. Quy tắc chung này độc lập với trạng thái cụ thể từng vụ.
