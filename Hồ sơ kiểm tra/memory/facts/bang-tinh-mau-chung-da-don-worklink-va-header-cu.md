---
type: decision
provenance: source
origin: javis-learned
created: 2026-11-06
updated: 2026-11-06
---
Theo yêu cầu người dùng, đã dọn file mẫu chung `1. TD\BIEN BAN\Bang tinh.xlsx` dùng làm phôi cho skill [[tao-bang-tinh-tu-ptdn]] (thực tế tính đến 2026-10-23):

- Xoá 5 công thức worklink ngoài dạng `[1]Sheet2!...` tại `GTGT!S96, S107, R120, S122` và `DTCHEP!C332` (số liệu chuyển kỳ cũ của Hoàng Danh Phát), thay bằng `0`; gỡ sạch phần `externalLink` khỏi gói file.
- Xoá header công ty cũ tại `CDKT!F5:F8` (kỳ báo cáo, MST, tên NNT, cơ quan thuế của Hoàng Danh Phát), để trống hẳn.
- Backup bản gốc trước khi sửa tại `1. TD\BIEN BAN\Bang tinh (backup truoc don 2026-10-23).xlsx`.
- Chưa xoá số liệu nghiệp vụ thật (doanh thu, GTGT...) còn sót ở các sheet `GTGT`, `DTCHEP`, `KQKDSXKD` - người dùng chưa quyết định có xoá tiếp hay không.
