---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-09
updated: 2026-10-09
---
Áp dụng đồng loạt trong các skill (đặc biệt nhóm Tài chính / xuất Excel thuế): **số tiền** định dạng chuẩn Việt Nam.

- **Hiển thị** (chat, Markdown, báo cáo, ô text): dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (`number_format` openpyxl): `#,##0` hoặc `#,##0.00` (mã invariant); Excel locale VN hiện thành `.` / `,`.
- Tỷ lệ % hiển thị dạng `12,34%`.
- Không dùng dạng Mỹ `1,234,567.89` trong báo cáo tiếng Việt.

**Vì sao:** Thống nhất cách đọc số trên hồ sơ kiểm tra thuế và tránh nhầm "dấu phẩy hàng nghìn" (chuẩn Mỹ) với chuẩn VN.

**Áp dụng:** Skill `phan-tich-ho-so-thue`, các skill đối chiếu GTGT/TNDN/NCC/511-632, tổng hợp CVCBRR, mở hồ sơ TTKT, QLRR, xử phạt VPHC; reference CD1 Excel.
