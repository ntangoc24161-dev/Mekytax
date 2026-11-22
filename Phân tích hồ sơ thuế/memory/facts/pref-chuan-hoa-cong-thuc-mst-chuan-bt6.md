---
type: preference
provenance: user
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
---
Quy chuẩn chuẩn hóa mã số thuế (MST_chuan - BT6) áp dụng cho bảng kê, HĐĐT và các biểu đối chiếu thuế:
1. 10 chữ số (hoặc 14 ký tự đã có dấu `-`): Giữ nguyên (vd: `3602996707`).
2. 9 chữ số: Thêm 1 số 0 ở đầu thành 10 số (vd: `602996707` → `0602996707`).
3. 12 chữ số: Thêm 1 số 0 ở đầu và dấu gạch ngang `-` trước 3 số cuối (nhánh) thành 14 ký tự (vd: `602996707001` → `0602996707-001`).
4. 13 chữ số: Giữ nguyên 10 số đầu và thêm dấu gạch ngang `-` trước 3 số cuối thành 14 ký tự (vd: `3602996707001` → `3602996707-001`).

Công thức Excel tham chiếu:
`=IF(LEN(A1)=10, A1, IF(LEN(A1)=9, "0"&A1, IF(LEN(A1)=12, "0"&LEFT(A1,9)&"-"&RIGHT(A1,3), IF(LEN(A1)=13, LEFT(A1,10)&"-"&RIGHT(A1,3), A1))))`
