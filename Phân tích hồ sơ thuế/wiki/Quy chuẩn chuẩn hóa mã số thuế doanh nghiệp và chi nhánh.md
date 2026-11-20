---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
source: [[conversations/2026-11-18]]
---
Quy chuẩn kỹ thuật chuẩn hóa mã số thuế (MST_chuan - BT6) xử lý lỗi định dạng số bị mất số 0 đầu trong Excel hoặc thiếu dấu phân cách chi nhánh [[conversations/2026-11-18]]:
- Mã số thuế 10 chữ số: Giữ nguyên định dạng chuỗi 10 ký tự [[conversations/2026-11-18]].
- Mã số thuế 9 chữ số: Thêm số '0' vào đầu để hoàn thiện mã số thuế doanh nghiệp 10 số (thực tế tính đến 18/11/2026) [[conversations/2026-11-18]].
- Mã số thuế chi nhánh 12 chữ số: Bổ sung số '0' vào đầu và chèn ký tự gạch nối '-' trước 3 chữ số mã chi nhánh, đưa về định dạng chuẩn 14 ký tự (`XXXXXXXXXX-XXX`) [[conversations/2026-11-18]].
- Mã số thuế chi nhánh 13 chữ số: Giữ nguyên 10 chữ số đầu của đơn vị chủ quản và chèn ký tự gạch nối '-' trước 3 chữ số mã chi nhánh (`XXXXXXXXXX-XXX`) [[conversations/2026-11-18]].

Quy chuẩn này được sử dụng làm trường khóa đồng bộ dữ liệu giữa bảng kê mua vào, bán ra, hóa đơn điện tử portal và danh mục cảnh báo rủi ro thuế [[conversations/2026-11-18]].
