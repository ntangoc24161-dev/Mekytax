---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
source: [[conversations/2026-10-07]]
---
Khi đối soát Nhà cung cấp trong bảng kê mua vào với sổ theo dõi công văn cảnh báo rủi ro F0, hệ thống phân biệt cụ thể giữa Mã số thuế 10 số (công ty mẹ) và Mã số thuế 13 số (chi nhánh, địa điểm kinh doanh) [[conversations/2026-10-07]].

Quy tắc đối soát theo cấp độ chi nhánh:
- Nếu mã 13 số khớp chính xác với bản ghi vi phạm trong Sổ F0: xác định chi nhánh có cảnh báo rủi ro F0 trực tiếp [[conversations/2026-10-07]].
- Nếu công ty mẹ (10 số) bị cảnh báo F0: xác định rủi ro liên đới đối với toàn bộ các chi nhánh trực thuộc [[conversations/2026-10-07]].
- Nếu cảnh báo F0 chỉ nêu đích danh một chi nhánh cụ thể khác (ví dụ `-030`) mà công ty mẹ và chi nhánh đang xét (ví dụ `-006`) không nằm trong danh sách vi phạm: xác định chi nhánh đang xét là an toàn/bình thường, không đưa vào danh sách cảnh báo F0 [[conversations/2026-10-07]].
