---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-27
updated: 2026-10-27
source: chat 2026-10-27 (65cbc484)
---
Khi yêu cầu bổ sung một năm mới (vd 2026) vào `Bang tinh.xlsx`, người dùng muốn: (1) bổ sung **mọi sheet có theo năm**, không chỉ vài sheet chính; (2) **đổ số thật**, không để trống hay điền 0; (3) căn cứ **BCTC và tờ khai quyết toán TNDN năm đó, file XML**. [[conversations/2026-10-27]]

**Vì sao:** Bảng tính dùng để điền biên bản và phân tích, nên cần số liệu thật đủ năm.
**Áp dụng:** lấy số từ XML BCTC + TNDN (có thể dùng skill xem-ho-so-thue để đọc XML), sau đó điền các sheet theo năm; chỗ nào thiếu nguồn thì hỏi lại người dùng, đừng tự bịa số.
