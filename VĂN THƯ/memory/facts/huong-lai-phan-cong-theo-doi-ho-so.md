---
source: chat 2026-11-23 (8aa1855e)
---
﻿---
type: decision
created: 2026-11-23
updated: 2026-11-23
source: chat 2026-11-23 (8aa1855e)
---
Người dùng chọn **hướng lai** để phân công trợ lý theo dõi hồ sơ kiểm tra.

**Mô hình:** Hub Văn thư + Chủ loại hồ sơ + Khâu dùng chung. Nguồn sự thật duy nhất là sheet `DANH_SACH_THEO_DOI_CHUNG` trên file Master.

**Vì sao:** Thuần theo khâu làm loãng trách nhiệm theo DN; thuần theo đầu việc làm trùng logic nhắc hạn/xử lý cuối. Hướng lai giữ chủ sở hữu theo loại và tái dùng thao tác chuẩn.

**Áp dụng:**
- Hub: `phong-van-thu` tiếp nhận, đọc CONFIG, ghi Master, chống trùng MST+loại.
- Chủ loại (2 nhóm): `theo-doi-ke-hoach-kiem-tra` (KHBS + quyết toán), `theo-doi-chuyen-de` (chuyên đề trường học và chuyên đề khác).
- Khâu dùng chung: `theo-doi-tien-do`, `xu-ly-cuoi`, skill soạn TB/bóc tách khi cần.
- Không tạo sổ theo dõi riêng theo loại.
