---
source: chat 2026-11-23 (8aa1855e)
---
﻿---
type: decision
created: 2026-11-23
updated: 2026-11-23
source: chat 2026-11-23 (8aa1855e)
---
Taxonomy chủ loại trên Master công tác kiểm tra:

1. **Kế hoạch kiểm tra** = gộp **KHBS (bổ sung)** và **Kiểm tra quyết toán** thành một nhóm chủ loại.
2. **Chuyên đề** = gồm **chuyên đề trường học** (Loại trên sheet: `Chuyên đề trường công lập`) và các chuyên đề khác khi phát sinh.

**Vì sao:** Người dùng xác định bổ sung và quyết toán là một luồng kế hoạch; chuyên đề là nhóm riêng.

**Áp dụng:**
- CONFIG có cột `Nhóm chủ loại`.
- Agent `theo-doi-ke-hoach-kiem-tra` ôm KHBS + Kiểm tra quyết toán.
- Agent `theo-doi-chuyen-de` ôm chuyên đề trường học và chuyên đề mới thuộc nhóm Chuyên đề.
- Giữ nhãn `Loại công việc` chi tiết trên sheet để lọc/báo cáo phụ; không tách sổ riêng.
