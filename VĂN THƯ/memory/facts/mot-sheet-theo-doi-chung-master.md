---
type: decision
created: 2026-11-20
updated: 2026-11-20
source: chat 2026-11-20 (8aa1855e)
---
Người dùng quyết định chỉ theo dõi trên **một sheet chung** trong file Master công tác kiểm tra.

**Vì sao:** Hai sheet `DANH_SACH_THEO_DOI_CHUNG` và `KT_TONG_HOP` trùng DN, dễ lệch khi đồng bộ.

**Áp dụng:** Chỉ dùng `DANH_SACH_THEO_DOI_CHUNG` (đã gộp cột địa chỉ, SĐT, cán bộ, giai đoạn, L1/L2, trạng thái, cấp độ, hành động). Sheet `KT_TONG_HOP` đã xóa. Agent/workflow/reminder chỉ trỏ sheet chung.
