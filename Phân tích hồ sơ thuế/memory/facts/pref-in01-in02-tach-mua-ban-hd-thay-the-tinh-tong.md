---
type: preference
created: 2026-10-18
updated: 2026-10-18
---
IN01 và IN02 chia riêng khối Mua vào / Bán ra, mỗi khối có dòng tổng. Font Times New Roman cỡ 11, ngày `dd/MM/yyyy`, số Excel `#,##0` / `#,##0.00`, tự giãn cột.

**Vì sao:** User chốt khi hoàn thiện bộ mẫu v2 T&T Biên Hòa.

**Áp dụng:** Hóa đơn thay thế vẫn tính vào tổng IN01; hóa đơn đã bị thay thế và hóa đơn hủy/xóa bỏ không tính. Dòng tổng dùng `TONG_MUA` / `TONG_BAN` để công thức BM04 COUNTIFS `Chieu=MUA` không cộng nhầm.
