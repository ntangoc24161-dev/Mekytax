---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-05
updated: 2026-11-18
source: chat 2026-11-18 (mwlp8ddh)
---
**Quyết định:** Lệnh in hàng loạt "in phân công KT1" lấy công văn có Tổ KT = KT1, trạng thái = "Mới" từ sheet `NHAP_HO_SO` (file `Bao_cao_cong_van.xlsx`). Mỗi cán bộ nhận 1 phiếu PDF riêng, kèm toàn văn công văn gốc, dòng đầu trang ghi cán bộ/MST/tên công ty/ngày phân công; tổng số trang mỗi phiếu phải là số chẵn. Dòng chưa xác định được cán bộ thì loại ra, báo lại, không tự gán bừa. Script: `scripts/in_phan_cong_kt1.py`, output `exports/in-phan-cong-kt1/`.

**Chốt trạng thái (18/11/2026):** Sau khi xem trước và người dùng **chốt in** thành công, cột O cập nhật sang **`Đang xử lý`** (không dùng "Đang thực hiện").

**Vì sao:** Cần phiếu phân công rõ cho từng cán bộ KT1; thống nhất nhãn trạng thái với KT2/KT3 Zalo và skill `in-phan-cong-kt1`.

**Áp dụng:** Mọi lệnh "in phân công KT1": gửi PDF xem trước qua **Telegram** → chốt in → ghi cột O = `Đang xử lý`. Macro Excel (.xlsm) vẫn chưa triển khai, chờ chủ đồng ý riêng.
