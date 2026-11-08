---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-05
updated: 2026-11-05
---
**Quyết định (2026-11-05):** Lệnh in hàng loạt "in phân công KT1" lấy công văn có Tổ KT = KT1, trạng thái = "Mới" từ sheet `NHAP_HO_SO` (file `Bao_cao_cong_van.xlsx`). Mỗi cán bộ nhận 1 phiếu PDF riêng, kèm toàn văn công văn gốc, dòng đầu trang ghi cán bộ/MST/tên công ty/ngày phân công; tổng số trang mỗi phiếu phải là số chẵn. Dòng chưa xác định được cán bộ (vd nhiều DN gộp chung chưa gán tổ) thì loại ra, báo lại cho người dùng bổ sung, không tự gán bừa. Script thực thi: `scripts/in_phan_cong_kt1.py`, output tại `exports/in-phan-cong-kt1/`.

**Vì sao:** Người dùng cần phiếu phân công rõ ràng cho từng cán bộ Tổ KT1 để giao việc xử lý công văn Công an, tránh bịa cán bộ khi dữ liệu chưa đủ.

**Áp dụng:** Khi nhận lệnh "in phân công KT1" qua chat/Telegram, chạy lại script trên. Sau khi in cần hỏi người dùng xác nhận tên trạng thái cập nhật ("Đang thực hiện" hay "Đang xử lý") trước khi ghi cột trạng thái - việc này còn mở (cần xác minh), chưa chốt tới 2026-11-05. Việc thêm nút macro trên Excel (cần đổi sang .xlsm) cũng chưa triển khai, chờ người dùng đồng ý.
