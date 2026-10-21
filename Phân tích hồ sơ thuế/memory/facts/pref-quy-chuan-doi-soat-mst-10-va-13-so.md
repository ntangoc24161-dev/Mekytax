---
type: preference
created: 2026-10-07
updated: 2026-10-07
---
Quy chuẩn đối soát Mã số thuế (MST) 10 số và 13 số trong công tác kiểm tra hồ sơ thuế:

1. **Chuẩn hóa định dạng MST:**
   - Mã số thuế 10 số: Định dạng `XXXXXXXXXX` (Doanh nghiệp, Tổ chức độc lập).
   - Mã số thuế 13 số: Định dạng `XXXXXXXXXX-YYY` (Chi nhánh, Địa điểm kinh doanh, Hộ KD trực thuộc). Nếu dữ liệu đầu vào thiếu dấu gạch ngang (dạng `XXXXXXXXXXYYY`), hệ thống tự động chuẩn hóa chèn dấu gạch ngang thành `XXXXXXXXXX-YYY`.

2. **Quy tắc đối soát rủi ro F0 (Nhà cung cấp / Bên bán):**
   - **Đối tác là MST 13 số (`XXXXXXXXXX-YYY`):**
     - So khớp chính xác mã 13 số trong danh mục F0. Nếu đúng mã chi nhánh có trong F0 -> Đưa vào danh sách cảnh báo F0.
     - Nếu không có mã 13 số trong F0, kiểm tra MST 10 số công ty mẹ `XXXXXXXXXX`. Nếu công ty mẹ bị cảnh báo F0 toàn diện -> Cảnh báo F0 liên đới theo công ty mẹ.
     - Nếu công ty mẹ không có trong F0 và F0 chỉ cảnh báo một chi nhánh cụ thể khác (ví dụ `-030`), thì chi nhánh đang xét (ví dụ `-006`) là **AN TOÀN / BÌNH THƯỜNG**, tuyệt đối không gán nhầm F0.
   - **Đối tác là MST 10 số (`XXXXXXXXXX`):**
     - So khớp chính xác mã 10 số trong F0. Nếu F0 chỉ cảnh báo chi nhánh trực thuộc mà không cảnh báo công ty mẹ thì pháp nhân mẹ không bị gán F0 trực tiếp.

3. **Quy tắc đối soát Danh bạ NNT (Trạng thái ĐKKD):**
   - Tra cứu chính xác mã 13 số trước để lấy trạng thái hoạt động của chi nhánh.
   - Nếu Danh bạ chỉ quản lý cấp công ty mẹ, tra cứu theo 10 số và ghi chú trạng thái `(Theo cty mẹ)`.

**Vì sao:** Tránh gán nhầm rủi ro hóa đơn bất hợp pháp giữa các chi nhánh độc lập thuộc cùng một Tổng công ty (ví dụ: các cây xăng thuộc Petrolimex, Saigon Petro) và đảm bảo tính pháp lý chính xác tuyệt đối khi lập biên bản kiểm tra thuế.
**Áp dụng:** Áp dụng bắt buộc cho toàn bộ 13 doanh nghiệp kiểm tra năm 2026 và mọi module phân tích rủi ro hóa đơn của Javis.
