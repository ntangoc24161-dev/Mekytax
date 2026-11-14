---
type: decision
created: 2026-11-13
updated: 2026-11-13
source: chat 2026-11-13 (ce012a55)
---
Khi thêm công ty mới hoặc khi cập nhật số liệu từ Bảng tính (`Bang tinh.xlsx`) lên sổ theo dõi `Danh sach kiem tra quyet toan pro.xlsm` (sheet `01_DA_BAN_HANH`), hệ thống tự động sao chép/điền công thức đọc số thành chữ `=DocSo(...)` cho cả 6 cột tiền bằng chữ cuối bảng (từ cột 53 đến cột 58 / BA đến BF: `Tong_truy_thu_chu`, `Tong_cham_nop_chu`, `Tong_tien_phat_chu`, `Tong_phat_sinh_chu`, `Giam_lo_chu`, `Giam_khau_tru_chu`).

**Vì sao:**
Sheet `01_DA_BAN_HANH` hiện được kết nối trực tiếp với phôi Word Mail Merge. Nếu thiếu công thức đọc số chữ cho dòng mới hoặc dòng cập nhật, khi trộn thư văn bản sẽ bị trống các trường số tiền bằng chữ trong Biên bản kiểm tra và Quyết định xử phạt.

**Áp dụng:**
- Áp dụng trong script đồng bộ số liệu `scripts/dong_bo_tinh_phat_sang_so.py` và quy trình tạo/cập nhật vụ việc của skill `cap-nhat-so-tu-tinh-phat-cham-nop`, `lap-ho-so-kiem-tra`.
- Tự động gán công thức tương ứng theo số hàng dữ liệu `r`:
  - Cột 53 (BA): `=DocSo(AP{r})` (Tổng truy thu bằng chữ)
  - Cột 54 (BB): `=DocSo(AS{r})` (Tổng chậm nộp bằng chữ)
  - Cột 55 (BC): `=DocSo(AV{r})` (Tổng phạt bằng chữ)
  - Cột 56 (BD): `=DocSo(AW{r})` (Tổng phát sinh bằng chữ)
  - Cột 57 (BE): `=DocSo(AX{r})` (Giảm lỗ bằng chữ)
  - Cột 58 (BF): `=DocSo(AY{r})` (Giảm khấu trừ bằng chữ)
