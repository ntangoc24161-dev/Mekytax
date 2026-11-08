---
type: preference
provenance: user
origin: javis-learned
created: 2026-09-30
updated: 2026-11-06
---
Cấu trúc file theo dõi kiểm tra quyết toán theo yêu cầu người dùng:

- **File làm việc duy nhất có macro:** `Danh sach kiem tra quyet toan pro.xlsm` tại `00. Theo doi chung\` (tự chuyển dòng khi đổi trạng thái).
- **File gốc tham chiếu (giữ nguyên):** `Danh sach kiem tra quyet toan.xlsx` cùng thư mục.
- **Không tạo bak** trong `00. Theo doi chung\`.
- **Sheet giữ lại:** `00_DASHBOARD`, `01_DA_BAN_HANH`, `02_DANG_THUC_HIEN`, `DANH_MUC`.
- **Sheet đã bỏ:** dữ liệu lịch sử (`LICH_SU_2024`, `LICH_SU_2025`) và dữ liệu in ấn (`DU_LIEU_IN_AN`).
- **Sổ theo dõi chia 2 trạng thái:**
  1. `01_DA_BAN_HANH` - Đã ban hành (có Số QĐKT).
  2. `02_DANG_THUC_HIEN` - Đang thực hiện (chưa có Số QĐKT).
- **Tự chuyển dòng:** tại sheet `02_DANG_THUC_HIEN`, chọn trạng thái `1. Đã ban hành` sẽ tự chuyển hồ sơ sang `01_DA_BAN_HANH`; chiều ngược lại cũng vậy. Cần mở file `.xlsm` và bật macro (Enable Content).
- **Cột thông tin công ty theo mẫu (đã bổ sung lại):** Thời kỳ kiểm tra, Ngành nghề KD, Giám đốc, Giới tính, Số CCCD, Ngày cấp CCCD, Nơi cấp CCCD, Kế toán trưởng, CCCD Kế toán, Số GCN ĐKKD, Ngày cấp ĐKKD, Chi nhánh, Tài khoản ngân hàng.
- **Cột phân công nhiệm vụ đoàn:** `Phân công trưởng đoàn` (cột 27 / AA) và `Phân công thành viên` (cột 28 / AB).
- **Cột đã bỏ:** Hạn công bố QĐ, Hạn ký biên bản KT, Hạn nộp tiền, Tình trạng nộp, và toàn bộ phần theo dõi nộp ngân sách nhà nước.
- **Cảnh báo:** cột `Thông báo quá hạn` hiện thông báo khi quá hạn mà chưa thực hiện.
- **Định dạng:** font chữ Times New Roman toàn bộ file.
- **Script:** `scripts/create_redesigned_excel.py` (tạo cấu trúc) và `scripts/inject_status_move_vba.py` (gắn macro chuyển dòng).
