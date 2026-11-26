---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-25
updated: 2026-11-25
---
Trên hai sheet `01_DA_BAN_HANH` và `02_DANG_THUC_HIEN` của `Danh sach kiem tra quyet toan pro.xlsm`, các cột ngày (`Ngay_cap_CCCD`, `Ngay_cap_DKKD`, `Ngay_QDKT`, `Ngay_cong_bo`, `Ngay_bat_dau`, `Ngay_ket_thuc`, `Ngay_ky_BBKT`, `Ngay_bao_cao`, `Ngay_QDXPHC`) được chuẩn hóa về định dạng ngày `dd/mm/yyyy`, các cột tiền (cột AN đến AY từ Truy thu đến Giảm khấu trừ) được định dạng `#.##0` hiển thị dấu chấm phân cách hàng nghìn (ví dụ: `17.850.636`). Nhờ đó khi Word kết nối qua OleDB/ODBC (`IMEX=1`), dữ liệu được trả về dạng chuỗi định dạng sẵn thay vì số sê-ri ngày (serial number) hay số tiền có đuôi thập phân lỗi.
