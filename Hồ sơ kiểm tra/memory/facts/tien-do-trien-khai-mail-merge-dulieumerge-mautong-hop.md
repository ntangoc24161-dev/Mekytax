---
type: business
provenance: assistant
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
---
Thực tế tính đến 12/11/2026:
- Sổ `Danh sach kiem tra quyet toan pro.xlsm` đã có sheet `DuLieuMerge` với 52 trường. Hàng 1 là tên trường không dấu, hàng 2 là công thức link sang dòng vụ việc ở `01_DA_BAN_HANH` (mặc định đang trỏ dòng 5). Các cột chữ (`Tong_truy_thu_chu`, `Tong_cham_nop_chu`, `Tong_tien_phat_chu`, `Tong_phat_sinh_chu`, `Giam_lo_chu`, `Giam_khau_tru_chu`) gọi `=DocSo(...)`.
- Các sheet `01_DA_BAN_HANH`, `02_DANG_THUC_HIEN` chưa có cột bằng chữ. File chưa có module VBA `DocSo`; người dùng phải tự dán mã qua Alt+F11 > Insert > Module.
- Phôi `templates/MAUTONG HOP.docx` lấy theo bản người dùng chỉnh sửa gửi lên, đã đổi trường cũ có dấu sang tên chuẩn (`Ten_NNT`, `Truong_doan`, `Thanh_vien`, `So_QDKT`, `Ngay_QDKT`, `Thoi_ky_KT`...). Tổng phôi có 336 thẻ MERGEFIELD.
- 20 vị trí số QĐKT bị gán nhầm `Ke_toan_truong` đã sửa thành `So_QDKT` và bỏ màu đỏ/xanh. 8 vị trí Kế toán trưởng thật vẫn giữ nguyên.
- Ngày tháng tách theo switch: sau chữ "ngày" dùng `\@ "dd"`, sau "tháng" dùng `\@ "MM"`, sau "năm" dùng `\@ "yyyy"`.
- Bản sao lưu: `templates/MAUTONG HOP.bak-truoc-mergefield-20261112.docx` và `.staging/MAUTONG HOP.bak-truoc-sua-qdkt.docx`.
- Còn tồn đọng: nạp VBA DocSo, chuyển hàng 2 DuLieuMerge sang đúng dòng vụ đang làm, chạy thử Preview Results.
