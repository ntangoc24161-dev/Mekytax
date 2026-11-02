---
type: business
provenance: user
created: 2026-10-23
updated: 2026-10-23
---
Thông tin công ty Công Ty TNHH Thương Mại Dịch Vụ Tâm Hoàng Phát (MST 3603721832), đã cập nhật vào sổ JAVIS PRO (`G:\...\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm`, sheet `01_DA_BAN_HANH`, hàng 14):

- Kế toán trưởng: **Phạm Thị Bích Duyên** (mới thêm, trước đó cột trống; chưa có CCCD kế toán).
- Giám đốc: Vũ Văn Sĩ (đã khớp sẵn, không đổi).
- Tài khoản ngân hàng: TK số 114604422888 tại Ngân hàng Vietinbank - CN KCN Biên Hòa, PGD Hố Nai (đã khớp sẵn, không đổi).
- Sổ hiện đã có số thật: Số QĐKT `6899/QĐ-TCS3`, Số QĐGS `6900/QĐ-TCS3`, ngày ban hành `15/10/2026`.
- Backup sổ trước khi sửa: `Danh sach kiem tra quyet toan pro.bak-20261023-010811.xlsm` (cùng thư mục).
- **Đã sửa 2 lỗi trong `scripts/build_ho_so_di_tam_hoang_phat.py` và chạy lại xuất hồ sơ đi:**
  1. `EXCEL_SO` trỏ sai tên file (`...- JAVIS PRO.xlsm` không tồn tại) -> sửa thành tên file thật `Danh sach kiem tra quyet toan pro.xlsm`.
  2. `load_company_from_so()` đọc lệch cột: `so_qdkt`/`so_qdgs`/`ngay_ban_hanh` lấy nhầm cột 27/28/29 (Phân công trưởng đoàn/thành viên/Số QĐKT) thay vì cột 29/30/31 (Số QĐKT/Số QĐGS/Ngày ban hành QĐ thật) - do sheet từng chèn thêm 2 cột "Phân công trưởng đoàn/thành viên" mà script chưa cập nhật theo. Đã sửa index đúng + xử lý `ngay_ban_hanh` dạng chuỗi "dd/mm/yyyy" (không chỉ datetime).
  - Hồ sơ đi xuất lại tại `0. BIEN BAN\TONG HOP.docx` + 2 QĐ + `hoso.zip` nay đã có đúng Số QĐKT 6899/QĐ-TCS3, Số QĐGS 6900/QĐ-TCS3, ngày QĐ 15/10/2026 (đã grep xác nhận trong file).
- **Tác dụng phụ cần biết:** thao tác ghi cột 16 bằng openpyxl (không `data_only`) làm sổ Excel mất CACHE hiển thị của 3 ô công thức trên hàng 14 (cột 1 STT, cột 8 Mã CQT, cột 34 Ngày bắt đầu) - công thức vẫn còn nguyên, chỉ cần **mở lại file bằng Excel một lần** để nó tự tính lại (Excel recalc on open mặc định).

[[ho-so-kiem-tra-mst-3603721832-chuyen-thu-muc-va-so-thu-tu]]
