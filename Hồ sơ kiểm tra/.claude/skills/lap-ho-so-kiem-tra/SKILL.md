---
name: Lập hồ sơ kiểm tra thuế
description: "Vận hành sổ Excel JAVIS PRO và app hồ sơ kiểm tra: thêm vụ từ DB NNT, xuất Word DI trình ký, kết nối Mail Merge trực tiếp từ 01_DA_BAN_HANH."
group: Tài chính
updated: '2026-11-13'
learned_updated: '2026-11-13'
---

# Lập hồ sơ kiểm tra thuế

Chuỗi nghiệp vụ từ sổ theo dõi Excel (JAVIS PRO) tới xuất bộ Word hồ sơ đi trên app nội bộ Thuế cơ sở 3. App không ký số, không đóng dấu, không cấp số quyết định, không xuất PDF.

## Khi nào dùng

- "thêm hồ sơ kiểm tra", "cập nhật sổ Excel", "thêm MST từ DB NNT", "xuất hồ sơ đi", "xuất DI", "trình ký", "hoso.zip".
- User đưa MST, thời kỳ, nội dung kiểm tra, đoàn kiểm tra, hoặc bảo xuất bộ Word.
- Không dùng để chấm XML/Excel hóa đơn (`phan-tich-ho-so-thue`), mở tờ khai XML (`xem-ho-so-thue`), hay xử phạt VPHC (`xu-phat-vphc-thue-hoa-don`).

## Chuẩn bị

**Thư mục `0. BIEN BAN` và file `Bang tinh.xlsx` (làm trước khi nhập liệu / xuất hồ sơ đi):**

Đường gốc vụ theo năm, ví dụ:
`G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\<số>. <TÊN CÔNG TY>\`
hoặc thư mục kế hoạch `...\KE HOACH\<TEN>_<MST>\`.

1. Trong thư mục công ty, tạo `0. BIEN BAN\` nếu chưa có (nơi lưu chung `Bang tinh.xlsx`, `TONG HOP.docx`, QĐ riêng, `hoso.zip`).
2. Nhân bản mẫu `1. TD\BIEN BAN\Bang tinh.xlsx` vào `0. BIEN BAN\Bang tinh.xlsx`; đổi header `F5/F6/F7`, xóa số công ty mẫu. Chi tiết: skill `tao-bang-tinh-tu-ptdn`. Không tạo file trống.
3. Sau đó chạy `nhap-bang-tinh-tu-bao-cao-hdkd` → `boc-tach-bang-tinh-phan-tich` / `dien-bien-ban-tu-bang-tinh` → xuất hồ sơ đi.
4. Việc tạo phôi Bang tinh thuộc **Trợ lý khởi tạo hồ sơ** (`khoi-tao-ho-so`), không phải bước xuất Word của agent Lập hồ sơ kiểm tra.

**Sổ Excel theo dõi** (thư mục sổ theo dõi):

`G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\`

- File làm việc: **`Danh sach kiem tra quyet toan pro.xlsm`** (bật Enable Content).
- File gốc tham chiếu (giữ nguyên): `Danh sach kiem tra quyet toan.xlsx`.
- **Không tạo** file `.bak` / bản sao sổ trong `00. Theo doi chung\`.
- Không dùng tên cũ `Danh sach kiem tra quyet toan - JAVIS PRO.xlsm` / `.xlsx`.

**DB NNT:** `I:\My Drive\DB NNT.xls` (file lớn; ưu tiên parse SpreadsheetML / openpyxl, tránh Excel COM nếu chậm).

**App xuất Word:** `D:\01. Cong tac kiem tra\HO SO KIEM TRA\app` - mở `chay.bat`, trang `http://127.0.0.1:8000`.

**Script app** (không cài thêm gói):

```
python "D:\SECOND BRAIN\javis-os\shared-skills\lap-ho-so-kiem-tra\scripts\chay_vu.py" status
```

Thiếu lệnh `python` thì dùng `C:\Users\NGOC\AppData\Local\Programs\Python\Python312\python.exe`.

Đăng nhập script: tài khoản `canbo` (`HOSO_USER` / `HOSO_PASS` nếu đã đổi). Không dùng admin.

**Script sổ Excel trong brain:** `scripts/create_redesigned_excel.py`, `scripts/inject_status_move_vba.py`, `scripts/javis_tax_tracker.py`.

Chi tiết cột sheet: `references/so-excel-javis-pro.md`. Danh mục mẫu Word: `D:\SECOND BRAIN\javis-os\shared-skills\lap-ho-so-kiem-tra\references\mau-van-ban.md`.

## Quy trình

### A. Thêm / cập nhật trên sổ Excel

1. Xác định sheet trên sổ `Danh sach kiem tra quyet toan pro.xlsm`:
   - Chưa tạo bộ hồ sơ đi / chưa có Số QĐKT → lưu tại `02_DANG_THUC_HIEN` (trạng thái `2. Đang thực hiện`).
   - Đã tạo bộ hồ sơ đi (trình ký) HOẶC đã có Số QĐKT chính thức → tự động chuyển/lưu tại `01_DA_BAN_HANH` (trạng thái `1. Đã ban hành`), đồng thời di chuyển thư mục của công ty sang thư mục `1. TD` (đặt STT nối tiếp, ví dụ `1. TD\<STT>. <TÊN CÔNG TY>\`), không cần hỏi lại.
2. Tra DB NNT theo MST: tên NNT, địa chỉ, địa bàn (tên đầy đủ để VLOOKUP Mã CQT khớp), ngành nghề, GĐ, KTT, CCCD nếu có.
3. Điền thời kỳ dạng `Từ năm YYYY đến năm YYYY`.
4. Cột **Nội dung kiểm tra** xuống hàng, gạch đầu dòng; **giá vốn hàng bán trước** chi phí QLDN khi cả hai cùng có:

```
- Giá vốn hàng bán.
- Chi phí quản lý doanh nghiệp.
```

5. Điền Trưởng đoàn, Thành viên, các cột thông tin công ty còn thiếu theo mẫu (xem reference).
6. Điền / sao chép công thức 6 cột tiền bằng chữ `=DocSo(...)` (cột 53 đến 58 / BA đến BF: `Tong_truy_thu_chu`, `Tong_cham_nop_chu`, `Tong_tien_phat_chu`, `Tong_phat_sinh_chu`, `Giam_lo_chu`, `Giam_khau_tru_chu`) cho dòng vụ việc tại sheet `01_DA_BAN_HANH`.
   - Lưu ý cấu trúc kết nối Mail Merge: Sheet `01_DA_BAN_HANH` có **Dòng 1 ẩn** chứa đúng 58 mã trường chuẩn không dấu (`Ten_NNT`, `MST`, `So_QDKT`...). Phôi Word `templates/MAUTONG HOP.docx` kết nối trực tiếp vào bảng `01_DA_BAN_HANH$`. Sheet đệm `DuLieuMerge` đã bị bãi bỏ và xóa hoàn toàn.
7. Chỉ lưu `Danh sach kiem tra quyet toan pro.xlsm`. File đang mở bị khóa thì lưu bản tạm **ngoài** `00. Theo doi chung\` rồi thay thế khi Excel nhả file; không để bak trong thư mục sổ.
8. Báo dòng sheet đã ghi và các trường còn trống (Số QĐKT, ngày ban hành...).

### B. Xuất bộ hồ sơ đi (Word)

1. Lấy MST. Có Số QĐKT chính thức thì dùng số đó. **Chưa có mà user chọn xuất trình ký** → trên văn bản Word trình bày để khoảng trống dài trước ký hiệu dạng `Số:              /QĐ-TCS3` để điền tay sau, **TUYỆT ĐỐI KHÔNG dùng mã tạm `DU_THAO-TRINH-KY` trên văn bản xuất**. Ngày tháng năm ban hành dự thảo để trống toàn bộ cả ngày, tháng và năm `Trấn Biên, ngày       tháng       năm      ` (không giữ năm 2026). Khi QĐKT/QĐGS đã có số và ngày ban hành chính thức thì giữ nguyên, không ghi đè lại số/ngày. Căn cứ QĐ 1377/QĐ-CT chốt thống nhất năm 2025 (ngày 30/06/2025).
2. `chay_vu.py status`. In `chua chay` → mở `chay.bat` nền, gọi lại `status` đến `dang chay`. Giữ cổng **8000**.
3. Đảm bảo vụ đã có trên app (import Excel mẫu app hoặc nhập tay). Cùng MST + số quyết định = cập nhật vụ cũ.
4. Xuất nhóm DI hồ sơ đi khi user yêu cầu hồ sơ đi / trình ký:

```
python "...\chay_vu.py" xuat --mst <MST> --qd "DU_THAO-TRINH-KY" --docs DI-01,DI-02,DI-03,DI-04,DI-05,DI-06,DI-07
```

Có số thật thì thay `--qd` bằng số chính thức. Bỏ `--docs` thì lấy đúng các ô đang tick trên trang xuất.

5. `pick_vu` khớp sai (vd nhãn trạng thái "Đang làm" lệch chuỗi QĐ) → liệt kê `id=` từ stdout, xuất theo `vu_id` đúng vụ (HTTP `/vu/<id>/xuat`) hoặc sửa số QĐ trên app cho khớp rồi chạy lại script.
6. Script in `LOI` / `CANH BAO` / `CHUA XUAT` → đưa nguyên chữ cho user, không tự xác nhận, không xuất lần hai.
7. Thành công → trả thư mục `output\<MST>\<số-QĐ, dấu / thành ->\<mốc thời gian>\`, danh sách file, `hoso.zip`. Nhắc mở Word điền chỗ ký. Khi có Số QĐKT thật: cập nhật Excel + app rồi xuất lại.
8. **Tự động chuyển trạng thái sổ theo dõi và di chuyển thư mục**: Ngay khi hoàn tất tạo/xuất bộ hồ sơ đi (trình ký) hoặc khi vụ việc chuyển sang trạng thái `1. Đã ban hành`, tự động cập nhật sổ `Danh sach kiem tra quyet toan pro.xlsm` chuyển hồ sơ sang sheet `01_DA_BAN_HANH` và đồng thời di chuyển thư mục công ty sang thư mục `1. TD` (đặt STT nối tiếp), không cần chờ đến khi có Số QĐKT chính thức và không hỏi lại người dùng. Luôn giữ nguyên file gốc `Danh sach kiem tra quyet toan.xlsx` để tham chiếu.

## Bẫy

- Căn cứ pháp lý hiện hành: **QĐ 1161/QĐ-CT**, **Luật 108/2025/QH15**. Không còn dùng QĐ 970 / Luật 38.
- Thời hạn kiểm tra tại trụ sở NNT: **không quá 20 ngày làm việc** (gia hạn 1 lần ≤ 20 ngày).
- Sheet lịch sử / in ấn đã bỏ; không tạo lại `LICH_SU_*`, `DU_LIEU_IN_AN`.
- Macro chuyển dòng chỉ chạy trên `.xlsm` + Enable Content; khớp trạng thái bằng tiền tố `1.` / `2.` (tránh lỗi font tiếng Việt trong VBA).
- `HUONG_DAN.md` cũ còn ghi `D:\HO SO KIEM TRA\app` - đường đúng là `D:\01. Cong tac kiem tra\HO SO KIEM TRA\app`.
- Không sửa `HO SO NGUON`. Không ký số, không xuất PDF, không bịa số tiền / số quyết định.
- Quên chuyển sheet sang `01_DA_BAN_HANH` khi tạo bộ hồ sơ đi: quy tắc đã chốt là tự động chuyển ngay khi tạo hồ sơ đi, không chờ có số QĐ.
- Thiếu 6 cột công thức tiền bằng chữ `=DocSo(...)` khi thêm vụ mới trên `01_DA_BAN_HANH` làm trống số tiền bằng chữ khi Word trộn thư.
- Xóa nhầm Dòng 1 ẩn trên `01_DA_BAN_HANH` làm Word Mail Merge mất tên trường chuẩn không dấu.

## Kiểm chứng

- Excel: MST nằm đúng sheet (ở `01_DA_BAN_HANH` nếu đã tạo hồ sơ đi hoặc có QĐKT; ở `02_DANG_THUC_HIEN` nếu mới khởi tạo); nội dung có xuống hàng và gạch đầu dòng; `.xlsm` mở được macro.
- 6 cột tiền bằng chữ (cột 53-58 / BA-BF) trên dòng vụ việc có công thức `=DocSo(...)` hoạt động.
- File gốc `Danh sach kiem tra quyet toan.xlsx` được bảo toàn nguyên vẹn, không có file `.bak` trong thư mục `00. Theo doi chung\`.
- `status` → `dang chay` khi app mở.
- Sau `xuat` thành công: stdout có `OK`, đường dẫn dưới `output\`, có `hoso.zip`.
- Xuất trình ký chưa có QĐKT: thư mục chứa `DU_THAO-TRINH-KY` và câu trả lời nói rõ đó là mã tạm.

## Lịch sử (tự học)
- [2026-11-12] Thống nhất vận hành duy nhất sổ macro `Danh sach kiem tra quyet toan pro.xlsm`, giữ file gốc `.xlsx`, cấm tạo `.bak` trong thư mục sổ; tự động chuyển sang sheet `01_DA_BAN_HANH` và chuyển thư mục sang `1. TD` ngay khi tạo xong hồ sơ đi trình ký.
- [2026-11-13] Thiết lập Dòng 1 ẩn trên sheet `01_DA_BAN_HANH` chứa 58 mã trường chuẩn không dấu để kết nối Word Mail Merge trực tiếp và xóa bỏ hoàn toàn sheet đệm `DuLieuMerge`. Tự động sao chép 6 cột công thức chữ `=DocSo(...)` (cột 53 đến 58 / BA đến BF) khi thêm vụ mới hoặc đồng bộ số liệu từ `Bang tinh.xlsx`.
