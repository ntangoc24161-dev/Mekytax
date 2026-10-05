---
name: Lập hồ sơ kiểm tra thuế
description: "Vận hành sổ Excel JAVIS PRO và app hồ sơ kiểm tra: thêm vụ từ DB NNT, xuất Word DI trình ký."
group: Tài chính
---

# Lập hồ sơ kiểm tra thuế

Chuỗi nghiệp vụ từ sổ theo dõi Excel (JAVIS PRO) tới xuất bộ Word hồ sơ đi trên app nội bộ Thuế cơ sở 3. App không ký số, không đóng dấu, không cấp số quyết định, không xuất PDF.

## Khi nào dùng

- "thêm hồ sơ kiểm tra", "cập nhật sổ Excel", "thêm MST từ DB NNT", "xuất hồ sơ đi", "xuất DI", "trình ký", "hoso.zip".
- User đưa MST, thời kỳ, nội dung kiểm tra, đoàn kiểm tra, hoặc bảo xuất bộ Word.
- Không dùng để chấm XML/Excel hóa đơn (`phan-tich-ho-so-thue`), mở tờ khai XML (`xem-ho-so-thue`), hay xử phạt VPHC (`xu-phat-vphc-thue-hoa-don`).

## Chuẩn bị

**Sổ Excel JAVIS PRO** (thư mục sổ theo dõi):

`G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\`

- Làm việc: `Danh sach kiem tra quyet toan - JAVIS PRO.xlsm` (bật Enable Content).
- Bản không macro: `Danh sach kiem tra quyet toan - JAVIS PRO.xlsx`.
- File gốc giữ nguyên: `Danh sach kiem tra quyet toan.xlsx`.

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

1. Xác định sheet:
   - Có Số QĐKT thật → `01_DA_BAN_HANH`.
   - Chưa có Số QĐKT → `02_DANG_THUC_HIEN` (trạng thái `2. Đang thực hiện`).
   - User đã chọn trình ký / coi như đã ban hành dù chưa có số → có thể ghi `01_DA_BAN_HANH` theo yêu cầu tường minh.
2. Tra DB NNT theo MST: tên NNT, địa chỉ, địa bàn (tên đầy đủ để VLOOKUP Mã CQT khớp), ngành nghề, GĐ, KTT, CCCD nếu có.
3. Điền thời kỳ dạng `Từ năm YYYY đến năm YYYY`.
4. Cột **Nội dung kiểm tra** xuống hàng, gạch đầu dòng; **giá vốn hàng bán trước** chi phí QLDN khi cả hai cùng có:

```
- Giá vốn hàng bán.
- Chi phí quản lý doanh nghiệp.
```

5. Điền Trưởng đoàn, Thành viên, các cột thông tin công ty còn thiếu theo mẫu (xem reference).
6. Lưu cả `.xlsm` và `.xlsx`. File đang mở bị khóa thì lưu bản tạm rồi thay thế khi Excel nhả file.
7. Báo dòng sheet đã ghi và các trường còn trống (Số QĐKT, ngày ban hành...).

### B. Xuất bộ hồ sơ đi (Word)

1. Lấy MST. Có Số QĐKT chính thức thì dùng số đó. **Chưa có mà user chọn xuất trình ký** → dùng mã tạm `DU_THAO-TRINH-KY` (không phải số đã cấp). Không tự bịa số kiểu `xxxx/QĐ-TCS3`.
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

## Bẫy

- Căn cứ pháp lý hiện hành: **QĐ 1161/QĐ-CT**, **Luật 108/2025/QH15**. Không còn dùng QĐ 970 / Luật 38.
- Thời hạn kiểm tra tại trụ sở NNT: **không quá 20 ngày làm việc** (gia hạn 1 lần ≤ 20 ngày).
- Sheet lịch sử / in ấn đã bỏ; không tạo lại `LICH_SU_*`, `DU_LIEU_IN_AN`.
- Macro chuyển dòng chỉ chạy trên `.xlsm` + Enable Content; khớp trạng thái bằng tiền tố `1.` / `2.` (tránh lỗi font tiếng Việt trong VBA).
- `HUONG_DAN.md` cũ còn ghi `D:\HO SO KIEM TRA\app` - đường đúng là `D:\01. Cong tac kiem tra\HO SO KIEM TRA\app`.
- Không sửa `HO SO NGUON`. Không ký số, không xuất PDF, không bịa số tiền / số quyết định.

## Kiểm chứng

- Excel: MST nằm đúng sheet; nội dung có xuống hàng và gạch đầu dòng; `.xlsm` mở được macro.
- `status` → `dang chay` khi app mở.
- Sau `xuat` thành công: stdout có `OK`, đường dẫn dưới `output\`, có `hoso.zip`.
- Xuất trình ký chưa có QĐKT: thư mục chứa `DU_THAO-TRINH-KY` và câu trả lời nói rõ đó là mã tạm.
