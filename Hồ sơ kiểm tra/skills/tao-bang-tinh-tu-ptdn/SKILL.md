---
name: Tạo Bang tinh từ PTDN
description: 'Tạo Bang tinh.xlsx trong 0. BIEN BAN từ mẫu: đổi header F5/F6/F7, xóa số mẫu; CDKT tùy chọn.'
group: Tài chính
origin: javis-learned
status: active
created: 2026-10-16
updated: '2026-11-09'
learned_updated: '2026-11-09'
---
# Tạo Bang tinh từ PTDN

Chiều **TẠO PHÔI** đứng trước [[nhap-bang-tinh-tu-bao-cao-hdkd]]: nhân bản mẫu `Bang tinh.xlsx` vào thư mục vụ, đổi header đúng NNT, xóa số công ty mẫu. Sau đó mới nhập KQKDSXKD / GTGT / EBITDA bằng skill nhập. Thuộc **Trợ lý khởi tạo hồ sơ** (`khoi-tao-ho-so`).

## Khi nào dùng

- "tạo bảng tính", "copy Bang tinh", "lập phôi Bang tinh", "tạo 0. BIEN BAN".
- "cập nhật CDKT", "đổ CDKT từ PTDN01" khi `0. BIEN BAN\Bang tinh.xlsx` đã có: mở file hiện có, không tạo workbook mới.
- Vụ mới chưa có `0. BIEN BAN\Bang tinh.xlsx`.
- Không dùng để nhập số liệu PTDN02/KQ2.1 (đó là `nhap-bang-tinh-tu-bao-cao-hdkd`).
- Không dùng để bóc tách tỷ lệ / viết DI-01 (`boc-tach-bang-tinh-phan-tich`) hay xuất Word hồ sơ đi.

## Chuẩn bị

1. Xác định thư mục vụ (kế hoạch hoặc `1. TD\<số>. <TÊN>\`, hoặc `KE HOACH\<TEN>_<MST>\`).
2. Mẫu chuẩn: `...\1. TD\BIEN BAN\Bang tinh.xlsx` (đủ sheet `KQKDSXKD`, `DTCHEP`, `KQKD`, `CĐKT`/`CDKT`, `GTGT`, `EBITDA`...).
3. MST, tên NNT, thời kỳ kiểm tra do user nêu hoặc có trên sổ. Chưa có trên sổ thì không copy kỳ của vụ kế hoạch khác.
4. Nếu user yêu cầu điền CDKT: cần thêm `PTDN01_*.xlsx`. Việc nhập KQKD/GTGT/EBITDA vẫn để skill nhập sau bước này.

## Quy trình

### A. Tạo thư mục và nhân bản

1. Trong thư mục vụ, tạo `0. BIEN BAN\` nếu chưa có.
2. Nếu đã có `0. BIEN BAN\Bang tinh.xlsx` thì không tạo file mới; mở file hiện có. User đang yêu cầu CDKT thì làm tiếp mục C trên chính file này.
3. Chưa có file thì copy mẫu `1. TD\BIEN BAN\Bang tinh.xlsx` vào `0. BIEN BAN\Bang tinh.xlsx`. Không tạo workbook trống.

### B. Đổi header và dọn số mẫu (bắt buộc)

Trên sheet `KQKDSXKD` (và các sheet còn tên công ty mẫu nếu thấy):

- `F5` = `Từ năm YYYY đến năm YYYY` (đúng kỳ vụ)
- `F6` = MST vụ
- `F7` = tên NNT đúng vụ

Nếu `KQKDSXKD!F5:F7` là công thức trỏ sang `CDKT`, sửa header ở `CDKT!F5:F8` rồi kiểm tra công thức ăn theo. Không ghi đè công thức bằng chữ.

Xóa số liệu công ty mẫu còn sót theo danh tính (tên/MST cũ, ví dụ Hoàng Danh Phát), không xóa theo mốc năm cố định. Không bịa số 0 vào năm/quý chưa có nguồn.

### C. CDKT từ PTDN01 (chỉ khi user yêu cầu)

1. Header ghi ở `CDKT!F5:F8`:
   - `F5` = `Từ năm YYYY đến năm YYYY` (đúng kỳ user yêu cầu, không mở rộng thành mọi năm PTDN có số)
   - `F6` = MST
   - `F7` = tên NNT
   - `F8` = địa bàn / cơ quan thuế lấy từ header PTDN01, không bịa
2. Phôi `CDKT` là form TT133, PTDN01 là form TT200. Map theo tên chỉ tiêu. Không copy theo số thứ tự dòng và không bịa bảng mã.
3. Trước khi ghi số: đặt định dạng số cho ô đích. Ô phôi để định dạng ngày thì số lớn lưu thành `#VALUE!` (vụ Thanh Thanh Phương, MST 3600840461, lần ghi 09/11/2026). Gán number format rồi ghi lại. Không bàn giao file còn `#VALUE!`.
4. Điền sheet `CDKT` và sheet `CĐKT ` (có dấu, có thể thừa dấu cách cuối) chỉ ở các cột năm phôi đã có. Năm không có cột thì nói rõ, không tự chèn cột. User yêu cầu tường minh tạo cột năm thiếu thì chèn ngay sau cột năm cuối, trước `Dọc - Ngang`; copy merge 2 dòng và định dạng số từ cột năm liền trước. Chỉ chèn khi user ra lệnh rõ (ví dụ "tạo cột 2025"): chèn sau cột năm cuối, trước `Dọc - Ngang`, copy merge 2 dòng và định dạng `#,##0`, đổ cùng nguồn đã có trên `CDKT`.
5. Đối chiếu từng năm: Tổng tài sản = Nợ phải trả + Vốn chủ sở hữu. Lệch thì dừng, không bàn giao.
6. Backup trước khi ghi CDKT: `Bang tinh_backup_before_cdkt_YYYYMMDD_HHMMSS.xlsx` trong cùng `0. BIEN BAN\`.

### D. Bàn giao

1. Báo đường dẫn tuyệt đối `...\0. BIEN BAN\Bang tinh.xlsx`.
2. Chạy tiếp [[nhap-bang-tinh-tu-bao-cao-hdkd]] để đổ KQKDSXKD từ PTDN02 (thiếu thì bù BCTC `NamNay`); có `KQ2.1_*_FINAL` thì đổ thêm GTGT và EBITDA.
3. Sau nhập: [[boc-tach-bang-tinh-phan-tich]] để tỷ lệ chi phí và Mục III DI-01.

## Phân vai với skill khác

| Việc | Skill |
| :--- | :--- |
| Tạo `0. BIEN BAN` + copy mẫu + header F5/F6/F7 + CDKT khi user yêu cầu | **`tao-bang-tinh-tu-ptdn`** (skill này) |
| Nhập KQKDSXKD / GTGT / EBITDA | `nhap-bang-tinh-tu-bao-cao-hdkd` |
| Tỷ lệ chi phí, phân tích DI-01 | `boc-tach-bang-tinh-phan-tich` |
| Sổ Excel / xuất Word DI | `lap-ho-so-kiem-tra` |
| KQ2.1 (GTGT, EBITDA) phục vụ đổ vào Bang tinh | pipeline `phan-tich-ho-so-thue` (GĐ2) |

Đích chuẩn vụ kế hoạch / trụ sở NNT: **`0. BIEN BAN\Bang tinh.xlsx`**. Bản vault `exports/<slug-mst>/` hoặc đẩy Drive chỉ làm khi user yêu cầu tường minh thêm bản sao; không thay file làm việc trong `0. BIEN BAN`.

## Bẫy

- File nhân bản còn header/số công ty mẫu: bắt buộc đổi `F5/F6/F7` trước khi bàn giao nhập.
- Trộn vai: skill này không map mã PTDN 24 sang 26, không bù BCTC `NamNay`, không đổ GTGT/EBITDA (đều thuộc skill nhập).
- Tạo file trống hoặc copy sai mẫu thiếu sheet: DTCHEP/GTGT/EBITDA hỏng ở bước sau.
- Phôi CDKT hoặc `CĐKT ` thiếu cột năm (đã gặp 2014-2015, và cột 2025 trên sheet `CĐKT `): mặc định điền năm còn cột, báo năm thiếu, không tự chèn. User bảo tạo cột thì chèn sau năm cuối, trước `Dọc - Ngang`, clone merge 2 dòng (vụ Đại Huy Hoàng MST 3602179712, 09/11/2026). User bảo tạo cột thì mới chèn.
- Ô CDKT định dạng ngày: ghi số lớn ra `#VALUE!`. Gán định dạng số trước khi ghi.
- Map TT200 (PTDN01) sang TT133 (phôi) theo tên chỉ tiêu, không theo vị trí dòng.
- PTDN01 không luôn là TT200: vụ Nông sản Lưu Hoài Anh (MST 3603057410, 09/11/2026) xuất form TT133 (mã 200/300/400/500). Nhận form trước khi map; form TT133 thì ghi 1-1 theo mã, không cộng mã TT200 (270/440) vì sẽ ra 0 và lệch cân đối.
- MST chưa có trên sổ: không copy kỳ của vụ kế hoạch vừa làm. Phôi Thanh Thanh Phương 09/11/2026 để 2022-2025 theo Nhất Vinh và Huy Thành Phát; user sửa thành 2021-2025.
- Sau khi đẩy bản sao Drive/Sheet: nhắc user F5 nếu còn cache bản cũ.

## Kiểm chứng

- Tồn tại `0. BIEN BAN\Bang tinh.xlsx` trong đúng thư mục vụ.
- `KQKDSXKD!F6` = MST vụ; `F7` = tên NNT vụ; không còn tên công ty mẫu.
- `F5` khớp thời kỳ user chốt, không phải kỳ copy từ vụ khác.
- Nếu đã đổ CDKT: `CDKT!F5:F8` đúng vụ; vùng số không còn `#VALUE!`; mỗi năm Tổng tài sản = Nợ phải trả + Vốn chủ sở hữu.
- Đã nêu rõ bước tiếp theo là `nhap-bang-tinh-tu-bao-cao-hdkd` (và CDKT đã hoặc chưa làm theo yêu cầu).

## Lịch sử (tự học)
- [2026-11-09] Vụ Thanh Thanh Phương 09/11/2026: kỳ phôi 2022-2025 bị user sửa thành 2021-2025. Ghi CDKT lần đầu ra #VALUE! vì ô định dạng ngày. PTDN01 (TT200) map theo tên chỉ tiêu sang phôi TT133. Header CDKT!F5:F8; F5:F7 của KQKDSXKD là công thức. Sheet CĐKT có dấu không có cột 2025.
- [2026-11-09] Vụ Nông sản Lưu Hoài Anh: PTDN01 đã là TT133 nên đổ 1-1 theo mã; năm 2025 PTDN = 0 bù BCTC `CTieuTKhaiChinh/SoCuoiNam` (không lấy `PL_CDTK`). Sheet `CĐKT ` không có cột 2025.
- [2026-11-09] Vụ Đại Huy Hoàng (MST 3602179712): user yêu cầu tạo cột 2025 trên `CĐKT `. Chèn sau 2024, trước `Dọc - Ngang`, clone merge 2 dòng, đổ `SoCuoiNam` BCTC 2025. Tổng TS = NV 34.145.950.083.
- [2026-11-09] Vụ Thành Sen MST 3603403068: user bảo tạo cột 2025 cho CDKT. Sheet `CDKT` đã có O=2025. Sheet `CĐKT ` chèn P sau O=2024, trước Dọc-Ngang, copy merge 2 dòng, đổ số từ CDKT/BCTC `SoCuoiNam`. Tổng tài sản 2025 = 37.260.675.631, lệch 0.
- [2026-11-09] Vụ Huy Khang MST 3603755214: user bảo tạo cột 2025 trên `CĐKT `. Chèn P sau O=2024, trước `Dọc - Ngang`, clone merge 2 dòng, đổ BCTC `SoCuoiNam` (đã khớp CDKT). Tổng TS 2025 = 2.674.407.729, lệch 0.
