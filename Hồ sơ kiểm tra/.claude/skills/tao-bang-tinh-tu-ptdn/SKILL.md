---
name: "Tạo Bang tinh từ PTDN"
description: "Tạo Bang tinh.xlsx trong 0. BIEN BAN từ mẫu: đổi header F5/F6/F7, xóa số mẫu; CDKT tùy chọn."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-10-16
updated: '2026-11-05'
---
# Tạo Bang tinh từ PTDN

Chiều **TẠO PHÔI** đứng trước [[nhap-bang-tinh-tu-bao-cao-hdkd]]: nhân bản mẫu `Bang tinh.xlsx` vào thư mục vụ, đổi header đúng NNT, xóa số công ty mẫu. Sau đó mới nhập KQKDSXKD / GTGT / EBITDA bằng skill nhập. Thuộc **Trợ lý khởi tạo hồ sơ** (`khoi-tao-ho-so`).

## Khi nào dùng

- "tạo bảng tính", "copy Bang tinh", "lập phôi Bang tinh", "tạo 0. BIEN BAN".
- Vụ mới chưa có `0. BIEN BAN\Bang tinh.xlsx`.
- Không dùng để nhập số liệu PTDN02/KQ2.1 (đó là `nhap-bang-tinh-tu-bao-cao-hdkd`).
- Không dùng để bóc tách tỷ lệ / viết DI-01 (`boc-tach-bang-tinh-phan-tich`) hay xuất Word hồ sơ đi.

## Chuẩn bị

1. Xác định thư mục vụ (kế hoạch hoặc `1. TD\<số>. <TÊN>\`, hoặc `KE HOACH\<TEN>_<MST>\`).
2. Mẫu chuẩn: `...\1. TD\BIEN BAN\Bang tinh.xlsx` (đủ sheet `KQKDSXKD`, `DTCHEP`, `KQKD`, `CĐKT`/`CDKT`, `GTGT`, `EBITDA`...).
3. MST, tên NNT, thời kỳ kiểm tra (ví dụ `Từ năm 2022 đến năm 2025`).
4. Nếu user yêu cầu điền luôn CDKT: cần thêm `PTDN01_*.xlsx`. Việc nhập KQKD/GTGT/EBITDA vẫn để skill nhập sau bước này.

## Quy trình

### A. Tạo thư mục và nhân bản

1. Trong thư mục vụ, tạo `0. BIEN BAN\` nếu chưa có.
2. Nếu đã có `0. BIEN BAN\Bang tinh.xlsx` → **không tạo file mới**; mở file hiện có, chỉ sửa header/số mẫu nếu còn sót, rồi bàn giao skill nhập.
3. Chưa có file → copy mẫu `1. TD\BIEN BAN\Bang tinh.xlsx` → `0. BIEN BAN\Bang tinh.xlsx`. Không tạo workbook trống.

### B. Đổi header và dọn số mẫu (bắt buộc)

Trên sheet `KQKDSXKD` (và các sheet còn tên công ty mẫu nếu thấy):

- `F5` = `Từ năm YYYY đến năm YYYY` (đúng kỳ vụ)
- `F6` = MST vụ
- `F7` = tên NNT đúng vụ

Xóa số liệu công ty mẫu còn sót theo **danh tính** (tên/MST cũ, ví dụ Hoàng Danh Phát), không xóa theo mốc năm cố định. Không bịa số 0 vào năm/quý chưa có nguồn.

### C. CDKT từ PTDN01 (chỉ khi user yêu cầu)

1. Điền sheet `CDKT`/`CĐKT` theo các cột năm phôi còn hỗ trợ.
2. Năm không có cột trên phôi → nói rõ, không tự chèn cột.
3. Kỳ báo cáo = union năm user yêu cầu và năm PTDN thật có số.

### D. Bàn giao

1. Báo đường dẫn tuyệt đối `...\0. BIEN BAN\Bang tinh.xlsx`.
2. Chạy tiếp [[nhap-bang-tinh-tu-bao-cao-hdkd]] để đổ KQKDSXKD từ PTDN02 (thiếu thì bù BCTC `NamNay`); có `KQ2.1_*_FINAL` thì đổ thêm GTGT và EBITDA.
3. Sau nhập: [[boc-tach-bang-tinh-phan-tich]] để tỷ lệ chi phí và Mục III DI-01.

## Phân vai với skill khác

| Việc | Skill |
| :--- | :--- |
| Tạo `0. BIEN BAN` + copy mẫu + header F5/F6/F7 | **`tao-bang-tinh-tu-ptdn`** (skill này) |
| Nhập KQKDSXKD / GTGT / EBITDA | `nhap-bang-tinh-tu-bao-cao-hdkd` |
| Tỷ lệ chi phí, phân tích DI-01 | `boc-tach-bang-tinh-phan-tich` |
| Sổ Excel / xuất Word DI | `lap-ho-so-kiem-tra` |
| KQ2.1 (GTGT, EBITDA) phục vụ đổ vào Bang tinh | pipeline `phan-tich-ho-so-thue` (GĐ2) |

Đích chuẩn vụ kế hoạch / trụ sở NNT: **`0. BIEN BAN\Bang tinh.xlsx`**. Bản vault `exports/<slug-mst>/` hoặc đẩy Drive chỉ làm khi user yêu cầu tường minh thêm bản sao; không thay file làm việc trong `0. BIEN BAN`.

## Bẫy

- File nhân bản còn header/số công ty mẫu → bắt buộc đổi `F5/F6/F7` trước khi bàn giao nhập.
- Trộn vai: skill này **không** map mã PTDN 24→26, không bù BCTC `NamNay`, không đổ GTGT/EBITDA (đều thuộc skill nhập).
- Tạo file trống hoặc copy sai mẫu thiếu sheet → DTCHEP/GTGT/EBITDA hỏng ở bước sau.
- Phôi CDKT thiếu cột năm sớm (đã gặp 2014-2015): điền năm còn cột, báo năm thiếu.
- Sau khi đẩy bản sao Drive/Sheet: nhắc user F5 nếu còn cache bản cũ.

## Kiểm chứng

- Tồn tại `0. BIEN BAN\Bang tinh.xlsx` trong đúng thư mục vụ.
- `KQKDSXKD!F6` = MST vụ; `F7` = tên NNT vụ; không còn tên công ty mẫu.
- `F5` khớp thời kỳ kiểm tra.
- Đã nêu rõ bước tiếp theo là `nhap-bang-tinh-tu-bao-cao-hdkd` (và CDKT đã/ chưa làm theo yêu cầu).
