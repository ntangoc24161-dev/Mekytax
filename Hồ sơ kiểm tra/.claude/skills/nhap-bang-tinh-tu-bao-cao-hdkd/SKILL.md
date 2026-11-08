---
name: Nhập số liệu Bảng tính từ báo cáo HĐKD ngành thuế
description: 'Điền KQKDSXKD từ PTDN02 (thiếu thì bù BCTC NamNay); có KQ2.1 thì đổ thêm GTGT và EBITDA.'
group: Tài chính
updated: '2026-11-05'
---
# Nhập số liệu Bảng tính từ báo cáo HĐKD ngành thuế

Chiều NHẬP LIỆU đứng trước [[boc-tach-bang-tinh-phan-tich]]: lấy số liệu kết quả hoạt động kinh doanh theo năm từ **báo cáo phân tích HĐKD của ngành thuế (mẫu PTDN02, xuất từ TMS)** rồi điền vào sheet `KQKDSXKD` của file `Bang tinh.xlsx`. Khi báo cáo thiếu năm hoặc ô, bù bằng **BCTC sau cùng** (khối `NamNay` trong XML). Nếu đã có `KQ2.1_Phan_tich_CQT_*.xlsx` thì đổ thêm sheet `GTGT` và `EBITDA` từ KQ2.1. Điền xong mới chuyển sang skill phân tích để viết DI-01.

## Khi nào dùng

- "bóc tờ khai điền vào bảng tính", "hoàn chỉnh file bảng tính", "nhập số liệu KQKD".
- "nhập GTGT/EBITDA từ KQ2.1", "đổ KQ2 vào Bảng tính".
- Có file `PTDN02_*.xlsx` trong `1. DU LIEU THO NGANH THUE\1.3. Nghia vu nop thue & Thong bao\` của NNT.
- Cần dựng số liệu nhiều năm cho một MST trước khi tính tỷ lệ và chọn chuyên đề kiểm tra.

## Chuẩn bị (trước khi nhập)

1. Thư mục vụ đã có `0. BIEN BAN\` và file `Bang tinh.xlsx` nhân bản từ mẫu, đã đổi header `F5/F6/F7` (skill `tao-bang-tinh-tu-ptdn`; mục Chuẩn bị của `lap-ho-so-kiem-tra`; thuộc Trợ lý khởi tạo hồ sơ). Chưa có phôi thì chạy `tao-bang-tinh-tu-ptdn` trước, không nhập đè lên file mẫu chung.
2. Xác định MST, tên NNT, thời kỳ kiểm tra (ví dụ 2022-2025).
3. Backup trước khi ghi: copy `Bang tinh.xlsx` → `Bang tinh_backup_before_nhap_YYYYMMDD_HHMMSS.xlsx` trong cùng `0. BIEN BAN\`.

## Nguồn dữ liệu

| Sheet đích | Nguồn chính | Nguồn bù / bổ sung |
| :--- | :--- | :--- |
| `KQKDSXKD` | `PTDN02_*.xlsx` (TMS) | BCTC XML/`QTT 03` năm thiếu - **chỉ lấy khối `NamNay`**, không lấy `NamTruoc` |
| `GTGT` | `KQ2.1_*_FINAL.xlsx` sheet `KQ2_Doi_chieu_GTGT` (Phần I theo quý) | Để trống quý không có tờ khai; không ghi 0 giả |
| `EBITDA` | `KQ2.1_*_FINAL.xlsx` sheet `KQ2_EBITDA_Phan_tich` | Giữ đúng số KQ2.1 dù lệch `KQKDSXKD` (báo lệch cho user) |

Đọc bằng `openpyxl` `read_only=True, data_only=True`. Ghi bằng `openpyxl` (hoặc Excel COM), không sửa file nguồn.

Ví dụ đường dẫn vụ TT Biên Hòa:
`...\KE HOACH\TT_BIEN_HOA_3603883505\0. BIEN BAN\Bang tinh.xlsx`

## Cấu trúc file PTDN02

- Hàng 5-8: kỳ báo cáo, MST, tên NNT, cơ quan thuế. Đối chiếu đúng MST và kỳ.
- Hàng 10: nhãn cột (Tên chỉ tiêu ở cột 4, Mã chỉ tiêu ở cột 8).
- Hàng 11: **năm** nằm rải theo cột (có thể từ 2015 đến 2025, có cột trống xen kẽ). **Đọc hàng 11 để lập `năm -> cột`**, không giả định cố định cột 10..13.
- Từ hàng 12 trở đi: **mỗi chỉ tiêu chiếm 2 hàng**. Hàng trên có số liệu (cột 2 = STT, cột 4 = tên, cột 8 = mã); hàng dưới là % Dọc/Ngang. **Chỉ lấy hàng số liệu** (có STT dạng `000001`...), bỏ hàng %.
- Giá trị thường là chuỗi có dấu phẩy (`205,622,500.00`) - parse về số trước khi ghi.

## Bản đồ mã chỉ tiêu (PTDN02 / BCTC -> KQKDSXKD)

Map theo **mã**, không theo vị trí dòng:

| Mã nguồn | Chỉ tiêu | Ghi vào KQKDSXKD |
| :---: | :--- | :--- |
| 01 | Doanh thu bán hàng và cung cấp dịch vụ | 01 |
| 02 | Các khoản giảm trừ doanh thu | 02 |
| 10 | Doanh thu thuần | 10 |
| 11 | Giá vốn hàng bán | 11 |
| 20 | Lợi nhuận gộp | 20 |
| 21 | Doanh thu hoạt động tài chính | 21 |
| 22 | Chi phí tài chính | 22 |
| 23 | Trong đó: chi phí lãi vay | 23 |
| **24** | Chi phí quản lý kinh doanh (gộp, form cũ/TT133) | **26** (CP QLDN); mã **25** = 0 nếu không tách được CP bán hàng |
| 25 | Chi phí bán hàng (khi PTDN có tách) | 25 |
| 26 | Chi phí quản lý doanh nghiệp | 26 |
| 30 | Lợi nhuận thuần từ HĐKD | 30 |
| 31 | Thu nhập khác | 31 |
| 32 | Chi phí khác | 32 |
| 40 | Lợi nhuận khác | 40 |
| 50 | Tổng lợi nhuận kế toán trước thuế | 50 |
| 51 | Chi phí thuế TNDN hiện hành | 51 |
| 52 | Chi phí thuế TNDN hoãn lại | 52 (không có thì 0) |
| 60 | Lợi nhuận sau thuế | 60 |

Cột năm trên `KQKDSXKD`: đọc hàng 11 sheet đích (thường 2022=cột 15 ... 2025=cột 18). Sửa mã `20` nếu ô mã bị hỏng thành `0`.

## Quy trình

### A. Sheet `KQKDSXKD`

1. **Đổi header vụ** ngay (tránh sót tên công ty mẫu như Hoàng Danh Phát):
   - `F5` = `Từ năm YYYY đến năm YYYY`
   - `F6` = MST
   - `F7` = tên NNT đúng vụ
2. Đọc PTDN02: MST/kỳ (hàng 5-6), map năm (hàng 11), lấy số theo mã.
3. Năm PTDN = trống hoặc toàn 0 trong khi DN có hoạt động → đánh dấu thiếu, **không ghi 0 giả**.
4. Bù BCTC XML năm thiếu: trong `PL_KQHDSXKD` chỉ lấy con của **`NamNay`** (năm báo cáo). `NamTruoc` là số so sánh năm trước - **cấm** ghi vào cột năm hiện tại.
5. Ghi số thô (đồng) đúng cột năm theo mã. Xóa số liệu công ty mẫu còn sót ở các năm ngoài kỳ (nếu đang làm vụ mới từ file nhân bản).
6. Báo rõ năm nào lấy PTDN02, năm nào lấy BCTC.

### B. Sheet `GTGT` (khi có KQ2.1)

1. Mở `KQ2_Doi_chieu_GTGT` Phần I (theo tờ khai GTGT).
2. Khối năm trên Bảng tính: `Năm 2022` ~ hàng 92, `2023` ~ 106, `2024` ~ 120, `2025` ~ 134 (Q1-Q4 ngay dưới header năm).
3. Map cột nhập (giữ công thức F/G/H/J nếu có):
   - E = doanh thu chịu 8%/10%
   - I = thuế GTGT đầu ra 8%/10%
   - K = giá trị mua vào
   - L = thuế GTGT đầu vào được khấu trừ
   - M/N = điều chỉnh giảm/tăng; O = nộp vãng lai; P = đề nghị hoàn
4. Quý không có tờ khai trên KQ2 → **để trống**, không điền 0.
5. Xóa số mẫu công ty cũ trong các khối năm thuộc kỳ trước khi đổ số mới.

### C. Sheet `EBITDA` (khi có KQ2.1)

1. Đổi header cột năm cho khớp kỳ (ví dụ `Năm 2022` ... `Năm 2025`, cột Tổng).
2. Copy từng dòng STT 1-9 từ `KQ2_EBITDA_Phan_tich` sang các hàng tương ứng trên sheet `EBITDA` (LN, khấu hao, lãi vay, lãi tiền gửi, tổng, EBITDA 30%, lãi vay không hợp lệ, đã loại trừ, chênh lệch).
3. Nếu lãi vay trên EBITDA (KQ2) khác mã 22/23 trên `KQKDSXKD` (BCTC) → **giữ đúng nguồn từng sheet** và nói rõ lệch cho user, không tự hòa.

### D. Bàn giao

Chạy [[boc-tach-bang-tinh-phan-tich]] để tính tỷ lệ chi phí/doanh thu và viết Mục III/IV DI-01.

## Ràng buộc

- Khối Q: không bịa số; không điền 0 giả cho năm/quý thiếu nguồn; thiếu cả PTDN02 lẫn BCTC thì ghi "chưa có dữ liệu".
- Ưu tiên PTDN02 cho `KQKDSXKD`; chỉ dùng BCTC khi PTDN thiếu/0; ghi rõ nguồn từng năm.
- Không sửa file nguồn PTDN02 / BCTC / KQ2.1.
- File Bảng tính đích nằm trong `0. BIEN BAN\` của đúng vụ đang làm.

## Pitfalls

- Lấy nhầm hàng % Dọc/Ngang của PTDN02 → số sai hoàn toàn. Khóa theo hàng có mã cột 8 + STT.
- Năm cuối kỳ = 0 trên PTDN02 vì chưa chốt TMS → phải bù BCTC, không bỏ qua.
- BCTC XML: lấy nhầm `NamTruoc` sẽ ghi trùng số năm trước vào năm hiện tại (đã xảy ra vụ TT Biên Hòa 2025).
- PTDN form cũ chỉ có mã **24** gộp: phải đổ sang **26**, không để trống cả 25 và 26.
- File nhân bản còn header/số của công ty mẫu (Hoàng Danh Phát...) → bắt buộc đổi `F5/F6/F7` và xóa số ngoài kỳ trước khi bàn giao.
- Cột năm PTDN và `KQKDSXKD` lệch nhau: luôn map theo mã + năm đọc từ hàng 11.
- Đổ GTGT/EBITDA từ bản KQ2 không phải `*_FINAL` / bản cũ hơn sẽ lệch tờ khai mới nhất.

## Kiểm chứng

- `KQKDSXKD!F6` = MST vụ; `F7` = tên NNT vụ; không còn tên công ty mẫu.
- Các năm trong kỳ có đủ mã 01/10/11/26/30/50/60 (hoặc đã ghi rõ năm thiếu nguồn).
- Năm bù BCTC: đối chiếu nhanh 1 chỉ tiêu (ví dụ mã 10 hoặc 50) với `NamNay` XML.
- `GTGT`: ít nhất các quý có số trên KQ2.1 đã vào đúng khối năm; quý trống vẫn trống.
- `EBITDA`: header năm đúng kỳ; dòng 1 và dòng lãi vay khớp KQ2.1.
- Có file backup `_before_nhap_` trong `0. BIEN BAN\`.
