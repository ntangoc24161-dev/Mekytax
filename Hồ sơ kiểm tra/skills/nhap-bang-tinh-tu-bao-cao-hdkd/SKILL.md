---
name: Nhập số liệu Bảng tính từ báo cáo HĐKD ngành thuế
description: Điền KQKDSXKD từ PTDN02 (thiếu thì bù BCTC NamNay hoặc QTT 03-1A); có KQ2.1 thì đổ GTGT và EBITDA.
group: Tài chính
updated: '2026-11-09'
learned_updated: '2026-11-09'
---
# Nhập số liệu Bảng tính từ báo cáo HĐKD ngành thuế

Chiều NHẬP LIỆU đứng trước [[boc-tach-bang-tinh-phan-tich]]: lấy số liệu kết quả hoạt động kinh doanh theo năm từ báo cáo phân tích HĐKD của ngành thuế (mẫu PTDN02, xuất từ TMS) rồi điền vào sheet `KQKDSXKD` của file `Bang tinh.xlsx`. Năm PTDN thiếu hoặc toàn 0 thì bù BCTC khối `NamNay`; không có BCTC XML thì bù QTT 03-1A/TNDN. Có `KQ2.1_Phan_tich_CQT_*_FINAL.xlsx` thì đổ thêm sheet `GTGT` và `EBITDA`. Điền xong mới chuyển sang skill phân tích để viết DI-01.

## Khi nào dùng

- "bóc tờ khai điền vào bảng tính", "hoàn chỉnh file bảng tính", "nhập số liệu KQKD", "nhập luôn", "bóc tách dữ liệu thiếu từ bctc".
- "nhập GTGT/EBITDA từ KQ2.1", "đổ KQ2 vào Bảng tính".
- Có file `PTDN02_*.xlsx` trong `1. DU LIEU THO NGANH THUE\1.3. Nghia vu nop thue & Thong bao\` của NNT.
- Cần dựng số liệu nhiều năm cho một MST trước khi tính tỷ lệ và chọn chuyên đề kiểm tra.
- Không dùng để tạo phôi hay đổ CDKT từ PTDN01 (đó là `tao-bang-tinh-tu-ptdn`).

## Chuẩn bị

1. Thư mục vụ đã có `0. BIEN BAN\Bang tinh.xlsx`, header `F5/F6/F7` đúng vụ (skill `tao-bang-tinh-tu-ptdn`). Chưa có phôi thì dừng, không nhập đè lên mẫu chung `1. TD\BIEN BAN\Bang tinh.xlsx`.
2. Xác định MST, tên NNT, thời kỳ user đã chốt. MST chưa có trên sổ thì không lấy kỳ của vụ kế hoạch khác.
3. Backup trước khi ghi: `Bang tinh_backup_before_nhap_YYYYMMDD_HHMMSS.xlsx` trong cùng `0. BIEN BAN\`.
4. Nguồn:

| Sheet đích | Nguồn chính | Nguồn bù |
| :--- | :--- | :--- |
| `KQKDSXKD` | `PTDN02_*.xlsx` (TMS) | BCTC XML khối `NamNay`, rồi QTT 03-1A/TNDN. Cấm lấy `NamTruoc` |
| `GTGT` | `KQ2.1_*_FINAL.xlsx` sheet `KQ2_Doi_chieu_GTGT` | Quý không có tờ khai để trống, không ghi 0 giả |
| `EBITDA` | `KQ2.1_*_FINAL.xlsx` sheet `KQ2_EBITDA_Phan_tich` | Giữ đúng số KQ2.1 dù lệch `KQKDSXKD` |

## Cách chạy

1. Đọc PTDN02 bằng openpyxl `read_only=True, data_only=True`. Nếu vùng số toàn trống trong khi file có dung lượng, đó là inline string của bản xuất TMS: parse XML thô của xlsx, không kết luận PTDN rỗng. Đã gặp `NHat vinh-PTDN02_20260524183525.xlsx` và `DAI HUY HOANG-PTDN02_20260524182807.xlsx` ngày 09/11/2026.
2. Ghi `Bang tinh.xlsx` bằng openpyxl hoặc Excel COM. Không sửa file nguồn PTDN02, BCTC, QTT, KQ2.1.
3. Đọc năm ở hàng 11 của cả PTDN02 và `KQKDSXKD` để lập `năm -> cột`. Không giả định cột cố định.

## Quy trình

### A. Đọc PTDN02

- Hàng 5-8: kỳ, MST, tên NNT, cơ quan thuế. Đối chiếu đúng MST.
- Hàng 10: tên chỉ tiêu cột 4, mã chỉ tiêu cột 8.
- Hàng 11: năm rải theo cột, có cột trống xen kẽ.
- Từ hàng 12: mỗi chỉ tiêu 2 hàng. Chỉ lấy hàng có STT dạng `000001` và mã ở cột 8. Bỏ hàng % Dọc/Ngang.
- Giá trị thường là chuỗi có dấu phẩy (`205,622,500.00`). Parse về số rồi mới ghi.

Map theo mã, không theo vị trí dòng:

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
| 24 | Chi phí quản lý kinh doanh (gộp, form cũ) | 26; mã 25 = 0 nếu không tách được CP bán hàng. Có `TMBCTC_*.xlsx` mục 6 thì tách: a) CPQLDN → 26, b) CPBH → 25; tổng 25+26 phải bằng mã 24 |
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

Sửa mã `20` trên phôi nếu ô mã bị hỏng thành `0`.

### B. Sheet KQKDSXKD

1. Nếu `KQKDSXKD!F5:F7` là công thức trỏ `CDKT`, không ghi đè bằng chữ. Header vụ đã nằm ở `CDKT!F5:F7`.
2. Ghi số thô (đồng) đúng cột năm theo mã. Xóa số công ty mẫu còn sót ở năm ngoài kỳ.
3. Năm PTDN trống hoặc toàn 0 trong khi DN có hoạt động: không ghi 0 giả. Thứ tự bù: BCTC XML chỉ lấy con của `NamNay` trong `PL_KQHDSXKD`; không có file BCTC XML thì lấy QTT 03-1A/TNDN; không có cả hai thì để trống và ghi "chưa có dữ liệu".
4. Chỉ tiêu phụ lục bù không có: để trống và xóa công thức mẫu còn sót ở ô đó. Vụ Nhất Vinh năm 2025 còn sót công thức mẫu ở mã 30; mã 22/23/30/31/32/40 để trống vì phụ lục không có.
5. Giữ đúng số nguồn. Không tự hòa khi các mã trên cùng nguồn lệch nhau (vụ Nhất Vinh 2025: mã 01 = 10.162.614.641, mã 10 = 7.898.994.937).
6. Phôi thiếu cột của một năm nằm trong kỳ user chốt (vụ Nhất Vinh, phôi nhảy 2019 sang 2021): được chèn cột năm đó, rồi sửa tham chiếu công thức từ hàng 34 trở đi cho đúng năm. Không bàn giao khi công thức còn trỏ cột cũ. Năm ngoài kỳ mà phôi không có cột (vụ Đại Huy Hoàng, năm 2020): không tự chèn, nói rõ năm chưa ghi.
7. Báo rõ từng năm lấy PTDN02, BCTC `NamNay` hay QTT 03-1A.

### C. Sheet GTGT (khi có KQ2.1 FINAL)

1. Mở `KQ2_Doi_chieu_GTGT` Phần I.
2. Khối năm trên Bảng tính: `Năm 2022` khoảng hàng 92, `2023` khoảng 106, `2024` khoảng 120, `2025` khoảng 134.
3. Map cột nhập, giữ công thức F/G/H/J nếu có: E doanh thu chịu 8%/10%; I thuế đầu ra 8%/10%; K giá trị mua vào; L thuế đầu vào được khấu trừ; M/N điều chỉnh giảm/tăng; O nộp vãng lai; P đề nghị hoàn.
4. Quý không có tờ khai để trống. Xóa số quý mẫu của công ty cũ trong các khối năm thuộc kỳ trước khi đổ số mới.

### D. Sheet EBITDA (khi có KQ2.1 FINAL)

1. Đổi header cột năm cho khớp kỳ, kể cả cột Tổng.
2. Copy dòng STT 1-9 từ `KQ2_EBITDA_Phan_tich` sang đúng hàng trên sheet `EBITDA`.
3. Lãi vay trên EBITDA khác mã 22/23 của `KQKDSXKD`: giữ đúng nguồn từng sheet và nói rõ lệch, không tự hòa.

### E. Bàn giao

Nêu đường dẫn `0. BIEN BAN\Bang tinh.xlsx`, nguồn từng năm, năm chưa ghi, và ô để trống vì không có nguồn. Chưa có `KQ2.1_*_FINAL` thì nói rõ GTGT và EBITDA chưa đổ. Bước sau, khi user yêu cầu, là [[boc-tach-bang-tinh-phan-tich]].

## Bẫy

- openpyxl `data_only` đọc PTDN02 xuất TMS ra toàn ô trống vì inline string. Parse XML thô trước khi kết luận file không có số.
- Lấy nhầm hàng % Dọc/Ngang làm sai cả cột. Khóa hàng có mã cột 8 và STT.
- Lấy `NamTruoc` ghi trùng số năm trước vào năm hiện tại (đã xảy ra vụ TT Biên Hòa 2025).
- Form cũ chỉ có mã 24: đổ sang mã 26, mã 25 = 0. Không để trống cả 25 và 26. Nếu có thuyết minh BCTC (`TMBCTC_*.xlsx` mục 6 Chi phí quản lý kinh doanh) thì tách 25/26, không để 25 = 0. Cột "Kỳ trước" của file năm N chỉ dùng để tách năm N-1 khi năm đó đã có mã 24 gộp khớp tổng; không lấy `NamTruoc` của `PL_KQHDSXKD` làm số năm N.
- Chèn cột năm làm lệch công thức từ hàng 34 trở đi. Sửa tham chiếu trước khi bàn giao.
- Ô năm bù còn công thức mẫu (mã 30 và các mã phụ lục không có). Xóa công thức, để trống, không ghi 0 giả.
- Không tự sửa số nguồn lệch nội bộ (giá vốn quá nhỏ so với doanh thu, mã 01 khác mã 10, mã 23 = 0 trong khi mã 22 có số). Giữ số và nêu trong bàn giao.
- Đổ GTGT/EBITDA từ bản KQ2 không phải `*_FINAL` sẽ lệch tờ khai mới nhất. Không có file FINAL thì bỏ hai sheet này.
- File nhân bản còn header hoặc số công ty mẫu: đổi `F5/F6/F7` và xóa số ngoài kỳ trước khi bàn giao.

## Kiểm chứng

- `KQKDSXKD!F6` = MST vụ; `F7` = tên NNT vụ; không còn tên công ty mẫu.
- Mỗi năm trong kỳ có mã 01/10/11/26/50 hoặc đã ghi rõ năm và mã nào thiếu nguồn.
- Năm bù BCTC: đối chiếu ít nhất mã 10 hoặc mã 50 với `NamNay`. Năm bù QTT: đối chiếu mã 10 với phụ lục 03-1A.
- Sau khi chèn cột năm: công thức từ hàng 34 trỏ đúng cột năm, không còn `#REF!` hay lệch năm.
- Ô không có nguồn không còn công thức mẫu và không bị ghi 0 giả.
- `GTGT`: quý có số trên KQ2.1 FINAL đã vào đúng khối năm; quý trống vẫn trống.
- `EBITDA`: header năm đúng kỳ; dòng 1 và dòng lãi vay khớp KQ2.1.
- Có file backup `_before_nhap_` trong `0. BIEN BAN\`.

## Lịch sử (tự học)
- [2026-11-09] Vụ Nông sản Lưu Hoài Anh (MST 3603057410): PTDN mã 24 gộp nên lần nhập đầu ghi 25 = 0. `TMBCTC_2025` mục 6 tách được CPBH/CPQLDN năm 2025 (Kỳ này) và 2024 (Kỳ trước, tổng khớp mã 24). Không có TMBCTC 2022-2023 nên hai năm đó giữ 25 = 0.
- [2026-11-09] Batch 09/11/2026 làm theo skill cũ bị thiếu ba bước. openpyxl data_only đọc trống PTDN02 Nhật Vinh và Đại Huy Hoàng vì inline string, phải parse XML thô. Chèn cột 2020 trên KQKDSXKD Nhật Vinh làm lệch công thức từ hàng 34. Năm 2025 Nhật Vinh không có BCTC XML nên bù QTT 03-1A, và mã 30 còn sót công
