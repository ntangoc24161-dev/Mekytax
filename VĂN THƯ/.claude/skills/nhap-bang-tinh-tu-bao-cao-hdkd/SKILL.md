---
name: Nhập số liệu Bảng tính từ báo cáo HĐKD ngành thuế
description: 'Điền sheet KQKDSXKD của Bảng tính từ báo cáo phân tích HĐKD ngành thuế (PTDN02) theo mã chỉ tiêu; năm hoặc ô thiếu thì bù bằng BCTC sau cùng.'
group: Tài chính
updated: '2026-10-15'
---
# Nhập số liệu Bảng tính từ báo cáo HĐKD ngành thuế

Chiều NHẬP LIỆU đứng trước [[boc-tach-bang-tinh-phan-tich]]: lấy số liệu kết quả hoạt động kinh doanh theo năm từ **báo cáo phân tích HĐKD của ngành thuế (mẫu PTDN02, xuất từ TMS)** rồi điền vào sheet `KQKDSXKD` của file `Bang tinh <kỳ>.xlsx`. Khi báo cáo thiếu năm hoặc ô, bù bằng **báo cáo tài chính (BCTC) sau cùng** của năm đó. Điền xong mới chuyển sang skill phân tích để viết DI-01.

## Khi nào dùng

- "bóc tờ khai điền vào bảng tính", "hoàn chỉnh file bảng tính", "nhập số liệu KQKD".
- Có file `PTDN02_*.xlsx` trong thư mục `1. DU LIEU THO NGANH THUE\1.3. Nghia vu nop thue & Thong bao\` của NNT.
- Cần dựng số liệu nhiều năm cho một MST trước khi tính tỷ lệ và chọn chuyên đề kiểm tra.

## Nguồn dữ liệu

**Nguồn chính - PTDN02** (Phân tích báo cáo kết quả hoạt động kinh doanh), ví dụ:
`G:\...\10. TAM HOANG PHAT\1. DU LIEU THO NGANH THUE\1.3. Nghia vu nop thue & Thong bao\PTDN02_20260828084112.xlsx`

**Nguồn bù khi thiếu - BCTC sau cùng:** báo cáo tài chính (bảng KQKD) hoặc QTT 03/TNDN của năm bị thiếu, lấy bản nộp sau cùng.

Đọc bằng `openpyxl` `read_only=True, data_only=True`.

## Cấu trúc file PTDN02

- Hàng 5-8: kỳ báo cáo, MST, tên NNT, cơ quan thuế. Lấy để đối chiếu đúng MST và kỳ.
- Hàng 10: nhãn cột (Tên chỉ tiêu ở cột 4, Mã chỉ tiêu ở cột 8, Số liệu trên BCTC từ cột 10).
- Hàng 11: **năm** nằm ở cột 10..13 (ví dụ 10=2022, 11=2023, 12=2024, 13=2025). Phải đọc hàng này để định vị cột năm, không đếm mò.
- Từ hàng 12 trở đi: **mỗi chỉ tiêu chiếm 2 hàng**. Hàng trên có số liệu (cột 2 = STT, cột 4 = tên, cột 8 = mã, cột 10..13 = số theo năm); hàng dưới chỉ là phân tích Dọc/Ngang (%). **Chỉ lấy hàng số liệu, bỏ hàng %.**

## Bản đồ mã chỉ tiêu (PTDN02 -> KQKDSXKD)

Cùng bộ mã ở cả hai file nên map thẳng theo mã, không map theo vị trí dòng:

| Mã | Chỉ tiêu |
| :---: | :--- |
| 01 | Doanh thu bán hàng và cung cấp dịch vụ |
| 02 | Các khoản giảm trừ doanh thu |
| 10 | Doanh thu thuần |
| 11 | Giá vốn hàng bán |
| 20 | Lợi nhuận gộp |
| 21 | Doanh thu hoạt động tài chính |
| 22 | Chi phí tài chính |
| 23 | Trong đó: chi phí lãi vay |
| 25 | Chi phí bán hàng |
| 26 | Chi phí quản lý doanh nghiệp |
| 30 | Lợi nhuận thuần từ HĐKD |
| 31 | Thu nhập khác |
| 32 | Chi phí khác |
| 40 | Lợi nhuận khác |
| 50 | Tổng lợi nhuận kế toán trước thuế |
| 51 | Chi phí thuế TNDN hiện hành |
| 52 | Chi phí thuế TNDN hoãn lại |
| 60 | Lợi nhuận sau thuế |

## Quy trình

1. **Xác định MST và kỳ** từ hàng 6 và 5 của PTDN02; đối chiếu với thời kỳ kiểm tra đang làm.
2. **Định vị cột năm** ở hàng 11; lập ánh xạ `năm -> cột`.
3. **Duyệt từng chỉ tiêu theo mã** (cột 8 hàng số liệu), đọc giá trị các năm ở cột 10..13. Bỏ qua hàng phân tích Dọc/Ngang.
4. **Phát hiện năm thiếu:** cột để trống hoặc toàn 0 ở năm mà DN thực tế có hoạt động (ví dụ năm cuối chưa vào TMS). Đánh dấu năm đó là thiếu, KHÔNG ghi 0 giả.
5. **Bù từ BCTC sau cùng:** với năm thiếu, mở BCTC/QTT 03 năm đó lấy đúng chỉ tiêu KQKD rồi điền. Ghi nguồn bù (BCTC ngày nộp nào) để truy vết.
6. **Ghi vào `KQKDSXKD`** đúng cột năm theo mã chỉ tiêu, giữ giá trị thô (đồng), không làm tròn. Dùng Excel COM hoặc `openpyxl` ghi, lưu lại file.
7. **Bàn giao** cho [[boc-tach-bang-tinh-phan-tich]] để tính tỷ lệ chi phí/doanh thu và viết Mục III, Mục IV của DI-01.

## Ràng buộc

- Khối Q: không bịa số, không điền 0 giả cho năm thiếu nguồn; thiếu cả PTDN02 lẫn BCTC thì ghi "chưa có dữ liệu" và nói rõ.
- Ưu tiên số liệu PTDN02 (ngành thuế); chỉ dùng BCTC khi PTDN02 thiếu, và ghi rõ dòng nào lấy từ BCTC.
- Không sửa file nguồn PTDN02/BCTC; chỉ ghi vào `KQKDSXKD` của Bảng tính.

## Pitfalls

- Mỗi chỉ tiêu 2 hàng: lấy nhầm hàng % Dọc/Ngang sẽ ra số sai hoàn toàn. Luôn khóa theo hàng có mã ở cột 8.
- Năm cuối kỳ thường = 0 trong PTDN02 vì quyết toán chưa chốt: đây đúng là ca phải bù BCTC sau cùng, không bỏ qua.
- PTDN02 và `KQKDSXKD` cùng mã nhưng khác vị trí cột năm: map theo mã + năm, không theo chỉ số dòng/cột cứng.
