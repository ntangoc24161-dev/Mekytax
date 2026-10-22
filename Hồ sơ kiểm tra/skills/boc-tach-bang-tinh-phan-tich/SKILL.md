---
name: Bóc tách Bảng tính phân tích tài chính NNT
description: 'Bóc tách Bảng tính NNT (sheet KQKDSXKD): lấy KQKD theo năm, tính tỷ lệ chi phí/doanh thu, điền phân tích Mục III và đề xuất nội dung kiểm tra DI-01.'
group: Tài chính
updated: '2026-10-15'
---
# Bóc tách Bảng tính phân tích tài chính NNT

Bóc số liệu kết quả kinh doanh theo năm từ file `Bang tinh <kỳ>.xlsx` của doanh nghiệp, tính tỷ lệ chi phí trên doanh thu, rồi điền phần phân tích tài chính (Mục III) và đề xuất nội dung kiểm tra (Mục IV) của Bản phân tích **DI-01**. Skill này đứng trước `dien-bien-ban-tu-bang-tinh` (điền BBXN/BBKT) và `ho-so-kiem-tra-ke-hoach` (chuẩn hóa/xuất).

## Khi nào dùng

- "bóc tách dữ liệu từ Bang tinh.xlsx", "lấy số liệu KQKD", "tỷ lệ chi phí trên doanh thu".
- "sửa DI-01 phần phân tích hồ sơ", "cập nhật đề xuất nội dung kiểm tra".
- Khi cần số liệu tài chính nhiều năm của một MST để chọn chuyên đề kiểm tra.

## Nguồn dữ liệu

File bảng tính của vụ, ví dụ:
`G:\...\VAN PHAT THANH\0. BIEN BAN\Bang tinh 2021-2025.xlsx` (và bản `2016-2025`).

Dùng `openpyxl` với `read_only=True, data_only=True` để lấy giá trị đã tính sẵn. Nếu ô trả về `None` hoặc `#DIV/0!` nghĩa là công thức tham chiếu sheet khác chưa có cache: KHÔNG được coi là 0 giả, phải mở file bằng Excel lưu lại hoặc đọc ô gốc, rồi mới bóc.

## Sheet trục và bản đồ chỉ tiêu

Sheet **`KQKDSXKD`** là nguồn chính (Phân tích bảng cân đối phát sinh và kết quả kinh doanh):
- Hàng 5-8: kỳ báo cáo, MST, tên NNT, cơ quan thuế.
- Hàng 10: tên chỉ tiêu (cột 2) + mã chỉ tiêu (cột 6); hàng 11: các năm bắt đầu từ cột 8 (năm 2014) chạy sang phải. Phải khớp đúng cột theo năm cần lấy, không đếm nhầm.

Bản đồ Mã chỉ tiêu cần bóc:

| Chỉ tiêu | Mã | Vai trò phân tích |
| :--- | :---: | :--- |
| Doanh thu thuần | 10 | Mẫu số cho mọi tỷ lệ |
| Giá vốn hàng bán | 11 | Tỷ lệ giá vốn/DT |
| Doanh thu hoạt động tài chính | 21 | Đối chiếu |
| Chi phí tài chính | 22 | Tỷ lệ CPTC/DT; mã 23 là chi phí lãi vay |
| Chi phí bán hàng | 25 | Tỷ lệ CPBH/DT |
| Chi phí quản lý doanh nghiệp | 26 | Tỷ lệ CPQLDN/DT |
| Tổng lợi nhuận kế toán trước thuế | 50 | Đánh giá lãi/lỗ |

Sheet phụ khi cần: `So sánh DThu` (đối chiếu doanh thu GTGT vs TNDN), `GTGT`, `CDKT`/`CĐKT` (cân đối kế toán), `GNT`/`GNT TNDN` (giấy nộp tiền), `TINH PHAT + CHAM NOP` (tính phạt, tiền chậm nộp).

## Quy trình

1. **Xác định kỳ và MST** từ hàng 5-6 của `KQKDSXKD`, đối chiếu với sổ `Danh sach kiem tra quyet toan - JAVIS PRO.xlsm` để lấy đoàn, trạng thái, địa bàn.
2. **Bóc số liệu từng năm** theo bản đồ mã chỉ tiêu ở trên, đúng khoảng năm của thời kỳ kiểm tra. Ghi lại giá trị thô (đồng), không làm tròn khi lưu.
3. **Tính tỷ lệ** cho từng năm: giá vốn/DT, CPQLDN/DT, CPBH/DT, CPTC/DT (phần trăm, 2 chữ số thập phân). Mẫu số là doanh thu thuần (mã 10).
4. **Viết Mục III - Phân tích tài chính** trong DI-01: bảng số liệu các năm + nhận xét xu hướng (doanh thu tăng/giảm, tỷ lệ giá vốn cao bất thường, chi phí biến động không tương ứng doanh thu). Mỗi khẳng định gắn số thật từ bảng tính.
5. **Viết Mục IV - Đề xuất nội dung kiểm tra**, mỗi nội dung một dòng theo đúng cấu trúc:
   `- <tên nội dung kiểm tra>, thời kỳ kiểm tra: Từ năm <đầu> đến năm <cuối>.`
   Chọn chuyên đề từ các tỷ lệ bất thường vừa tính (ví dụ giá vốn cao → kiểm giá vốn; CPQLDN giảm sâu khó lý giải → kiểm CPQLDN).
6. **Đồng bộ ngược sổ Excel**: cập nhật cột Nội dung kiểm tra và Thời kỳ của MST trên `Danh sach kiem tra quyet toan - JAVIS PRO.xlsm`/`.xlsx` cho khớp DI-01.
7. **Nối sang điền biên bản / xuất hồ sơ**: bàn giao `dien-bien-ban-tu-bang-tinh` để đổ DTCHEP vào BBXN/BBKT, đồng bộ QĐ và chèn ngày; rồi `ho-so-kiem-tra-ke-hoach` để chuẩn hóa phôi, lint và xuất trọn gói.

## Ràng buộc

- Không bịa số: thiếu năm/thiếu ô thì ghi "chưa có dữ liệu", không điền 0 giả.
- Giữ nguyên thứ tự nội dung kiểm tra theo yêu cầu người dùng (ví dụ giá vốn trước, hoặc theo chuyên đề chốt trong [[ho-so-kiem-tra-mst-3601409441-chot-so-lieu-va-phan-cong]]).
- Nội dung loại chi phí tài chính: nguyên nhân thống nhất là chi phí lãi vay không tương ứng doanh thu, căn cứ pháp lý theo kỳ (xem [[loai-tru-noi-dung-chi-phi-tai-chinh-va-nam-khong-chenh-lech-]]).
- Năm không chênh lệch thì ghi riêng `*/ Không chênh lệch.`, không đưa vào khung giải trình.

## Pitfalls (bài học thực tế)

- File rất lớn, mở bằng Excel COM chậm: ưu tiên parse XML/`openpyxl` read-only.
- Đếm nhầm cột năm là lỗi hay gặp nhất: luôn đọc hàng 11 để định vị cột trước khi lấy số.
- Có 2 file (`2016-2025` và `2021-2025`): lấy đúng bản khớp thời kỳ kiểm tra đang làm.
