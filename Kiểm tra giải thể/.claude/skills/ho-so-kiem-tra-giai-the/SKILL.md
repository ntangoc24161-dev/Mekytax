---
name: Hồ sơ kiểm tra giải thể
description: "Lập và chuẩn hóa bộ hồ sơ kiểm tra quyết toán giải thể: bộ 22 văn bản, Mục III NSNN, văn bản từ chối nộp thừa, bù trừ hoàn thuế."
group: Tài chính
updated: '2026-11-13'
learned_updated: '2026-11-13'
---

# Hồ sơ kiểm tra giải thể

Quy trình nghiệp vụ chuyên biệt lập, điền số liệu và chuẩn hóa bộ hồ sơ kiểm tra quyết toán thuế phục vụ giải thể doanh nghiệp (đóng mã số thuế), áp dụng cho cả diện giải thể theo thông báo ngừng hoạt động và diện giải thể trên danh mục kế hoạch năm.

Vận hành theo **mô hình Hybrid 2 tầng** dùng chung nguồn dữ liệu `01_DA_BAN_HANH$` từ sổ `Danh sach kiem tra quyet toan pro.xlsm` (phân loại qua trường `Loai_hinh = GT`), đi kèm phôi Word giải thể độc lập [templates/MAUTONG HOP GIAI THE.docx](templates/MAUTONG HOP GIAI THE.docx) gồm 22 văn bản, tích hợp đầy đủ nghĩa vụ ngân sách nhà nước, xử lý tiền nộp thừa, và thủ tục bù trừ hoàn thuế giải thể theo quy định pháp luật mới nhất.

## Khi nào dùng

- "lập hồ sơ kiểm tra giải thể", "hồ sơ về giải thể", "quyết toán giải thể đóng MST".
- "chuẩn hóa TONG HOP giải thể", "điền biên bản giải thể từ bảng tính".
- "bù trừ hoàn thuế giải thể", "văn bản từ chối nhận lại tiền nộp thừa", "điền Mục III nghĩa vụ NSNN giải thể".
- Xử lý các hồ sơ doanh nghiệp giải thể phát sinh trên nền kế hoạch năm (như Sơn Tiền Hải) hoặc theo thông báo ngừng hoạt động (như Mo Fa) hoặc giải thể kèm hoàn thuế (như Đại Phát Đạt).

## Chuẩn bị

- File phôi tổng hợp chuẩn giải thể: `templates/MAUTONG HOP GIAI THE.docx` (đã liên kết trực tiếp `01_DA_BAN_HANH$`, chuẩn hóa 198 thẻ MERGEFIELD, dọn sạch hàng trống thừa theo NĐ 30).
- Biểu mẫu phụ trợ giải thể: `templates/MAU CAM KET DOANH THU VA CHI PHI.docx`, `templates/MAU DON XIN QUYET TOAN GIAI THE.docx`, `templates/MAU DE XUAT HOAN GIAI THE.docx`, `templates/MAU MS-03-DCS.docx`, `templates/MAU QD XPHC GIAI THE.docx`.
- File Bảng tính: `Bang tinh.xlsx` (các sheet `BB`, `DTCHEP`, `TINH PHAT + CHAM NOP (C T)`).
- Sổ Excel theo dõi: `Danh sach kiem tra quyet toan pro.xlsm` (sheet `01_DA_BAN_HANH$` qua Dòng 1 ẩn).
- Căn cứ pháp lý: `references/nen-phap-ly-va-can-cu-giai-the.md`.
- Danh mục 22 văn bản: `references/danh-muc-22-van-ban-giai-the.md`.
- Nghiệp vụ chi tiết: `references/nghiep-vu-nghia-vu-nsnn-va-hoan-bu-tru.md`.

## Quy trình Hybrid 2 tầng giải thể

### Tầng 1: Mail Merge trực tiếp từ sheet 01_DA_BAN_HANH
1. Mở phôi `templates/MAUTONG HOP GIAI THE.docx` trong Microsoft Word (tab **Mailings** tự động kích hoạt kết nối tới `01_DA_BAN_HANH$`).
2. Lọc theo điều kiện: `Loai_hinh` = `GT` để chọn đúng các doanh nghiệp giải thể (Sơn Tiền Hải, Châm Pic, Đường Hải Thủy, Công Nghệ Hoàng Huy, Thúy Hồng...).
3. Toàn bộ 198 trường `{ MERGEFIELD }` chuẩn hóa tự động khớp dữ liệu công ty, MST, số quyết định, trưởng đoàn, thành viên và các cột tiền bằng chữ.

### Tầng 2: Điền biên bản từ Bảng tính và hoàn thiện đặc thù giải thể
1. Bóc tách số liệu từ sheet `DTCHEP` của `Bang tinh.xlsx`:
   - Trưởng đoàn kiểm tra Chi phí quản lý doanh nghiệp và Chi phí bán hàng (TT=1, TT=2); thống nhất chuẩn hóa tên gọi là **`Chi phí quản lý doanh nghiệp`**.
   - Thành viên kiểm tra Chi phí tài chính (hoặc khối GTGT).
2. Điền Mục III Nghĩa vụ NSNN (Bắt buộc cho diện giải thể):
   - Công thức suy luận: `(3)=0, (4)=Kê khai, (5)=Truy thu, (6)=(4), (7)=(3)+(4)+(5)-(6)=(5)`.
3. Lập các văn bản đặc thù giải thể:
   - QĐ xử phạt VPHC và Tờ trình theo NĐ 125/2020 và NĐ 310/2025.
   - Văn bản từ chối nhận lại số tiền nộp thừa (điền đúng 10 số MST vào ô lưới).
   - Thủ tục hoàn kiêm bù trừ thu NSNN (MS-03-DCS và dự thảo QĐ hoàn bù trừ) nếu còn thuế GTGT khấu trừ.
4. Thể thức hành chính NĐ 30:
   - Các bảng tiêu đề bảo đảm đúng 1 nét gạch chuẩn dưới cơ quan ban hành và tiêu ngữ, không lặp dòng đệm thừa.
   - Xuất 02 file QĐKT và QĐGS độc lập chuẩn 1 trang A4.

## Bẫy

- **Bỏ trống Mục III Nghĩa vụ NSNN**: Hồ sơ giải thể bị đình trệ, bộ phận KK&KTT và Quản lý nợ không khóa được mã số thuế cho người nộp thuế.
- **Sót MST phôi cũ trên Văn bản từ chối nộp thừa**: Do văn bản này nằm ở cuối bộ và dùng cấu trúc ô lưới từng số, script thông thường hay bỏ sót dẫn đến sai thông tin pháp lý.
- **Bên nhận trùng bên giao ở Biên bản giao nhận**: Trưởng đoàn vừa là bên giao vừa ký bên nhận hồ sơ.
- **Căn cứ ban hành sau ngày quyết định**: Viện dẫn số và ngày Thông báo ngừng hoạt động rơi vào ngày sau ngày ký Quyết định kiểm tra.
- **Áp sai mức phạt vi phạm hành chính**: Trường hợp sau khi loại trừ chi phí mà doanh nghiệp vẫn còn lỗ hoặc thuộc diện miễn thuế (thuế phải nộp = 0), máy móc phạt 20% là sai luật; phải phạt mức tiền cố định (6.500.000 đ) theo quy định tại NĐ 310/2025/NĐ-CP.
- **Lệch lịch kiểm tra chạm ngày nghỉ**: Lên lịch kiểm tra 3 ngày nhưng rơi vào thứ Năm, thứ Sáu và Chủ nhật.

## Kiểm chứng

- Mở `TONG HOP.docx` kiểm tra đủ 22 văn bản, không sót placeholder sống dạng `«...»` hay `....`.
- Mục III Nghĩa vụ NSNN có đầy đủ số phát sinh các năm kiểm tra.
- Văn bản từ chối nhận lại tiền nộp thừa có đúng 10 chữ số MST của doanh nghiệp.
- Checkbox Báo cáo giám sát hiển thị `☑` rõ nét sau chữ.
- 02 file Quyết định kiểm tra và Giám sát xuất riêng gói gọn trong 1 trang A4.

## Lịch sử (tự học)

- [2026-10-30] Vụ giải thể Sơn Tiền Hải chuyển từ phôi cũ sang phôi Mo Fa để cập nhật đồng bộ Luật Thanh tra 2025, QĐ 1614 ngày 13/10/2020, QĐ 1377 ngày 30/06/2025; Mục III Nghĩa vụ NSNN bắt buộc điền đủ số phát sinh 5 năm theo công thức suy luận `(3)=0, (6)=(4), (7)=(3)+(4)+(5)-(6)`.
- [2026-10-30] Phát hiện lỗi bên nhận hồ sơ trùng chính bên giao tại Biên bản giao nhận trong phôi gốc; đã sửa bên nhận thành Nguyễn Thành Đồng (Kiểm tra viên thuế).
- [2026-10-30] Phát hiện lỗi sót MST của vụ khác (3600275611) trên Văn bản từ chối nhận lại số tiền nộp thừa của phôi Mo Fa; đã chuẩn hóa điền đúng MST của doanh nghiệp theo từng ô lưới ký tự.
- [2026-10-31] Chuẩn hóa tên gọi chỉ tiêu: quét regex thay thế triệt để 18 vị trí từ "chi phí quản lý kinh doanh" sang "chi phí quản lý doanh nghiệp" trong hồ sơ kiểm tra.
- [2026-10-31] Đóng gói bộ hồ sơ đề xuất hoàn thuế kiêm bù trừ thu NSNN giải thể (mẫu MS-03-DCS và Tờ trình đề xuất hoàn giải thể) áp dụng Điều 32 NĐ 181/2025/NĐ-CP.
- [2026-11-11] Quy tắc kiểm tra số gõ cứng trên DTCHEP hoạt động dạng cảnh báo mềm (non-blocking warning), ghi nhận vào open-questions mà không làm gián đoạn luồng tự động đổ số vào biên bản.
