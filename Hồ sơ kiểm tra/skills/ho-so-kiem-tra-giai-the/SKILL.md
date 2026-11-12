---
name: Hồ sơ kiểm tra giải thể
description: "Lập và chuẩn hóa bộ hồ sơ kiểm tra quyết toán giải thể: bộ 22 văn bản, Mục III NSNN, văn bản từ chối nộp thừa, bù trừ hoàn thuế."
group: Tài chính
updated: '2026-11-11'
learned_updated: '2026-11-11'
---

# Hồ sơ kiểm tra giải thể

Quy trình nghiệp vụ chuyên biệt lập, điền số liệu và chuẩn hóa bộ hồ sơ kiểm tra quyết toán thuế phục vụ giải thể doanh nghiệp (đóng mã số thuế), áp dụng cho cả diện giải thể theo thông báo ngừng hoạt động và diện giải thể trên danh mục kế hoạch năm.

Bộ tài liệu tổng hợp hoàn chỉnh gồm 22 văn bản (`MAUTONG HOP GIAI THE.docx`), tích hợp đầy đủ nghĩa vụ ngân sách nhà nước, xử lý tiền nộp thừa, và thủ tục bù trừ hoàn thuế giải thể theo quy định pháp luật mới nhất.

## Khi nào dùng

- "lập hồ sơ kiểm tra giải thể", "hồ sơ về giải thể", "quyết toán giải thể đóng MST".
- "chuẩn hóa TONG HOP giải thể", "điền biên bản giải thể từ bảng tính".
- "bù trừ hoàn thuế giải thể", "văn bản từ chối nhận lại tiền nộp thừa", "điền Mục III nghĩa vụ NSNN giải thể".
- Xử lý các hồ sơ doanh nghiệp giải thể phát sinh trên nền kế hoạch năm (như Sơn Tiền Hải) hoặc theo thông báo ngừng hoạt động (như Mo Fa) hoặc giải thể kèm hoàn thuế (như Đại Phát Đạt).

## Chuẩn bị

- File phôi tổng hợp chuẩn giải thể: `templates/MAUTONG HOP GIAI THE.docx` (hoặc file `TONG HOP.docx` trong thư mục vụ `0. BIEN BAN`).
- Biểu mẫu hoàn bù trừ giải thể: `templates/MAU DE XUAT HOAN GIAI THE.docx`, `templates/MAU MS-03-DCS.docx`.
- File Bảng tính: `Bang tinh.xlsx` (các sheet `BB`, `DTCHEP`, `TINH PHAT + CHAM NOP (C T)`).
- Sổ Excel theo dõi: `Danh sach kiem tra quyet toan pro.xlsm` (thư mục `00. Theo doi chung`).
- Căn cứ pháp lý: `references/nen-phap-ly-va-can-cu-giai-the.md`.
- Danh mục 22 văn bản: `references/danh-muc-22-van-ban-giai-the.md`.
- Nghiệp vụ chi tiết: `references/nghiep-vu-nghia-vu-nsnn-va-hoan-bu-tru.md`.

## Quy trình

### 1. Xác định diện phát sinh và chuẩn hóa nền pháp lý
1. Xác định nguồn gốc cuộc kiểm tra:
   - **Diện 1 - Theo thông báo ngừng hoạt động (như Mo Fa):** Căn cứ thông báo của CQT hoặc cơ quan ĐKKD; thời hạn kiểm tra rút gọn (thường 03 ngày làm việc).
   - **Diện 2 - Giải thể trên nền kế hoạch năm (như Sơn Tiền Hải):** Vừa có căn cứ Quyết định kế hoạch năm (`QĐ 1412/QĐ-DON`), vừa chốt số đóng MST; thời hạn 20 ngày.
   - **Diện 3 - Giải thể kiêm hoàn thuế GTGT (như Đại Phát Đạt):** Kết hợp kiểm tra hoàn thuế theo Điều 32 NĐ 181/2025/NĐ-CP và chốt số giải thể.
2. Áp dụng nền pháp lý mới nhất:
   - Luật Quản lý thuế số 108/2025/QH15 ngày 10/12/2025; Luật Thanh tra số 84/2025/QH15 ngày 25/06/2025.
   - Quyết định 1614/QĐ-TCT ngày 13/10/2020; Quyết định 1377/QĐ-CT ngày 30/06/2025.
   - Thống nhất gọi tên cuộc kiểm tra là "kiểm tra thuế" trên Quyết định và Biên bản; lý do "kiểm tra quyết toán giải thể" chỉ để ở Tờ trình.

### 2. Bóc tách Bảng tính và điền số liệu (BBXN và BBKT)
1. Bóc tách số liệu từ sheet `DTCHEP`:
   - Trưởng đoàn kiểm tra Chi phí quản lý doanh nghiệp và Chi phí bán hàng (TT=1, TT=2); thống nhất chuẩn hóa tên gọi là **`Chi phí quản lý doanh nghiệp`** (loại bỏ hoàn toàn cụm từ cũ `Chi phí quản lý kinh doanh`).
   - Thành viên kiểm tra Chi phí tài chính (hoặc khối GTGT).
2. Kiểm tra số gõ cứng (Hardcoded Data Check - Cảnh báo mềm):
   - Mở file Excel kiểm tra công thức các ô chênh lệch lớn (như ô `DTCHEP!E372`).
   - Nếu là số gõ cứng không công thức và không có sheet kê chi tiết: **vẫn tiếp tục đổ số bình thường**, ghi nhận dòng `[LƯU Ý NGHIỆP VỤ]` ở cuối báo cáo và append vào `wiki/_open-questions.md`.
3. Điền BBKT giải thể:
   - Thể hiện đầy đủ chênh lệch các năm, giảm lỗ lũy kế, thu nhập tính thuế sau giảm lỗ, thuế TNDN truy thu, tiền phạt và tiền chậm nộp.
   - Căn cứ loại trừ chi phí: Điều 4 TT 96/2015/TT-BTC (trước 2025) hoặc Điều 10 NĐ 320/2025/NĐ-CP + TT 20/2026/TT-BTC (từ 2025).

### 3. Điền Mục III Nghĩa vụ NSNN (Bắt buộc cho diện giải thể)
1. Tuyệt đối không để trống hoặc dùng dấu gạch ngang.
2. Áp dụng công thức suy luận từ kết luận kê khai và nợ đọng:
   - `(3) = 0` (không có nợ cũ chuyển sang).
   - `(4)` = Số phát sinh theo kê khai của doanh nghiệp.
   - `(5)` = Số truy thu xác định qua kiểm tra.
   - `(6) = (4)` (đã nộp đủ số kê khai).
   - `(7) = (3) + (4) + (5) - (6) = (5)` (số còn phải nộp chính bằng số truy thu).

### 4. Lập các văn bản đặc thù giải thể
1. **Quyết định xử phạt VPHC (QĐXP) và Tờ trình**:
   - Định danh hành vi chuẩn theo NĐ 125/2020 sửa đổi bởi NĐ 310/2025.
   - Có phát sinh thiếu thuế: Phạt 20% trên số thuế thiếu.
   - Không phát sinh thiếu thuế (do lỗ hoặc được miễn giảm thuế theo NĐ 20/2026): Phạt tiền cố định 6.500.000 đồng (khoản 3 Điều 16 NĐ 125/2020, khoản 5 Điều 1 NĐ 310/2025).
2. **Văn bản từ chối nhận lại số tiền nộp thừa**:
   - Điền đầy đủ mã số thuế của doanh nghiệp theo từng ô lưới riêng biệt (10 ô).
   - Rà soát tránh sót MST của phôi cũ.
3. **Thủ tục hoàn kiêm bù trừ thu NSNN (nếu có thuế GTGT còn khấu trừ)**:
   - Lập Tờ trình đề xuất hoàn thuế, Phiếu đề xuất hoàn thuế `MS-03-DCS`, và Dự thảo QĐ hoàn thuế kiêm bù trừ thu NSNN theo Điều 32 NĐ 181/2025/NĐ-CP.
   - Ưu tiên bù trừ toàn bộ số thuế GTGT được hoàn với số thuế TNDN truy thu, tiền phạt và chậm nộp.

### 5. Chuẩn hóa thể thức hành chính theo NĐ 30 và xuất hồ sơ
1. **Thể thức hành chính**:
   - Tờ trình, QĐKT, QĐGS, QĐXP: Chữ ký cách chức danh mặc định 5 dòng trong 1 ô (nếu rớt sang trang mới thì tự động giảm xuống 4 dòng theo NĐ 30 để cùng trang in).
   - Dự thảo Quyết định: Chữ 'Số' thẳng dưới chữ 'sở'; ký hiệu để khoảng trống dài `Số:              /QĐ-TCS3` (tuyệt đối không dùng mã tạm `DU_THAO-TRINH-KY`); ngày tháng căn giữa để trống toàn bộ `Trấn Biên, ngày       tháng       năm      ` (bỏ năm 2026); các Điều 1, 2, 3... tô đậm. Căn cứ QĐ 1377/QĐ-CT chốt thống nhất ngày 30/06/2025. Khi đã ban hành chính thức: không ghi đè lại số/ngày.
   - Biên bản giao nhận hồ sơ: Bên nhận phải là thành viên khác (Nguyễn Thành Đồng), tuyệt đối không để trùng tên Trưởng đoàn bên giao (Nguyễn Thị Ánh Ngọc).
2. **Báo cáo giám sát**:
   - Đổi toàn bộ checkbox sang ký tự Unicode `☐`/`☑` nằm SAU nhãn chữ.
   - Tick sẵn 9 tiêu chí đạt chuẩn; Mục 9 phải có đủ 3 lựa chọn với 3 ô checkbox riêng.
3. **Xuất file và đồng bộ**:
   - Xuất 02 file QĐKT và QĐGS độc lập chuẩn 1 trang A4.
   - Xuất file QĐXP độc lập và cập nhật dòng tương ứng trên sổ theo dõi `Danh sach kiem tra quyet toan pro.xlsm`.

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
