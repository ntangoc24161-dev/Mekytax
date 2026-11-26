---
name: Hồ sơ kiểm tra giải thể
description: "Lập và chuẩn hóa bộ hồ sơ kiểm tra quyết toán giải thể: bộ 22 văn bản, Mục III NSNN, văn bản từ chối nộp thừa, bù trừ hoàn thuế."
group: Tài chính
updated: '2026-11-23'
learned_updated: '2026-11-23'
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
- Biểu mẫu phụ trợ giải thể: `templates/MAU CAM KET DOANH THU VA CHI PHI.docx` (1 công văn DN `/ĐKCP`: cam kết chi phí và chi phí lương), `templates/MAU DON XIN QUYET TOAN GIAI THE.docx`, `templates/MAU DE XUAT HOAN GIAI THE.docx`, `templates/MAU MS-03-DCS.docx`, `templates/MAU QD XPHC GIAI THE.docx` (khi còn thời hiệu 5 năm), `templates/MAU QD KHAC PHUC HAU QUA GIAI THE.docx` (khi hết thời hiệu xử phạt nhưng còn thời hạn truy thu 6-10 năm theo Mẫu 15/NĐ 118).
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
   - Phân công BBXN **theo sổ JAVIS PRO / chỉ đạo vụ** (cột `Phan_cong_TD` / `Phan_cong_TV`), không cứng TD = QLDN + bán hàng. Vụ Kiến Thành Tân: TD = Giá vốn hàng bán (1 chỉ tiêu); TV = Chi phí bán hàng + Chi phí quản lý doanh nghiệp (2 chỉ tiêu, TT=1/2 trong cùng bảng).
   - Tên chỉ tiêu thống nhất **`Chi phí quản lý doanh nghiệp`** khi sổ/user ghi vậy. Quét bảng chữ ký: sót tên kế toán phôi (`Lê Hoài Phương`) phải thay theo cột `Ke_toan_truong` trên sổ.
2. Điền Mục III Nghĩa vụ NSNN (Bắt buộc cho diện giải thể):
   - Công thức suy luận: `(3)=0, (4)=Kê khai, (5)=Truy thu, (6)=(4), (7)=(3)+(4)+(5)-(6)=(5)`.
3. Lập các văn bản đặc thù giải thể:
   - **Lựa chọn giữa ban hành QĐ xử phạt VPHC hay QĐ buộc thực hiện biện pháp khắc phục hậu quả** theo Điều 8 Nghị định 125/2020/NĐ-CP:
     + **Từ năm thứ 1 đến hết năm thứ 5**: Còn thời hiệu xử phạt VPHC. Doanh nghiệp bị phạt tiền (20% trên số thuế thiếu hoặc phạt tiền cố định) đồng thời nộp đủ thuế thiếu, tiền chậm nộp (0,03%/ngày) và điều chỉnh giảm lỗ -> Ban hành **Quyết định xử phạt VPHC (Mẫu số 02 ban hành kèm NĐ 68/2025/NĐ-CP)** theo phôi `templates/MAU QD XPHC GIAI THE.docx` và Tờ trình đề nghị ban hành QĐXP.
     + **Từ năm thứ 6 đến hết năm thứ 10**: Đã hết thời hiệu xử phạt tiền, nhưng còn trong thời hạn truy thu 10 năm. Doanh nghiệp KHÔNG bị phạt tiền (không mất 20% tiền phạt), nhưng vẫn bị áp dụng biện pháp khắc phục hậu quả: truy thu đủ tiền thuế thiếu, tiền chậm nộp (0,03%/ngày) và điều chỉnh giảm lỗ -> Ban hành **Quyết định buộc thực hiện biện pháp khắc phục hậu quả (Mẫu số 15 ban hành kèm NĐ 118/2021/NĐ-CP)** theo phôi chuẩn `templates/MAU QD KHAC PHUC HAU QUA GIAI THE.docx` và Tờ trình đề nghị ban hành QĐ buộc thực hiện biện pháp KPHQ.
     + **Từ năm thứ 11 trở về trước (quá 10 năm)**: Không bị phạt tiền và không bị truy thu (trừ trường hợp người nộp thuế không đăng ký thuế).
     + *Ví dụ cụ thể (phát hiện sai sót năm 2026)*: Sai sót 2021 - 2026 (trong 5 năm): Phạt 20% + Truy thu + Chậm nộp/Giảm lỗ (QĐ xử phạt VPHC); sai sót 2016 - 2020 (năm 6 đến năm 10): Không phạt 20%, chỉ Truy thu + Chậm nộp/Giảm lỗ (QĐ buộc khắc phục hậu quả); sai sót 2015 trở về trước: Không phạt, không truy thu.
     + **Xử lý xung đột khi kiểm tra nhiều năm (Hồ sơ hỗn hợp niên độ)**:
       * Khi vụ việc gồm nhiều năm, vừa có năm còn thời hiệu xử phạt (trong 5 năm), vừa có năm hết thời hiệu xử phạt nhưng còn thời hạn truy thu (năm thứ 6 đến năm thứ 10): **Không tách làm 02 quyết định**, vẫn ban hành **01 Quyết định xử phạt VPHC duy nhất (Mẫu số 02 - NĐ 68/2025/NĐ-CP)** cho toàn bộ vụ việc.
       * *Hình thức phạt tiền*: Chỉ tính và áp dụng phạt tiền (20% hoặc cố định) đối với số thuế khai sai của các năm còn thời hiệu (trong 5 năm gần nhất).
       * *Biện pháp khắc phục hậu quả*: Gộp toàn bộ số thuế truy thu và tiền chậm nộp của cả 2 giai đoạn (bao gồm cả các năm từ năm thứ 6 đến năm thứ 10).
       * *Điều kiện duy nhất áp dụng Mẫu số 15 (QĐ buộc KPHQ)*: Chỉ khi **toàn bộ** các hành vi vi phạm và niên độ kiểm tra đều đã hết thời hiệu phạt tiền (từ năm thứ 6 trở về trước, hoặc thuộc các trường hợp tại khoản 1 Điều 65 Luật XLVPHC).
     + *Thời hiệu theo từng loại hành vi*: Vi phạm thủ tục thuế, hóa đơn là **02 năm**; Khai sai thiếu thuế, trốn thuế là **05 năm**; Truy thu thuế và tiền chậm nộp là **10 năm**.
   - Văn bản từ chối nhận lại số tiền nộp thừa (điền đúng 10 số MST vào ô lưới). Cuối bộ chỉ **1** Công văn xin quyết toán giải thể; không nhân đôi.
   - Khi user yêu cầu mẫu cam kết chi phí / chi phí lương: **1** công văn DN `/ĐKCP` từ phôi `MAU CAM KET DOANH THU VA CHI PHI.docx` (hàng hóa dịch vụ + tiền lương thực tế trong cùng văn bản). Không gắn `CAM KẾT KHÔNG PHÁT SINH LAO ĐỘNG CHI TRẢ LƯƠNG` trừ khi DN thật sự không có lao động.
   - Thủ tục hoàn kiêm bù trừ thu NSNN (MS-03-DCS và dự thảo QĐ hoàn bù trừ) nếu còn thuế GTGT khấu trừ.
4. Thể thức hành chính NĐ 30:
   - Các bảng tiêu đề bảo đảm đúng 1 nét gạch chuẩn dưới cơ quan ban hành và tiêu ngữ, không lặp dòng đệm thừa.
   - **Mẫu 15/NĐ 118:** Quốc hiệu `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` nằm **trọn 1 dòng**, canh giữa từng ô; tiêu ngữ đậm + gạch dưới; tên cơ quan đậm + gạch dưới; không nhét cả khối header vào 1 run/`\n` (sẽ rớt chữ `NAM` / `ĐỒNG NAI`).
   - **Ngày ban hành:** dự thảo QĐ / tờ trình / báo cáo kết quả = ngày báo cáo trên sổ (`Ngay_bao_cao`); hồ sơ phân tích và kế hoạch giám sát = ngày QĐKT; các mốc còn lại theo sổ. Báo cáo kết quả phải có số QĐKT, không để `SỐ          /QĐ-TCS3`, không lặp năm (`năm 2026 2026`).
   - Tờ trình / QĐ / văn bản thẩm định: chức danh (CÁN BỘ ĐỀ XUẤT, TỔ TRƯỞNG, PHÓ TRƯỞNG) cách tên người ký **4 dòng**. Mỗi văn bản mới bắt đầu trang mới.
   - Xuất 02 file QĐKT và QĐGS độc lập chuẩn 1 trang A4.

## Bẫy

- **Bỏ trống Mục III Nghĩa vụ NSNN**: Hồ sơ giải thể bị đình trệ, bộ phận KK&KTT và Quản lý nợ không khóa được mã số thuế cho người nộp thuế.
- **Sót MST phôi cũ trên Văn bản từ chối nộp thừa**: Do văn bản này nằm ở cuối bộ và dùng cấu trúc ô lưới từng số, script thông thường hay bỏ sót dẫn đến sai thông tin pháp lý.
- **Bên nhận trùng bên giao ở Biên bản giao nhận**: Trưởng đoàn vừa là bên giao vừa ký bên nhận hồ sơ.
- **Căn cứ ban hành sau ngày quyết định**: Viện dẫn số và ngày Thông báo ngừng hoạt động rơi vào ngày sau ngày ký Quyết định kiểm tra.
- **Áp sai mức phạt vi phạm hành chính**: Trường hợp sau khi loại trừ chi phí mà doanh nghiệp vẫn còn lỗ hoặc thuộc diện miễn thuế (thuế phải nộp = 0), máy móc phạt 20% là sai luật; phải phạt mức tiền cố định (6.500.000 đ) theo quy định tại NĐ 310/2025/NĐ-CP.
- **Nhầm lẫn giữa QĐ xử phạt VPHC và QĐ buộc khắc phục hậu quả**: Doanh nghiệp có hành vi vi phạm ở các kỳ kiểm tra thuộc giai đoạn từ năm thứ 6 đến hết năm thứ 10 (ví dụ kiểm tra năm 2026 đối với sai sót giai đoạn 2016-2020) đã hết thời hiệu xử phạt tiền; nếu áp dụng Mẫu 02 xử phạt tiền hoặc đưa điều khoản phạt 20%, phạt chậm nộp tiền phạt 0,05%/ngày vào quyết định là trái Điều 8 Nghị định 125/2020/NĐ-CP và Điều 65 Luật Xử lý VPHC. Bắt buộc phải áp dụng Mẫu số 15 (Quyết định buộc khắc phục hậu quả).
- **Lệch lịch kiểm tra chạm ngày nghỉ**: Lên lịch kiểm tra 3 ngày nhưng rơi vào thứ Năm, thứ Sáu và Chủ nhật.
- **Header Mẫu 15 dồn 1 run:** Quốc hiệu rớt `NAM`, cơ quan rớt `ĐỒNG NA` - phải tách đoạn, cỡ 12, canh giữa ô.
- **Báo cáo kết quả lặp năm / thiếu số QĐ:** phôi hay ra `ngày 26 tháng 10 năm 2026 2026` và `SỐ          /QĐ-TCS3`. Điền từ sổ (`So_QDKT`, `Ngay_bao_cao`).
- **Tên kế toán phôi sót ở bảng ký BBXN/BBKT:** `Lê Hoài Phương` nằm trong ô gộp Giám đốc/Kế toán, replace paragraph không đụng tới. Phải quét `cell.text`.
- **Nội dung kiểm tra trên Báo cáo kết quả:** mặc định `Kiểm tra chấp hành pháp luật thuế`. Chi tiết phân công chỉ ở BBXN và Biên bản phân công.
- **Nhầm cam kết lương:** `CAM KẾT KHÔNG PHÁT SINH LAO ĐỘNG CHI TRẢ LƯƠNG` chỉ dùng khi DN không có lao động. User xin "cam kết chi phí lương" thì lấy đoạn tiền lương trong `MAU CAM KET DOANH THU VA CHI PHI.docx`.
- **Công văn xin quyết toán nhân đôi** ở cuối `TONG HOP` (thừa 1 bản + mảnh từ chối nộp thừa). Giữ 1 công văn, xóa phần thừa sau chữ ký.
- **QĐ Mẫu 15 nhảy trang giữa tiêu ngữ và tên quyết định:** `pageBreakBefore` nằm trên chữ `QUYẾT ĐỊNH` trong khi bảng quốc hiệu/tiêu ngữ thiếu hoặc đứng trước chỗ ngắt trang. Ngắt trang phải ở bảng header; tên quyết định không được `pageBreakBefore`. Giữ `keepWithNext` trên các dòng header.
- **Văn bản đề nghị thẩm định thiếu ngày / nhảy trang:** header còn `ngày ..... tháng ..... năm`; bảng hồ sơ đính kèm còn `..../..../2026` và số QĐKT phôi. Điền ngày báo cáo trên header, ngày từng tài liệu theo sổ; chỉ giữ 1 ngắt trang ở bảng header, bỏ `pageBreakBefore` trên dòng tiêu đề.
- **Công văn DN nhảy trang giữa tiêu ngữ và chữ CÔNG VĂN:** xin quyết toán (`/CV`) và cam kết (`/ĐKCP`) để `pageBreakBefore` trên dòng `CÔNG VĂN` sau bảng quốc hiệu. Bỏ `pageBreakBefore` trên tiêu đề; giữ 1 ngắt trang ở đoạn trống trước bảng header; `keepWithNext` trên header + `CÔNG VĂN` + `V/v`.

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
- [2026-11-22] Thiết lập nguyên tắc rẽ nhánh ban hành QĐ xử phạt VPHC (1-5 năm: phạt tiền 20% + truy thu + chậm nộp/giảm lỗ) vs QĐ buộc thực hiện biện pháp khắc phục hậu quả (6-10 năm: không phạt tiền, chỉ truy thu + chậm nộp/giảm lỗ; quá 10 năm không phạt, không truy thu) theo Điều 8 NĐ 125/2020/NĐ-CP; chuẩn hóa phôi gốc `templates/MAU QD KHAC PHUC HAU QUA GIAI THE.docx` (chuẩn Mẫu 15/NĐ 118).
- [2026-11-23] Vụ Kiến Thành Tân: BBXN bám phân công sổ (TD giá vốn, TV bán hàng + QLDN); Mẫu 15 tách Quốc hiệu/tiêu ngữ canh giữa NĐ 30; Báo cáo KQ điền số QĐ + ngày báo cáo, bỏ năm lặp; quét tên kế toán trong bảng ký; tờ trình/thẩm định cách ký 4 dòng; mỗi văn bản sang trang mới.
- [2026-11-23] Hồ sơ về DN: giữ 1 Công văn xin quyết toán; 1 công văn Cam kết chi phí và chi phí lương (`/ĐKCP`) từ phôi Tâm Hoàng Phát, điền pháp nhân vụ; không dùng mẫu không phát sinh lao động.
- [2026-11-24] Vụ Kiến Thành Tân: QĐ Mẫu 15 mất bảng quốc hiệu và `pageBreakBefore` trên chữ QUYẾT ĐỊNH; văn bản đề nghị thẩm định thiếu ngày header/bảng đính kèm và bị đôi ngắt trang. Đã chèn header QĐ, chuyển ngắt trang lên bảng, điền ngày 27/11/2026 và số QĐKT 7486.
- [2026-11-24] Hai công văn DN (xin quyết toán `/CV`, cam kết `/ĐKCP`) vẫn `pageBreakBefore` trên chữ `CÔNG VĂN` nên tiêu ngữ tách trang. Đã bỏ ngắt trang trên tiêu đề, dính header với `CÔNG VĂN` + `V/v`.
- [2026-11-25] Áp dụng thêm các quy tắc chung trong `ho-so-kiem-tra-ke-hoach`: Mail Merge (ngày thật dd/mm/yyyy, tiền #.##0, không switch \@ \#), ngày/năm quyết định (dự thảo bỏ trống, còn lại điền đủ), phân công theo sổ, cảnh báo phân công QLDN/bán hàng, và checklist đối soát của agent giám sát.
