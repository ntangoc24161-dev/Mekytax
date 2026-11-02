---
name: Theo dõi cung cấp hồ sơ Công an
description: Bóc tách công văn điều tra, tra cứu NNT, kiểm tra cơ quan quản lý rồi phân tổ KT, cập nhật Bao_cao_cong_van.xlsx
group: Tài chính
updated: '2026-10-28'
learned_updated: '2026-10-28'
---
## Khi nào dùng
- Khi tiếp nhận file PDF công văn yêu cầu cung cấp hồ sơ từ cơ quan Công an trong thư mục `CV MOI` hoặc vừa tải từ iOffice.
- Cần bóc tách thông tin, tra cứu địa bàn/người nộp thuế và ghi nhận vào sổ theo dõi.
- Cần xuất danh sách tổng hợp chi tiết hồ sơ tài liệu hoặc rà soát đồng bộ thư mục đã hoàn tất có bản ký số (`-R`).

## Chuẩn bị
- Thứ tự đường vào Drive (thử lần lượt, dừng khi đọc được):
  1. **Composio Google Drive, tài khoản `cungcapthongtin.tcs3tdn@gmail.com`** (đang ACTIVE, là mặc định cho hồ sơ Công an). Đọc, tải, tạo thư mục, chuyển file và ghi đè file (kể cả .xlsx/.docx) đều làm được qua đường này. Có nhiều tài khoản Drive cùng đấu, phải chọn đúng tài khoản này.
  2. `/mnt/i/HS CONG AN-2025-2026/HS CUNG CAP CA 2026/` (VPS sau rclone; `My Drive` = `/mnt/i`).
  3. `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\` (Google Drive Desktop trên Windows).
- File theo dõi đích: `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`) trong thư mục trên.
- File danh bạ NNT: `DB NNT.xls` (hoặc `.xlsx`, `DB NNT 03.06.2026.xls`) - sheet `NNT` tra cứu pháp lý NNT; sheet `KTR1` tra cứu CBQL cho Tổ KT1. Không đọc được bản Drive thì dùng bản sao vault `sources/DB NNT 03.06.2026.xlsx`.
- Thư mục công văn mới: `1. CUNG CAP CONG AN NAM 2026/CV MOI`.
- Thư mục lưu trữ công văn theo năm: `1. CUNG CAP CONG AN NAM 2026`.
- Thư mục lưu trữ file Markdown: `D:\HSO CONG AN FILE MD` (đĩa máy Windows; VPS không có). Không có thì lưu vault `ho-so-cv-moi-YYYY-MM/`.
- Không tự gắn rclone từ khung chat. Hướng dẫn rclone trên VPS: `ho-so-cv-moi-2026-10/huong-dan-rclone-google-drive-vps.md`.

## Quy trình
0. **Phân loại trước khi nhập sổ:**
   - `YC-CSKT` / `CSKT` gửi thẳng Thuế cơ sở, có MST địa bàn: nhập sổ.
   - `DON-KTr1` chỉ đạo Thuế TP cung cấp hồ sơ: nhập số hiệu; thiếu bảng kê MST thì ghi chú thiếu, tìm bản gốc Công an, không bịa MST.
   - Thuế TP báo **đã cung cấp** thay (ví dụ 42073/YC-CSKT đã trả qua 20543/DON-KTr1): không nhập sổ cung cấp, vẫn lưu folder riêng.
   - `DON-NVDTPC` (phòng ngừa tội phạm): không nhập sổ cung cấp hồ sơ.
   - `DON-KTr1` xác minh đối tượng truy nã: không nhập sổ theo mẫu cung cấp hồ sơ DN.
1. **Quét và bóc tách PDF:**
   - Liệt kê file trong `CV MOI`, tải về vault `ho-so-cv-moi-YYYY-MM/pdf/`.
   - **Đối chiếu sổ trước:** CV nào đã có trên `NHAP_HO_SO` thì không nhập lại, chỉ chuyển PDF về đúng folder đã có.
   - Nếu 1 file PDF chứa nhiều công văn đính kèm: tách riêng từng công văn để theo dõi độc lập.
   - Tên file có thể không khớp số CV (ví dụ `5684.pdf` là CV 20568/DON-VP): đọc nội dung, không tin tên file.
   - PDF có lớp chữ: trích xuất thành `.md`. PDF scan không lớp chữ: xuất ảnh từng trang rồi đọc ảnh; không bỏ hồ sơ, không bịa MST. OCR (skill `pdf-sang-md`) chỉ dùng khi venv OCR thật sự có trên máy.
   - **Kiểm chéo bản đọc ảnh scan (bắt buộc):** sau khi đọc ảnh, ghi lại nguyên văn 4 trường: cơ quan ban hành (dòng đầu header), số/ngày CV, tên DN, MST. Tra MST trong `DB NNT`: nếu MST không có nhưng DB có DN **cùng tên hoặc cùng địa chỉ** với MST khác, coi là nghi đọc sai số; xuất lại trang đó ở độ phân giải cao hơn và đọc lại từng chữ số trước khi ghi sổ. Tên cơ quan ban hành phải khớp với ký hiệu và nơi gửi trên CV. Còn nghi ngờ thì ghi sổ kèm ghi chú 'cần xác minh' và báo người dùng, không khẳng định.
2. **Tra cứu CSDL và cổng thông tin, kiểm tra cơ quan quản lý, phân Tổ Kiểm tra (kỹ năng `tra-cuu-nnt-tu-dong`):**
   - **Ưu tiên 1 (CSDL nội bộ):** Tra `MST` trong `DB NNT` (sheet `NNT`) lấy tên, địa chỉ, Phường/Xã, tình trạng ĐKKD. Có MST trên sheet `NNT` = thuộc TCS3.
   - **Ưu tiên 2 (Cổng thông tin):** CSDL nội bộ không có MST thì tra `https://masothue.com/` lấy pháp lý NNT **và trường Quản lý bởi**. Không suy đoán.
   - **Cổng trước khi phân tổ:** NNT do Thuế thành phố Đồng Nai hoặc CQT khác quản lý thì ghi rõ cơ quan quản lý, **không gán** KT1/KT2/KT3, **không gán** CBQL, **không** lập dự thảo TCS3. MST không có trong DB NNT và chưa xác nhận CQT thì không phân tổ theo phường, ghi chú chờ xác minh.
   - **Phân Tổ Kiểm tra (chỉ NNT đã xác nhận thuộc TCS3):** Rút gọn địa chỉ về cấp Phường/Xã và đối chiếu:
     - **Tổ KT1**: Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước (gồm địa bàn cũ: Tân Biên, Tân Hòa, Hố Nai 3, An Hòa, Long Bình Tân...).
     - **Tổ KT2**: Tam Hiệp, Trảng Dài, Trị An, Tân An, Tân Triều, Phú Lý (gồm địa bàn cũ: Bình Đa, Tam Hòa, Tân Hiệp, Tân Mai, Tân Tiến, Thiện Tân, Vĩnh An, Hiếu Liêm, Mã Đà, Vĩnh Tân, Tân Phong, Bình Hòa, Bình Lợi, Tân Bình, Thạnh Phú...).
     - **Tổ KT3**: Biên Hòa, Trấn Biên (gồm địa bàn cũ: Bửu Hòa, Hóa An, Tân Hạnh, Tân Vạn, An Bình, Bửu Long, Hiệp Hòa, Hòa Bình, Quang Vinh, Quyết Thắng, Thanh Bình, Thống Nhất, Trung Dũng...).
3. **Xác định cán bộ xử lý Tổ KT1:**
   - Chỉ khi công ty đã xác nhận thuộc TCS3 và đã gán Tổ KT1: tra MST trong sheet `KTR1`. Khớp đúng dòng thì lấy tên CBQL.
   - MST không có trên `KTR1`: để trống cột cán bộ, không suy tên.
   - Ngày phân công/ngày giao CBXL là ngày thực hiện lệnh.
4. **Lưu trữ hồ sơ:**
   - Tìm số folder lớn nhất hiện có trong thư mục lưu trữ năm 2026 trên Drive. Số folder khác STT trên sổ (ví dụ STT 201 nằm folder 272), không lấy lẫn.
   - Tạo thư mục theo cấu trúc: `STT. CV SO CV-CQBAN HANH-TO KT XU LY`.
   - Chuyển PDF khỏi `CV MOI` vào thư mục tương ứng (cả CV mới lẫn CV đã có trên sổ). Xong thì `CV MOI` phải trống.
5. **Cập nhật dữ liệu:**
   - Đọc được Drive (Composio, `/mnt/i` hoặc `I:`): tải `Bao_cao_cong_van.xlsx` bản live.
   - **Tìm dòng cuối thật:** dò ngược cột Số folder/Số CV tìm dòng cuối có dữ liệu (dòng của folder lớn nhất). KHÔNG dùng `max_row`/append: sheet có vùng hàng trống đã định dạng tới khoảng hàng 990.
   - Ghi CV và danh sách DN ngay dòng kế tiếp. Cột A (STT) là công thức: chép công thức từ dòng trên xuống, không ghi số cứng.
   - Đẩy ghi đè lên đúng file trên Drive, tải lại bản live và kiểm tra dòng mới nằm sát ngay dưới folder trước đó, đọc thời điểm sửa để xác nhận. NNT ngoài TCS3 nếu cần ghi sổ thì ghi cơ quan quản lý, để trống Tổ KT và CBQL.
   - Luôn lưu kèm bản nháp `ho-so-cv-moi-YYYY-MM/nhap-ho-so-draft-YYYY-MM-DD.xlsx` và file tóm tắt bóc tách trong vault.
   - Không đọc/ghi được Drive bằng mọi đường: chỉ lập bản nháp vault, **báo chưa ghi sổ chính**. Không nói đã cập nhật xong.
   - Người dùng báo nội dung CV khác bản đã ghi: bản người dùng đang xem là căn cứ. Đọc lại ảnh/PDF họ gửi, sửa dòng sổ và đổi tên folder cho khớp (cơ quan ban hành, tên DN), rồi báo rõ đã sửa gì.
6. **Tự động lập dự thảo hồ sơ trả lời (nếu có phân công Tổ KT1 thuộc TCS3):**
   - Chỉ khi đã ghi được sổ hoặc người dùng yêu cầu soạn ngay, và chỉ với DN TCS3 (KT1/KT2/KT3). Người dùng nói "không dự thảo" thì bỏ bước này. Kích hoạt kỹ năng `tra-loi-cong-van`:
     - Công văn trả lời: `TRA LOI CV <Số CV>_<CQ>_du thao.docx`
     - Tờ trình lãnh đạo: `to trinh <Số CV>.docx`
   - Lưu vào thư mục hồ sơ công văn tương ứng trên Drive.
7. **Xuất danh sách tổng hợp chi tiết tài liệu (khi có yêu cầu):** kỹ năng `xuat-danh-sach-ho-so-ca`.
8. **Rà soát và đồng bộ thư mục đã hoàn tất (khi có yêu cầu):** kỹ năng `dong-bo-thu-muc-ca`.

## Bẫy
- **Đọc nhầm ảnh scan ra công văn khác hẳn:** ngày 28/10/2026 file `12853CSHS 14347.pdf` (CA Cần Thơ, Công ty TNHH ĐTPT Sunset) bị đọc thành 'CA Hưng Yên, Giày Dép Hưng Hồng, MST 3603998255' và ghi sổ, người dùng phải gửi ảnh để chỉ ra. Lần đọc lại vẫn sai MST (3603999937 thay vì 3604037657). Dấu hiệu bỏ qua: MST không có trong DB NNT trong khi DB có DN cùng tên/cùng địa chỉ. Luôn kiểm chéo theo bước 1.
- Khi người dùng chất vấn nội dung, đừng khẳng định lại bản đọc cũ là đúng file; mở lại ảnh trang ở độ phân giải cao và đối chiếu với bản người dùng gửi.
- Kết luận "không đọc được Drive" chỉ vì thiếu `/mnt/i` và `I:`: sai. Ngày 26-28/10/2026 Drive cungcapthongtin đọc/ghi được đầy đủ qua Composio dù `/mnt/i` không tồn tại. Thử Composio trước.
- Chọn nhầm tài khoản Drive (có `ntangoc2416.1@gmail.com` chỉ đọc và tài khoản khác cùng đấu): hồ sơ Công an luôn dùng `cungcapthongtin.tcs3tdn@gmail.com`.
- Nhầm lẫn file theo dõi: dùng `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`), không ghi vào file cũ `File theo doi hs CUNG CAP CA-01.04.2025.xls` (sheet `Phan cong`).
- Nối dòng bằng `max_row`/append trên `NHAP_HO_SO`: dòng mới rơi xuống hàng ~990, cách dòng cuối thật gần 180 hàng trống, người dùng mở sổ không thấy. Ngày 28/10/2026 folder 272-273 bị ghi vào hàng 990-991 kèm STT số cứng, phải chuyển lên hàng 812-813.
- Ghi số cứng vào cột A (STT): phá công thức. Chép công thức từ dòng trên.
- Chỉ kiểm "có dòng mới trong file" mà không kiểm vị trí: phải xác nhận dòng mới nằm sát ngay dưới folder trước.
- Nhập lại CV đã có trên sổ: phải đối chiếu sổ trước khi nhập.
- File gộp nhiều công văn: không bóc tách từng công văn sẽ sót việc và sai STT.
- Tự suy đoán địa chỉ khi CSDL nội bộ không có: bắt buộc tra `masothue.com` hoặc để trống.
- Quên tạo dự thảo khi có Tổ KT1 của TCS3 và đã ghi sổ (trừ khi người dùng bảo không dự thảo).
- Thiếu Drive mọi đường: không ghi được sổ chính. Phải có bản nháp vault và nói rõ chưa nhập sổ.
- PDF scan không lớp chữ: không có OCR thì đọc ảnh trang, không bỏ hồ sơ.
- `DON-NVDTPC` và xác minh truy nã lọt lọc ANĐT/CSKT/KTr1: không nhập sổ cung cấp hồ sơ.
- `DON-KTr1` Thuế TP thường chỉ có công văn chỉ đạo, thiếu bảng kê MST: ghi chú thiếu, tải bản gốc CA, không bịa MST.
- Phân tổ theo Phường/Xã khi NNT do Thuế thành phố quản lý (ví dụ FDI/KCN Amata tại Long Bình): sai. Phải đọc cơ quan quản lý trước.
- MST không có trên sheet `KTR1` mà vẫn điền tên cán bộ xử lý: bịa. Để trống.

## Kiểm chứng
- Số dòng thêm vào `NHAP_HO_SO` khớp danh sách CV và DN; thời điểm sửa file trên Drive là lượt này.
- Mỗi dòng mới: cơ quan ban hành, số CV, tên DN, MST khớp ảnh trang 1 của PDF; MST có trong DB NNT hoặc có ghi chú lý do.
- Dòng mới nằm ngay dưới dòng folder trước đó (không có khoảng hàng trống chen giữa); cột A của dòng mới là công thức.
- `CV MOI` trống; mỗi PDF nằm đúng folder (mới hoặc đã có), tên folder khớp cơ quan ban hành thật.
- Nếu chỉ có bản nháp vault: báo chưa cập nhật sổ chính, dẫn đường dẫn bản nháp.
- Thư mục công văn có Tổ KT1 thuộc TCS3 đã có đủ 02 file `TRA LOI CV..._du thao.docx` và `to trinh...docx` (khi bước dự thảo đã chạy).
- MST ngoài TCS3 hoặc chưa xác nhận CQT không bị gán Tổ KT/CBQL và không có trong dự thảo.

## Lịch sử (tự học)
- [2026-10-12] Thêm fallback vault, đọc ảnh PDF scan, loại DON-NVDTPC/truy nã, DON-KTr1 thiếu bảng kê không bịa MST.
- [2026-10-14] CV 1625 / MST 3601051082 (SAITEX, Thuế TP quản lý) bị gán KT1 và CBQL. Thêm cổng đọc cơ quan quản lý trước khi phân tổ.
- [2026-10-15] Thêm thứ tự /mnt/i trước I:\, D: không theo rclone, cấm gắn Drive từ chat.
- [2026-10-28] Lượt xử lý CV MOI chạy được trọn vẹn qua Composio Drive cungcapthongtin (tải PDF, ghi sổ, tạo folder 272-274, chuyển file) trong khi skill chỉ ghi /mnt/i và I:. Thêm Composio làm đường ưu tiên, bước đối chiếu sổ trước khi nhập, chuyển cả CV đã có về folder cũ, CV Thuế TP báo đã cung cấp thì không nhập sổ.
- [2026-10-28] Ngày 28/10/2026 quy trình CV MOI chạy thành công hoàn toàn qua Composio Google Drive (cungcapthongtin.tcs3tdn@gmail.com) dù /mnt/i không tồn tại; skill cũ chỉ liệt kê /mnt/i và I:, dễ dẫn tới kết luận sai 'không có Drive, chỉ ghi nháp' (lượt 26/10 đã từng trả lời sai hướng /mnt/i). Bổ sung thêm các
- [2026-10-28] Lượt xử lý CV MOI chạy trọn vẹn qua Composio Drive cungcapthongtin dù /mnt/i không tồn tại. Thêm Composio làm đường ưu tiên, đối chiếu sổ trước khi nhập, chuyển cả CV đã có về folder cũ, CV Thuế TP báo đã cung cấp thì không nhập sổ.
- [2026-10-28] Dòng 272-273 bị append xuống hàng 990-991 với STT số cứng, người dùng mở sổ chỉ thấy tới folder 271. Thêm bước tìm dòng cuối thật, giữ công thức cột A, kiểm vị trí dòng sau khi ghi.
- [2026-10-28] Ngày 28/10/2026 làm theo bước 5 cũ (chỉ ghi 'ghi CV vào NHAP_HO_SO, đẩy ghi đè, đọc thời điểm sửa') đã append 2 dòng folder 272-273 xuống hàng 990-991 với STT số cứng; người dùng mở sổ vẫn chỉ thấy tới folder 271 và phải báo lại, Javis phải chuyển lên hàng 812-813 và khôi phục công thức cột A. Bổ su
- [2026-10-28] CV 12853/CSHS (CA Cần Thơ, ĐTPT Sunset, MST 3604037657) bị đọc ảnh scan thành CA Hưng Yên / Hưng Hồng và ghi sổ sai; lần đọc lại vẫn sai MST. Thêm bước kiểm chéo 4 trường với DB NNT, xử lý khi người dùng báo sai nội dung.
- [2026-10-28] Ngày 28/10/2026 làm theo skill cũ (chỉ ghi 'xuất ảnh từng trang rồi đọc ảnh') đã đọc nhầm CV 12853/CSHS thành CA Hưng Yên / Giày Dép Hưng Hồng / MST 3603998255 và ghi lên sổ; người dùng phải gửi ảnh CV thật (CA Cần Thơ, Sunset). Lần đọc lại vẫn ra MST 3603999937, người dùng xác nhận MST đúng 3604037
