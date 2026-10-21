---
name: Theo dõi cung cấp hồ sơ Công an
description: Bóc tách công văn điều tra, tra cứu NNT, kiểm tra cơ quan quản lý rồi phân tổ KT, cập nhật Bao_cao_cong_van.xlsx
group: Tài chính
updated: '2026-10-15'
learned_updated: '2026-10-15'
---
## Khi nào dùng
- Khi tiếp nhận file PDF công văn yêu cầu cung cấp hồ sơ từ cơ quan Công an trong thư mục `CV MOI` hoặc vừa tải từ iOffice.
- Cần bóc tách thông tin, tra cứu địa bàn/người nộp thuế và ghi nhận vào sổ theo dõi.
- Cần xuất danh sách tổng hợp chi tiết hồ sơ tài liệu hoặc rà soát đồng bộ thư mục đã hoàn tất có bản ký số (`-R`).

## Chuẩn bị
- Thứ tự đường dẫn Drive (thử lần lượt, dừng khi thấy file):
  1. `/mnt/i/HS CONG AN-2025-2026/HS CUNG CAP CA 2026/` (VPS sau rclone; `My Drive` = `/mnt/i`).
  2. `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\` (Google Drive Desktop trên Windows).
- File theo dõi đích: `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`) trong thư mục trên.
- File danh bạ NNT: `DB NNT.xls` (hoặc `.xlsx`, `DB NNT 03.06.2026.xls`) tại `/mnt/i/` hoặc `I:\My Drive\` - sheet `NNT` tra cứu pháp lý NNT; sheet `KTR1` tra cứu CBQL cho Tổ KT1. Nếu không có cả `/mnt/i` lẫn `I:`, dùng bản sao vault `sources/DB NNT 03.06.2026.xlsx`.
- Thư mục công văn mới: `1. CUNG CAP CONG AN NAM 2026/CV MOI` trên `/mnt/i` hoặc `I:\My Drive`.
- Thư mục lưu trữ công văn theo năm: `1. CUNG CAP CONG AN NAM 2026` trên cùng gốc Drive.
- Thư mục lưu trữ file Markdown: `D:\HSO CONG AN FILE MD` (đĩa máy Windows; rclone không mang sang VPS).
- Kiểm tra `/mnt/i` rồi `I:` rồi `D:` trước khi ghi sổ. Không gắn được thì chuẩn bị thư mục vault `ho-so-cv-moi-YYYY-MM/`.
- Không tự gắn Google Drive từ khung chat. Hướng dẫn rclone trên VPS: `ho-so-cv-moi-2026-10/huong-dan-rclone-google-drive-vps.md`. Gắn xong, người dùng nhắn `kiểm tra /mnt/i`.

## Quy trình
0. **Phân loại trước khi nhập sổ:**
   - `YC-CSKT` / `CSKT` gửi thẳng Thuế cơ sở, có MST địa bàn: nhập sổ.
   - `DON-KTr1` chỉ đạo Thuế TP cung cấp hồ sơ: nhập số hiệu; thiếu bảng kê MST thì ghi chú thiếu, tìm bản gốc Công an, không bịa MST.
   - `DON-NVDTPC` (phòng ngừa tội phạm): không nhập sổ cung cấp hồ sơ.
   - `DON-KTr1` xác minh đối tượng truy nã: không nhập sổ theo mẫu cung cấp hồ sơ DN.
1. **Quét và bóc tách PDF:**
   - Kiểm tra các file trong thư mục `CV MOI` (hoặc vault `ho-so-cv-moi-YYYY-MM/pdf/` nếu vừa tải khi thiếu `/mnt/i` và `I:`).
   - Nếu 1 file PDF chứa nhiều công văn đính kèm: tách riêng từng công văn để theo dõi độc lập như nhận riêng lẻ.
   - PDF có lớp chữ: trích xuất thành file `.md`. PDF scan không lớp chữ: xuất ảnh từng trang rồi đọc ảnh; không bỏ qua hồ sơ, không bịa MST. OCR (skill `pdf-sang-md`) chỉ dùng khi venv OCR thật sự có trên máy.
   - File `.md` lưu tại `D:\HSO CONG AN FILE MD` khi có ổ `D:`; không thì lưu `ho-so-cv-moi-YYYY-MM/md/`.
2. **Tra cứu CSDL và cổng thông tin, kiểm tra cơ quan quản lý, phân Tổ Kiểm tra (kỹ năng `tra-cuu-nnt-tu-dong`):**
   - **Ưu tiên 1 (CSDL nội bộ):** Tra cứu `MST` trong `DB NNT.xls` / `/mnt/i/DB NNT.xls` / bản sao vault (sheet `NNT`) để lấy tên doanh nghiệp, địa chỉ trụ sở, Phường/Xã, tình trạng ĐKKD, số/ngày thông báo đóng MST. Có MST trên sheet `NNT` = thuộc TCS3.
   - **Ưu tiên 2 (Cổng thông tin trực tuyến):** Trường hợp CSDL nội bộ không có MST hoặc thiếu dữ liệu, tra cứu `https://masothue.com/` lấy pháp lý NNT **và trường Quản lý bởi**. Tuyệt đối không suy đoán hay bịa đặt.
   - **Cổng trước khi phân tổ:** NNT do Thuế thành phố Đồng Nai hoặc CQT khác quản lý thì ghi rõ cơ quan quản lý, **không gán** KT1/KT2/KT3, **không gán** CBQL, **không** lập dự thảo TCS3. MST không có trong DB NNT không được phân tổ theo phường.
   - **Phân Tổ Kiểm tra (chỉ NNT đã xác nhận thuộc TCS3):** Rút gọn địa chỉ về cấp Phường/Xã và đối chiếu theo bảng phân công địa bàn:
     - **Tổ KT1**: Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước (gồm các địa bàn cũ: Tân Biên, Tân Hòa, Hố Nai 3, An Hòa, Long Bình Tân...).
     - **Tổ KT2**: Tam Hiệp, Trảng Dài, Trị An, Tân An, Tân Triều, Phú Lý (gồm các địa bàn cũ: Bình Đa, Tam Hòa, Tân Hiệp, Tân Mai, Tân Tiến, Thiện Tân, Vĩnh An, Hiếu Liêm, Mã Đà, Vĩnh Tân, Tân Phong, Bình Hòa, Bình Lợi, Tân Bình, Thạnh Phú...).
     - **Tổ KT3**: Biên Hòa, Trấn Biên (gồm các địa bàn cũ: Bửu Hòa, Hóa An, Tân Hạnh, Tân Vạn, An Bình, Bửu Long, Hiệp Hòa, Hòa Bình, Quang Vinh, Quyết Thắng, Thanh Bình, Thống Nhất, Trung Dũng...).
3. **Xác định cán bộ xử lý Tổ KT1:**
   - Chỉ khi công ty đã xác nhận thuộc TCS3 và đã gán Tổ KT1: tra cứu MST trong sheet `KTR1`. Khớp đúng dòng thì lấy tên CBQL.
   - MST không có trên `KTR1`: để trống cột cán bộ, không suy tên.
   - Ngày phân công/ngày giao CBXL là ngày thực hiện lệnh.
4. **Lưu trữ hồ sơ:**
   - Tìm STT lớn nhất hiện có trong thư mục lưu trữ năm 2026 trên Drive (`/mnt/i` hoặc `I:`). Không đọc được Drive thì chưa gán STT/folder, ghi chú trên bản nháp.
   - Tạo thư mục theo cấu trúc: `STT. CV SO CV-CQBAN HANH-TO KT XU LY`.
   - Di chuyển file PDF vào thư mục tương ứng.
5. **Cập nhật dữ liệu:**
   - Nếu có `/mnt/i` hoặc ổ `I:`: ghi thông tin công văn và danh sách doanh nghiệp vào `Bao_cao_cong_van.xlsx` sheet `NHAP_HO_SO`; sao lưu `D:\HSO CONG AN FILE MD\Bao_cao_cong_van_backup.xlsx` khi có ổ `D:`. NNT ngoài TCS3 nếu cần ghi sổ thì ghi cơ quan quản lý, để trống Tổ KT và CBQL.
   - Nếu không có `/mnt/i` và không có ổ `I:`: lập bản nháp `ho-so-cv-moi-YYYY-MM/nhap-ho-so-draft.xlsx` kèm file tóm tắt; **báo chưa ghi sổ chính**. Không nói đã cập nhật xong.
6. **Tự động lập dự thảo hồ sơ trả lời (nếu có phân công Tổ KT1 thuộc TCS3):**
   - Chỉ khi đã ghi được sổ hoặc người dùng yêu cầu soạn ngay, và chỉ với doanh nghiệp TCS3 (KT1/KT2/KT3). Không soạn dự thảo cho NNT do Thuế thành phố hoặc CQT khác quản lý. Kích hoạt kỹ năng `tra-loi-cong-van`:
     - Công văn trả lời: `TRA LOI CV <Số CV>_<CQ>_du thao.docx`
     - Tờ trình lãnh đạo: `to trinh <Số CV>.docx`
   - Lưu vào thư mục hồ sơ công văn tương ứng trên Drive khi có `/mnt/i` hoặc `I:`.
7. **Xuất danh sách tổng hợp chi tiết tài liệu (khi có yêu cầu):**
   - Kỹ năng `xuat-danh-sach-ho-so-ca`.
8. **Rà soát và đồng bộ thư mục đã hoàn tất (khi có yêu cầu):**
   - Kỹ năng `dong-bo-thu-muc-ca`.

## Bẫy
- Nhầm lẫn file theo dõi: người dùng đã đổi sang `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`), không ghi vào file cũ `File theo doi hs CUNG CAP CA-01.04.2025.xls` (sheet `Phan cong`).
- File gộp nhiều công văn: nếu không bóc tách từng công văn sẽ bị sót việc và sai STT theo dõi.
- Tự suy đoán địa chỉ khi CSDL nội bộ không có: bắt buộc tra cứu cổng thông tin `masothue.com` hoặc để trống.
- Quên tạo dự thảo hồ sơ trả lời khi có Tổ KT1: khi phát hiện có doanh nghiệp thuộc Tổ KT1 **của TCS3** và đã ghi sổ, phải tạo ngay cả file công văn phúc đáp và tờ trình lãnh đạo.
- Thiếu Drive: không ghi được sổ chính. Lần 12/10/2026 đã bóc 11 CV rồi kẹt bước này. Phải có bản nháp vault và nói rõ chưa nhập sổ.
- Lần 15/10/2026: người dùng yêu cầu rà dòng 761-810; skill chỉ đường `I:\` nên không mở được sổ dù nguồn thật là Google Drive. Trên VPS phải thử `/mnt/i` trước, không kết luận hết sổ khi chưa kiểm tra đường Linux.
- PDF scan không lớp chữ: không có OCR thì đọc ảnh trang, không bỏ hồ sơ.
- `DON-NVDTPC` và xác minh truy nã lọt lọc ANĐT/CSKT/KTr1: không nhập sổ cung cấp hồ sơ.
- `DON-KTr1` Thuế TP thường chỉ có công văn chỉ đạo, thiếu bảng kê MST: ghi chú thiếu, tải bản gốc CA, không bịa MST.
- Phân tổ theo Phường/Xã khi NNT do Thuế thành phố quản lý (ví dụ FDI/KCN Amata tại Long Bình): sai. Phải đọc cơ quan quản lý trước.
- MST không có trên sheet `KTR1` mà vẫn điền tên cán bộ xử lý: bịa. Để trống.

## Kiểm chứng
- Kiểm tra số dòng được thêm vào sheet `NHAP_HO_SO` khớp với danh sách công văn và doanh nghiệp **chỉ khi đã ghi được file trên `/mnt/i` hoặc ổ `I:`**.
- Nếu chỉ có bản nháp vault: báo chưa cập nhật sổ chính, dẫn đường dẫn `nhap-ho-so-draft.xlsx`.
- File `.md` đã xuất hiện đầy đủ trong `D:\HSO CONG AN FILE MD` hoặc thư mục vault tương ứng.
- Thư mục công văn có Tổ KT1 thuộc TCS3 đã có đủ 02 file `TRA LOI CV..._du thao.docx` và `to trinh...docx` (khi bước dự thảo đã chạy).
- MST ngoài TCS3 không bị gán Tổ KT/CBQL và không có trong dự thảo.
- Tệp Excel tổng hợp tại `CV MOI` và các thư mục có đuôi `-R` trên Drive phản ánh đúng trạng thái thực tế.

## Lịch sử (tự học)
- [2026-09-22] Người dùng yêu cầu đổi file và sheet cập nhật từ `File theo doi hs CUNG CAP CA-01.04.2025.xls` (sheet `Phan cong`) sang `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`). Cần cập nhật skill để quy trình phản ánh đúng file và sheet đích mới.
- [2026-09-23] Tích hợp bảng phân công địa bàn mới và cũ của các Tổ KT1, KT2, KT3 vào quy trình xác định địa bàn để tự động phân Tổ kiểm tra phụ trách.
- [2026-09-25] Bổ sung quy định tự động tạo dự thảo hồ sơ trả lời công văn (gồm công văn phúc đáp và tờ trình lãnh đạo kèm theo) ngay khi cập nhật công văn công an có phân công doanh nghiệp cho Tổ KT1 phụ trách, lưu trực tiếp vào thư mục công văn tương ứng.
- [2026-09-26] Bổ sung quy trình xuất danh sách tổng hợp chi tiết hồ sơ tài liệu ra Excel tại thư mục CV MOI (kỹ năng `xuat-danh-sach-ho-so-ca`) và quy trình rà soát đồng bộ hậu tố `-R` cho toàn bộ thư mục công văn đã có bản ký số scan (kỹ năng `dong-bo-thu-muc-ca`).
- [2026-09-26] Tích hợp kỹ năng `tra-cuu-nnt-tu-dong`: chuẩn hóa cơ chế tra cứu 2 tầng ưu tiên CSDL nội bộ DB NNT.xls (sheet NNT và KTR1), fallback cổng thông tin khi thiếu dữ liệu để phân Tổ KT và gán CBQL KT1.
- [2026-10-12] Lần rà 01-12/10/2026: máy không gắn I:/D: nên không ghi sổ chính; PDF scan phải đọc ảnh; loại DON-NVDTPC và xác minh truy nã; DON-KTr1 thiếu bảng kê MST thì không bịa; fallback vault `ho-so-cv-moi-YYYY-MM/` và `sources/DB NNT 03.06.2026.xlsx`.
- [2026-10-12] Lần rà 12/10/2026 làm theo skill cũ thiếu bước kiểm tra ổ I:/D: nên không ghi được sổ, thiếu fallback vault; PDF scan không lớp chữ không có bước đọc ảnh; lọc xong vẫn suýt nhập 18979/DON-NVDTPC và 19022 truy nã; DON-KTr1 thiếu bảng kê MST không có bẫy.
- [2026-10-14] CV 1625 / MST 3601051082 (SAITEX, Long Bình, Thuế thành phố quản lý) bị gán KT1 và CBQL dù không có trong DB NNT/KTR1. Thêm cổng đọc cơ quan quản lý trước khi phân tổ; cấm gán CBQL khi không khớp KTR1; không dự thảo NNT ngoài TCS3.
- [2026-10-14] Cùng sự cố 14/10/2026 với MST 3601051082: skill này tự lặp bước phân tổ theo phường và gán CBQL KT1, không đọc cơ quan quản lý, rồi còn kích hoạt dự thảo khi thấy KT1. Cần cùng cổng với tra-cuu-nnt-tu-dong.
- [2026-10-15] 15/10/2026 làm theo skill chỉ đường I:\ nên không mở được Bao_cao_cong_van.xlsx trên VPS Docker dù user xác nhận I: là Google Drive và chọn rclone. Thêm thứ tự /mnt/i trước I:\, ghi rõ D: không theo rclone, cấm gắn Drive từ chat.
