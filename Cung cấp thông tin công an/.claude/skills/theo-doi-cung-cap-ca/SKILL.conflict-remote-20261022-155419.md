---
name: Theo dõi cung cấp hồ sơ Công an
description: Bóc tách CV điều tra, tra cứu CQT quản lý, phân tổ KT (Vĩnh Cửu vào KT2; CQT khác ghi tên vào Tổ KT), gán CBQL KT1 và cập nhật Bao_cao_cong_van.xlsx.
group: Tài chính
updated: '2026-10-16'
learned_updated: '2026-10-16'
---
## Khi nào dùng
- Khi tiếp nhận file PDF công văn yêu cầu cung cấp hồ sơ từ cơ quan Công an trong thư mục `CV MOI`.
- Cần bóc tách thông tin, tra cứu CQT quản lý, địa bàn/người nộp thuế và ghi nhận vào sổ theo dõi `Bao_cao_cong_van.xlsx`.
- Cần xuất danh sách tổng hợp chi tiết hồ sơ tài liệu hoặc rà soát đồng bộ thư mục đã hoàn tất có bản ký số (`-R`).

## Chuẩn bị
- File theo dõi đích: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).
- File danh bạ NNT: `I:\My Drive\DB NNT.xls` (hoặc `.xlsx`, `DB NNT 03.06.2026.xls` - sheet `NNT` tra cứu CQT quản lý, pháp lý NNT; sheet `KTR1` tra cứu CBQL cho Tổ KT1).
- Thư mục công văn mới: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI`.
- Thư mục lưu trữ công văn theo năm: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026`.
- Thư mục lưu trữ file Markdown: `D:\HSO CONG AN FILE MD`.

## Quy trình
1. **Quét và bóc tách PDF:**
   - Kiểm tra các file trong thư mục `CV MOI`.
   - Nếu 1 file PDF chứa nhiều công văn đính kèm: tách riêng từng công văn để theo dõi độc lập như nhận riêng lẻ.
   - Trích xuất nội dung thành file `.md` lưu tại `D:\HSO CONG AN FILE MD` (tuyệt đối không lưu lẫn vào thư mục chứa PDF).
2. **Tra cứu CSDL & Cổng thông tin, kiểm tra CQT quản lý và phân Tổ Kiểm tra (kỹ năng `tra-cuu-nnt-tu-dong`):**
   - **Tra cứu CSDL & Cổng thông tin:** Tra cứu `MST` trong file `I:\My Drive\DB NNT.xls` (sheet `NNT`, cột `Tên cơ quan thuế`) hoặc fallback qua `https://masothue.com/` (trường Cơ quan thuế quản lý) để lấy tên doanh nghiệp, CQ quản lý, địa chỉ trụ sở, Phường/Xã, tình trạng ĐKKD, số/ngày thông báo đóng MST.
   - **Kiểm tra Cơ quan thuế quản lý:**
     + Thuộc TCS3 khi tên CQ khớp `TP.Biên Hòa-TCS3 Đồng Nai`, `Thuế cơ sở 3 thành phố Đồng Nai`, `TCS3`, hoặc `Chi cục Thuế khu vực Biên Hòa - Vĩnh Cửu`.
     + Nếu **KHÔNG PHẢI** TCS3: ghi nguyên văn tên CQ vào ô "Tổ KT" (Cột L sheet `NHAP_HO_SO`). Không phân tổ và không gán CBQL.
     + Nếu **DO** Thuế cơ sở 3 TP Đồng Nai (Biên Hòa - Vĩnh Cửu) quản lý: Tiến hành phân Tổ KT theo địa bàn:
       * Địa bàn **Vĩnh Cửu**: Phân công cho **Tổ KT2**.
       * Địa bàn TP Biên Hòa: Rút gọn Phường/Xã và đối chiếu theo bảng phân công địa bàn:
         - **Tổ KT1**: Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước (gồm các địa bàn cũ: Tân Biên, Tân Hòa, Hố Nai 3, An Hòa, Long Bình Tân...).
         - **Tổ KT2**: Tam Hiệp, Trảng Dài, Trị An, Tân An, Tân Triều, Phú Lý (gồm các địa bàn cũ: Bình Đa, Tam Hòa, Tân Hiệp, Tân Mai, Tân Tiến, Thiện Tân, Vĩnh An, Hiếu Liêm, Mã Đà, Vĩnh Tân, Tân Phong, Bình Hòa, Bình Lợi, Tân Bình, Thạnh Phú...).
         - **Tổ KT3**: Biên Hòa, Trấn Biên (gồm các địa bàn cũ: Bửu Hòa, Hóa An, Tân Hạnh, Tân Vạn, An Bình, Bửu Long, Hiệp Hòa, Hòa Bình, Quang Vinh, Quyết Thắng, Thanh Bình, Thống Nhất, Trung Dũng...).
3. **Xác định cán bộ xử lý Tổ KT1:**
   - Đối với công ty thuộc Thuế cơ sở 3 và phân về Tổ KT1, tra cứu MST trong sheet `KTR1` của file `DB NNT.xls` để lấy tên Cán bộ quản lý trực tiếp.
   - Ngày phân công/ngày giao CBXL là ngày thực hiện lệnh.
   - Doanh nghiệp thuộc KT2, KT3 hoặc cơ quan thuế khác quản lý thì cột CBQL để trống.
4. **Lưu trữ hồ sơ:**
   - Tìm STT lớn nhất hiện có trong thư mục lưu trữ năm 2026.
   - Tạo thư mục theo cấu trúc: `STT. CV SO CV-CQBAN HANH-TO KT XU LY`.
   - Di chuyển file PDF vào thư mục tương ứng.
5. **Cập nhật dữ liệu:**
   - Ghi thông tin công văn và danh sách doanh nghiệp vào file `Bao_cao_cong_van.xlsx` tại sheet `NHAP_HO_SO`.
   - Tự động tạo bản sao lưu tại `D:\HSO CONG AN FILE MD\Bao_cao_cong_van_backup.xlsx`.
6. **Tự động lập dự thảo hồ sơ trả lời (nếu có phân công Tổ KT1):**
   - Nếu công văn có doanh nghiệp được phân công cho Tổ KT1, tự động kích hoạt kỹ năng `tra-loi-cong-van` để soạn thảo 02 file chuẩn thể thức: Công văn trả lời và Tờ trình lãnh đạo, lưu trực tiếp vào thư mục công văn tương ứng.
7. **Xuất danh sách tổng hợp chi tiết tài liệu & Đồng bộ thư mục hoàn tất:**
   - Kích hoạt kỹ năng `xuat-danh-sach-ho-so-ca` hoặc `dong-bo-thu-muc-ca` khi có yêu cầu.

## Bẫy
- **Bỏ qua bước tra cứu cơ quan thuế quản lý:** Phải xác định CQT quản lý trước; nếu không thuộc Thuế cơ sở 3 TP Đồng Nai thì ghi tên CQT vào ô Tổ KT, tuyệt đối không phân tổ KT1/KT2/KT3.
- Quên quy tắc địa bàn Vĩnh Cửu: Mọi đơn vị thuộc Vĩnh Cửu do Thuế cơ sở 3 quản lý đều phân công cho Tổ KT2.
- Nhầm lẫn file theo dõi: Đảm bảo ghi đúng vào `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).
- File gộp nhiều công văn: nếu không bóc tách từng công văn sẽ bị sót việc và sai STT theo dõi.

## Kiểm chứng
- Kiểm tra cột Tổ KT: NNT ngoài địa bàn quản lý hiển thị rõ tên CQT quản lý; NNT thuộc Thuế cơ sở 3 hiển thị đúng KT1, KT2, KT3 (Vĩnh Cửu là KT2).
- Cán bộ quản lý chỉ gán cho doanh nghiệp thuộc KT1; các trường hợp khác để trống.
- File `.md` xuất hiện đầy đủ trong `D:\HSO CONG AN FILE MD`.

## Lịch sử (tự học)
- [2026-09-17] Chuẩn hóa toàn bộ quy trình: bóc tách độc lập file PDF đa công văn; trích xuất địa chỉ từ phụ lục hoặc tra cứu masothue.com; tra cứu cán bộ quản lý Tổ KT1 qua sheet KTR1 file DB NNT; quy định ngày phân công là ngày thực hiện lệnh; lưu PDF vào folder riêng theo STT và tập trung toàn bộ file .md tại D:\HSO CONG AN FILE MD.
- [2026-09-22] Người dùng yêu cầu đổi file và sheet cập nhật từ `File theo doi hs CUNG CAP CA-01.04.2025.xls` (sheet `Phan cong`) sang `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`). Cần cập nhật skill để quy trình phản ánh đúng file và sheet đích mới.
- [2026-09-23] Tích hợp bảng phân công địa bàn mới và cũ của các Tổ KT1, KT2, KT3 vào quy trình xác định địa bàn để tự động phân Tổ kiểm tra phụ trách.
- [2026-09-25] Bổ sung quy định tự động tạo dự thảo hồ sơ trả lời công văn (gồm công văn phúc đáp và tờ trình lãnh đạo kèm theo) ngay khi cập nhật công văn công an có phân công doanh nghiệp cho Tổ KT1 phụ trách, lưu trực tiếp vào thư mục công văn tương ứng.
- [2026-09-26] Bổ sung quy trình xuất danh sách tổng hợp chi tiết hồ sơ tài liệu ra Excel tại thư mục CV MOI (kỹ năng `xuat-danh-sach-ho-so-ca`) và quy trình rà soát đồng bộ hậu tố `-R` cho toàn bộ thư mục công văn đã có bản ký số scan (kỹ năng `dong-bo-thu-muc-ca`).
- [2026-09-26] Tích hợp kỹ năng `tra-cuu-nnt-tu-dong`: chuẩn hóa cơ chế tra cứu 2 tầng ưu tiên CSDL nội bộ DB NNT.xls (sheet NNT và KTR1), fallback cổng thông tin khi thiếu dữ liệu để phân Tổ KT và gán CBQL KT1.
- [2026-10-16] Cửa kiểm CQ thuế quản lý trước khi phân tổ: chỉ chia KT1-KT3 và gán CBQL khi thuộc TCS3 (tên DB `TP.Biên Hòa-TCS3 Đồng Nai`); Vĩnh Cửu -> KT2; CQ khác ghi nguyên văn tên CQ vào ô Tổ KT.
