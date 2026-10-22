---
name: Theo dõi cung cấp hồ sơ Công an
description: Bóc tách công văn điều tra, tách file gộp, tra cứu địa chỉ/masothue, phân tổ KT, cán bộ KT1 và cập nhật Bao_cao_cong_van.xlsx sheet NHAP_HO_SO.
group: Tài chính
updated: '2026-09-22'
learned_updated: '2026-09-22'
---
## Khi nào dùng
- Khi tiếp nhận file PDF công văn yêu cầu cung cấp hồ sơ từ cơ quan Công an trong thư mục `CV MOI`.
- Cần bóc tách thông tin, tra cứu địa bàn/người nộp thuế và ghi nhận vào sổ theo dõi.

## Chuẩn bị
- File theo dõi đích: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).
- File danh bạ NNT: `I:\My Drive\DB NNT 03.06.2026.xls` (hoặc `.xlsx`, sheet `KTR1` để tra cứu CBQL cho Tổ KT1).
- Thư mục công văn mới: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI`.
- Thư mục lưu trữ công văn theo năm: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026`.
- Thư mục lưu trữ file Markdown: `D:\HSO CONG AN FILE MD`.

## Quy trình
1. **Quét và bóc tách PDF:**
   - Kiểm tra các file trong thư mục `CV MOI`.
   - Nếu 1 file PDF chứa nhiều công văn đính kèm: tách riêng từng công văn để theo dõi độc lập như nhận riêng lẻ.
   - Trích xuất nội dung thành file `.md` lưu tại `D:\HSO CONG AN FILE MD` (tuyệt đối không lưu lẫn vào thư mục chứa PDF).
2. **Xác định địa bàn & phân Tổ Kiểm tra:**
   - Kiểm tra địa chỉ trụ sở ghi trên công văn hoặc phụ lục/bảng kê.
   - Nếu thiếu địa chỉ và không có lịch sử: tra cứu MST trên `https://masothue.com/` để lấy thông tin Phường/Xã, từ đó phân về Tổ KT (KT1, KT2, KT3).
   - Tuyệt đối không suy đoán hoặc bịa đặt; nếu không tra cứu được thì để trống để xin ý kiến xử lý.
3. **Xác định cán bộ xử lý Tổ KT1:**
   - Đối với công ty thuộc Tổ KT1, đối chiếu MST trong sheet `KTR1` của file `DB NNT 03.06.2026.xls` để lấy tên Cán bộ quản lý.
   - Ngày phân công/ngày giao CBXL là ngày thực hiện lệnh.
4. **Lưu trữ hồ sơ:**
   - Tìm STT lớn nhất hiện có trong thư mục lưu trữ năm 2026.
   - Tạo thư mục theo cấu trúc: `STT. CV SO CV-CQBAN HANH-TO KT XU LY`.
   - Di chuyển file PDF vào thư mục tương ứng.
5. **Cập nhật dữ liệu:**
   - Ghi thông tin công văn và danh sách doanh nghiệp vào file `Bao_cao_cong_van.xlsx` tại sheet `NHAP_HO_SO`.

## Bẫy
- Nhầm lẫn file theo dõi: người dùng đã đổi sang `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`), không ghi vào file cũ `File theo doi hs CUNG CAP CA-01.04.2025.xls` (sheet `Phan cong`).
- File gộp nhiều công văn: nếu không bóc tách từng công văn sẽ bị sót việc và sai STT theo dõi.
- Tự suy đoán địa chỉ khi công văn không ghi: bắt buộc tra cứu `masothue.com` hoặc để trống.

## Kiểm chứng
- Kiểm tra số dòng được thêm vào sheet `NHAP_HO_SO` khớp với danh sách công văn và doanh nghiệp.
- File `.md` đã xuất hiện đầy đủ trong `D:\HSO CONG AN FILE MD`.

## Lịch sử (tự học)
- [2026-09-17] Chuẩn hóa toàn bộ quy trình: bóc tách độc lập file PDF đa công văn; trích xuất địa chỉ từ phụ lục hoặc tra cứu masothue.com; tra cứu cán bộ quản lý Tổ KT1 qua sheet KTR1 file DB NNT; quy định ngày phân công là ngày thực hiện lệnh; lưu PDF vào folder riêng theo STT và tập trung toàn bộ file .md tại D:\HSO CONG AN FILE MD.
- [2026-09-22] Người dùng yêu cầu đổi file và sheet cập nhật từ `File theo doi hs CUNG CAP CA-01.04.2025.xls` (sheet `Phan cong`) sang `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`). Cần cập nhật skill để quy trình phản ánh đúng file và sheet đích mới.
