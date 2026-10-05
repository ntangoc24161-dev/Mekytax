---
name: Tra cứu NNT và phân Tổ Kiểm tra tự động
description: Tra cứu MST qua DB NNT.xls và cổng thông tin, chuẩn hóa địa chỉ, phân Tổ KT1-KT2-KT3 và xác định cán bộ quản lý KT1 tự động.
group: Tài chính
updated: '2026-09-26'
learned_updated: '2026-09-26'
---
## Khi nào dùng
- Khi tiếp nhận công văn điều tra của cơ quan Công an có danh sách MST hoặc Doanh nghiệp cần tra cứu thông tin pháp lý, tình trạng hoạt động và phân Tổ Kiểm tra phụ trách.
- Khi bóc tách công văn mới trong thư mục `CV MOI` hoặc nhận qua Zalo.
- Khi cần hoàn thiện dữ liệu thông tin NNT (địa chỉ, người đại diện theo pháp luật, CCCD, tình trạng đóng/mở MST, số tài khoản...) để cập nhật sổ theo dõi `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`) hoặc lập hồ sơ dự thảo Word trả lời Công an.

## Chuẩn bị
- **Nguồn CSDL nội bộ chính:**
  + File cơ sở dữ liệu NNT: `I:\My Drive\DB NNT.xls` (hoặc `DB NNT.xlsx`, `DB NNT 03.06.2026.xls`).
  + Sheet `NNT`: CSDL tổng thể về NNT (MST, Tên NNT, Địa chỉ trụ sở, Phường/Xã, Tình trạng hoạt động, Trạng thái đóng MST, Số & Ngày Thông báo đóng MST/bỏ địa chỉ, Người đại diện PL, CCCD, Kế toán trưởng, Cán bộ quản lý...).
  + Sheet `KTR1`: Bảng phân công chi tiết Cán bộ quản lý (CBQL) cho các doanh nghiệp thuộc Tổ Kiểm tra 1 (Tổ QLDN 1).
- **Nguồn Cổng thông tin trực tuyến (Fallback / Bù đắp dữ liệu):**
  + Cổng tra cứu thông tin doanh nghiệp & MST: `https://masothue.com/`.
  + Cổng thông tin Đăng ký doanh nghiệp quốc gia: `https://dangkykinhdoanh.gov.vn/`.
  + Cổng thông tin Người nộp thuế ngành Thuế: `https://tracuunnt.gdt.gov.vn/`.
- **Bảng phân công địa bàn quản lý kiểm tra doanh nghiệp:**
  + **Tổ KT1 (Tổ QL Doanh nghiệp 1):** Địa bàn mới gồm Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước (gồm các địa bàn cũ: Hố Nai, Long Bình, Tân Biên, Tân Hòa, Hố Nai 3, An Hòa, Long Bình Tân, Long Hưng, Phước Tân, Tam Phước).
  + **Tổ KT2 (Tổ QL Doanh nghiệp 1):** Địa bàn mới gồm Tam Hiệp, Trảng Dài, Trị An, Tân An, Tân Triều, Phú Lý (gồm các địa bàn cũ: Bình Đa, Tam Hiệp, Tam Hòa, Tân Hiệp, Tân Mai, Tân Tiến, Trảng Dài, Thiện Tân, TT Vĩnh An, Hiếu Liêm, Mã Đà, Trị An, Tân An, Vĩnh Tân, Tân Phong, Bình Hòa, Bình Lợi, Tân Bình, Thạnh Phú, Phú Lý).
  + **Tổ KT3 (Tổ QL Doanh nghiệp 2):** Địa bàn mới gồm Biên Hòa, Trấn Biên (gồm các địa bàn cũ: Bửu Hòa, Hóa An, Tân Hạnh, Tân Vạn, An Bình, Bửu Long, Hiệp Hòa, Hòa Bình, Quang Vinh, Quyết Thắng, Thanh Bình, Thống Nhất, Trung Dũng).

## Quy trình
1. **Bước 1: Tra cứu CSDL nội bộ `DB NNT.xls` (Ưu tiên số 1):**
   - Đọc và tìm kiếm theo `MST` trên sheet `NNT`.
   - Trích xuất các trường thông tin: Tên NNT, Địa chỉ trụ sở, Phường/Xã, Tình trạng hoạt động (Đang hoạt động / Ngừng hoạt động / Tạm ngừng KD), Người ĐDPL, CCCD/CMND, Kế toán, Số và ngày Thông báo không hoạt động tại địa chỉ đăng ký (từ cột Lý do đóng MST tổ chức).
2. **Bước 2: Bù đắp dữ liệu qua Cổng thông tin (Fallback khi CSDL thiếu dữ liệu):**
   - Nếu `MST` không có trong file `DB NNT.xls` hoặc bị thiếu thông tin địa chỉ/pháp lý:
   - Tự động tra cứu trực tuyến trên cổng thông tin `https://masothue.com/` (hoặc cổng đăng ký kinh doanh).
   - Trích xuất: Tên doanh nghiệp chuẩn hóa, Địa chỉ chi tiết (xác định Phường/Xã), Người đại diện pháp luật, Trạng thái hoạt động.
   - *Nguyên tắc trung thực dữ liệu:* Tuyệt đối không suy đoán hay bịa đặt thông tin; nếu cả CSDL và cổng thông tin đều không có thì để trống địa chỉ và ghi chú cần xác minh bổ sung.
3. **Bước 3: Chuẩn hóa địa chỉ và Phân Tổ Kiểm tra (KT1, KT2, KT3):**
   - Rút gọn địa chỉ về cấp Phường/Xã (Cột K trên sheet `NHAP_HO_SO`).
   - Đối chiếu Phường/Xã với Bảng phân công địa bàn quản lý để xác định Tổ KT:
     + Khớp địa bàn KT1 -> Gán `KT1` (Cột L).
     + Khớp địa bàn KT2 -> Gán `KT2` (Cột L).
     + Khớp địa bàn KT3 -> Gán `KT3` (Cột L).
     + Trường hợp địa bàn ngoài tỉnh/thành phố hoặc hộ kinh doanh -> Ghi rõ ngoài địa bàn / chuyển Tổ HKD.
4. **Bước 4: Xác định Cán bộ quản lý và Ngày phân công (Áp dụng cho Tổ KT1):**
   - Nếu doanh nghiệp thuộc **Tổ KT1**:
     + Tra cứu `MST` trên sheet `KTR1` của file `DB NNT.xls` để lấy tên Cán bộ quản lý trực tiếp (Cột M trên sheet `NHAP_HO_SO`).
     + Gán Ngày phân công / Ngày giao CBXL (Cột O, P) là **ngày thực hiện lệnh** (ngày hiện tại).
   - Nếu thuộc Tổ KT2, KT3: Cột cán bộ xử lý để trống theo phân cấp quản lý nội bộ.

## Bẫy
- Nhầm lẫn giữa mã số thuế doanh nghiệp (10 số) và mã số thuế chi nhánh (13 số): cần tra cứu đúng mã số để lấy địa chỉ thực tế của đơn vị trực thuộc.
- Bỏ qua dữ liệu Thông báo đóng MST: Đối với doanh nghiệp ngừng hoạt động, bắt buộc lấy chính xác Số và Ngày ban hành Thông báo không hoạt động tại địa chỉ đã đăng ký từ cột Lý do đóng MST tổ chức trong sheet `NNT`.
- Sai lệch địa bàn do thay đổi địa giới hành chính: Luôn đối chiếu danh sách đối ứng giữa tên phường mới và tên phường cũ trước khi phân Tổ KT.
- Bịa đặt thông tin khi thiếu dữ liệu: Nghiêm cấm tự điền địa chỉ hoặc phân tổ khi không có căn cứ từ CSDL hoặc cổng thông tin.

## Kiểm chứng
- 100% doanh nghiệp có MST được tra cứu đầy đủ tên, địa chỉ rút gọn Phường/Xã, Tổ KT.
- Doanh nghiệp thuộc Tổ KT1 có đầy đủ tên Cán bộ quản lý từ sheet `KTR1` và ngày phân công là ngày thực hiện lệnh.
- Dữ liệu chuẩn hóa sẵn sàng nhập vào sheet `NHAP_HO_SO` của `Bao_cao_cong_van.xlsx` và cung cấp tham số cho tệp dự thảo Word.

## Lịch sử (tự học)
- [2026-09-26] Khởi tạo kỹ năng tra cứu NNT tự động: chuẩn hóa quy trình tra cứu ưu tiên CSDL nội bộ DB NNT.xls (sheet NNT và KTR1), tự động fallback tra cứu cổng thông tin (masothue.com) khi thiếu dữ liệu, phân Tổ KT1-KT2-KT3 theo địa bàn và gán CBQL Tổ KT1.
