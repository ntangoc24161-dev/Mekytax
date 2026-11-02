---
type: decision
created: '2026-09-26'
updated: '2026-09-26'
---
Quy chuẩn kỹ năng tra cứu NNT và phân Tổ Kiểm tra tự động:
- **Tên kỹ năng:** `tra-cuu-nnt-tu-dong` (Tra cứu NNT và phân Tổ Kiểm tra tự động).
- **Mục tiêu:** Tự động hóa toàn diện quy trình tra cứu mã số thuế (MST), bóc tách thông tin pháp lý doanh nghiệp, chuẩn hóa địa chỉ Phường/Xã, phân bổ đúng Tổ Kiểm tra (KT1, KT2, KT3) và gán Cán bộ quản lý trực tiếp cho Tổ KT1.
- **Cơ chế tra cứu 2 tầng chuẩn mực:**
  1. **Tầng 1 (Ưu tiên số 1 - CSDL nội bộ):** Tra cứu trực tiếp từ file `I:\My Drive\DB NNT.xls` (hoặc `DB NNT.xlsx`, `DB NNT 03.06.2026.xls`):
     - Sheet `NNT`: Lấy Tên NNT, Địa chỉ trụ sở, Phường/Xã, Tình trạng hoạt động, Số/Ngày Thông báo không hoạt động tại địa chỉ đăng ký (cột Lý do đóng MST tổ chức), Người ĐDPL, CCCD/CMND, Kế toán trưởng, Cán bộ quản lý.
     - Sheet `KTR1`: Tra cứu tên Cán bộ quản lý (CBQL) phân công trực tiếp cho các doanh nghiệp thuộc địa bàn Tổ KT1.
  2. **Tầng 2 (Bù đắp / Fallback qua Cổng thông tin trực tuyến):**
     - Khi MST không có trong file `DB NNT.xls` hoặc bị thiếu địa chỉ, tình trạng, người ĐDPL: Tự động tra cứu trực tuyến qua `https://masothue.com/`, `https://dangkykinhdoanh.gov.vn/`, `https://tracuunnt.gdt.gov.vn/` để lấy thông tin Phường/Xã và pháp lý NNT.
     - *Nguyên tắc trung thực dữ liệu:* Tuyệt đối không suy đoán hay bịa đặt; nếu cả hai nguồn đều không có thì để trống địa chỉ và ghi chú cần xác minh.
- **Quy tắc phân Tổ Kiểm tra:** Rút gọn địa chỉ về cấp Phường/Xã và đối chiếu Bảng phân công địa bàn (`facts/bang-phan-cong-dia-ban-to-kt.md`):
  - **Tổ KT1:** Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước (và địa bàn cũ: Tân Biên, Tân Hòa, Hố Nai 3, An Hòa, Long Bình Tân...).
  - **Tổ KT2:** Tam Hiệp, Trảng Dài, Trị An, Tân An, Tân Triều, Phú Lý (và địa bàn cũ: Bình Đa, Tam Hiệp, Tam Hòa, Tân Hiệp, Tân Mai, Tân Tiến, Trảng Dài, Thiện Tân, TT Vĩnh An, Hiếu Liêm, Mã Đà, Trị An, Tân An, Vĩnh Tân, Tân Phong, Bình Hòa, Bình Lợi, Tân Bình, Thạnh Phú, Phú Lý...).
  - **Tổ KT3:** Biên Hòa, Trấn Biên (và địa bàn cũ: Bửu Hòa, Hóa An, Tân Hạnh, Tân Vạn, An Bình, Bửu Long, Hiệp Hòa, Hòa Bình, Quang Vinh, Quyết Thắng, Thanh Bình, Thống Nhất, Trung Dũng...).
- **Xác định Cán bộ quản lý Tổ KT1:**
  - Đối với doanh nghiệp thuộc Tổ KT1: Tra cứu sheet `KTR1` của `DB NNT.xls` để điền Cán bộ quản lý (Cột M).
  - Ngày phân công / Ngày giao CBXL (Cột O, P trên sheet `NHAP_HO_SO`): Là ngày thực hiện lệnh (ngày hiện tại).
- **Vì sao:** Đảm bảo dữ liệu đầu vào cho sổ theo dõi `Bao_cao_cong_van.xlsx` và các tệp Word dự thảo (`TRA LOI CV..._du thao.docx`, `to trinh...docx`) luôn đầy đủ, chính xác, không bị gián đoạn khi CSDL nội bộ chưa kịp cập nhật doanh nghiệp mới.
- **Áp dụng:** Tự động kích hoạt trong mọi lượt tiếp nhận, bóc tách công văn mới từ `CV MOI` hoặc Zalo, và khi thực hiện lệnh xử lý toàn diện công văn công an.
