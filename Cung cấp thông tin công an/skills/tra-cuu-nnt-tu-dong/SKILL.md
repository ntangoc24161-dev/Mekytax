---
name: Tra cứu NNT và phân Tổ Kiểm tra tự động
description: "Tra cứu MST, CQ thuế quản lý; chỉ phân KT1-KT3/CBQL nếu thuộc TCS3 (Vĩnh Cửu=KT2); CQ khác ghi tên vào ô Tổ KT."
group: Tài chính
updated: '2026-10-16'
learned_updated: '2026-10-16'
---
## Khi nào dùng
- Khi tiếp nhận công văn điều tra của cơ quan Công an có danh sách MST hoặc Doanh nghiệp cần tra cứu thông tin pháp lý, tình trạng hoạt động và phân Tổ Kiểm tra phụ trách.
- Khi bóc tách công văn mới trong thư mục `CV MOI` hoặc nhận qua Zalo.
- Khi cần hoàn thiện dữ liệu thông tin NNT (địa chỉ, người đại diện theo pháp luật, CCCD, tình trạng đóng/mở MST, số tài khoản, cơ quan thuế quản lý) để cập nhật sổ theo dõi `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`) hoặc lập hồ sơ dự thảo Word trả lời Công an.

## Chuẩn bị
- **Nguồn CSDL nội bộ chính:**
  + File cơ sở dữ liệu NNT: `I:\My Drive\DB NNT.xls` (hoặc `DB NNT.xlsx`, `DB NNT 03.06.2026.xls`).
  + Sheet `NNT`: CSDL tổng thể về NNT. Cột bắt buộc khi phân tổ: `Cơ quan thuế`, `Tên cơ quan thuế`, `Tên Phường Xã`, `Tên Quận/Huyện`, `Tổ kiểm tra quản lý`, `CBQL`, MST, Tên NNT, địa chỉ, Trạng thái ĐKT tổ chức, Lý do đóng MST tổ chức.
  + Sheet `KTR1`: Bảng phân công chi tiết Cán bộ quản lý (CBQL) cho các doanh nghiệp thuộc Tổ Kiểm tra 1 (Tổ QLDN 1).
- **Nguồn Cổng thông tin trực tuyến (Fallback / Bù đắp dữ liệu):**
  + Cổng tra cứu thông tin doanh nghiệp & MST: `https://masothue.com/` (trường **Cơ quan thuế quản lý**).
  + Cổng thông tin Đăng ký doanh nghiệp quốc gia: `https://dangkykinhdoanh.gov.vn/`.
  + Cổng thông tin Người nộp thuế ngành Thuế: `https://tracuunnt.gdt.gov.vn/`.
- **Tên CQ thuộc Thuế cơ sở 3 thành phố Đồng Nai (Biên Hòa - Vĩnh Cửu):** khớp một trong các dạng (không phân biệt hoa thường):
  + `TP.Biên Hòa-TCS3 Đồng Nai` (đúng tên cột `Tên cơ quan thuế` trong DB NNT)
  + `Thuế cơ sở 3 thành phố Đồng Nai`, `TCS3`
  + `Chi cục Thuế khu vực Biên Hòa - Vĩnh Cửu` (tên cũ)
- **Bảng phân công địa bàn (chỉ dùng sau khi đã xác nhận CQ thuộc TCS3):**
  + **Tổ KT1 (Tổ QL Doanh nghiệp 1):** Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước (địa bàn cũ: Tân Biên, Tân Hòa, Hố Nai 3, An Hòa, Long Bình Tân...).
  + **Tổ KT2 (Tổ QL Doanh nghiệp 1):** Tam Hiệp, Trảng Dài, Trị An, Tân An, Tân Triều, Phú Lý (địa bàn cũ: Bình Đa, Tam Hòa, Tân Hiệp, Tân Mai, Tân Tiến, Thiện Tân, Vĩnh An, Hiếu Liêm, Mã Đà, Vĩnh Tân, Tân Phong, Bình Hòa, Bình Lợi, Tân Bình, Thạnh Phú...). **Toàn bộ địa bàn Vĩnh Cửu gán KT2.**
  + **Tổ KT3 (Tổ QL Doanh nghiệp 2):** Biên Hòa, Trấn Biên (địa bàn cũ: Bửu Hòa, Hóa An, Tân Hạnh, Tân Vạn, An Bình, Bửu Long, Hiệp Hòa, Hòa Bình, Quang Vinh, Quyết Thắng, Thanh Bình, Thống Nhất, Trung Dũng).

## Quy trình
1. **Bước 1: Tra cứu CSDL nội bộ `DB NNT.xls` (Ưu tiên số 1):**
   - Đọc và tìm kiếm theo `MST` trên sheet `NNT`.
   - Trích xuất: Tên NNT, địa chỉ trụ sở, `Tên Phường Xã`, `Tên Quận/Huyện`, **`Tên cơ quan thuế`**, `Tổ kiểm tra quản lý`, `CBQL`, tình trạng ĐKT tổ chức, Người ĐDPL, CCCD/CMND, Kế toán, Số và ngày Thông báo không hoạt động tại địa chỉ đăng ký (cột Lý do đóng MST tổ chức).
2. **Bước 2: Bù đắp dữ liệu qua Cổng thông tin (khi CSDL thiếu):**
   - Nếu MST không có trong `DB NNT.xls` hoặc thiếu địa chỉ / **thiếu cơ quan thuế quản lý**:
   - Tra `https://masothue.com/` (hoặc cổng ĐKKD / tra cứu NNT).
   - Trích xuất: tên DN, địa chỉ (Phường/Xã, Quận/Huyện), Người ĐDPL, trạng thái, **Cơ quan thuế quản lý**.
   - Không suy đoán. Cả hai nguồn đều trống thì để trống và ghi `[CHƯA XÁC ĐỊNH]`.
3. **Bước 3: Cửa kiểm cơ quan thuế quản lý (BẮT BUỘC trước khi phân Tổ KT):**
   - Lấy tên CQ từ cột `Tên cơ quan thuế` (DB) hoặc trường Cơ quan thuế quản lý (cổng).
   - **Không thuộc TCS3** (tên CQ không khớp danh sách TCS3 ở mục Chuẩn bị): ghi **nguyên văn tên cơ quan thuế quản lý** vào ô **Tổ KT** (Cột L sheet `NHAP_HO_SO`). Để trống CBQL (Cột M). **Không** gán KT1/KT2/KT3, **không** tra sheet `KTR1`. DN này không đưa vào gói bàn giao Soạn thảo.
   - **Thuộc TCS3:** mới sang Bước 4.
   - Thiếu tên CQ ở cả DB và cổng: ô Tổ KT ghi `[CHƯA XÁC ĐỊNH]`, không phân tổ.
4. **Bước 4: Chuẩn hóa địa chỉ và Phân Tổ Kiểm tra (chỉ DN thuộc TCS3):**
   - Rút gọn địa chỉ về cấp Phường/Xã (Cột K trên sheet `NHAP_HO_SO`).
   - **Nếu địa bàn Vĩnh Cửu** (`Tên Quận/Huyện` chứa `Vĩnh Cửu`, kể cả dạng `Huyện Vĩnh Cửu(Hết HL)`, hoặc phường/xã thuộc danh sách Vĩnh Cửu trong bảng KT2): gán `KT2` (Cột L). Không xét tiếp bảng KT1/KT3.
   - **Nếu địa bàn Biên Hòa / Trấn Biên / còn lại thuộc TCS3:** đối chiếu Phường/Xã với bảng phân công:
     + Khớp KT1 -> `KT1` (Cột L).
     + Khớp KT2 -> `KT2` (Cột L).
     + Khớp KT3 -> `KT3` (Cột L).
   - Cột `Tổ kiểm tra quản lý` trong DB chỉ dùng để đối chiếu khi đã xác nhận TCS3; nếu lệch với bước này thì theo cửa kiểm + bảng địa bàn, ghi chú lệch.
5. **Bước 5: Cán bộ quản lý và Ngày phân công (chỉ Tổ KT1 thuộc TCS3):**
   - Chỉ khi ô Tổ KT = `KT1`: tra MST trên sheet `KTR1` lấy CBQL (Cột M); Ngày phân công / Ngày giao CBXL (Cột O, P) = ngày thực hiện lệnh.
   - KT2, KT3: để trống cột cán bộ xử lý.
   - Ô Tổ KT đang là tên CQ khác: không gán CBQL.

## Bẫy
- Phân KT1/KT2/KT3 theo phường khi CQ quản lý không phải TCS3: sai. Phải đọc `Tên cơ quan thuế` trước.
- Ghi `ngoài địa bàn` chung chung vào ô Tổ KT khi đã biết tên CQ: phải ghi đúng tên cơ quan thuế quản lý.
- Nhầm MST chi nhánh 13 số với MST 10 số: tra đúng mã để lấy CQ và địa chỉ đơn vị trực thuộc.
- Bỏ qua Thông báo đóng MST: DN ngừng hoạt động phải lấy Số và Ngày TB từ cột Lý do đóng MST tổ chức.
- Sai địa giới hành chính: đối chiếu tên phường mới/cũ; Vĩnh Cửu luôn KT2 khi đã thuộc TCS3.
- Bịa địa chỉ, CQ hoặc tổ khi thiếu dữ liệu.

## Kiểm chứng
- 100% MST có tên CQ quản lý (từ DB hoặc cổng) hoặc nhãn `[CHƯA XÁC ĐỊNH]`.
- Ô Tổ KT chỉ mang `KT1`/`KT2`/`KT3` khi CQ là TCS3; các dòng khác là tên CQ nguyên văn.
- DN Vĩnh Cửu thuộc TCS3 đều là `KT2`.
- DN KT1 (TCS3) có CBQL từ sheet `KTR1` và ngày phân công = ngày thực hiện lệnh.
- Dữ liệu sẵn sàng nhập `NHAP_HO_SO`; chỉ DN `KT1`/`KT2`/`KT3` vào gói Soạn thảo.

## Lịch sử (tự học)
- [2026-10-16] Bổ sung cửa kiểm cơ quan thuế quản lý trước khi phân tổ: chỉ chia KT1-KT3 và gán CBQL khi CQ là Thuế cơ sở 3 TP Đồng Nai (Biên Hòa-Vĩnh Cửu); Vĩnh Cửu gán KT2; CQ khác ghi nguyên văn tên CQ vào ô Tổ KT. Nguồn: cột `Tên cơ quan thuế` trên DB NNT.xls.
- [2026-09-26] Khởi tạo kỹ năng tra cứu NNT tự động: chuẩn hóa quy trình tra cứu ưu tiên CSDL nội bộ DB NNT.xls (sheet NNT và KTR1), tự động fallback tra cứu cổng thông tin (masothue.com) khi thiếu dữ liệu, phân Tổ KT1-KT2-KT3 theo địa bàn và gán CBQL Tổ KT1.
