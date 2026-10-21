---
name: Tra cứu NNT và phân Tổ Kiểm tra tự động
description: Tra cứu MST qua DB NNT và cổng thông tin, đọc cơ quan quản lý, phân Tổ KT1-KT3; MST ngoài TCS3 thì không gán CBQL.
group: Tài chính
updated: '2026-10-15'
learned_updated: '2026-10-15'
---
## Khi nào dùng
- Khi tiếp nhận công văn điều tra của cơ quan Công an có danh sách MST hoặc Doanh nghiệp cần tra cứu thông tin pháp lý, tình trạng hoạt động và phân Tổ Kiểm tra phụ trách.
- Khi bóc tách công văn mới trong thư mục `CV MOI` hoặc nhận qua Zalo.
- Khi cần hoàn thiện dữ liệu thông tin NNT (địa chỉ, người đại diện theo pháp luật, CCCD, tình trạng đóng/mở MST, số tài khoản...) để cập nhật sổ theo dõi `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`) hoặc lập hồ sơ dự thảo Word trả lời Công an.

## Chuẩn bị
- **Nguồn CSDL nội bộ chính:**
  + File cơ sở dữ liệu NNT, thử lần lượt: `/mnt/i/DB NNT.xls` (VPS sau rclone), `I:\My Drive\DB NNT.xls` (Windows), hoặc `DB NNT.xlsx` / `DB NNT 03.06.2026.xls` cùng gốc đó. Nếu không có cả `/mnt/i` lẫn `I:`, dùng bản sao vault `sources/DB NNT 03.06.2026.xlsx`.
  + Sheet `NNT`: CSDL tổng thể về NNT (MST, Tên NNT, Địa chỉ trụ sở, Phường/Xã, Tình trạng hoạt động, Trạng thái đóng MST, Số và Ngày Thông báo đóng MST/bỏ địa chỉ, Người đại diện PL, CCCD, Kế toán trưởng, Cán bộ quản lý...). Có MST trên sheet này = thuộc phạm vi Thuế cơ sở 3.
  + Sheet `KTR1`: Bảng phân công chi tiết Cán bộ quản lý (CBQL) cho các doanh nghiệp thuộc Tổ Kiểm tra 1 (Tổ QLDN 1). Chỉ lấy tên CBQL khi MST khớp đúng dòng trên sheet này.
- **Nguồn Cổng thông tin trực tuyến (Fallback / Bù đắp dữ liệu):**
  + Cổng tra cứu thông tin doanh nghiệp và MST: `https://masothue.com/`.
  + Cổng thông tin Đăng ký doanh nghiệp quốc gia: `https://dangkykinhdoanh.gov.vn/`.
  + Cổng thông tin Người nộp thuế ngành Thuế: `https://tracuunnt.gdt.gov.vn/`.
  + Khi tra masothue.com, bắt buộc lấy cả trường **Quản lý bởi** (cơ quan thuế quản lý), không chỉ Phường/Xã.
- **Bảng phân công địa bàn quản lý kiểm tra doanh nghiệp (chỉ dùng sau khi đã xác nhận NNT thuộc TCS3):**
  + **Tổ KT1 (Tổ QL Doanh nghiệp 1):** Địa bàn mới gồm Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước (gồm các địa bàn cũ: Hố Nai, Long Bình, Tân Biên, Tân Hòa, Hố Nai 3, An Hòa, Long Bình Tân, Long Hưng, Phước Tân, Tam Phước).
  + **Tổ KT2 (Tổ QL Doanh nghiệp 1):** Địa bàn mới gồm Tam Hiệp, Trảng Dài, Trị An, Tân An, Tân Triều, Phú Lý (gồm các địa bàn cũ: Bình Đa, Tam Hiệp, Tam Hòa, Tân Hiệp, Tân Mai, Tân Tiến, Trảng Dài, Thiện Tân, TT Vĩnh An, Hiếu Liêm, Mã Đà, Trị An, Tân An, Vĩnh Tân, Tân Phong, Bình Hòa, Bình Lợi, Tân Bình, Thạnh Phú, Phú Lý).
  + **Tổ KT3 (Tổ QL Doanh nghiệp 2):** Địa bàn mới gồm Biên Hòa, Trấn Biên (gồm các địa bàn cũ: Bửu Hòa, Hóa An, Tân Hạnh, Tân Vạn, An Bình, Bửu Long, Hiệp Hòa, Hòa Bình, Quang Vinh, Quyết Thắng, Thanh Bình, Thống Nhất, Trung Dũng).

## Quy trình
1. **Bước 1: Tra cứu CSDL nội bộ `DB NNT.xls` (Ưu tiên số 1):**
   - Đọc và tìm kiếm theo `MST` trên sheet `NNT`.
   - Trích xuất các trường thông tin: Tên NNT, Địa chỉ trụ sở, Phường/Xã, Tình trạng hoạt động (Đang hoạt động / Ngừng hoạt động / Tạm ngừng KD), Người ĐDPL, CCCD/CMND, Kế toán, Số và ngày Thông báo không hoạt động tại địa chỉ đăng ký (từ cột Lý do đóng MST tổ chức).
   - Có MST trên sheet `NNT` thì đánh dấu thuộc TCS3 và sang Bước 3.
2. **Bước 2: Bù đắp dữ liệu qua Cổng thông tin (khi CSDL thiếu hoặc không có MST):**
   - Nếu `MST` không có trong file `DB NNT.xls` hoặc bị thiếu thông tin địa chỉ/pháp lý/cơ quan quản lý:
   - Tra cứu trực tuyến trên `https://masothue.com/` (hoặc cổng đăng ký kinh doanh).
   - Trích xuất: Tên doanh nghiệp chuẩn hóa, Địa chỉ chi tiết (Phường/Xã), Người đại diện pháp luật, Trạng thái hoạt động, **Quản lý bởi**.
   - *Nguyên tắc trung thực dữ liệu:* Tuyệt đối không suy đoán hay bịa đặt; nếu cả CSDL và cổng thông tin đều không có thì để trống và ghi chú cần xác minh.
3. **Bước 3: Xác định cơ quan thuế quản lý (bắt buộc trước khi phân Tổ KT):**
   - Nguồn: có trong sheet `NNT` của DB TCS3 → thuộc TCS3. Không có thì lấy trường Quản lý bởi trên masothue.com.
   - Nếu Quản lý bởi là **Thuế thành phố Đồng Nai** hoặc thuế cơ sở/chi cục khác (không phải Thuế cơ sở 3): ghi rõ tên cơ quan quản lý, đánh dấu **ngoài phạm vi TCS3**, **dừng** không gán KT1/KT2/KT3, không gán CBQL, không đưa vào dự thảo cung cấp hồ sơ.
   - MST không có trong DB NNT **không** được coi là thiếu dữ liệu để rồi phân tổ theo phường; phải đọc cơ quan quản lý trước.
4. **Bước 4: Chuẩn hóa địa chỉ và Phân Tổ Kiểm tra (chỉ với NNT đã xác nhận thuộc TCS3):**
   - Rút gọn địa chỉ về cấp Phường/Xã (Cột K trên sheet `NHAP_HO_SO`).
   - Đối chiếu Phường/Xã với Bảng phân công địa bàn quản lý để xác định Tổ KT:
     + Khớp địa bàn KT1 -> Gán `KT1` (Cột L).
     + Khớp địa bàn KT2 -> Gán `KT2` (Cột L).
     + Khớp địa bàn KT3 -> Gán `KT3` (Cột L).
     + Địa bàn ngoài tỉnh/thành phố hoặc hộ kinh doanh -> Ghi rõ ngoài địa bàn / chuyển Tổ HKD.
5. **Bước 5: Xác định Cán bộ quản lý và Ngày phân công (chỉ Tổ KT1 đã xác nhận thuộc TCS3):**
   - Tra cứu `MST` trên sheet `KTR1`. Khớp đúng dòng thì lấy tên CBQL (Cột M trên sheet `NHAP_HO_SO`) và gán Ngày phân công / Ngày giao CBXL (Cột O, P) là **ngày thực hiện lệnh**.
   - MST không có trên `KTR1`: để trống cột cán bộ xử lý, ghi chú chưa có CBQL trên KTR1. **Không suy tên, không lấy tên cán bộ của MST khác.**
   - Tổ KT2, KT3: cột cán bộ xử lý để trống theo phân cấp quản lý nội bộ.

## Bẫy
- Nhầm lẫn giữa mã số thuế doanh nghiệp (10 số) và mã số thuế chi nhánh (13 số): cần tra cứu đúng mã số để lấy địa chỉ thực tế của đơn vị trực thuộc.
- Bỏ qua dữ liệu Thông báo đóng MST: Đối với doanh nghiệp ngừng hoạt động, bắt buộc lấy chính xác Số và Ngày ban hành Thông báo không hoạt động tại địa chỉ đã đăng ký từ cột Lý do đóng MST tổ chức trong sheet `NNT`.
- Sai lệch địa bàn do thay đổi địa giới hành chính: Luôn đối chiếu danh sách đối ứng giữa tên phường mới và tên phường cũ trước khi phân Tổ KT.
- Bịa đặt thông tin khi thiếu dữ liệu: Nghiêm cấm tự điền địa chỉ hoặc phân tổ khi không có căn cứ từ CSDL hoặc cổng thông tin.
- Phân tổ chỉ theo Phường/Xã khi NNT do Thuế thành phố hoặc CQT khác quản lý (điển hình FDI / KCN Amata tại Phường Long Bình vẫn là địa bàn KT1 về mặt phường). Phường khớp không đủ để gán Tổ KT.
- MST không có trong DB NNT mà vẫn gán KT theo phường rồi lấy CBQL: sai. Không có trên `KTR1` thì để trống, không điền tên cán bộ.
- Trên VPS, đừng dừng ở hết ổ `I:` khi chưa thử `/mnt/i/DB NNT.xls` (rclone). Lần 15/10/2026 skill chỉ đường `I:\` nên không đọc được CSDL live.

## Kiểm chứng
- 100% MST đã ghi cơ quan quản lý hoặc ghi chú ngoài phạm vi TCS3.
- MST ngoài TCS3: Tổ KT trống, CBQL trống, có tên CQT quản lý, không nằm trong danh sách dự thảo.
- MST thuộc TCS3 có địa chỉ rút gọn Phường/Xã và Tổ KT.
- CBQL Tổ KT1 chỉ xuất hiện khi MST khớp sheet `KTR1`; ngày phân công là ngày thực hiện lệnh.
- Dữ liệu chuẩn hóa sẵn sàng nhập vào sheet `NHAP_HO_SO` của `Bao_cao_cong_van.xlsx`.

## Lịch sử (tự học)
- [2026-09-26] Khởi tạo kỹ năng tra cứu NNT tự động: chuẩn hóa quy trình tra cứu ưu tiên CSDL nội bộ DB NNT.xls (sheet NNT và KTR1), tự động fallback tra cứu cổng thông tin (masothue.com) khi thiếu dữ liệu, phân Tổ KT1-KT2-KT3 theo địa bàn và gán CBQL Tổ KT1.
- [2026-10-14] Làm theo skill cũ: MST 3601051082 (SAITEX, Phường Long Bình, Thuế thành phố quản lý, không có trong DB NNT/KTR1) vẫn bị gán KT1 và bịa CBQL. Thêm bước đọc cơ quan quản lý trước khi phân tổ; cấm gán CBQL khi không khớp KTR1.
- [2026-10-14] Hội thoại 14/10/2026: user hỏi MST 3601051082 tại CV 1625 do Thuế thành phố quản lý sao lại gán KT1 và Nguyễn Văn Thông. Làm theo skill cũ chỉ nhìn Phường Long Bình nên gán KT1; MST không có trên KTR1 mà vẫn điền tên CBQL. Thiếu bước đọc cơ quan quản lý và cấm bịa CBQL.
- [2026-10-15] 15/10/2026 skill chỉ trỏ I:\My Drive\DB NNT.xls nên trên VPS Docker không thấy CSDL dù Drive là nguồn thật. Thêm /mnt/i/DB NNT.xls vào thứ tự đọc, giữ cổng CQT và cấm gán CBQL ngoài KTR1.
