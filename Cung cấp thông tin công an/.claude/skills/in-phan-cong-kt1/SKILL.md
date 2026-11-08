---
name: in-phan-cong-kt1
description: In đồng loạt công văn Công an giao Tổ KT1 trạng thái Mới, mỗi cán bộ một PDF chẵn trang, có dòng phân công đầu trang, đổi sang Đang xử lý.
group: Năng suất
---

# In phân công KT1

## Khi nào dùng
- Người dùng yêu cầu: "in phân công KT1", "in phiếu phân công", "xuất file in KT1" (qua chat Javis hoặc Telegram).
- Cần in hàng loạt công văn yêu cầu cung cấp thông tin của Công an giao cho Tổ Kiểm tra 1 (KT1) đang ở trạng thái "Mới" trong sổ theo dõi.
- Chưa tạo macro trong Excel (tạm hoãn theo yêu cầu, file vẫn giữ định dạng chuẩn `.xlsx`).

## Chuẩn bị
- File sổ theo dõi: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).
  - Cột B (index 1): Số thư mục lưu trữ (folder STT).
  - Cột C (index 2): Số công văn.
  - Cột D (index 3): Ngày công văn.
  - Cột J (index 9): Mã số thuế (MST).
  - Cột K (index 10): Tên doanh nghiệp / Người nộp thuế.
  - Cột M (index 12): Tổ Kiểm tra phụ trách (lọc `KT1`).
  - Cột N (index 13): Cán bộ xử lý.
  - Cột O (index 14): Tình trạng hồ sơ (lọc `Mới`, cập nhật sang `Đang xử lý`).
  - Cột Q (index 16): Ngày phân công xử lý.
- Thư mục công văn gốc năm 2026: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026`.
- Script thực thi: `scripts/in_phan_cong_kt1.py`.
- Thư mục lưu trữ kết quả:
  - Google Drive: `I:\My Drive\KIEM SOAT HSO CUNG CAP\IN PHAN CONG\` (tài khoản Drive `cungcapthongtin.tcs3tdn`).
  - Cục bộ: `D:\SECOND BRAIN\javis-os\brains\Cung cấp thông tin công an\exports\in-phan-cong-kt1\`.

## Quy trình
1. **Quét dữ liệu sổ theo dõi:**
   - Mở sheet `NHAP_HO_SO` trong file `Bao_cao_cong_van.xlsx`.
   - Lọc các dòng có Tổ KT = `KT1` và Tình trạng = `Mới`.
   - Kiểm tra thông tin: Nếu thiếu Tổ KT hoặc thiếu tên Cán bộ xử lý (như các dòng "N cty đính kèm", hồ sơ chưa rõ người phụ trách), **dừng lại và lập danh sách báo người dùng bổ sung**, tuyệt đối không tự ý in hoặc suy đoán cán bộ.
2. **Tạo PDF riêng cho từng cán bộ:**
   - Gom các công văn theo từng cán bộ xử lý.
   - Tìm file PDF công văn gốc trong thư mục tương ứng (`STT. CV SO CV...`).
   - In toàn văn 100% nội dung công văn.
   - Chèn dòng phân công trên đầu trang 1 của mỗi công văn: **duy nhất 1 dòng**, nền trắng che đè, nội dung: `CB: <Cán bộ> | MST: <MST> | <Tên công ty> | Ngày PC: <dd/mm/yyyy> | CV <Số CV>`.
   - Cỡ chữ tự động co giãn theo chiều ngang của trang scan thực tế (kể cả trang scan độ phân giải cao ~1945x2738 pt) để đảm bảo không bị cụt hay tràn chữ.
   - **Đảm bảo tổng số trang in của mỗi cán bộ luôn là số chẵn:** Nếu tổng số trang của cán bộ là số lẻ, tự động bổ sung thêm 01 trang trắng ở cuối để khi in 2 mặt không bị lẫn lộn giữa các cán bộ.
   - Xuất file riêng cho từng cán bộ lưu vào `exports/in-phan-cong-kt1/<Tên_cán_bộ>.pdf`.
3. **Ghép và xuất file in tổng hợp:**
   - Ghép toàn bộ file của các cán bộ thành một file PDF tổng hợp duy nhất.
   - Đặt tên file theo chuẩn: `Phan cong xu ly KT1 tu dong <dòng đầu> den dong <dòng cuối> ngay in <dd-mm-yyyy>.pdf`.
   - Lưu đồng thời vào:
     - `I:\My Drive\KIEM SOAT HSO CUNG CAP\IN PHAN CONG\`
     - `exports/in-phan-cong-kt1/`
4. **Báo cáo và chờ xác nhận chốt in:**
   - Báo cáo số lượng dòng, số công văn, danh sách cán bộ và số trang in tương ứng.
   - Đợi người dùng kiểm tra bản xem trước và ra lệnh "chốt in".
5. **Cập nhật trạng thái sang "Đang xử lý":**
   - Đảm bảo đóng hoàn toàn các ứng dụng hoặc tiến trình Excel đang mở trên máy để tránh lock file.
   - Chạy script với tham số `--apply` để cập nhật cột O từ "Mới" sang "Đang xử lý" bằng Excel COM (`win32com.client.DispatchEx`), đảm bảo bảo toàn cấu trúc workbook và biểu đồ Dashboard.
   - Đọc kiểm tra lại dữ liệu sau khi ghi để xác nhận hoàn tất.

## Các lệnh thực thi
- `python scripts/in_phan_cong_kt1.py` : Quét sổ, tạo PDF riêng và file tổng hợp, chưa đổi trạng thái.
- `python scripts/in_phan_cong_kt1.py --apply` : Chốt in, cập nhật trạng thái cột O sang "Đang xử lý".
- `python scripts/in_phan_cong_kt1.py --reprint` : In lại các dòng đã ở trạng thái "Đang xử lý" (phục vụ in lại khi cần).
- `python scripts/in_phan_cong_kt1.py --reprint --row=<số dòng>` : In lại giới hạn cho các dòng cụ thể (ví dụ: `--row=816 --row=817`).

## Bẫy
- **Trang scan kích thước lớn:** Các file scan công văn có khổ rất lớn (khoảng 1945 x 2738 pt). Nếu dùng font cố định 7.5 - 11 pt chữ sẽ bị tí hon không đọc được. Luôn tính toán tỷ lệ co giãn theo chiều ngang của trang và chiều dài chuỗi chữ.
- **Tiến trình Excel giữ lock file:** Khi ghi bằng COM, nếu Excel đang mở file hoặc có process ẩn (`EXCEL.EXE`), lệnh sẽ treo vô hạn. Phải kiểm tra và đóng/tắt tiến trình trước khi gọi `--apply`.
- **Lỗi cache win32com:** Nếu gặp lỗi COM cache, xóa thư mục `gen_py` trong `%LOCALAPPDATA%\Temp\gen_py` trước khi chạy lại.
- **Ghi đè trạng thái trước khi kiểm tra:** Tuyệt đối không tự ý chạy `--apply` khi mới tạo file xem thử; chỉ thực hiện khi người dùng xác nhận "chốt in".
- **Trùng lặp / In lại nhầm:** Khi dùng `--reprint`, luôn kèm `--row` nếu chỉ muốn in lại một vài dòng cụ thể để tránh tạo ra PDF cho toàn bộ các dòng cũ.

## Lịch sử (tự học)
- [2026-11-06] Xây dựng kỹ năng in đồng loạt công văn giao Tổ KT1: tạo PDF chẵn trang cho từng cán bộ, chèn dòng phân công tự co giãn font trên đầu trang 1.
- [2026-11-06] Chuẩn hóa tên file in tổng hợp dạng `Phan cong xu ly KT1 tu dong <đầu> den dong <cuối> ngay in <dd-mm-yyyy>.pdf` và tự động đẩy lên Google Drive `I:\My Drive\KIEM SOAT HSO CUNG CAP\IN PHAN CONG\`.
- [2026-11-06] Chốt trạng thái sau khi in là "Đang xử lý" (không dùng "Đang thực hiện"); thêm cơ chế đóng an toàn Excel trước khi cập nhật COM để tránh treo tiến trình.
