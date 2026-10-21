---
name: Rà soát và đồng bộ thư mục Công văn hoàn tất
description: Rà soát kiểm tra file công văn trả lời đã ký số trong toàn bộ thư mục công văn năm 2026 và tự động cập nhật hậu tố -R vào tên thư mục.
group: Tài chính
updated: '2026-09-26'
learned_updated: '2026-09-26'
---
## Khi nào dùng
- Khi cần rà soát toàn bộ các thư mục lưu trữ hồ sơ công văn năm 2026 trên Google Drive để đồng bộ trạng thái hoàn tất xử lý.
- Tự động nhận diện các thư mục đã có tệp Công văn trả lời ký số (`.signed.pdf`, `.signed.signed.pdf` hoặc bản scan ký đóng dấu) nhưng tên thư mục chưa có hậu tố `-R`, và tiến hành đổi tên thư mục thêm `-R` ở cuối.

## Chuẩn bị
- Thư mục gốc lưu trữ công văn năm 2026: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026`.
- Tiêu chí nhận diện tệp Công văn trả lời đã ký số:
  + Tệp có hậu tố chữ ký số điện tử: `.signed.pdf`, `.signed.signed.pdf`.
  + Tệp chứa các cụm từ nhận diện văn bản phúc đáp: `TRA LOI CV...`, `CV TRẢ LỜI...`, `Cung cấp thông tin...`, `Công văn trả lời...`, `Tra loi...`.
  + Tệp bản scan công văn trả lời đã có đầy đủ chữ ký và dấu mộc đỏ của cơ quan thuế.

## Quy trình
1. **Quét danh mục thư mục năm 2026:**
   - Liệt kê toàn bộ các thư mục con trong `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026`.
   - Bỏ qua các thư mục đặc biệt không phải hồ sơ công văn (ví dụ: `CV MOI`, các thư mục tạm).
   - Lọc ra danh sách các thư mục **chưa có hậu tố `-R` hoặc `-r` ở cuối tên thư mục**.

2. **Kiểm tra tệp công văn trả lời ký số bên trong:**
   - Quét tất cả các tệp `.pdf` bên trong từng thư mục cần kiểm tra.
   - Phân biệt chính xác giữa:
     + *Công văn đến của cơ quan Công an* (văn bản yêu cầu cung cấp thông tin, có thể mang chữ ký số của Công an).
     + *Công văn trả lời của cơ quan Thuế* (văn bản phúc đáp cung cấp tài liệu gửi cơ quan Công an mang chữ ký số của Thuế Thành phố Đồng Nai / Thuế Cơ sở 3).
   - Xác định danh sách các tệp công văn trả lời đã ký số hợp lệ.

3. **Thực hiện đổi tên thư mục bổ sung hậu tố `-R`:**
   - Giữ nguyên toàn bộ phần tên thư mục gốc hiện hữu (bao gồm số thứ tự, số công văn, cơ quan ban hành, ghi chú cán bộ...).
   - Bổ sung chuỗi `-R` vào cuối tên thư mục (ví dụ: `132. CV 17932-CA DONG NAI` -> `132. CV 17932-CA DONG NAI-R`).
   - Đảm bảo việc đổi tên an toàn, không làm gián đoạn hoặc thất thoát bất kỳ tệp dữ liệu nào bên trong.

4. **Tổng hợp và báo cáo kết quả:**
   - Thống kê tổng số thư mục đã quét, số thư mục đã có đuôi `-R` từ trước, và số thư mục mới được cập nhật thêm `-R`.
   - Lập bảng chi tiết gồm 3 cột: `STT`, `Tên thư mục gốc`, `Tên thư mục mới (đã thêm -R)`, `Tệp Công văn trả lời đã ký số`.

## Bẫy
- Nhầm lẫn tệp công văn yêu cầu của Công an (có ký số của Công an) thành công văn trả lời của Thuế, dẫn đến việc đánh dấu `-R` sai khi hồ sơ chưa thực sự xử lý xong.
- Thư mục vốn đã có đuôi `-R` (hoặc `-R, ...`) bị gắn thêm lần nữa thành `-R-R`. Phải kiểm tra điều kiện kết thúc bằng `-R` hoặc biến thể trước khi đổi tên.
- Xóa nhầm hoặc thay đổi cấu trúc tệp bên trong thư mục trong quá trình đổi tên.

## Kiểm chứng
- Quét lại toàn bộ thư mục năm 2026: 100% các thư mục có tệp công văn trả lời ký số đều có hậu tố `-R` ở cuối.
- Tất cả các tệp bên trong thư mục sau khi đổi tên vẫn mở và truy cập bình thường.

## Lịch sử (tự học)
- [2026-09-26] Khởi tạo skill rà soát và đồng bộ thư mục công văn hoàn tất: tự động kiểm tra chữ ký số công văn trả lời trong toàn bộ thư mục năm 2026 và cập nhật hậu tố `-R` an toàn, chính xác.
