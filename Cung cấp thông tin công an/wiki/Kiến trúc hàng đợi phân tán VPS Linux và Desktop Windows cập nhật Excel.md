---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-22
updated: 2026-11-22
source: [[conversations/2026-11-22]]
---
Mô hình kiến trúc phân tán cập nhật file Excel phức tạp chứa Dashboard và Macro giữa môi trường Linux và Windows [[conversations/2026-11-22]]:

- **Tách tầng xử lý (Staging / Queue):** Node chạy trên nền tảng Linux (VPS/Server) thực hiện các tác vụ trích xuất, OCR và bóc tách dữ liệu nhưng không ghi trực tiếp vào tệp `.xlsx` chính để tránh rủi ro hỏng cấu trúc XML, mất cache tính toán và gãy biểu đồ từ các thư viện bên thứ ba (thực tế tính đến 22/11/2026) [[conversations/2026-11-22]]. Dữ liệu bóc tách được lưu vào tệp hàng đợi trung gian dạng JSON hoặc CSV [[conversations/2026-11-22]].
- **Tầng nạp dữ liệu bằng Excel COM:** Node chạy trên hệ điều hành Windows (Desktop) định kỳ kiểm tra trạng thái khóa tệp (`is_file_locked`), sau đó sử dụng tiến trình Microsoft Excel thông qua COM (`win32com.client`) để mở và ghi dữ liệu từ hàng đợi vào sổ chính nhằm bảo toàn nguyên vẹn 100% đồ họa và định dạng [[conversations/2026-11-22]].
- **Tối ưu hóa bảng biểu Excel Table (ListObject):** Chuyển đổi các vùng dữ liệu thuần sang cấu trúc Excel Table cho phép tự động dàn trải công thức (`Calculated Column`) và định dạng sang các dòng mới thêm mà không cần thao tác chèn/kéo công thức thủ công [[conversations/2026-11-22]].
- **Quy tắc an toàn công thức và tối ưu hiệu năng:** Tuyệt đối không xóa hàng vật lý (`Delete Row`) tại các vùng có công thức tham chiếu nối tiếp để phòng ngừa lỗi `#REF!`, thay vào đó sử dụng lệnh xóa nội dung (`ClearContents`) [[conversations/2026-11-22]]; chuyển các dòng dữ liệu quá khứ sang giá trị tĩnh (Paste Values) và loại bỏ các hàm biến động (`OFFSET`, `INDIRECT`) [[conversations/2026-11-22]].
