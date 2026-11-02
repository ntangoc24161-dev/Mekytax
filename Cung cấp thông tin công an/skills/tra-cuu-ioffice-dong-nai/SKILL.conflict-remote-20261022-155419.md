---
name: Tra cứu và tải văn bản iOffice Đồng Nai
description: Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
group: Năng suất
updated: '2026-09-26'
learned_updated: '2026-09-26'
---

# Tra cứu và tải văn bản iOffice Đồng Nai

Kỹ năng tự động hóa đăng nhập, duy trì phiên làm việc, tra cứu hồ sơ văn bản và tải toàn bộ tệp đính kèm từ Hệ thống Quản lý văn bản và điều hành (VNPT iOffice) của Thuế Thành phố Đồng Nai về máy tính cá nhân.

## Khi nào dùng
- Khi cần đăng nhập tự động vào hệ thống VNPT iOffice Đồng Nai (`https://dongnai.vnptioffice.vn/qlvbdh_dni/main?lang=vi`).
- Khi cần tra cứu danh sách công văn, văn bản đến, văn bản đi theo các điều kiện lọc: trích yếu (ví dụ: *"cảnh báo"*, *"xác minh"*, *"kết luận"*), số ký hiệu văn bản, khoảng thời gian ban hành hoặc ngày văn bản đến.
- Khi cần tải hàng loạt tệp đính kèm (file PDF, DOCX, ZIP) của các văn bản tìm thấy về các thư mục lưu trữ chuyên đề (như `D:\HO SO CONG VAN CANH BAO`, `D:\HSO CONG AN FILE MD`).
- Khi cần tích hợp luồng lấy công văn cảnh báo mới nhất từ iOffice để chuyển tiếp sang pipeline bóc tách OCR và nạp sổ theo dõi Excel của skill `theo-doi-cv-canh-bao-rui-ro`.
- Khi tra cứu văn bản iOffice và cập nhật Sổ Công văn Công an: kết quả tra cứu và cập nhật sổ theo dõi được ghi nhận trực tiếp vào file `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).

## Thông tin hệ thống & Cấu hình tài khoản
- **Đường dẫn truy cập:** `https://dongnai.vnptioffice.vn/qlvbdh_dni/main?lang=vi`
- **Tài khoản (Username):** `ntangoc3.dni.don`
- **Mật khẩu (Password):** `Dni@1234`
- **Người dùng xác thực:** Nguyễn Thị Ánh Ngọc (Thuế Thành phố Đồng Nai)
- **Tệp lưu phiên làm việc (Session state):** `ioffice_session.json` (lưu trữ cookies và localStorage để tái sử dụng, không phải gửi lại mật khẩu mỗi lần thao tác).

## Cấu trúc công cụ & Script thực thi
Toàn bộ mã nguồn tự động hóa được đóng gói tại thư mục `scripts/` của skill:
- **`scripts/ioffice_client.py`:** Thư viện lõi Playwright điều khiển trình duyệt ngầm (headless Chrome), xử lý đăng nhập, quản lý session, lọc tìm kiếm, bắt sự kiện tải tệp `allFileDownload(doc_id)` và giải nén ZIP tự động.
- **`scripts/run_ioffice.py`:** Giao diện dòng lệnh (CLI) cho phép agent hoặc người dùng chạy trực tiếp các tác vụ `login`, `search`, `download`, `status`.

## Quy trình vận hành chuẩn 5 bước

### Bước 1: Khởi tạo và kiểm tra phiên làm việc (Session Management)
- Hệ thống kiểm tra tệp `ioffice_session.json` tại thư mục làm việc.
- Nếu phiên còn hiệu lực (truy cập thẳng `Trang chủ`), tái sử dụng ngay lập tức mà không cần gõ lại tài khoản/mật khẩu.
- Nếu phiên hết hạn hoặc tệp chưa tồn tại: Tự động điều hướng đến form đăng nhập, điền `#userName`, `#passWord`, nhấn `#submitBtn`, chờ hệ thống xác thực và lưu lại session mới.

### Bước 2: Điều hướng đến phân hệ tra cứu
- Kích hoạt menu **Tra cứu văn bản** thông qua lời gọi hàm hoặc click vào liên kết menu `m2286`.
- Chờ form tìm kiếm nâng cao hiển thị đầy đủ các trường nhập liệu và nút chức năng `#btnClickLoading`.

### Bước 3: Thiết lập điều kiện lọc dữ liệu
Tùy theo câu lệnh của người dùng, điền các tham số tìm kiếm vào các trường tương ứng:
- **Trích yếu:** Điền vào `#txt_trichyeu` (ví dụ: `cảnh báo`).
- **Số ký hiệu văn bản:** Điền vào `#txt_so_kyhieu` (ví dụ: `5977/TCS2-QLDN`).
- **Thời gian ban hành:** Điền khoảng từ ngày `#txt_start_date_banhanh` đến ngày `#txt_end_date_banhanh` theo định dạng `DD/MM/YYYY`.
- **Thời gian văn bản đến:** Điền khoảng từ ngày `#txt_start_date_ngayden` đến ngày `#txt_end_date_ngayden` theo định dạng `DD/MM/YYYY`.
- **Số lượng bản ghi hiển thị:** Thiết lập `#selectPageRec` lên mức tối đa (50 hoặc 100) để lấy toàn bộ dữ liệu trên một trang.

### Bước 4: Thực thi tra cứu và trích xuất danh sách
- Nhấn nút tìm kiếm `#btnClickLoading` (`check_button_click();btnSearchAdvanceClick();`).
- Bóc tách bảng kết quả hiển thị gồm các trường: Số thứ tự (STT), Số ký hiệu văn bản, Số đến, Trích yếu nội dung.
- Phân tích thuộc tính `href` hoặc `onclick` của nút tải file để lấy mã định danh tài liệu nội bộ: `doc_id` từ hàm `allFileDownload(doc_id)`.

### Bước 5: Tải tệp và phân loại lưu trữ
- Với mỗi văn bản có `doc_id`, kích hoạt hàm `allFileDownload(doc_id)` kết hợp với bộ lắng nghe sự kiện tải xuống `page.expect_download()`.
- Tự động đổi tên tệp tải về theo cấu trúc chuẩn: `[<Số đến>]_<Số ký hiệu>_<Tên tệp gốc>`.
- Nếu tệp tải về là dạng nén `.zip`, hệ thống tự động giải nén ra thư mục con cùng tên để giải phóng trực tiếp các tệp văn bản `.pdf`, `.docx`.
- Lưu trữ vào đúng thư mục đích theo yêu cầu của người dùng (ví dụ: `D:\HO SO CONG VAN CANH BAO\2026\2026-09` hoặc thư mục chuyển đổi).

---

## Hướng dẫn giao việc cho Agent (Cách đưa câu lệnh)

Khi cần tra cứu và lưu văn bản trên hệ thống VNPT iOffice, bạn chỉ cần gửi câu lệnh theo các mẫu dưới đây tùy theo mục đích:

### Mẫu 1: Tra cứu và tải công văn cảnh báo theo khoảng thời gian (Phổ biến nhất)
```text
Tra cứu và tải các công văn cảnh báo trên iOffice:
- Trích yếu: cảnh báo
- Thời gian: từ ngày 01/05/2026 đến 22/09/2026
- Thư mục lưu: D:\HO SO CONG VAN CANH BAO\Chuyen_doi
- Tự động giải nén các file zip sau khi tải xong và báo cáo danh sách văn bản đã lưu.
```

### Mẫu 2: Tra cứu và tải theo số ký hiệu công văn cụ thể
```text
Tra cứu trên iOffice văn bản có số ký hiệu "5977/TCS2-QLDN", tải toàn bộ tệp đính kèm về thư mục D:\HO SO CONG VAN CANH BAO\2026\2026-09.
```

### Mẫu 3: Quét kiểm tra văn bản đến mới nhất trong ngày/tuần
```text
Kiểm tra trên iOffice các văn bản đến từ ngày 20/09/2026 đến hôm nay có trích yếu chứa từ khóa "công an" hoặc "xác minh", liệt kê số ký hiệu và trích yếu vào khung chat trước khi tải.
```

### Mẫu 4: Giao việc nền trọn gói tự động 100% (Kanban không dừng lại hỏi)
```text
Giao việc nền Kanban chạy trọn gói skill tra-cuu-ioffice-dong-nai:
1. Đăng nhập iOffice bằng tài khoản ntangoc3.dni.don.
2. Tra cứu văn bản có trích yếu "cảnh báo" phát sinh trong tháng 09/2026.
3. Tải toàn bộ tệp đính kèm về D:\HO SO CONG VAN CANH BAO\2026\2026-09.
4. Tự động giải nén các tệp zip, sau đó chuyển tiếp sang skill theo-doi-cv-canh-bao-rui-ro để bóc tách thông tin và nạp vào Sổ theo dõi Excel.
5. Thông báo kết quả tổng hợp vào khung chat khi hoàn tất.
```

---

## Cú pháp lệnh chạy thủ công bằng CLI
Trong trường hợp cần kiểm tra nhanh trực tiếp từ dòng lệnh:
- **Kiểm tra trạng thái đăng nhập:**
  ```powershell
  python "skills/tra-cuu-ioffice-dong-nai/scripts/run_ioffice.py" --action status
  ```
- **Thực hiện tra cứu văn bản:**
  ```powershell
  python "skills/tra-cuu-ioffice-dong-nai/scripts/run_ioffice.py" --action search --keyword "cảnh báo" --max-docs 10
  ```
- **Tải văn bản về thư mục chỉ định:**
  ```powershell
  python "skills/tra-cuu-ioffice-dong-nai/scripts/run_ioffice.py" --action download --keyword "cảnh báo" --from-date "01/09/2026" --to-date "26/09/2026" --output-dir "D:\HO SO CONG VAN CANH BAO\2026\2026-09" --max-docs 20
  ```

---

## Bẫy nghiệp vụ & Lưu ý kỹ thuật
- **Cơ chế tải file một gói (`allFileDownload`):** Cổng iOffice đóng gói toàn bộ tài liệu của văn bản thành 1 tệp `.zip` khi nhấn nút tải. Script đã được cấu hình tự động giải nén ra thư mục con để người dùng mở trực tiếp tệp PDF mà không cần un-zip thủ công.
- **Bảo mật và duy trì Session:** Không lưu mật khẩu lộ thiên trong lịch sử nếu không cần thiết; phiên làm việc được lưu trong `ioffice_session.json` có hiệu lực nhiều ngày, hạn chế tối đa số lần phải nhập lại mật khẩu.
- **Thời gian chờ mạng (Timeout):** Với các văn bản có tệp đính kèm dung lượng lớn (trên 20MB), bộ lắng nghe `expect_download` được đặt thời gian chờ tối thiểu 30 giây để tránh bị ngắt kết nối giữa chừng.

## Lịch sử (tự học)
- [2026-10-02] Cập nhật chuẩn hóa file đích cho luồng kết quả tra cứu iOffice và cập nhật Sổ Công văn Công an thành `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).

