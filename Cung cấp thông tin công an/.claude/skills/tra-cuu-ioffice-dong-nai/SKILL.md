---
name: Tra cứu và tải văn bản iOffice Đồng Nai
description: Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
group: Năng suất
updated: '2026-10-12'
learned_updated: '2026-10-12'
---
# Tra cứu và tải văn bản iOffice Đồng Nai

Kỹ năng tự động hóa đăng nhập, duy trì phiên làm việc, tra cứu hồ sơ văn bản và tải toàn bộ tệp đính kèm từ Hệ thống Quản lý văn bản và điều hành (VNPT iOffice) của Thuế Thành phố Đồng Nai về máy tính cá nhân.

## Khi nào dùng
- Khi cần đăng nhập tự động vào hệ thống VNPT iOffice Đồng Nai (`https://dongnai.vnptioffice.vn/qlvbdh_dni/main?lang=vi`).
- Khi cần tra cứu danh sách công văn, văn bản đến, văn bản đi theo các điều kiện lọc: trích yếu, số ký hiệu văn bản, khoảng thời gian ban hành hoặc ngày văn bản đến.
- Khi người dùng yêu cầu rà công văn cung cấp thông tin Công an chưa cập nhật sổ: lọc ANĐT/CSKT/KTr1 theo khoảng ngày, đối chiếu `Bao_cao_cong_van.xlsx`, tải file mới.
- Khi cần tải hàng loạt tệp đính kèm (file PDF, DOCX, ZIP) về thư mục chuyên đề (như `D:\HO SO CONG VAN CANH BAO`, `D:\HSO CONG AN FILE MD`, `CV MOI`).
- Khi cần tích hợp luồng lấy công văn cảnh báo mới nhất từ iOffice để chuyển tiếp sang pipeline bóc tách OCR và nạp sổ theo dõi Excel của skill `theo-doi-cv-canh-bao-rui-ro`.
- Khi tra cứu văn bản iOffice và cập nhật Sổ Công văn Công an: kết quả ghi vào `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).

## Chuẩn bị
- Playwright + Chromium: nếu chưa có thì cài trước khi login (`pip install playwright` rồi `playwright install chromium`). Không giả định máy đã có sẵn.
- Script: `skills/tra-cuu-ioffice-dong-nai/scripts/ioffice_client.py` và `scripts/run_ioffice.py`.
- Tệp phiên: `ioffice_session.json`. Gọi `--action status` trước; cookie hết hạn thì login lại, không dùng session cũ.
- Tài khoản: `ntangoc3.dni.don` / `Dni@1234` (Nguyễn Thị Ánh Ngọc).

## Thông tin hệ thống & Cấu hình tài khoản
- **Đường dẫn truy cập:** `https://dongnai.vnptioffice.vn/qlvbdh_dni/main?lang=vi`
- **Tài khoản (Username):** `ntangoc3.dni.don`
- **Mật khẩu (Password):** `Dni@1234`
- **Người dùng xác thực:** Nguyễn Thị Ánh Ngọc (Thuế Thành phố Đồng Nai)
- **Tệp lưu phiên làm việc (Session state):** `ioffice_session.json` (cookies và localStorage; tái sử dụng khi còn hiệu lực).

## Cấu trúc công cụ & Script thực thi
Toàn bộ mã nguồn tự động hóa được đóng gói tại thư mục `scripts/` của skill:
- **`scripts/ioffice_client.py`:** Thư viện lõi Playwright điều khiển trình duyệt ngầm (headless Chrome), xử lý đăng nhập, quản lý session, lọc tìm kiếm, bắt sự kiện tải tệp `allFileDownload(doc_id)` và giải nén ZIP tự động.
- **`scripts/run_ioffice.py`:** Giao diện dòng lệnh (CLI) cho phép agent hoặc người dùng chạy trực tiếp các tác vụ `login`, `search`, `download`, `status`.

## Quy trình vận hành chuẩn 5 bước

### Bước 1: Khởi tạo và kiểm tra phiên làm việc (Session Management)
- Kiểm tra tệp `ioffice_session.json` tại thư mục làm việc.
- Nếu phiên còn hiệu lực (truy cập thẳng Trang chủ), tái sử dụng ngay, không gõ lại mật khẩu.
- Nếu phiên hết hạn hoặc tệp chưa tồn tại: điều hướng form đăng nhập, điền `#userName`, `#passWord`, nhấn `#submitBtn`, chờ xác thực và lưu session mới.
- Captcha có thể nằm trong HTML nhưng form vẫn hiện ô tài khoản/mật khẩu: điền và submit. Chỉ xử lý captcha khi ô captcha thật sự hiện và chặn đăng nhập.

### Bước 2: Điều hướng đến phân hệ tra cứu
- Kích hoạt menu **Tra cứu văn bản** (`m2286`).
- **Bắt buộc mở Tìm kiếm nâng cao** trước khi điền ngày. Form mặc định ẩn ô ngày (`#txt_start_date_ngayden`, `#txt_end_date_ngayden`, ngày ban hành). Không tra cứu theo khoảng ngày khi ô ngày còn ẩn.
- Chờ form nâng cao hiện đủ trường và nút `#btnClickLoading`.

### Bước 3: Thiết lập điều kiện lọc dữ liệu
- **Trích yếu:** `#txt_trichyeu`.
  - Công văn cảnh báo: `cảnh báo`.
  - Công văn cung cấp thông tin Công an: **không** dùng từ khóa `công an`. Lọc số ký hiệu/trích yếu chứa `ANĐT`, `CSKT`, `KTr1`.
- **Số ký hiệu văn bản:** `#txt_so_kyhieu`.
- **Thời gian ban hành:** `#txt_start_date_banhanh` đến `#txt_end_date_banhanh` định dạng `DD/MM/YYYY`.
- **Thời gian văn bản đến:** `#txt_start_date_ngayden` đến `#txt_end_date_ngayden` định dạng `DD/MM/YYYY`.
- **Số lượng bản ghi:** luôn đẩy `#selectPageRec` lên 50 hoặc 100. CLI `--max-docs 10` cắt danh sách; khi quét theo khoảng ngày đặt tối thiểu 50.

### Bước 4: Thực thi tra cứu và trích xuất danh sách
- Nhấn `#btnClickLoading` (`check_button_click();btnSearchAdvanceClick();`).
- Bóc tách bảng: STT, Số ký hiệu, Số đến, Trích yếu.
- Lấy `doc_id` từ `allFileDownload(doc_id)`.
- Với luồng Công an: loại bản ghi không phải ANĐT/CSKT/KTr1 trước khi tải.

### Bước 5: Tải tệp và phân loại lưu trữ
- Với mỗi văn bản có `doc_id`, gọi `allFileDownload(doc_id)` kèm `page.expect_download()`.
- Đổi tên: `[<Số đến>]_<Số ký hiệu>_<Tên tệp gốc>`.
- File `.zip` thì giải nén ra thư mục con.
- Công văn Công an mới: lưu vào `CV MOI` rồi chuyển skill `theo-doi-cung-cap-ca`. Công văn cảnh báo: `D:\HO SO CONG VAN CANH BAO\...`.

## Hướng dẫn giao việc cho Agent (Cách đưa câu lệnh)

### Mẫu 1: Tra cứu và tải công văn cảnh báo theo khoảng thời gian
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

### Mẫu 3: Quét công văn cung cấp thông tin Công an chưa có trên sổ
```text
Đăng nhập iOffice, mở tìm kiếm nâng cao, nới số bản ghi tối đa.
Tra văn bản đến từ ngày [từ] đến hôm nay, lọc ANĐT/CSKT/KTr1 (không dùng từ khóa "công an").
Đối chiếu số ký hiệu với Bao_cao_cong_van.xlsx sheet NHAP_HO_SO; liệt kê số chưa có rồi tải đính kèm về CV MOI.
```

### Mẫu 4: Giao việc nền trọn gói tự động 100% (Kanban không dừng lại hỏi)
```text
Giao việc nền Kanban chạy trọn gói skill tra-cuu-ioffice-dong-nai:
1. Đăng nhập iOffice bằng tài khoản ntangoc3.dni.don.
2. Tra cứu văn bản có trích yếu "cảnh báo" phát sinh trong tháng chỉ định.
3. Tải toàn bộ tệp đính kèm về D:\HO SO CONG VAN CANH BAO\...
4. Tự động giải nén zip, chuyển tiếp skill theo-doi-cv-canh-bao-rui-ro.
5. Thông báo kết quả vào khung chat khi hoàn tất.
```

## Cú pháp lệnh chạy thủ công bằng CLI
- Kiểm tra trạng thái đăng nhập:
  ```powershell
  python "skills/tra-cuu-ioffice-dong-nai/scripts/run_ioffice.py" --action status
  ```
- Tra cứu văn bản (nới max-docs, không để 10):
  ```powershell
  python "skills/tra-cuu-ioffice-dong-nai/scripts/run_ioffice.py" --action search --keyword "CSKT" --max-docs 50
  ```
- Tải văn bản về thư mục chỉ định:
  ```powershell
  python "skills/tra-cuu-ioffice-dong-nai/scripts/run_ioffice.py" --action download --keyword "ANĐT" --from-date "01/10/2026" --to-date "12/10/2026" --output-dir "I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI" --max-docs 50
  ```

## Bẫy
- **Phiên cookie hết hạn:** `ioffice_session.json` không dùng được thì login lại. Đừng lặp search trên session chết.
- **Ô ngày ẩn:** form mặc định giấu khoảng ngày. Phải mở Tìm kiếm nâng cao rồi mới điền ngày văn bản đến/ban hành.
- **Từ khóa "công an" quá rộng:** lần quét 01/10/2026 khớp văn bản không thuộc luồng cung cấp hồ sơ. Lọc ANĐT/CSKT/KTr1.
- **`--max-docs 10` cắt danh sách:** luôn nới `#selectPageRec` và `--max-docs` tối thiểu 50 khi quét theo khoảng ngày.
- **Captcha trên HTML:** nếu ô tài khoản/mật khẩu đang hiện thì submit bình thường; chỉ dừng khi captcha chặn thật.
- **Playwright chưa cài:** cài Chromium trước, không báo "không vào được iOffice" vì thiếu thư viện.
- **Cơ chế tải `allFileDownload`:** cổng đóng gói ZIP; script giải nén ra thư mục con.
- **Timeout file lớn:** `expect_download` chờ tối thiểu 30 giây với tệp trên 20MB.

## Kiểm chứng
- `status` ra phiên sống, hoặc login thành công với tài khoản Nguyễn Thị Ánh Ngọc.
- Ô ngày đã hiện trước khi bấm tìm.
- Số dòng kết quả phản ánh `#selectPageRec` 50/100, không kẹt 10 dòng đầu.
- Danh sách Công an chỉ còn ANĐT/CSKT/KTr1.
- File đính kèm nằm đúng thư mục đích; số hiệu mới đã đối chiếu sheet `NHAP_HO_SO` trước khi tải.

## Lịch sử (tự học)
- [2026-10-02] Cập nhật chuẩn hóa file đích cho luồng kết quả tra cứu iOffice và cập nhật Sổ Công văn Công an thành `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).
- [2026-10-02] Chuẩn hóa file đích kết quả tra cứu iOffice và Sổ Công văn Công an thành `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).
- [2026-10-12] Lần rà từ 01/10/2026: phiên cookie hết hạn; captcha có trong HTML nhưng form login vẫn dùng được; ô ngày ẩn đến khi mở tìm kiếm nâng cao; từ khóa "công an" khớp quá rộng phải lọc ANĐT/CSKT/KTr1; `--max-docs 10` cắt danh sách; Playwright có thể chưa cài trên máy chạy.
- [2026-10-12] Lần chạy 12/10/2026 làm theo skill cũ ra kết quả sai/thiếu: session cookie hết hạn phải login lại; ô ngày ẩn nên form mặc định không lọc được từ 01/10; mẫu CLI --max-docs 10 và mẫu 3 dùng từ khóa "công an" lấy tập quá rộng; phải lọc ANĐT/CSKT/KTr1 và nới số bản ghi. Bổ sung bẫy captcha trên HTML, cà
