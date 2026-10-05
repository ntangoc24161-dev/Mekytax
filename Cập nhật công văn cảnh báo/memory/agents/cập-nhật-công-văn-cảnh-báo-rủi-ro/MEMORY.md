## Bài học (tự học)
- Khi tra cứu văn bản trên VNPT iOffice, gọi trực tiếp hàm JavaScript `zipfileDownload_('1', docId)` để tải toàn bộ tệp đính kèm dạng ZIP về máy một cách tự động.
- Form tra cứu VNPT iOffice ẩn các ô lọc ngày trong div_timkiem_nangcao, cần kích hoạt showDivTimKiem() trước khi điền ngày ban hành và ngày đến.
- Khi báo tiến độ lô mới, đối chiếu mtime sổ Excel với mtime tệp vừa tải; nếu Excel cũ hơn giờ tải thì Bước 5 chưa chạy dù file `.bak` cùng ngày vẫn còn.
- Cookie iOffice trong `ioffice_session.json` hết hạn sau khoảng một ngày; phải đăng nhập lại trước khi tra cứu tiếp.
- `doc2md --list-backends` báo paddle OK nhưng PPStructureV3 vẫn vướng WDAC (DLL bidi/sklearn); với bản scan nên đọc ảnh trang đầu hoặc gọi PaddleOCR trực tiếp sau khi stub module bidi.
- Cùng số ký hiệu tải hai lần (văn bản đến và đi) thì lọc theo hash; riêng `608/TCS4-QLDN` có hai file scan khác hash thì giữ bản nhiều trang hơn.
