# Wiki Log

Nhật ký thời gian (append-only). Mỗi entry: `## [YYYY-MM-DD] loại | tiêu đề` (loại: init/ingest/update/query/lint/migrate).

## [2026-10-01] ingest | tự học từ chat

Đã tạo:
- [[Định dạng nội dung kiểm tra trên sổ theo dõi]]

## [2026-10-01] ingest | tự học từ chat

Đã tạo:
- [[Xuất hồ sơ kiểm tra để trình ký trước khi có Số QĐKT]]

## [2026-10-01] update | Chuẩn hóa bộ mẫu Hồ sơ kiểm tra kế hoạch và tạo skill mới

Đã cập nhật:
- Chuyển toàn bộ nhóm mẫu Hồ sơ đi (DI-01 đến DI-07) từ mẫu giải thể sang mẫu chuẩn kiểm tra theo kế hoạch năm (QĐ 1412/QĐ-DON, Luật 108/2025/QH15, QĐ 1161/QĐ-CT).
- Tạo file Word tổng hợp mẫu Hồ sơ đi: `MAU TONG HOP HO SO DI.docx` và file đã điền dữ liệu cho Vạn Phát Thành `TONG HOP HO SO DI.docx`.
- Tạo skill mới `ho-so-kiem-tra-ke-hoach`.
- [[Hồ sơ kiểm tra thuế theo kế hoạch năm]]

## [2026-10-01] ingest | tự học từ chat

Đã tạo:
- [[Quy chuẩn căn lề văn bản hành chính thuế theo Nghị định 30/2020/NĐ-CP]]

## [2026-10-02] ingest | tự học từ chat

Đã tạo:
- [[Kỹ thuật điền hồ sơ Word kiểm tra thuế từ phôi tổng hợp]]

## [2026-10-02] update | Đồng bộ phôi gốc MAUTONG HOP.docx và chuẩn hóa xuất file Quyết định

Đã cập nhật:
- Đồng bộ các thay đổi mới nhất từ phôi gốc `MAUTONG HOP.docx` vào Brain (`templates/MAUTONG HOP.docx`).
- Cập nhật quy cách xuất: Không xuất các file hồ sơ đơn lẻ, chỉ xuất các file Quyết định (`QD Kiem tra...`, `QDgiam sat...`), lưu toàn bộ các phần/biểu mẫu còn lại vào file tổng hợp `TONG HOP.docx`.

## [2026-10-02] ingest | tự học từ chat

Đã tạo:
- [[Quy cách xuất văn bản hồ sơ kiểm tra thuế theo kế hoạch]]

## [2026-10-04] ingest | tự học từ chat

Đã tạo:
- [[Kỹ thuật chống tràn trang và định dạng khung mẫu biểu hồ sơ thuế]]

## [2026-10-07] ingest | tự học từ chat

Đã tạo:
- [[Kỹ thuật trình bày văn bản hành chính Tờ trình và Quyết định thuế trong 1 trang in]]

## [2026-10-07] ingest | tự học từ chat

Đã tạo:
- [[Quy cách trình bày tên người nộp thuế và bố cục chữ ký hồ sơ kiểm tra]]

## [2026-10-08] ingest | tự học từ chat

Đã tạo:
- [[Quy chuẩn trình bày tiêu ngữ trong văn bản hành chính]]

## [2026-10-08] ingest | tự học từ chat

Đã tạo:
- [[Phân biệt thể thức văn bản hành chính NĐ 30 và biểu mẫu nghiệp vụ chuyên ngành thuế]]

## [2026-10-08] ingest | tự học từ chat

Đã tạo:
- [[Quy chuẩn trình bày Quyết định kiểm tra thuế]]

## [2026-10-08] ingest | tự học từ chat

Đã tạo:
- [[Quy cách lập Biên bản xác nhận số liệu kiểm tra thuế]]

## [2026-10-08] ingest | tự học từ chat

Đã tạo:
- [[Quy chuẩn trình bày Biên bản kiểm tra thuế]]

## [2026-10-08] update | Triển khai Pipeline 1-Click và Module Linter kiểm định chất lượng

Đã cập nhật:
- Xây dựng module `scripts/lint_ho_so_docx.py`: Tự động kiểm định thể thức NĐ 30, quét sạch font Calibri, kiểm tra bảo toàn checkbox Wingdings (\xa8/\xfe), kiểm tra 7 floating tables và cảnh báo tràn trang.
- Xây dựng pipeline `scripts/xuat_ho_so_tron_goi.py`: Tích hợp xuất trọn gói 1-click từ phôi `TONG HOP CHUAN DINH DANG.docx`, trích xuất 02 Quyết định 1 trang A4, nén `hoso.zip` và tự động kích hoạt Linter trước khi bàn giao.

## [2026-10-11] ingest | tự học từ chat

Đã tạo:
- [[Quy chuẩn thể hiện ngày tháng báo cáo và số liệu bằng 0 trong hồ sơ kiểm tra thuế]]

## [2026-10-15] ingest | tự học từ chat

Đã tạo:
- [[Quy trình hoàn thiện hồ sơ kiểm tra thuế từ Bảng tính]]

## [2026-10-16] ingest | tự học từ chat

Đã tạo:
- [[Quy trình phối hợp 4 trợ lý chuyên trách kiểm tra thuế]]

## [2026-10-16] ingest | tự học từ chat

Đã tạo:
- [[Cấu trúc cột và phân công đoàn kiểm tra trên sổ theo dõi JAVIS PRO]]

## [2026-10-19] ingest | tự học từ chat

Đã tạo:
- [[Quy trình khởi tạo Bảng tính và thư mục Biên bản vụ kiểm tra thuế]]
