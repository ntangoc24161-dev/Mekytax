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

## [2026-10-23] ingest | tự học từ chat

Đã tạo:
- [[Quy chuẩn thể hiện số văn bản khi xuất hồ sơ kiểm tra thuế chưa có số quyết định]]

## [2026-10-23] update | Cắt token phòng
- Tạo: [[Cat-token-2026-10-23]]
- orchestration off; 11 skill -> .disabled; còn 15 skill; 2 workflow vẫn off.

## [2026-10-26] ingest | tự học từ chat

Đã tạo:
- [[Giải trình loại chi phí lương trong Biên bản xác nhận số liệu kiểm tra thuế]]

## [2026-10-27] ingest | tự học từ chat

Đã tạo:
- [[Kỹ thuật căn chỉnh chữ ký trong ô và chống rớt trang theo Nghị định 30]]

## [2026-10-30] ingest | tự học từ chat

Đã tạo:
- [[Diễn giải loại chi phí giá vốn không tương ứng doanh thu]]

## [2026-10-30] cập nhật hồ sơ | Sơn Tiền Hải lấy bộ Mo Fa làm phôi nền

Đã cập nhật `TONG HOP.docx` và `QD XPHC` vụ Sơn Tiền Hải (MST 3603859012) theo nền pháp lý
bộ Mo Fa (MST 3603961087): Luật Thanh tra 2025, QĐ 1614/QĐ-TCT ngày 13/10/2020,
QĐ 1377/QĐ-CT ngày 30/06/2025; chuẩn checkbox ☐/☑ Unicode; điền Mục III nghĩa vụ NSNN,
nhật ký kiểm tra, 4 nội dung Điều 1; xuất file dữ liệu trộn.
Liên quan: [[Chuẩn định dạng TONG HOP hồ sơ kiểm tra thuế]],
[[Mẫu hồ sơ kiểm tra thuế đối với doanh nghiệp giải thể]],
[[Diễn giải loại chi phí giá vốn không tương ứng doanh thu]]

## [2026-10-30] update | Cập nhật số QĐKT, QĐGS, kế toán, TKNH và đổi chỉ tiêu QLDN vụ Sơn Tiền Hải

Đã cập nhật toàn diện hồ sơ vụ Sơn Tiền Hải (MST 3603859012):
- Đổi toàn bộ "chi phí quản lý kinh doanh" -> "chi phí quản lý doanh nghiệp" (18 vị trí trên các bảng BBXN, BBKT và văn bản).
- Điền Số QĐKT `7092/QĐ-TCS3` và QĐGS `7093/QĐ-TCS3` ngày ban hành `26/10/2026`.
- Điền Kế toán trưởng: `Lê Hoài Phương` (6 trường và 2 chữ ký tại BBXN Trưởng đoàn, Thành viên).
- Điền Tài khoản ngân hàng: `6722949824 tại Ngân hàng TMCP Đầu tư và phát triển Việt Nam`.
- Giữ trống các trường chưa cung cấp (CCCDKT).
- Xuất riêng 2 văn bản: `QD Kiem tra tai cty Son Tien Hai.docx` và `QDgiam sat tai cty Son Tien Hai.docx`.
- Đồng bộ dữ liệu lên sổ theo dõi `Danh sach kiem tra quyet toan pro.xlsm` dòng 16 và thư mục live G:.
Liên quan: [[Chuẩn định dạng TONG HOP hồ sơ kiểm tra thuế]], [[Kỹ thuật điền hồ sơ Word kiểm tra thuế từ phôi tổng hợp]]

## [2026-10-31] ingest | tự học từ chat

Đã tạo:
- [[Chuẩn hóa tên chỉ tiêu chi phí quản lý doanh nghiệp trong hồ sơ kiểm tra]]

## [2026-10-31] ingest | tự học từ chat

Đã tạo:
- [[Hồ sơ đề xuất hoàn thuế kiêm bù trừ thu ngân sách đối với doanh nghiệp giải thể]]

## [2026-11-09] ingest | tự học từ chat

Đã tạo:
- [[Cấu trúc bộ hồ sơ xác minh NCC rủi ro và TNCN nhiều nguồn thu nhập]]

## [2026-11-11] update | Cập nhật quy tắc Hardcoded Data Check (cảnh báo mềm) và hồ sơ giải thể vào skill

Đã cập nhật hệ thống skill liên quan đến Bảng tính và hoàn thiện hồ sơ kiểm tra thuế:
- `dien-bien-ban-tu-bang-tinh`: Thêm bước kiểm tra số liệu gõ cứng trên sheet DTCHEP. Khi phát hiện ô chênh lệch lớn không có công thức / không sheet chi tiết, phát cảnh báo mềm (non-blocking warning), ghi vào `wiki/_open-questions.md`, hoàn toàn không chặn luồng tự động đổ số vào biên bản. Thống nhất tên "Chi phí quản lý doanh nghiệp" và quy chuẩn điền đủ Mục III Nghĩa vụ NSNN các năm cho hồ sơ giải thể.
- `boc-tach-bang-tinh-phan-tich`: Bổ sung ràng buộc kiểm tra số gõ cứng dạng cảnh báo mềm khi rà soát Bảng tính.
- `ho-so-kiem-tra-ke-hoach`: Ghi nhận bài học về hồ sơ giải thể trên nền kế hoạch năm.
- Đồng bộ toàn bộ sang `.claude/skills/` và cập nhật fact memory.
Liên quan: [[Chuẩn định dạng TONG HOP hồ sơ kiểm tra thuế]], [[Quy cách lập Biên bản xác nhận số liệu kiểm tra thuế]]

## [2026-11-11] tạo năng lực | Đóng gói quy trình và biểu mẫu tạo skill chuyên biệt cho diện Giải thể

Đã tạo skill chuyên biệt `ho-so-kiem-tra-giai-the` (Hồ sơ kiểm tra giải thể):
- Đóng gói trọn vẹn quy trình lập, bóc tách và chuẩn hóa bộ hồ sơ giải thể 22 văn bản.
- Biểu mẫu chuẩn trong `templates/` và `skills/ho-so-kiem-tra-giai-the/templates/`: `MAUTONG HOP GIAI THE.docx`, `MAU DE XUAT HOAN GIAI THE.docx`, `MAU MS-03-DCS.docx`, `MAU QD XPHC GIAI THE.docx`.
- Tài liệu nghiệp vụ tham chiếu trong `references/`: danh mục 22 văn bản, nghiệp vụ Mục III NSNN & hoàn bù trừ NĐ 181/2025, nền tảng pháp lý và căn cứ giải thể.
- Đồng bộ đệ quy toàn bộ skill sang `.claude/skills/ho-so-kiem-tra-giai-the/`.
Liên quan: [[Quy trình lập hồ sơ kiểm tra thuế đối với doanh nghiệp giải thể]], [[Hồ sơ đề xuất hoàn thuế kiêm bù trừ thu ngân sách đối với doanh nghiệp giải thể]], [[Chuẩn định dạng TONG HOP hồ sơ kiểm tra thuế]]

## [2026-11-11] update | Chuẩn hóa khoảng cách chữ ký 5 dòng, bỏ năm dự thảo QĐ và đóng gói cập nhật skill

Đã rà soát và đóng gói toàn bộ thay đổi thể thức trong hội thoại vào hệ thống phôi và skill:
- **Áp dụng trực tiếp vào phôi TONG HOP** (`templates/TONG HOP CHUAN DINH DANG.docx`, `templates/MAUTONG HOP.docx`, `templates/MAU TONG HOP HO SO DI.docx`, `output/3603859012_HO-SO-VE_son-tien-hai/TONG HOP.docx`): Toàn bộ các ô chữ ký chức danh và họ tên người ký được căn chỉnh mặc định đúng 5 dòng trống (trường hợp tràn trang tự động giảm 4 dòng theo NĐ 30 để khống chế 1 trang in).
- **Đóng gói cập nhật vào 5 skill**: `ho-so-kiem-tra-ke-hoach`, `soan-thao-van-ban`, `dien-bien-ban-tu-bang-tinh`, `lap-ho-so-kiem-tra`, `ho-so-kiem-tra-giai-the` (đồng bộ sang `.claude/skills/`).
- **Nội dung chuẩn hóa**: (1) QĐ 1377/QĐ-CT chốt thống nhất năm 2025; (2) Dự thảo QĐ để trống toàn bộ ngày tháng năm `Trấn Biên, ngày       tháng       năm      `; (3) Ký hiệu dự thảo để khoảng trống `Số:              /QĐ-TCS3`, tuyệt đối không dùng mã tạm `DU_THAO-TRINH-KY`; (4) Không cập nhật lại số/ngày khi QĐ đã ban hành; (5) Khoảng cách chữ ký mặc định 5 dòng trong cùng 1 ô, co 4 dòng khi tràn trang; (6) BBXN Trưởng đoàn đánh số TT 1 và 2 cho QLDN và Bán hàng, chuẩn hóa giải trình.
Liên quan: [[Chuẩn định dạng TONG HOP hồ sơ kiểm tra thuế]], [[Kỹ thuật căn chỉnh chữ ký trong ô và chống rớt trang theo Nghị định 30]]

