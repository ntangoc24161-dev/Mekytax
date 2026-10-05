# Cấu trúc và Quy tắc dữ liệu Sheet Phan cong

Tài liệu tham chiếu chi tiết cấu trúc các cột và quy tắc nhập liệu cho sheet `Phan cong` tại file Excel:
`I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\File theo doi hs CUNG CAP CA-01.04.2025.xls`

## Bảng ánh xạ các cột (Dòng tiêu đề 5, dữ liệu từ dòng 6)

| Ký hiệu cột | Vị trí (1-based) | Index (0-based) | Tên tiêu đề trên Excel | Kiểu dữ liệu | Quy tắc điền & Lưu ý |
| :--- | :---: | :---: | :--- | :--- | :--- |
| **A** | 1 | 0 | STT | Số nguyên / Float | Số thứ tự hồ sơ. Nếu một công văn có nhiều doanh nghiệp thì tất cả các dòng của công văn đó dùng chung 1 số STT. |
| **B** | 2 | 1 | SỐ CV | Chuỗi văn bản | Số công văn yêu cầu của cơ quan Công an (ví dụ: `31819/YC-CSKT`). Chỉ điền tại dòng đầu tiên của cụm STT; các dòng sau để trống. |
| **C** | 3 | 2 | NGÀY CV | Ngày / Serial Date | Ngày ban hành ghi trên công văn yêu cầu. Chỉ điền tại dòng đầu tiên của cụm STT. |
| **D** | 4 | 3 | CƠ QUAN BAN HÀNH | Chuỗi văn bản | Tên cơ quan điều tra ban hành văn bản (ví dụ: `CÔNG AN TP ĐỒNG NAI`). Chỉ điền tại dòng đầu tiên. |
| **E** | 5 | 4 | Số CV Thuế TP | Chuỗi / Số | Số công văn đến theo dấu tiếp nhận của Cơ quan Thuế (ví dụ: `12181`). Chỉ điền tại dòng đầu tiên của cụm STT. |
| **F** | 6 | 5 | NGÀY CV ĐẾN | Ngày / Serial Date | Ngày ghi trên dấu công văn đến (ví dụ: `26/08/2026`). Chỉ điền tại dòng đầu tiên của cụm STT. |
| **G** | 7 | 6 | CBT ĐÃ CC | Chuỗi văn bản | Cán bộ thuế đã thực hiện cung cấp hồ sơ. Thường để trống lúc tiếp nhận phân công. |
| **H** | 8 | 7 | Mã số thuế CTY CẦN CUNG CẤP HS | Chuỗi 10 hoặc 13 số | Mã số thuế của người nộp thuế được yêu cầu cung cấp hồ sơ. Tuyệt đối không để định dạng số khoa học (scientific notation). |
| **I** | 9 | 8 | Tên NNT | Chuỗi văn bản | Tên đầy đủ của doanh nghiệp / người nộp thuế theo đăng ký kinh doanh. |
| **J** | 10 | 9 | Địa chỉ | Chuỗi văn bản | Địa chỉ trụ sở doanh nghiệp rút gọn về cấp Phường/Xã để phân Tổ KT. Nếu thiếu địa chỉ: tra cứu MST trên `masothue.com`; không có thì để trống, tuyệt đối không bịa đặt. |
| **K** | 11 | 10 | Tổ KT | Chuỗi văn bản | Tổ kiểm tra phụ trách thụ lý (ví dụ: `KT1`, `KT2`, `KT3`). Phân bổ theo địa bàn Phường/Xã. |
| **L** | 12 | 11 | CÁN BỘ XỬ LÝ | Chuỗi văn bản | Tên cán bộ thuế được phân công xử lý hồ sơ. Đối với Tổ KT1: tra cứu MST trong sheet `KTR1` của file `DB NNT 03.06.2026.xlsx` để lấy tên CBQL. |
| **M** | 13 | 12 | NGÀY GIAO CBXL | Ngày / Serial Date | Ngày bàn giao hồ sơ cho cán bộ xử lý: Ghi nhận ngày phân công là **ngày thực hiện lệnh**. |
| **N** | 14 | 13 | Tình trạng xử lý | Chuỗi văn bản | Ký hiệu trạng thái theo dõi hồ sơ (ví dụ: `R` - đã phân công / đang rà soát). |
| **O** | 15 | 14 | Ghi chú | Chuỗi văn bản | Các lưu ý phối hợp nghiệp vụ hoặc ghi chú phát sinh. |
| **P** | 16 | 15 | Ngày giấy phép | Ngày / Chuỗi | Ngày cấp giấy phép đăng ký kinh doanh (nếu có). |
| **Q** | 17 | 16 | Đại diện thep pháp luật | Chuỗi văn bản | Họ tên người đại diện theo pháp luật của doanh nghiệp. |
| **R** | 18 | 17 | *(Nội dung yêu cầu)* | Chuỗi văn bản | Nội dung chi tiết cơ quan Công an yêu cầu: thời kỳ hồ sơ (từ năm... đến năm...), danh mục tài liệu (BCTC, VAT, TNDN, TNCN), số lượng hóa đơn, số hóa đơn mua vào/bán ra cần kiểm tra. |

## Quy tắc xử lý công văn nhiều doanh nghiệp
1. **Dòng đầu tiên của vụ việc (Primary Row)**:
   - Điền đầy đủ: `STT`, `SỐ CV`, `NGÀY CV`, `CƠ QUAN BAN HÀNH`, `Số CV Thuế TP`, `NGÀY CV ĐẾN`.
   - Điền thông tin doanh nghiệp số 1: `Mã số thuế`, `Tên NNT`, `Địa chỉ`, `Tổ KT`, `CÁN BỘ XỬ LÝ`, `NGÀY GIAO CBXL` (ngày thực hiện lệnh), `Tình trạng xử lý` ('R'), `Nội dung yêu cầu` (Cột R).
2. **Các dòng phụ kế tiếp (Subsequent Rows)**:
   - Giữ nguyên giá trị `STT` giống dòng đầu tiên.
   - Để trống: `SỐ CV`, `NGÀY CV`, `CƠ QUAN BAN HÀNH`, `Số CV Thuế TP`, `NGÀY CV ĐẾN`.
   - Điền riêng rẽ cho từng doanh nghiệp: `Mã số thuế`, `Tên NNT`, `Địa chỉ`, `Tổ KT`, `CÁN BỘ XỬ LÝ`, `NGÀY GIAO CBXL`, `Nội dung yêu cầu` (Cột R).
