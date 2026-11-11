---
name: Điền biên bản từ Bảng tính
description: Điền BBXN/BBKT từ Bảng tính (DTCHEP), đồng bộ QĐ riêng, chèn ngày công bố/báo cáo và dọn bak/backup.
group: Tài chính
updated: '2026-10-29'
learned_updated: '2026-10-29'
---
# Điền biên bản từ Bảng tính

Hoàn thiện `TONG HOP.docx` sau khi đã có số trên Bảng tính: điền Biên bản xác nhận (BBXN) và Biên bản kiểm tra (BBKT) từ sheet `DTCHEP`/`DCHEP`, đồng bộ file QĐ riêng với khối trong TONG HOP, chèn ngày công bố / kiểm tra / báo cáo, rồi dọn file bak/backup thừa trong thư mục vụ.

Skill đứng **sau** `nhap-bang-tinh-tu-bao-cao-hdkd` / `boc-tach-bang-tinh-phan-tich` và **trước hoặc song song** bước chuẩn hóa/xuất của `ho-so-kiem-tra-ke-hoach`. Không dựng Word từ trang trắng; chỉ điền trên phôi / `TONG HOP` đang có.

## Khi nào dùng

- "điền biên bản từ Bảng tính", "đổ DTCHEP vào BBXN/BBKT", "cập nhật biên bản xác nhận / biên bản kiểm tra".
- "làm hết tất cả các biên bản và hồ sơ về".
- "soạn biên bản và hồ sơ về theo mẫu <vụ khác>" (vd theo mẫu DNTN Thúy Hồng).
- "đồng bộ QĐ riêng với TONG HOP", "xuất lại QĐ sau khi canh tay".
- "chèn ngày công bố / ngày báo cáo vào hồ sơ".
- "dọn bak", "xóa backup Excel/Word trong thư mục BIEN BAN".

## Chuẩn bị

- Thư mục vụ `...\0. BIEN BAN\` có `TONG HOP.docx`, `Bang tinh*.xlsx` (ưu tiên bản khớp thời kỳ), và (nếu đã xuất) `QD Kiem tra....docx`, `QDgiam sat....docx`.
- **Bảng tính và TONG HOP thường nằm trên Google Drive** `ntangoc2416.1@gmail.com`, không có trong vault. Không thấy trong vault thì KHÔNG bảo user tải tay: đọc qua Composio (`GOOGLEDRIVE_FIND_FILE` / `GOOGLEDRIVE_GET_FILE_METADATA` / `GOOGLEDRIVE_DOWNLOAD_FILE`, truyền `account` của user), xem memory `doc-file-google-drive-qua-composio`.
- **So ngày sửa Drive với vault** cho `TONG HOP.docx` và `Bang tinh.xlsx`: bản Drive mới hơn thì lấy bản Drive làm gốc (chứa phần user canh tay). Mỗi lượt chốt lại thì tải lại Bảng tính vì user có thể vừa sửa trên Drive (vd ngày lập BBKT).
- Chuẩn mắt: `templates/TONG HOP CHUAN DINH DANG.docx` (bản tay VPT).
- Phân công đoàn + nội dung kiểm tra trên sổ JAVIS PRO (TD: QLDN + bán hàng; TV: tài chính - hoặc theo vụ đang làm).
- Ngày vụ: công bố, bắt đầu, kết thúc, ký BBKT, báo cáo (từ sổ `Danh sach kiem tra quyet toan pro.xlsm`, sheet `01_DA_BAN_HANH`, hoặc user chốt).
- QĐ chưa có số thì để trống ô số và ngày tháng năm, không chèn mã tạm. Riêng **QĐ xử phạt VPHC**: nếu chưa có ký hiệu thì ghi `Số: /QĐ-XPHC`.
- Ngày trên header **QĐKT và QĐGS** để trống (`Trấn Biên, ngày tháng năm ...`), kể cả khi sổ đã có ngày ban hành.
- Mục III BBKT cột "Số đã nộp": lấy từ sheet **`QHNS 970`** của Bảng tính (không lấy sheet `GNT`). Hỏi user Mục III gồm sắc thuế nào: có vụ user chỉ cho thể hiện Thuế TNDN (Sơn Tiền Hải 29/10/2026).
- **Kế hoạch giám sát**: ngày lập = ngày ban hành QĐKT trên sổ JAVIS PRO.
- Sau khi điền xong: ghi đè bộ file lên đúng thư mục vụ trên Drive, thay bản cũ (xem Bẫy về upload Composio).
- Backup `.bak-*` trước khi sửa XML/run lớn; sau khi xong mới dọn (bước 4).

Chi tiết cấu trúc từng năm BBXN/BBKT: `references/cau-truc-bbxn-bbkt.md`.

## Quy trình

### 0. Kiểm phôi có phải bản clone của vụ khác không

Khi phôi `TONG HOP.docx` của vụ mới được sao từ vụ mẫu (vd DNTN Thúy Hồng), phôi thường còn nguyên dữ liệu vụ cũ.

1. Lập danh sách chuỗi đặc trưng của vụ mẫu: tên DN, MST, người đại diện + CCCD, loại hình (DNTN/Công ty), ngành nghề, khung năm, thời hạn kiểm tra, ngày trên Báo cáo, số liệu.
2. Đổi khung năm theo thời kỳ vụ mới trên **cả** BBXN thành viên, BBXN trưởng đoàn và BBKT (thêm/bớt khối năm).
3. Sau khi điền, quét lại toàn bộ danh sách chuỗi vụ mẫu trong document.xml; chỉ coi là xong khi còn 0 chuỗi.
4. Quét thêm **văn bản đặc thù ngành của vụ mẫu** mà danh sách chuỗi không bắt được (vd "Cam kết hoạt động kinh doanh vàng" của Thúy Hồng sót trong hồ sơ Sơn Tiền Hải - bán buôn VLXD). Liệt kê các văn bản không hợp ngành vụ mới và hỏi user bỏ hay sửa.
5. Quét căn cứ pháp lý cũ còn sót trong từng văn bản (vd "Luật Quản lý thuế ngày 13/6/2019" trong BB công bố QĐ) và đổi cho khớp căn cứ đang dùng ở các văn bản khác (Luật 108/2025/QH15).
6. Căn cứ pháp lý/kết luận của vụ mẫu có thể không áp được cho vụ mới (vd hủy hóa đơn theo Thông tư 39/2014 không áp cho DN thành lập từ 2022, phải dùng hóa đơn điện tử theo Nghị định 123/2020): rà và hỏi user phần thực tế (đã ngừng sử dụng HĐĐT, đã thanh lý TSCĐ, còn hàng tồn kho không).

### 1. Điền BBXN và BBKT từ Bảng tính

1. Đọc sheet **`DTCHEP`** (còn gọi `DCHEP`) bằng `openpyxl` `data_only=True`: số báo cáo / số kiểm tra / chênh lệch theo năm và theo chỉ tiêu chi phí.
2. Khớp chỉ tiêu với phân công đoàn:
   - Thành viên (BBXN TV): Chi phí tài chính.
   - Trưởng đoàn (BBXN TD): Chi phí QLDN và Chi phí bán hàng (mỗi ý một dòng).
3. Với **mỗi năm** trong thời kỳ kiểm tra, trên BBXN giữ đúng khối phôi:
   - Dòng năm **đậm không nghiêng** + tab + `Đơn vị tính: Đồng` nghiêng cùng dòng (tab right ~8931); không tách ĐVT ra dòng riêng.
   - Bảng 5 cột phôi: `TT | CHỈ TIÊU | SỐ BÁO CÁO | SỐ KIỂM TRA | CHÊNH LỆCH`.
   - Có chênh: `*/` giải trình + căn cứ pháp lý theo năm (xem dưới).
   - Không chênh: chỉ `*/ Không chênh lệch.` - không giữ khung giải trình của năm có chênh.
4. **BBXN không ghi** TNTT tăng / thuế TNDN truy thu (phần đó thuộc BBKT).
5. Trên **BBKT**, năm có chênh: khối Giải trình → diễn giải TNTT → `*/` nguyên nhân → `=>Thuế TNDN truy thu...` (+ giảm lỗ nếu có). Năm không chênh: chỉ `*/ Không chênh lệch.`
6. Câu loại **chi phí tài chính** thống nhất: `Công ty hạch toán chi phí lãi vay không tương ứng doanh thu` (không dùng cụm cũ "không đúng quy định", không EBITDA). Giá vốn hàng bán: `Công ty hạch toán giá vốn hàng bán không tương ứng doanh thu` (user chấp nhận ở vụ Sơn Tiền Hải).
7. Căn cứ pháp lý theo năm:
   - Trước 2025 (vd 2021-2022): `Điều 4 Thông tư 96/2015/TT-BTC ngày 22/06/2015`.
   - Từ 2025: `Điều 10 Nghị định số 320/2025/NĐ-CP` (+ Thông tư hướng dẫn nếu có trong hồ sơ).
8. Giá trị bằng **0** ghi số `0` (hoặc `0 đồng`), không thay bằng gạch ngang `-`.
9. Giữ 100% layout phôi (tab, bảng, floating Mẫu số). Chỉ điền số và lời nghiệp vụ; không normalize hàng loạt giãn dòng/font toàn file.
10. Bảng tính không ghi nguyên nhân chênh lệch (vd giảm giá vốn) thì viết câu tạm và **hỏi user** căn cứ thực tế của đoàn, không tự chốt.
11. Giảm lỗ có thể lệch giữa sổ và cách ghi của vụ mẫu (sổ ghi 0, vụ mẫu ghi lỗ báo cáo bị xóa): hỏi user chốt số.
12. Cơ quan cấp ĐKKD ghi thống nhất ở BBKT, QĐ XPHC và công văn xin quyết toán; thông tin thiếu (vd kế toán trưởng) để trống, không giữ của vụ mẫu.

### 2. Đồng bộ QĐ riêng sau canh tay

1. Sau khi user (hoặc lượt trước) đã canh `TONG HOP.docx`, **không** lấy bak-r cũ làm chuẩn mắt.
2. Sao khối **Quyết định kiểm tra** và **Quyết định giám sát** (và QĐ XPHC nếu có) trong TONG HOP ra file riêng:
   - `QD Kiem tra tai cty <Ten>.docx`
   - `QDgiam sat tai cty <Ten>.docx`
3. Khớp nội dung + format với khối tương ứng trong TONG HOP (Điều, thụt 1 cm, gạch VML QĐKT canh giữa, QĐGS căn cứ nghiêng, Điều 4 đậm). Mỗi file QĐ gọn **1 trang A4**.
4. Ngày tháng trên header QĐ để trống dạng `..., ngày .... tháng .... năm ....` nếu chưa ban hành / đang trình ký; có số QĐ thật thì điền số, vẫn để trống ngày trên QĐ nếu user yêu cầu.
5. Có thể dùng `scripts/xuat_ho_so_tron_goi.py` hoặc script đồng bộ QĐ của vụ; sau khi chạy phải **đối chiếu lại** với TONG HOP vừa canh tay (tránh lệch giờ file như vụ VPT).

### 3. Chèn ngày công bố / kiểm tra / báo cáo

Lấy mốc từ sổ JAVIS PRO hoặc từ user, rồi điền đúng chỗ trong TONG HOP (và file liên quan nếu đã tách):

| Mốc | Điền vào |
| :--- | :--- |
| Ngày công bố / bắt đầu kiểm tra | BBCB, BB phân công, đoạn mở BBKT, Báo cáo KQ, Báo cáo giám sát |
| Ngày kết thúc | BB công khai dự thảo BBKT, BBXN thành viên, BBXN trưởng đoàn |
| Ngày ký BBKT | BBKT, "tính đến ngày" Mục III, căn cứ BBKT trong QĐ XPHC và tờ trình, mốc tính chậm nộp |
| Ngày báo cáo | Báo cáo kết quả, Báo cáo giám sát, Biên bản làm việc 04-TTKT |

Không tự bịa ngày. Thiếu mốc thì ghi chỗ trống theo phôi (`..../..../<năm>`). Sổ không có cột ngày BBXN thì tạm lấy ngày kết thúc kiểm tra và báo user. Tiền chậm nộp tính theo ngày ký BBKT: nếu Bảng tính dùng ngày tạm thì báo user và tính lại khi có ngày thật (đếm ngày từ ngày sau hạn nộp, vd 01/04 năm sau cho TNDN năm trước). Sau khi chèn, rà lại checkbox BCGS: **giữ** trạng thái đã tích; mục 5 thường tích "Không có phản ánh", để trống "Không vi phạm" / "Có vi phạm" trừ khi user chốt khác.

### 4. Dọn bak / backup

Chỉ chạy khi user bảo dọn hoặc sau khi đã chốt file live và có bằng chứng lưu thành công:

- Xóa trong thư mục `0. BIEN BAN\`: `*.bak*`, `TONG HOP.bak-*`, `*_backup*.xlsx`, file `.md` tạm không thuộc hồ sơ.
- **Giữ** file live: `TONG HOP.docx`, QĐ riêng, PDF đã ký (nếu có), `Bang tinh*.xlsx` đang dùng, `hoso.zip` mới nhất.
- Liệt kê lại thư mục bằng tool sau khi xóa; chỉ xác nhận đã xóa khi lệnh trả thành công.

## Bẫy

- Không thấy Bảng tính trong vault rồi kết luận "không đọc được Drive" và bắt user tải tay: Drive đã đấu qua Composio (lỗi đã xảy ra 26/10/2026).
- Lấy `TONG HOP.docx` trong vault làm gốc khi bản Drive mới hơn: mất phần user canh tay.
- **Upload Drive qua Composio** (`GOOGLEDRIVE_UPLOAD_UPDATE_FILE`) không nhận đường dẫn file trên máy: lỗi 403 "file key does not belong to this project" (29/10/2026). Không báo là đã ghi đè Drive; nói rõ bản Drive vẫn là bản cũ và nhờ user tải file từ vault lên, hoặc tìm đường khác.
- Phôi TONG HOP là bản clone của vụ mẫu: sót tên, CCCD, loại hình, ngành nghề, khung năm, thời hạn, ngày Báo cáo của vụ cũ (gặp ở vụ Sơn Tiền Hải clone từ Thúy Hồng, 29/10/2026). Luôn chạy bước 0, kể cả quét văn bản đặc thù ngành và căn cứ luật cũ.
- Sheet `QHNS 970` có thể lỗi (GTGT để 0 dù có số phải nộp, cột còn phải nộp khóa về 0): để `......` ở cột chưa chắc và hỏi user, không tự chốt.
- Sheet phụ (vd `BANG KE KO HOP LE`) báo `#REF!` hoặc dòng năm không khớp thời kỳ: không dùng làm căn cứ.
- Phạt khai sai cho kỳ quyết toán giải thể chưa đến hạn khai: hỏi user có áp hay chỉ truy thu (vụ Sơn Tiền Hải user cho áp).
- Máy không có LibreOffice nên không đếm được số trang: để `.... trang` ở đoạn kết BBXN/BBKT và báo user điền tay, không đoán số.
- Điền BBXN nhầm sang kiểu BBKT (thêm TNTT/truy thu) làm lệch phôi.
- Tách ĐVT sang dòng riêng hoặc bỏ tab ~8931.
- Đồng bộ QĐ từ bản TONG cũ hơn bản tay vừa canh.
- Normalize NĐ 30 hàng loạt lên biên bản/kế hoạch/báo cáo (phá floating, giãn dòng).
- Thay checkbox BCGS mà reset hết ô đã tích.
- Xóa nhầm `Bang tinh 2021-2025.xlsx` / PDF đã ký khi dọn backup.

## Kiểm chứng

- File gốc dùng để điền là bản mới nhất giữa Drive và vault.
- Quét chuỗi đặc trưng của vụ mẫu còn 0 kết quả; không còn văn bản đặc thù ngành của vụ mẫu chưa được user quyết.
- Mỗi năm trong thời kỳ có đủ khối năm + bảng (nếu có số) + `*/` đúng loại (giải trình hoặc Không chênh lệch).
- BBXN không có dòng truy thu TNDN; BBKT năm chênh có `=>Thuế TNDN truy thu`.
- Tiền chậm nộp tính đúng theo ngày ký BBKT đã chốt.
- Hai file QĐ riêng khớp khối trong TONG HOP và mỗi file 1 trang.
- Ngày công bố/kết thúc/ký BBKT/báo cáo khớp sổ (hoặc đúng số user chốt).
- Chỉ báo đã ghi đè Drive khi tool upload trả thành công.
- Thư mục vụ không còn `*.bak*` / `*_backup*` sau bước dọn; còn đủ file live.

## Nối skill

1. `nhap-bang-tinh-tu-bao-cao-hdkd` - điền KQKDSXKD từ PTDN02/BCTC.
2. `boc-tach-bang-tinh-phan-tich` - tỷ lệ + Mục III/IV DI-01.
3. **`dien-bien-ban-tu-bang-tinh`** (skill này) - BBXN/BBKT + QĐ + ngày + dọn.
4. `ho-so-kiem-tra-ke-hoach` - font/checkbox/lint/xuất trọn gói / NĐ 30 cho hành chính.

## Lịch sử (tự học)
- [2026-10-26] User chốt 5 điểm khi duyệt dự thảo Tâm Hoàng Phát: để trống ngày QĐKT/QĐGS; QĐ XPHC thêm `/QĐ-XPHC` khi chưa có số; Mục III lấy `QHNS 970`; ngày Kế hoạch giám sát = ngày ban hành QĐKT trên sổ; điền xong thì ghi đè Drive. Đã đưa vào Chuẩn bị và fact `quy-uoc-soan-bien-ban-ho-so-ve`.
- [2026-10-26] Phần Chuẩn bị chỉ coi thư mục vụ nằm trên máy, nên ngày 26/10/2026 Javis không thấy Bảng tính trong vault, nói sai là không đọc được Drive và bắt user tải tay qua 3 lượt. Sau đó mới phát hiện TONG HOP trên Drive (22/10) mới hơn vault (16/10) và user chốt lấy bản Drive. Thêm bước đọc Drive qua Compos
- [2026-10-29] Vụ Sơn Tiền Hải: phôi TONG HOP là bản clone Thúy Hồng còn 22 chuỗi vụ cũ (khung năm 2017-2024, CCCD, loại hình DNTN, ngành nghề, thời hạn 10 ngày, ngày Báo cáo). Thêm bước 0 quét chuỗi vụ mẫu, và bẫy QHNS 970 lỗi, sheet phụ #REF!, phạt khai sai kỳ giải thể.
- [2026-10-29] Ngày 29/10/2026 vụ Sơn Tiền Hải: phôi TONG HOP clone từ Thúy Hồng còn nguyên dữ liệu vụ cũ (22 chuỗi), skill cũ không có bước kiểm/quét chuỗi vụ mẫu. Skill cũ còn bảo lấy Số đã nộp từ QHNS 970 mà không cảnh báo sheet này có thể lỗi (GTGT để 0, cột khóa 0). Thêm bước 0 và các mục Bẫy tương ứng, giữ n
- [2026-10-26] User chốt 5 điểm khi duyệt dự thảo Tâm Hoàng Phát: để trống ngày QĐKT/QĐGS; QĐ XPHC thêm `/QĐ-XPHC` khi chưa có số; Mục III lấy `QHNS 970`; ngày Kế hoạch giám sát = ngày ban hành QĐKT trên sổ; điền xong thì ghi đè Drive.
- [2026-10-26] Javis không thấy Bảng tính trong vault, nói sai là không đọc được Drive. Thêm bước đọc Drive qua Composio và so ngày Drive/vault.
- [2026-10-29] Vụ Sơn Tiền Hải: phôi clone Thúy Hồng còn 22 chuỗi vụ cũ. Thêm bước 0 quét chuỗi vụ mẫu, bẫy QHNS 970 lỗi, sheet phụ #REF!, phạt khai sai kỳ giải thể.
- [2026-10-29] Lượt chốt Sơn Tiền Hải: quét 22 chuỗi vẫn sót văn bản "Cam kết kinh doanh vàng" và căn cứ Luật QLT 2019 trong BB công bố; upload Drive qua Composio lỗi 403 dù skill bảo ghi đè Drive. Thêm bước 0.4-0.5, bảng mốc ngày chi tiết, bẫy upload và số trang.
- [2026-10-29] Batch 29/10/2026 vụ Sơn Tiền Hải cho thấy skill cũ thiếu: (1) bước 0 chỉ quét danh sách chuỗi nên sót văn bản 'Cam kết hoạt động kinh doanh vàng' của Thúy Hồng và căn cứ 'Luật Quản lý thuế ngày 13/6/2019' trong BB công bố; (2) skill bảo ghi đè Drive bằng GOOGLEDRIVE_UPLOAD_UPDATE_FILE nhưng thực tế
