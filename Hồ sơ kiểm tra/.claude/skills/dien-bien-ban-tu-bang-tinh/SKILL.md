---
name: Điền biên bản từ Bảng tính
description: Điền BBXN/BBKT từ DTCHEP, đối chiếu sổ khi lệch TINH PHAT, đồng bộ QĐ/thu hồi hoàn, chèn ngày, dọn bak.
group: Tài chính
updated: '2026-11-09'
learned_updated: '2026-11-09'
---
# Điền biên bản từ Bảng tính

Hoàn thiện `TONG HOP.docx` sau khi đã có số trên Bảng tính: điền Biên bản xác nhận (BBXN) và Biên bản kiểm tra (BBKT) từ sheet `DTCHEP`/`DCHEP`, đối chiếu sổ khi lệch sheet tính phạt, đồng bộ file QĐ riêng với khối trong TONG HOP, chèn ngày công bố / kiểm tra / báo cáo, rồi dọn file bak/backup thừa trong thư mục vụ.

Skill đứng **sau** `nhap-bang-tinh-tu-bao-cao-hdkd` / `boc-tach-bang-tinh-phan-tich` và **trước hoặc song song** bước chuẩn hóa/xuất của `ho-so-kiem-tra-ke-hoach`. Không dựng Word từ trang trắng; chỉ điền trên phôi / `TONG HOP` đang có.

## Khi nào dùng

- "điền biên bản từ Bảng tính", "đổ DTCHEP vào BBXN/BBKT", "cập nhật biên bản xác nhận / biên bản kiểm tra".
- "cập nhật TONG HOP theo Bảng tính và sổ", "hỏi chỗ lệch trước khi sửa".
- "đồng bộ QĐ riêng với TONG HOP", "xuất lại QĐ sau khi canh tay".
- "chèn ngày công bố / ngày báo cáo vào hồ sơ".
- "dọn bak", "xóa backup Excel/Word trong thư mục BIEN BAN".

## Chuẩn bị

- Thư mục vụ `...\0. BIEN BAN\` có `TONG HOP.docx`, `Bang tinh*.xlsx` (ưu tiên bản khớp thời kỳ), và (nếu đã xuất) `QD Kiem tra....docx`, `QDgiam sat....docx`, QĐ XPHC / QĐ thu hồi hoàn.
- Chuẩn mắt: `templates/TONG HOP CHUAN DINH DANG.docx` (bản tay VPT).
- Phân công đoàn + nội dung kiểm tra trên sổ JAVIS PRO (TD: QLDN + bán hàng; TV: tài chính - hoặc theo vụ đang làm).
- Ngày vụ: công bố, bắt đầu, kết thúc, ký BBKT, báo cáo (từ sổ hoặc user chốt).
- Ba nguồn số: sheet `DTCHEP`, sheet `TINH PHAT + CHAM NOP (C T)`, sổ `Danh sach kiem tra quyet toan pro.xlsm`. Lệch tổng truy thu / chậm nộp / phạt / thu hồi hoàn thì hỏi user trước khi ghi.
- Backup `.bak-*` trước khi sửa XML/run lớn; sau khi xong mới dọn (bước 4).

Chi tiết cấu trúc từng năm BBXN/BBKT: `references/cau-truc-bbxn-bbkt.md`.

## Quy trình

### 1. Điền BBXN và BBKT từ Bảng tính

1. Đối chiếu `TONG HOP` đang in với `DTCHEP` và sổ. Sheet `TINH PHAT` cộng trùng chậm nộp GTGT (dòng sắc thuế đã gom rồi dòng thu hồi hoàn cộng lại) thì lấy tổng trên sổ. Chỗ lệch khác: hỏi user, rồi mới ghi.
2. Đọc sheet **`DTCHEP`** (còn gọi `DCHEP`) bằng `openpyxl` `data_only=True`: số báo cáo / số kiểm tra / chênh lệch theo năm và theo chỉ tiêu chi phí.
3. Khớp chỉ tiêu với phân công đoàn:
   - Thành viên (BBXN TV): mặc định Chi phí tài chính (1 chỉ tiêu/năm). Vụ đã bổ sung khối GTGT (đầu vào được khấu trừ, đầu ra) thì **giữ nguyên số chỉ tiêu đó**, không thêm "Thuế GTGT đã được hoàn".
   - Trưởng đoàn (BBXN TD): mặc định Chi phí QLDN và Chi phí bán hàng - **2 chỉ tiêu/năm trong cùng một bảng**, đánh **TT = 1** cho QLDN và **TT = 2** cho bán hàng (không tách hai bảng riêng theo chỉ tiêu). Vụ có thêm giá vốn hàng bán thì giữ đúng số chỉ tiêu đã có trên phôi/TONG HOP.
   - Tên chỉ tiêu trong cột CHỈ TIÊU (và trong mục I/II nói chung) **không gạch đầu dòng `-`**, chỉ ghi thẳng tên nội dung (ví dụ `Chi phí quản lý doanh nghiệp`, không phải `- Chi phí quản lý doanh nghiệp`).
4. Với **mỗi năm** trong thời kỳ kiểm tra, trên BBXN giữ đúng khối phôi:
   - Dòng năm **đậm không nghiêng** + tab + `Đơn vị tính: Đồng` nghiêng cùng dòng (tab right ~8931); không tách ĐVT ra dòng riêng.
   - Bảng phôi: `TT | CHỈ TIÊU | SỐ BÁO CÁO | SỐ KIỂM TRA | CHÊNH LỆCH` (TT là số thứ tự chỉ tiêu trong năm đó, 1/2 như trên - không phải số thứ tự năm).
   - Có chênh: thêm dòng tiêu đề **`Giải trình số liệu chênh lệch qua kiểm tra:`**, xuống hàng mới bắt đầu bằng `*/` rồi nội dung giải trình + căn cứ pháp lý theo năm (xem dưới).
   - Không chênh: chỉ một dòng `*/ Không chênh lệch.` viết thường - không có dòng tiêu đề "Giải trình..." phía trên, không giữ khung giải trình của năm có chênh.
5. **BBXN không ghi** TNTT tăng / thuế TNDN truy thu (phần đó thuộc BBKT).
6. Trên **BBKT**, năm có chênh: khối Giải trình → diễn giải TNTT → `*/` nguyên nhân → `=>Thuế TNDN truy thu...` (+ giảm lỗ nếu có). Năm không chênh: chỉ `*/ Không chênh lệch.` Năm có miễn giảm TNDN (ND398): chèn **2 dòng** vào bảng năm đó (miễn giảm; thuế TNDN còn phải nộp); dòng truy thu lấy số **sau miễn giảm**; tổng nhiều năm lấy ô sổ khi cộng tay lệch 1 đồng. Phần GTGT user đã bổ sung trên BBKT giữ nguyên.
7. Câu loại **chi phí tài chính** thống nhất: `Công ty hạch toán chi phí lãi vay không tương ứng doanh thu` (không dùng cụm cũ "không đúng quy định", không EBITDA).
8. Căn cứ pháp lý theo năm:
   - Trước 2025 (vd 2021-2022): `Điều 4 Thông tư 96/2015/TT-BTC ngày 22/06/2015`.
   - Từ 2025: `Điều 10 Nghị định số 320/2025/NĐ-CP` (+ Thông tư hướng dẫn nếu có trong hồ sơ).
9. Giá trị bằng **0** ghi số `0` (hoặc `0 đồng`), không thay bằng gạch ngang `-`.
10. Giữ 100% layout phôi (tab, bảng, floating Mẫu số). Chỉ điền số và lời nghiệp vụ; không normalize hàng loạt giãn dòng/font toàn file.

### 2. Đồng bộ QĐ riêng sau canh tay

1. Sau khi user (hoặc lượt trước) đã canh `TONG HOP.docx`, **không** lấy bak-r cũ làm chuẩn mắt.
2. Sao khối **Quyết định kiểm tra** và **Quyết định giám sát** trong TONG HOP ra file riêng:
   - `QD Kiem tra tai cty <Ten>.docx`
   - `QDgiam sat tai cty <Ten>.docx`
3. Khớp nội dung + format với khối tương ứng trong TONG HOP (Điều, thụt 1 cm, gạch VML QĐKT canh giữa, QĐGS căn cứ nghiêng, Điều 4 đậm). Mỗi file QĐ gọn **1 trang A4**.
4. Ngày tháng trên header QĐ để trống dạng `..., ngày .... tháng .... năm ....` nếu chưa ban hành / đang trình ký; có số QĐ thật thì điền số, vẫn để trống ngày trên QĐ nếu user yêu cầu.
5. QĐ XPHC và QĐ thu hồi hoàn: cập nhật tiền theo sổ (truy thu, thu hồi hoàn, chậm nộp, phạt). QĐ thu hồi hoàn điền số tài khoản và tên ngân hàng của DN nếu hồ sơ/sổ đã có; thiếu thì để trống và hỏi.
6. Có thể dùng `scripts/xuat_ho_so_tron_goi.py` hoặc script đồng bộ QĐ của vụ; sau khi chạy phải **đối chiếu lại** với TONG HOP vừa canh tay (tránh lệch giờ file như vụ VPT).

### 3. Chèn ngày công bố / kiểm tra / báo cáo

Lấy mốc từ sổ JAVIS PRO hoặc từ user, rồi điền đúng chỗ trong TONG HOP (và file liên quan nếu đã tách):

| Mốc | Điền vào |
| :--- | :--- |
| Ngày công bố / bắt đầu kiểm tra | BBCB, đoạn mở BBKT, chỗ "công bố QĐKT" |
| Ngày kết thúc / ký BBKT | BBKT, chỗ kết thúc kiểm tra |
| Ngày báo cáo | Báo cáo kết quả, Báo cáo giám sát, Biên bản làm việc 04-TTKT |

Không tự bịa ngày. Thiếu mốc thì hỏi hoặc ghi chỗ trống theo phôi. Sau khi chèn, rà lại checkbox BCGS:
- **Chuẩn ô hiện tại (từ 2026-10-27, phôi `MAUTONG HOP.docx` đã cập nhật sẵn): ☐ (trống) / ☑ (đã chọn)**, ký tự Unicode đặt **ngay sau nhãn chữ** đúng vị trí phôi gốc. Không dùng lại Wingdings (`¨`/`þ`), không đặt ô ở đầu chữ.
- **Giữ nguyên trạng thái đã tích** của bản cũ khi sửa/chèn thêm nội dung; mục 5 thường tích "Không có phản ánh", để trống "Không vi phạm" / "Có vi phạm" trừ khi user chốt khác.
- Mục 9 có **3 lựa chọn** cùng hàng: "Có thực hiện kiểm tra", "Không thực hiện kiểm tra", "Không có thông tin thu thập từ bên ngoài" - cả 3 đều phải có ô riêng (phôi cũ từng thiếu ô ở lựa chọn thứ 3, đã vá).

### 4. Rà thể thức sau khi điền số (BẮT BUỘC, kể cả vụ giải thể)

Điền số xong **chưa phải là xong**: vụ lập từ phôi khác (GT giải thể, hoàn thuế) thường chưa có các chuẩn của `ho-so-kiem-tra-ke-hoach`. Rà đúng 4 việc này trên TONG HOP + từng file QĐ/tờ trình riêng:

1. **Khoảng ký 5 dòng**: giữa dòng chức danh cuối và dòng tên phải có đúng 5 đoạn trống (tờ trình, báo cáo kết quả, phân công, đề nghị ý kiến pháp lý, bàn giao, tờ trình hoàn). Quét bằng script: đi từ dòng tên ngược lên, đếm đoạn trống liền trước.
2. **Chống rớt trang**: sau khi tăng dòng phải mở Word, `Repaginate`, so `Information(3)` của dòng chức danh và dòng tên. Rớt trang thì giảm 5 -> 4 -> 3; vẫn rớt thì giữ 5 dòng và set `keep_with_next` cho dòng chức danh + các dòng trống để Word đẩy trọn khối sang trang sau.
3. **Quyết định**: `Số` thẳng dưới chữ `sở`; `Trấn Biên, ngày       tháng       năm 2026` (bỏ trống ngày tháng, **giữ năm**, canh giữa ô); các `Điều 1, 2, 3...` đậm.
4. **Đếm lại trang thật** rồi điền đoạn kết từng biên bản (số trang đổi mỗi khi thêm/bớt dòng ký).

### 5. Dọn bak / backup

Chỉ chạy khi user bảo dọn hoặc sau khi đã chốt file live và có bằng chứng lưu thành công:

- Xóa trong thư mục `0. BIEN BAN\`: `*.bak*`, `TONG HOP.bak-*`, `*_backup*.xlsx`, file `.md` tạm không thuộc hồ sơ.
- **Giữ** file live: `TONG HOP.docx`, QĐ riêng, PDF đã ký (nếu có), `Bang tinh*.xlsx` đang dùng, `hoso.zip` mới nhất.
- Liệt kê lại thư mục bằng tool sau khi xóa; chỉ xác nhận đã xóa khi lệnh trả thành công.

## Bẫy

- Điền BBXN nhầm sang kiểu BBKT (thêm TNTT/truy thu) làm lệch phôi.
- Tách ĐVT sang dòng riêng hoặc bỏ tab ~8931.
- Đồng bộ QĐ từ bản TONG cũ hơn bản tay vừa canh.
- Normalize NĐ 30 hàng loạt lên biên bản/kế hoạch/báo cáo (phá floating, giãn dòng).
- Thay checkbox BCGS mà reset hết ô đã tích.
- Dùng lại Wingdings hoặc đặt ô checkbox ở đầu chữ - chuẩn từ 2026-10-27 là ☐/☑ Unicode nằm sau chữ.
- BBXN trưởng đoàn chỉ ghi 1 chỉ tiêu/năm hoặc tách 2 chỉ tiêu thành 2 bảng riêng - phải là 1 bảng, TT 1/2 (trừ khi phân công vụ đã có thêm chỉ tiêu trên phôi).
- Gạch đầu dòng `-` trước tên chỉ tiêu trong BBXN/BBKT - phôi không gạch đầu dòng.
- Xóa nhầm `Bang tinh 2021-2025.xlsx` / PDF đã ký khi dọn backup.
- Cộng tổng `TINH PHAT` khi sheet đã gom chậm nộp GTGT vào dòng TNDN rồi cộng lại ở dòng thu hồi hoàn; lấy sổ.
- In truy thu TNDN năm ưu đãi bằng số trước miễn giảm; thiếu 2 dòng miễn giảm / còn phải nộp trên BBKT.
- Tự thêm chỉ tiêu vào BBXN khi user đã bổ sung khối GTGT.
- Bỏ sót số tài khoản trên Quyết định thu hồi hoàn.
- Sửa TONG HOP khi 3 nguồn lệch mà chưa hỏi user.
- Coi "điền xong số" là xong mà bỏ rà thể thức: vụ giải thể/hoàn thuế lập từ phôi khác hay còn khoảng ký 2-3 dòng, dòng `Trấn Biên, ngày tháng năm` mất `2026`, `Số` chưa canh dưới `sở`.
- Canh `Số` bằng cách đổi `LeftIndent` mà **không xóa 16-21 space đầu dòng** của phôi: indent đúng nhưng chữ vẫn lệch phải. Xóa whitespace đầu dòng trước, rồi mới set indent.
- Tin `Range.Information(5)` sau khi vừa sửa indent trong cùng session Word: giá trị không cập nhật (kể cả Repaginate, kể cả mở lại). Tính indent = x(chữ `SỞ`) - (x(đầu đoạn `Số`) - indent cũ) rồi set một lần, không đo lặp.
- `ExportAsFixedFormat` trên file ổ Google Drive (G:) treo vô hạn - copy về local trước nếu cần PDF, hoặc bỏ bước PDF.
- Script giảm dòng ký nhắm "lần khớp đầu tiên": tài liệu có nhiều khối cùng chức danh, phải nhắm đúng khối rớt trang (thường là khối cuối).
- Đoạn `Hôm nay, hồi 8 giờ 00 ngày DD/MM/YYYY` thường tách nhiều `w:t` (`Hôm nay, hồ` / `i ` / `8` / ` giờ` / `00`). Replace chuỗi liền sẽ miss; replace hàng loạt ngày ký BBKT sẽ đè nhầm ngày mở Báo cáo kết quả. Đổi ngày báo cáo theo ngữ cảnh (Tổ Kiểm tra / 8 giờ) riêng, trước khi đổi hàng loạt ngày BBKT.

## Kiểm chứng

- Mỗi năm trong thời kỳ có đủ khối năm + bảng (nếu có số) + `*/` đúng loại (`Giải trình số liệu chênh lệch qua kiểm tra:` + `*/` hoặc `*/ Không chênh lệch.`).
- BBXN trưởng đoàn mỗi năm: 1 bảng, TT 1 = QLDN, TT 2 = bán hàng (hoặc đúng số chỉ tiêu phân công vụ); tên chỉ tiêu không gạch đầu dòng.
- BBXN không có dòng truy thu TNDN; BBKT năm chênh có `=>Thuế TNDN truy thu`.
- Năm có miễn giảm TNDN: bảng BBKT đủ 2 dòng miễn giảm và còn phải nộp; truy thu = số sau miễn giảm.
- Tổng truy thu + chậm nộp + phạt + thu hồi hoàn khớp sổ.
- BBXN giữ đúng số chỉ tiêu bản user đã bổ sung; phần GTGT trên BBKT giữ nguyên.
- QĐ thu hồi hoàn có STK nếu hồ sơ/sổ đã có số TK.
- Checkbox BCGS toàn bộ là ☐/☑ Unicode sau chữ, không còn Wingdings/`□`; mục 9 đủ 3 ô.
- Hai file QĐ riêng khớp khối trong TONG HOP và mỗi file 1 trang.
- Ngày công bố/báo cáo khớp sổ (hoặc đúng số user chốt).
- Mọi khối ký: chức danh và tên cùng một trang (Word `Information(3)` trùng nhau), khoảng cách 5 dòng (hoặc 4/3 nếu đã phải giảm để cùng trang).
- Mọi quyết định: `Số` dưới `sở`, dòng địa danh có `năm 2026` và canh giữa, các `Điều` đậm.
- Thư mục vụ không còn `*.bak*` / `*_backup*` sau bước dọn; còn đủ file live.

## Nối skill

1. `nhap-bang-tinh-tu-bao-cao-hdkd` - điền KQKDSXKD từ PTDN02/BCTC.
2. `boc-tach-bang-tinh-phan-tich` - tỷ lệ + Mục III/IV DI-01.
3. **`dien-bien-ban-tu-bang-tinh`** (skill này) - BBXN/BBKT + QĐ + ngày + dọn.
4. `ho-so-kiem-tra-ke-hoach` - font/checkbox/lint/xuất trọn gói / NĐ 30 cho hành chính.

## Lịch sử (tự học)
- [2026-11-06] Lượt 06/11/2026, làm theo skill cũ sẽ cộng trùng chậm nộp GTGT từ TINH PHAT, in TNDN 2026 trước miễn giảm, và có thể thêm chỉ tiêu vào BBXN. User chốt 4 điểm: sổ 48.899.062, 2 dòng ND398 trên BBKT 2026, giữ BBXN GTGT, STK trên QĐ thu hồi hoàn. Bổ sung bước đối chiếu 3 nguồn, các bẫy đó, và kiểm chứn
- [2026-11-06] Vụ Sơn Tiền Hải (GT giải thể): điền số xong nhưng hồ sơ vẫn thiếu chuẩn thể thức (khoảng ký 2-3 dòng, mất `năm 2026`, `Số` chưa canh). Bổ sung mục 4 rà thể thức, 5 bẫy về canh `Số`/đo Word COM/PDF trên G:, và 2 điểm kiểm chứng khối ký + quyết định.
- [2026-11-09] Vụ Tâm Hoàng Phát: sổ đã có ngày ký BBKT 12/11 nhưng AR14 còn chậm nộp tính đến 30/10. Lấy TINH PHAT (4.747.340) vì khớp ngày BBKT; không dùng 4.670.434. Đoạn mở Báo cáo kết quả bị tách run nên replace `30/10` hàng loạt suýt đổi nhầm thành ngày BBKT.
