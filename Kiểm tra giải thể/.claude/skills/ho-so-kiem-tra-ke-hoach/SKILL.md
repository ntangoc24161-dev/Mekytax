---
name: Hồ sơ kiểm tra theo kế hoạch
description: 'Lập hồ sơ kiểm tra kế hoạch từ MAUTONG HOP: giữ phôi gốc nghiệp vụ, chuẩn NĐ 30 văn bản hành chính, chuẩn checkbox ☐/☑ và xuất QĐ riêng.'
group: Tài chính
updated: '2026-11-11'
learned_updated: '2026-11-11'
---
## Khi nào dùng

- Khi lập, xuất hoặc chuẩn hóa bộ hồ sơ kiểm tra thuế theo kế hoạch năm từ file phôi `MAUTONG HOP` hoặc `TONG HOP.docx`.
- Khi rà soát định dạng, chuẩn hóa font chữ Times New Roman và chuẩn hóa hệ thống ô đánh dấu (checkbox) trong Báo cáo giám sát.
- Điền số BBXN/BBKT từ Bảng tính, chèn ngày công bố/báo cáo, đồng bộ QĐ riêng sau canh tay, hoặc dọn bak: dùng skill `dien-bien-ban-tu-bang-tinh` trước (hoặc xen kẽ), rồi quay lại skill này để lint/xuất trọn gói.

## Chuẩn bị

- File `TONG HOP.docx` đang xử lý và file phôi mẫu `templates/TONG HOP CHUAN DINH DANG.docx`.
- Bản sao lưu backup trước khi can thiệp cấu trúc XML/run text của file docx.

## Quy trình

1. **Bảo tồn phôi nghiệp vụ và chuẩn hóa văn bản hành chính**:
   - Áp dụng chuẩn Nghị định 30/2020/NĐ-CP cho Tờ trình, Quyết định kiểm tra (QĐKT), Quyết định giám sát (QĐGS):
     + Khoảng cách chữ ký giữa chức danh và tên người ký cách ít nhất 5 dòng trong 1 ô. Nếu rớt sang trang mới thì tự động giảm còn 4 dòng để cùng trang in.
     + Quyết định: chữ 'Số' nằm thẳng dưới chữ 'sở' của dòng tiêu đề cơ quan; ngày tháng căn giữa để trống 'Trấn Biên, ngày       tháng       năm 2026'; các Điều 1, 2, 3... tô đậm.
     + Giấy xác nhận giao quyết định: dòng ngày tháng dùng chuẩn 'ngày ...... tháng ...... năm 2026' (đều 6 chấm).
   - Giữ nguyên cấu trúc, bảng biểu và nội dung phôi gốc cho Biên bản kiểm tra (BBKT), Biên bản xác nhận số liệu (BBXN), Kế hoạch và Báo cáo giám sát.
2. **Rà soát font chữ toàn diện**:
   - Quét và chuyển toàn bộ các đoạn văn bản (run) bị lẫn Calibri sang font Times New Roman tường minh.
3. **Chuẩn hóa Checkbox trong Báo cáo giám sát**:
   - Checkbox đặt ở vị trí nằm SAU nội dung chữ.
   - Sử dụng ô vuông Unicode hiện đại: ô trống dùng `☐` (U+2610), ô đã chọn dùng `☑` (U+2611).
   - Gỡ bỏ hoàn toàn các nét vẽ X thủ công (VML pict/shape đè lên ô).
   - **Đối chiếu bản cũ và phôi**: Mục 1 đến 8 đánh dấu theo kết quả giám sát; Mục 9 đánh dấu `☑` ở cả hai dòng 'Có thực hiện kiểm tra' và 'Không có thông tin thu thập từ bên ngoài'.
4. **Đồng bộ và xuất file 1-Click**:
   - Tự động xuất 02 file QĐKT và QĐGS độc lập (khống chế chuẩn 1 trang A4).
   - Lưu toàn bộ các biểu mẫu nghiệp vụ trong `TONG HOP.docx`, dọn dẹp các file bak thừa.
5. **Kiểm định chất lượng trước xuất xưởng (Pre-flight Lint)**:
   - 100% Times New Roman, bảo toàn trạng thái checkbox `☐`/`☑` sau chữ, không còn nét vẽ VML thừa, chữ ký không rớt trang in, giữ nguyên 7 khung floating table (`tblpPr`).

> Áp dụng cho brain giải thể: phôi là `MAUTONG HOP GIAI THE.docx`; Mục III NSNN giải thể vẫn điền đủ theo `ho-so-kiem-tra-giai-the`; văn bản đóng MST/từ chối nộp thừa theo skill đó.

## Quy tắc kỹ thuật Mail Merge (chốt 2026-11-25, đã sửa bằng thực chiến vụ T&T Biên Hòa)

**Phía sổ `Danh sach kiem tra quyet toan pro.xlsm` (làm TRƯỚC khi trộn):**
1. Các cột ngày (`Ngay_cap_CCCD`, `Ngay_cap_DKKD`, `Ngay_QDKT`, `Ngay_cong_bo`, `Ngay_bat_dau`, `Ngay_ket_thuc`, `Ngay_ky_BBKT`, `Ngay_bao_cao`, `Ngay_QDXPHC`) phải là **ngày thật** (không phải chữ như `13/11/2026`, `12.08.2021`, `<tab>21/09/2026`), định dạng ô `dd/mm/yyyy`. Ô công thức `=AHxx` giữ nguyên. Với kết nối `IMEX=1`, Word sẽ nhận chuỗi `dd/MM/yyyy` đúng ngay từ nguồn.
2. Các cột tiền (cột AN đến AY: truy thu đến giảm khấu trừ) định dạng `#.##0` (hiện `17.850.636`). Đặt đúng chuỗi này, không đặt `#.##0,0` (sẽ hiện `17850636,0`).
3. Sửa sổ bằng Excel COM (giữ macro/công thức); sổ đang mở trong Excel thì bị khóa ghi, phải nhờ user đóng trước.
4. Trước khi lập hồ sơ, kiểm tra các cột mốc ngày của vụ đã có dữ liệu, nhất là `Ngay_cong_bo` (hay bị trống). Trống thì hỏi user rồi ghi vào sổ, không tự bịa.

**Phía TONG HOP.docx:**
1. Tên MERGEFIELD phải đúng tên không dấu của Dòng 1 ẩn (`Ten_NNT`, `Giam_doc`, `So_QDKT`...). Tên có dấu cũ (`Tên_công_ty`, `TĐ`, `GĐ`) làm Word chỉ hiện `«...»`. Sau khi sửa, quét: không còn trường nào ngoài danh sách cột.
2. Bật `viewMergedData`; `activeRecord` = số thứ tự bản ghi (dòng Excel trừ 1 dòng tiêu đề ẩn: dòng 18 thì bản ghi 17); `table val` khớp tên bảng `01_DA_BAN_HANH$` trong câu SQL; đường dẫn nguồn phụ phải là ổ `G:` (không để `D:\`).
3. **Trường ngày đầy đủ**: chỉ `{ MERGEFIELD Ngay_xxx }`, KHÔNG gắn `\@ "dd/MM/yyyy"` (Word đọc nhầm tháng/ngày). **Trường tiền**: chỉ `{ MERGEFIELD ... }`, KHÔNG gắn `\# "#,##0"` (sinh `2804719,  0`).
4. **Ngày tách `dd` / `MM` / `yyyy`** không tách được từ chuỗi: nếu phôi cần dạng tách thì thêm cột phụ trong sổ (vd `Ngay_bao_cao_dd`, `_MM`, `_yyyy`) rồi trộn; chưa có cột phụ thì điền giá trị tĩnh và báo user cập nhật lại khi đổi ngày trong sổ.
5. Mở Word bằng COM để thử Mail Merge trên file ở ổ `G:` hay treo: chỉ kiểm chứng bằng đọc XML + đọc dữ liệu qua OLEDB; luôn dặn user tự mở Word, bấm Yes ở hộp SQL, đóng Excel trước. Không hiểu nguyên nhân thì lùi về ngày tĩnh, đừng để user bị bắt chọn lại bảng.

**Bảng ánh xạ ngày (user chốt):**

| Văn bản | Trường |
| :--- | :--- |
| Công bố QĐKT, giấy giao quyết định | `Ngay_cong_bo` |
| BBXN số liệu, công khai BBKT (Mẫu 15/KTT) | `Ngay_ket_thuc` |
| Báo cáo kết quả kiểm tra, báo cáo giám sát, biên bản làm việc | `Ngay_bao_cao` |
| Kế hoạch giám sát | `Ngay_QDKT` |
| BBKT (ký) | `Ngay_ky_BBKT` |

Không dùng nhầm `Ngay_ky_BBKT` cho các mục trên.

**Phân công đoàn theo sổ (chốt 2026-11-25):** nguồn duy nhất là cột `Phân công trưởng đoàn` (AA) và `Phân công thành viên` (AB) của sổ `Danh sach kiem tra quyet toan pro.xlsm`, không dùng mặc định cứng.
- Mỗi người được phân công N nội dung thì BBXN của người đó có N dòng chỉ tiêu mỗi năm (TT 1..N), tách riêng từng nội dung. Ví dụ Trưởng đoàn phụ trách `Chi phí quản lý doanh nghiệp` và `Chi phí bán hàng` thì tách 2 dòng; phụ trách 1 nội dung thì 1 dòng.
- Ngoại lệ do user dặn theo vụ (vd gộp bán hàng vào quản lý kinh doanh thành một dòng) ghi trong memory của vụ và được ưu tiên hơn quy tắc tách dòng.
- Tên chỉ tiêu lấy đúng chữ trên sổ; phôi và biên bản phải khớp, không dựng khung cố định "1. QLDN, 2. Bán hàng".

**Nội dung và định dạng khi đổ số:**
- Phân tích Mục III DI-01 và bảng biên bản chỉ theo **nội dung kiểm tra được phân công** (vd chỉ chi phí và chi phí tài chính, bỏ giá vốn nếu không kiểm tra). Chi phí quản lý (QLDN hoặc QLKD) = chi phí bán hàng + chi phí quản lý doanh nghiệp, ghi thành **một dòng**, không đưa dòng chi phí bán hàng riêng khi user dặn. Tên chỉ tiêu theo đúng **chốt của vụ** (vụ T&T giữ `Chi phí quản lý kinh doanh`, vụ khác `quản lý doanh nghiệp`): hỏi hoặc đọc memory của vụ, không tự đổi.
- Công thức DTCHEP chỉ lấy dòng QLDN (KQKDSXKD dòng 21) chưa cộng bán hàng: nếu bán hàng > 0 phải sửa công thức hoặc cộng tay, báo user.
- Năm chênh lệch: đọc **DTCHEP cột chênh lệch** (E) chứ không chỉ KQKDSXKD; vd CPTC bị loại 100% thì số kiểm tra = 0, chênh lệch = -số báo cáo, giải trình đúng mẫu `*/ ... Căn cứ ...` (BBXN không ghi thuế truy thu).
- Mục III NSNN: lấy từ sheet **QHNS** của Bảng tính (vd `QHNS 970`) đủ cột (3) kỳ trước, (4) phát sinh báo cáo, (5) qua kiểm tra, (6) đã nộp, (7) còn phải nộp; số đã nộp gộp cả nợ năm trước là đúng. Không tự điền số không có trong sheet.
- Cỡ chữ: nội dung 13, bảng số liệu các năm 12, tiêu đề 14, phần còn lại giữ như phôi; Times New Roman toàn file; bỏ dòng trắng thừa giữa các bảng năm.
- Số QĐKT/QĐGS/ngày QĐ đã có trên sổ: điền ngay, không để trống (chỉ dự thảo chưa có số mới để trống).
- Khi thay chữ trong dòng Word, giữ nguyên `rPr` (nghiêng, cỡ) và `xml:space="preserve"`; không ghi lại cả đoạn bằng một run mới (làm mất nghiêng và dính `số7516`).
- **Công văn cam kết doanh thu, chi phí và tiền lương** (khi user yêu cầu): thêm cuối TONG HOP, trang riêng, từ `templates/MAU CAM KET DOANH THU VA CHI PHI.docx`; đổi tên công ty, MST, địa chỉ, người ký, ngành nghề (câu mô tả hàng hóa) sang trường Mail Merge; số văn bản và ngày để trống cho DN điền.

**Quy tắc ngày/năm (chốt 2026-11-25):** dự thảo quyết định chưa có số trên sổ để trống số và toàn bộ dòng địa danh `ngày       tháng       năm      `; mọi trường hợp còn lại (QĐ đã có số, giấy giao quyết định, biên bản, báo cáo, hồ sơ về) điền đầy đủ số, ngày, tháng, năm theo sổ.

## Cảnh báo bắt buộc báo user (không chặn đổ số)

- **Lệch phân công và Bảng tính:** Bảng tính có 2 dòng riêng chi phí quản lý doanh nghiệp và chi phí bán hàng (hoặc bán hàng khác 0) mà sổ phân công là `Chi phí quản lý kinh doanh`. Cảnh báo dạng: `[CẢNH BÁO PHÂN CÔNG]: Bảng tính có QLDN = X, bán hàng = Y, nhưng phân công ghi chi phí quản lý kinh doanh. Gộp thành một dòng (X+Y) hay tách hai dòng?` Chưa có trả lời thì tạm gộp X+Y một dòng và nêu rõ trong báo cáo. Bán hàng bằng 0 thì không cần cảnh báo.

## Bẫy

- Gán cứng ô checkbox bằng Wingdings hoặc đặt ở đầu câu làm sai lệch phôi gốc mẫu giám sát.
- Để nguyên nét vẽ VML X khi đổi sang ô `☑` làm nét vẽ bị lệch đè lên chữ.
- Khoảng cách chữ ký 5 dòng cố định kéo dài ô làm rớt tên người ký sang trang mới (phải giảm xuống 4 dòng theo NĐ 30).

## Kiểm chứng

- Mở file bằng Word: toàn bộ checkbox hiển thị rõ nét `☐`/`☑` nằm sau chữ, các khối ký nằm gọn trong 1 trang, QĐKT và QĐGS gói gọn đúng 1 trang A4.

## Lịch sử (tự học)
- [2026-10-07] Format lệch mẫu vì script ép `line_spacing=1.15` + `space_after=1pt` và gỡ floating header. Cách đúng: thao tác trên phôi, giữ nguyên spacing/floating, không normalize hàng loạt.
- [2026-10-08] Thống nhất nguyên tắc chuẩn hóa: Tờ trình, QĐKT, QĐGS và Quyết định xử phạt VPHC áp dụng chuẩn thể thức NĐ 30/2020/NĐ-CP; toàn bộ Biên bản, Kế hoạch, Báo cáo và Bảng xác nhận giữ nguyên 100% layout, bảng biểu, tab stop và khung Mẫu số của phôi gốc `MAUTONG HOP.docx`.
- [2026-10-08] Cập nhật bẫy over-normalization làm lệch trích yếu các biểu mẫu nghiệp vụ và bổ sung nguyên tắc phân tách rõ giữa biểu mẫu phôi gốc và văn bản hành chính NĐ 30 kèm các quy tắc format đơn lẻ.
- [2026-10-08] Biên bản xác nhận số liệu (TV/TD): mỗi năm theo phôi = **cùng một dòng** năm đậm + tab + `Đơn vị tính: Đồng` nghiêng cỡ nhỏ hơn (tab right ~8931 twips); không tách ĐVT ra dòng riêng. Phần nội dung ngoài bảng số liệu thống nhất cỡ 13, canh đều; mục `I. Nội dung kiểm tra` / `II. Kết quả kiểm tra` đậm. Bảng 5 cột từ DTCHEP; `*/` giải trình hoặc `*/ Không chênh lệch.`; không ghi TNTT tăng / thuế truy thu (thuộc biên bản kiểm tra). Căn cứ giải trình từ năm 2025: Điều 10 NĐ 320/2025/NĐ-CP và Thông tư 20/2026/TT-BTC.
- [2026-10-08] QĐKT: gạch VML trên dòng `TRƯỞNG THUẾ CƠ SỞ` canh giữa với dòng trên (không lệch phải như phôi thô); Điều 1 và các dòng nội dung canh đều, thụt đầu dòng 1 cm. QĐGS: đoạn `Căn cứ Quyết định kiểm tra số...` in nghiêng; nhãn `Điều 4.` tô đậm, phần còn lại thường.
- [2026-10-08] Người dùng chốt tay `TONG HOP.docx` VPT làm chuẩn hồ sơ sau: lưu `templates/TONG HOP CHUAN DINH DANG.docx`; câu loại CPTC đổi thành `không tương ứng doanh thu` + TT 96/2015 hoặc NĐ 320/2025; BBKT năm dạng `+Năm 20xx` + tab ĐVT; đủ mục III/IV. Không lấy bak-r4 làm chuẩn mắt.
- [2026-10-08] Rà soát font: gỡ Calibri lẫn trong Báo cáo kết quả / kết luận BBKT / header bảng; checkbox Báo cáo giám sát đổi từ `□` Unicode sang hộp Wingdings chuẩn (nhãn Times New Roman).
- [2026-10-08] Khi thay checkbox phải giữ trạng thái đã tích: Wingdings trống `\xa8`, đã chọn `\xfe`. Với VPT BCGS: tích Công bố đúng hạn, Báo cáo GS đúng hạn, Có giao tài liệu, Đúng thời hạn, Tốt, Không phản ánh, Không phát sinh, Có theo tiêu chí rủi ro/phân tích, Không có thông tin bên ngoài; mục 5 **không** tích ô `Không vi phạm`.
- [2026-10-08] Cập nhật quy trình chuẩn hóa checkbox trong Báo cáo giám sát: ngăn chặn lỗi reset toàn bộ ô đã tích chọn về ô trống khi thay thế ký tự sang Wingdings.
- [2026-10-27] Cập nhật chuẩn checkbox hiện đại ☐/☑ nằm sau chữ (bỏ Wingdings và nét vẽ VML X), tích 2 ô mục 9 Báo cáo giám sát, chuẩn canh chữ ký 5/4 dòng chống rớt trang, căn 'Số' dưới 'sở' và in đậm các Điều theo phản hồi thực tế từ user.
- [2026-11-11] Vụ giải thể trên nền kế hoạch (như Sơn Tiền Hải): Mục III Nghĩa vụ NSNN bắt buộc điền đủ số phát sinh các năm; rà soát thay thế triệt để cụm từ cũ 'chi phí quản lý kinh doanh' thành 'chi phí quản lý doanh nghiệp'; bổ sung văn bản từ chối nhận lại tiền nộp thừa nếu có.
- [2026-11-25] Đồng bộ từ brain Hồ sơ kiểm tra: quy tắc Mail Merge (ngày thật dd/mm/yyyy, tiền #.##0, bỏ switch \@ và \#), phân công theo sổ (tách dòng), ngày/năm quyết định, cảnh báo phân công QLDN/bán hàng.
