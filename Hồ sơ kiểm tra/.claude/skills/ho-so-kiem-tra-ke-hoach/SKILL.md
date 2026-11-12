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
     + Khoảng cách chữ ký trong cùng 1 ô: mặc định chức danh và họ tên người ký cách nhau 5 dòng (5 dòng trống). Nếu rớt sang trang mới thì tự động giảm xuống 4 dòng (chuẩn NĐ 30) để cùng trang in.
     + Dự thảo Quyết định (chưa ban hành): chữ 'Số' nằm thẳng dưới chữ 'sở' của dòng tiêu đề cơ quan; ký hiệu để khoảng trống dài `Số:              /QĐ-TCS3` (tuyệt đối không dùng mã tạm `DU_THAO-TRINH-KY`); ngày tháng căn giữa để trống toàn bộ ngày, tháng và năm `Trấn Biên, ngày       tháng       năm      ` (không giữ năm cố định); các Điều 1, 2, 3... tô đậm.
     + Giấy xác nhận giao quyết định: dòng ngày tháng dùng chuẩn `ngày ...... tháng ...... năm      ` (đều 6 chấm, để trống năm).
     + Quy tắc sau ban hành: khi QĐKT/QĐGS đã có số và ngày chính thức thì không cập nhật hay ghi đè lại số/ngày đó ở các lần xử lý sau; chỉ cập nhật nội dung và số liệu nếu có thay đổi.
     + Căn cứ Quyết định 1377/QĐ-CT chốt thống nhất ngày 30/06/2025 (năm 2025, không dùng 2026).
   - Giữ nguyên cấu trúc, bảng biểu và nội dung phôi gốc cho Biên bản kiểm tra (BBKT), Biên bản xác nhận số liệu (BBXN), Kế hoạch và Báo cáo giám sát:
     + BBXN Trưởng đoàn: đánh số TT 1 và 2 cho Chi phí quản lý doanh nghiệp và Chi phí bán hàng, không dùng gạch đầu dòng ở tên chỉ tiêu; có chênh lệch thêm dòng `Giải trình số liệu chênh lệch qua kiểm tra:` rồi xuống hàng `*/ [nội dung]`; không chênh lệch ghi `*/ Không chênh lệch` (viết thường).
     + BBKT: giữ nguyên khối giải trình có sẵn của phôi, không chèn thêm khối giải trình trùng lặp.
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
- [2026-11-11] Thống nhất bộ quy tắc hoàn thiện thể thức hồ sơ đi và TONG HOP: (1) Căn cứ QĐ 1377/QĐ-CT chốt thống nhất năm 2025 (30/06/2025) cho toàn bộ phôi và script xuất văn bản; (2) Ngày tháng ban hành dự thảo QĐ để trống toàn bộ ngày, tháng, năm (bỏ năm 2026 cố định); (3) Ký hiệu dự thảo để khoảng trống dài `Số:              /QĐ-TCS3`, tuyệt đối không dùng mã tạm DU_THAO-TRINH-KY; (4) Sau khi QĐKT/QĐGS đã ban hành chính thức thì không cập nhật lại số/ngày; (5) Khoảng cách chữ ký trong ô mặc định 5 dòng trống, nếu rớt trang tự động giảm 4 dòng (chuẩn NĐ 30); (6) BBXN Trưởng đoàn đánh số TT 1 và 2, có chênh lệch ghi dòng tiêu đề Giải trình rồi xuống dòng `*/`, không chênh lệch ghi `*/ Không chênh lệch`.
