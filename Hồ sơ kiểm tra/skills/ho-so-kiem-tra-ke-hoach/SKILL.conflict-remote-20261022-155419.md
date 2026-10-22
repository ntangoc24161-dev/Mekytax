---
name: Hồ sơ kiểm tra theo kế hoạch
description: 'Lập hồ sơ kiểm tra kế hoạch từ MAUTONG HOP: giữ phôi gốc nghiệp vụ, chuẩn NĐ 30 văn bản hành chính, giữ checkbox và xuất QĐ riêng.'
group: Tài chính
updated: '2026-10-08'
learned_updated: '2026-10-08'
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
   - Áp dụng chuẩn Nghị định 30/2020/NĐ-CP cho Tờ trình, Quyết định kiểm tra (QĐKT), Quyết định giám sát (QĐGS).
   - Giữ nguyên cấu trúc, bảng biểu và nội dung phôi gốc cho Biên bản kiểm tra (BBKT), Biên bản xác nhận số liệu (BBXN), Kế hoạch và Báo cáo giám sát.
2. **Rà soát font chữ toàn diện**:
   - Quét và chuyển toàn bộ các đoạn văn bản (run) bị lẫn Calibri sang font Times New Roman tường minh.
3. **Chuẩn hóa Checkbox trong Báo cáo giám sát**:
   - Thay thế các ký tự checkbox thô hoặc thủ công (`□`) bằng hộp ký tự Wingdings chuẩn.
   - **Bắt buộc đối chiếu bản cũ**: Nhận diện các ô vốn đã được tích chọn (đánh dấu x hoặc tick) để chuyển thành ký tự hộp đã tích chọn chuẩn (Wingdings `þ` / `x`), tuyệt đối không để trắng toàn bộ các ô đã có dữ liệu chọn.
4. **Đồng bộ và xuất file 1-Click**:
   - Sử dụng công cụ xuất trọn gói `scripts/xuat_ho_so_tron_goi.py`:
     ```bash
     python "scripts/xuat_ho_so_tron_goi.py" --mst <MST> --ten-nnt "<Tên NNT>" --so-qd "DU_THAO-TRINH-KY"
     ```
   - Tự động xuất 02 file QĐKT và QĐGS độc lập (khống chế chuẩn 1 trang A4).
   - Tự động lưu toàn bộ các biểu mẫu nghiệp vụ trong `TONG HOP.docx` và đóng gói `hoso.zip`.
5. **Kiểm định chất lượng trước xuất xưởng (Pre-flight Lint)**:
   - Sử dụng công cụ kiểm định `scripts/lint_ho_so_docx.py`:
     ```bash
     python "scripts/lint_ho_so_docx.py" "<đường_dẫn_file.docx>"
     ```
   - Xác nhận: 100% Times New Roman (0 Calibri), bảo toàn trạng thái checkbox Wingdings, giữ nguyên 7 khung floating table (`tblpPr`), không tràn trang.

## Bẫy

- Khi duyệt thay thế chuỗi ký tự checkbox hàng loạt, việc gán cứng tất cả thành ô trống (`¨`) làm mất dữ liệu tích chọn người dùng đã đánh dấu trước đó.
- Thay đổi font của cả đoạn văn bản thay vì chỉ riêng ký tự checkbox làm lỗi font nhãn nội dung tiếng Việt đi kèm.

## Kiểm chứng

- Mở file kiểm tra bằng Word hoặc kiểm tra XML: số lượng ô tích chọn khớp chính xác với bản cũ trước khi thay thế.

## Lịch sử (tự học)
- [2026-10-07] Với vụ Vạn Phát Thành: nguyên nhân loại CPTC ghi "Công ty hạch toán chi phí lãi vay không đúng quy định"; năm không chênh lệch (2023, 2024) chỉ để 1 dòng "*/ Không chênh lệch.", tách riêng từng năm, không giữ khung giải trình của năm có chênh lệch; giữ toàn bộ floating của phôi (Mẫu số + quốc hiệu/đoàn) và giãn dòng theo phôi (chủ yếu 1.0), không ép 1.15.
- [2026-10-07] Format lệch mẫu vì script ép `line_spacing=1.15` + `space_after=1pt` và gỡ floating header. Cách đúng: thao tác trên phôi, giữ nguyên spacing/floating, không normalize hàng loạt.
- [2026-10-08] Thống nhất nguyên tắc chuẩn hóa: Tờ trình, QĐKT, QĐGS và Quyết định xử phạt VPHC áp dụng chuẩn thể thức NĐ 30/2020/NĐ-CP; toàn bộ Biên bản, Kế hoạch, Báo cáo và Bảng xác nhận giữ nguyên 100% layout, bảng biểu, tab stop và khung Mẫu số của phôi gốc `MAUTONG HOP.docx`.
- [2026-10-08] Cập nhật bẫy over-normalization làm lệch trích yếu các biểu mẫu nghiệp vụ và bổ sung nguyên tắc phân tách rõ giữa biểu mẫu phôi gốc và văn bản hành chính NĐ 30 kèm các quy tắc format đơn lẻ.
- [2026-10-08] Biên bản xác nhận số liệu (TV/TD): mỗi năm theo phôi = **cùng một dòng** năm đậm + tab + `Đơn vị tính: Đồng` nghiêng cỡ nhỏ hơn (tab right ~8931 twips); không tách ĐVT ra dòng riêng. Phần nội dung ngoài bảng số liệu thống nhất cỡ 13, canh đều; mục `I. Nội dung kiểm tra` / `II. Kết quả kiểm tra` đậm. Bảng 5 cột từ DTCHEP; `*/` giải trình hoặc `*/ Không chênh lệch.`; không ghi TNTT tăng / thuế truy thu (thuộc biên bản kiểm tra). Căn cứ giải trình từ năm 2025: Điều 10 NĐ 320/2025/NĐ-CP và Thông tư 20/2026/TT-BTC.
- [2026-10-08] QĐKT: gạch VML trên dòng `TRƯỞNG THUẾ CƠ SỞ` canh giữa với dòng trên (không lệch phải như phôi thô); Điều 1 và các dòng nội dung canh đều, thụt đầu dòng 1 cm. QĐGS: đoạn `Căn cứ Quyết định kiểm tra số...` in nghiêng; nhãn `Điều 4.` tô đậm, phần còn lại thường.
- [2026-10-08] Người dùng chốt tay `TONG HOP.docx` VPT làm chuẩn hồ sơ sau: lưu `templates/TONG HOP CHUAN DINH DANG.docx`; câu loại CPTC đổi thành `không tương ứng doanh thu` + TT 96/2015 hoặc NĐ 320/2025; BBKT năm dạng `+Năm 20xx` + tab ĐVT; đủ mục III/IV. Không lấy bak-r4 làm chuẩn mắt.
- [2026-10-08] Rà soát font: gỡ Calibri lẫn trong Báo cáo kết quả / kết luận BBKT / header bảng; checkbox Báo cáo giám sát đổi từ `□` Unicode sang hộp Wingdings chuẩn (nhãn Times New Roman).
- [2026-10-08] Khi thay checkbox phải giữ trạng thái đã tích: Wingdings trống `\xa8`, đã chọn `\xfe`. Với VPT BCGS: tích Công bố đúng hạn, Báo cáo GS đúng hạn, Có giao tài liệu, Đúng thời hạn, Tốt, Không phản ánh, Không phát sinh, Có theo tiêu chí rủi ro/phân tích, Không có thông tin bên ngoài; mục 5 **không** tích ô `Không vi phạm`.
- [2026-10-08] Cập nhật quy trình chuẩn hóa checkbox trong Báo cáo giám sát: ngăn chặn lỗi reset toàn bộ ô đã tích chọn về ô trống khi thay thế ký tự sang Wingdings.
