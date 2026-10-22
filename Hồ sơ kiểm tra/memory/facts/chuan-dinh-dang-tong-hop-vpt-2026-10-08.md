---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-08
updated: 2026-10-08
---
**Nguồn chuẩn:** file `TONG HOP.docx` vụ Vạn Phát Thành do người dùng canh chỉnh tay (15:44 08/10/2026). Bản tham chiếu lưu tại `templates/TONG HOP CHUAN DINH DANG.docx`. Đối chiếu với `TONG HOP.bak-r4-20261008-145422.docx`.

**Vì sao:** Bản script/auto còn lệch mắt (gạch ký, năm/ĐVT tách dòng, câu loại CPTC, một số dòng chi phí dính nhau). Bản tay là chuẩn áp dụng cho hồ sơ sau.

**Áp dụng:** Mọi hồ sơ kiểm tra kế hoạch xuất `TONG HOP.docx` + QĐ riêng phải bám các quy tắc dưới đây; không normalize hàng loạt.

### Meta giữ nguyên
- A4, **7** bảng nổi (`tblpPr`), khoảng **36** trang (tùy nội dung).
- Lề section chính: trái **3.0** cm, phải **2.0** cm (section 0: trên 1.8 / dưới 1.4; section khác có thể 2.0/2.0).
- Font Times New Roman toàn bộ nội dung; **không để lẫn Calibri** (hoặc 2 font trong cùng đoạn nội dung).
- Giãn dòng chủ yếu **1.0** (một số dòng `*/` giải trình ~1.2).
- Không dựng từ trang trắng; thao tác trên phôi `MAUTONG HOP.docx`.
- **Báo cáo giám sát - checkbox:** dùng hộp Wingdings `\xa8` (trống) / `\xfe` (đã chọn), không dùng ký tự Unicode `□` vẽ tay; chữ nhãn vẫn Times New Roman. Khi thay hộp phải **giữ nguyên ô đã tích**; không reset hết về trống.
- Mục 5 BCGS vụ VPT: tích `Không có phản ánh`; **không** tích `Không vi phạm` / `Có vi phạm`.

### QĐKT (văn bản hành chính NĐ 30, đúng 1 trang khi xuất riêng)
- Gạch VML trên `TRƯỞNG THUẾ CƠ SỞ...`: canh giữa dòng chữ, khoảng `from=0,-1pt` → `to=160pt,-1pt` (không lệch phải kiểu 157→299 của phôi thô).
- Khoảng trước/sau dòng chức danh ký nới vừa phải (sb ~12pt, sa ~6pt) để gạch không đè chữ.
- `QUYẾT ĐỊNH:` căn giữa; khoảng trước/sau ~8pt.
- `Điều 1.` / `Điều 2.` / `Điều 3.`: **canh đều**, thụt đầu dòng **1.0 cm**; nhãn `Điều x.` **đậm**, phần còn lại thường, cỡ 13.
- Các dòng nội dung dưới Điều 1 (`- Chi phí...`, `Thời kỳ...`, `Thời hạn...`): mỗi ý **một dòng**, canh đều, thụt đầu dòng **1.0 cm**, không đậm.
- Ngày tháng trên header QĐ để trống dạng `..., ngày .... tháng .... năm ....`.

### QĐGS (đúng 1 trang khi xuất riêng)
- Đoạn `Căn cứ Quyết định kiểm tra số...`: **nghiêng**, không đậm, cỡ **12**, canh đều, thụt đầu dòng **1.0 cm**.
- `Điều 1.`–`Điều 3.`: canh đều, thụt ~1.0–1.2 cm theo phôi.
- Nhãn `Điều 4.` **đậm**, phần sau thường; canh đều, thụt đầu dòng **1.0 cm**.

### BBXN (biểu mẫu nghiệp vụ - giữ tab/bảng phôi)
- `I. Nội dung kiểm tra:` / `II. Kết quả kiểm tra:`: **đậm**, canh đều, thụt đầu dòng **1.0 cm**.
- Mỗi chuyên đề chi phí một dòng riêng (`- Chi phí...`), canh đều, thụt 1.0 cm.
- Dòng năm + ĐVT **cùng một dòng**, có tab:
  - Thành viên (TV/CPTC): `+ NĂM 20xx:` đậm + tab + `Đơn vị tính: Đồng` nghiêng cỡ nhỏ hơn (~11); tab center 851 + **right 8931**.
  - Trưởng đoàn (TD/QLDN+CPBH): `+Năm 20xx` đậm + tab + ĐVT nghiêng; tab **right 8931**.
- Không tách ĐVT ra dòng riêng.
- `*/` giải trình hoặc `*/ Không chênh lệch.`: canh đều, thụt ~1.0–1.25 cm; BBXN **không** ghi TNTT tăng / thuế truy thu.
- Bảng số liệu 5 cột từ DTCHEP: giữ layout phôi, ô chỉ tiêu vừa 1 dòng.

### BBKT
- Dòng năm: `+Năm 20xx` + tab + `Đơn vị tính: Đồng` (cùng dòng), không dùng dạng `* NĂM 20xx Đơn vị tính: Đồng` dính không tab.
- Năm có chênh: khối Giải trình + diễn giải TNTT + `*/` nguyên nhân + `=>Thuế TNDN truy thu...`.
- Năm không chênh (2023, 2024): chỉ `*/ Không chênh lệch.`
- Gộp dòng tài khoản ngân hàng thành một dòng nếu bị tách.
- Đủ mục `III. TÌNH HÌNH THỰC HIỆN NGHĨA VỤ...` và `IV. KẾT LUẬN, KIẾN NGHỊ VÀ YÊU CẦU:`.

### Câu loại chi phí tài chính (nội dung nghiệp vụ)
- Thống nhất: `Công ty hạch toán chi phí lãi vay không tương ứng doanh thu`.
- Căn cứ theo năm:
  - 2021–2022 (và tương tự trước 2025): `Căn cứ Điều 4 Thông tư 96/2015/TT-BTC ngày 22/06/2015 của Bộ Tài chính.`
  - 2025 trở đi: `Căn cứ quy định tại Điều 10 Nghị định số 320/2025/NĐ-CP ...` (và Thông tư hướng dẫn nếu có trong hồ sơ).
- **Không** dùng cụm cũ `không đúng quy định` (đã bỏ).

### Xuất file
- Gom mọi biểu mẫu trong `TONG HOP.docx`.
- Chỉ xuất riêng QĐKT / QĐGS (và QĐXPHC khi đến bước xử phạt) để trình ký; đồng bộ format QĐ riêng với khối tương ứng trong TONG HOP.
