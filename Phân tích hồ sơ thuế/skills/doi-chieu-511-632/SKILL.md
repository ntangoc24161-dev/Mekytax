---
name: Đối chiếu doanh thu 511 và giá vốn 632
description: "Đối chiếu sổ chi tiết TK 511 với TK 632 theo Tool_DC_511_632: khớp HĐ/ngày CT/mặt hàng, lãi gộp và bán dưới giá."
group: Tài chính
updated: '2026-10-09'
---
> **Hợp đồng đầu ra v1.4:** module nội bộ. Kết quả ghi vào sheet đích của sáu tệp KQ (xem skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md). Không giao tệp chuyên đề Tong_hop_*/Doi_chieu_* / Master 8 sheet như sản phẩm.

# Đối chiếu doanh thu 511 và giá vốn 632

Đối chiếu từng dòng **sổ chi tiết TK 511 (Phát sinh Có)** với **sổ chi tiết TK 632 (Phát sinh Nợ)** bằng mini-app Excel `Tool_DC_511_632.xlsm` (macro `ChayDoiChieu`). Mục tiêu: xác định khớp chứng từ, lãi/lỗ từng món, bán dưới giá, và tổng hợp theo năm chứng từ để phục vụ kiểm tra thuế tại bàn / thực địa.

## Khi nào dùng

- "đối chiếu 511 632", "Tool_DC_511_632", "bán dưới giá", "lãi gộp từng hóa đơn", "khớp doanh thu giá vốn".
- Đã có sổ chi tiết bán hàng (511/5111...) và sổ chi tiết giá vốn hàng bán (632) trong hồ sơ NNT cung cấp.
- Cần bổ sung phân hệ doanh thu - giá vốn vào quy trình phân tích hồ sơ (sau bước giá thành / trước đề cương kiểm tra).

## Chuẩn bị

1. File tool: `skills/doi-chieu-511-632/assets/Tool_DC_511_632.xlsm` (hoặc bản đang làm việc của đoàn).
2. Sổ chi tiết TK **511** (hoặc 5111/511x): xuất Excel, đủ cột A→G đúng thứ tự bên dưới.
3. Sổ chi tiết TK **632**: cùng cấu trúc cột; cột G là **Phát sinh Nợ**.
4. Macro phải bật; module `DoiChieu511632` / macro `ChayDoiChieu` đã có trong file `.xlsm`.
5. Chi tiết cấu trúc sheet và nhãn trạng thái: xem `references/cau-truc-tool-dc-511-632.md`.

### Cột bắt buộc (đúng thứ tự A→G)

| Cột | Tên | 511 | 632 |
|---|---|---|---|
| A | Ngày hạch toán | có | có |
| B | Ngày chứng từ | có | có |
| C | Số chứng từ | có | có |
| D | Số hóa đơn | có | có |
| E | Diễn giải | có | có |
| F | Tài khoản | 511x | 632 |
| G | Số tiền | **Phát sinh Có** | **Phát sinh Nợ** |

Không đổi tên sheet `511` và `632`.

## Quy trình

### 1. Dán dữ liệu kỳ cần đối chiếu

1. Mở `Tool_DC_511_632.xlsm`, cho phép macro.
2. Sheet `511`: xóa từ dòng 2 trở đi, **giữ dòng tiêu đề**, dán sổ bán hàng.
3. Sheet `632`: xóa từ dòng 2 trở đi, **giữ dòng tiêu đề**, dán sổ giá vốn.
4. Không gộp ô từ vùng dữ liệu.

### 2. Chạy đối chiếu

1. `Alt + F8` → chọn `ChayDoiChieu` → Run.
2. Hộp thoại xuất kết quả:
   - **Yes**: chọn thư mục, xuất file mới `DoiChieu_511_632_KetQua_...`
   - **No**: ghi vào chính file tool
   - **Cancel**: chỉ xem trên sheet, chưa xuất
3. Macro ghi đè sheet `DCDTCP` theo dữ liệu kỳ vừa dán.

### 3. Đọc kết quả trên `DCDTCP`

Cấu trúc 3 khối:

- **A–G**: dòng 511 (doanh thu PS Có)
- **I–O**: dòng 632 ghép (giá vốn PS Nợ)
- **P–S**: kết quả đối chiếu

Chỉ tiêu kết quả:

- **P `Lãi/Lỗ (DT − GV)`** = PS Có 511 (G) − PS Nợ 632 (O). Số **âm hoặc 0** = bán dưới giá / hòa vốn.
- **Q `Khớp diễn giải`** = TRUE khi diễn giải hai bên giống nhau.
- **R `Tình trạng đối chiếu`** (màu lọc):
  - Xanh: `Khớp HĐ + ngày CT + mặt hàng` (Khóa 1)
  - Vàng: `Khớp HĐ + mặt hàng (khác ngày CT)` (Khóa 2)
  - Đỏ lệch HĐ/tên: `Ghép theo ngày CT (khác HĐ/mặt hàng)` (Khóa 3)
  - Đỏ thiếu GV: `Không khớp 632`
  - Đỏ thiếu DT: `Có giá vốn, không có doanh thu`
- **S `Năm CT`**: năm lấy theo ngày chứng từ, dùng cho `BaoCao`.

### 4. Quy tắc khớp đã chốt (3 khóa theo thứ tự ưu tiên)

1. **Khóa 1:** Số hóa đơn (bỏ số 0 đầu) + ngày chứng từ + diễn giải mặt hàng.
2. **Khóa 2:** Số hóa đơn + diễn giải; chấp nhận lệch ngày CT (ngày xuất kho khác ngày ghi nhận doanh thu).
3. **Khóa 3:** Cùng ngày chứng từ khi 632 thiếu số HĐ hoặc tên hàng lệch.

Số hóa đơn `0000044` và `44` được coi là cùng một số.

### 5. Đọc bảng tổng hợp `BaoCao`

Sheet tự tách theo năm chứng từ (2018–2030) + cột tổng:

1. Số dòng theo từng tình trạng R.
2. Doanh thu 511, giá vốn 632, lãi gộp, tỷ lệ lãi gộp.
3. Tổng PS Nợ 632 theo tình trạng đối chiếu + tỷ trọng GV khớp đủ 3 khóa.
4. Bán dưới giá: số dòng P < 0, P = 0, P ≤ 0; DT/GV/tổng lỗ của nhóm P ≤ 0; tỷ lệ DT dưới giá / tổng DT.

Công thức `BaoCao` loại dòng bắt đầu bằng `TỔNG`/`TONG` để tránh cộng nhầm dòng tổng.

### 6. Xuất vào hồ sơ kiểm tra

1. `DCDTCP` (chi tiết từng chứng từ) là bảng làm việc trung gian, giữ nội bộ (không phải sheet giao cuối trong bộ v1.4) - dùng để kiểm chứng và truy ngược khi cần.
2. Chép `BaoCao` (tổng hợp theo năm chứng từ) thành sheet `KQ2_511_632` của `KQ2.2_Phan_tich_ho_so_cty_<MST>_<KY>_v<n>.xlsx` (bổ sung 19/10/2026 vòng 2 - thay cho bản liệt kê từng dòng chứng từ trước đó). Cấu trúc: "I. BẢNG TỔNG HỢP ĐỐI CHIẾU THEO NĂM TÀI CHÍNH" - cột `Chỉ tiêu` | `Tổng cộng (kỳ kiểm tra)` | một cột mỗi năm trong kỳ kiểm tra | **`Đánh giá tuân thủ & Rủi ro`** (ngoại lệ 3.4 của skill `phan-tich-ho-so-thue` - sheet này được phép có cột đánh giá). Chỉ tiêu tối thiểu: doanh thu 511, giá vốn 632, lãi gộp, tỷ lệ lãi gộp, số dòng theo từng `Tình trạng đối chiếu` R, tổng PS Nợ 632 theo tình trạng, DT/GV/lỗ nhóm bán dưới giá (P ≤ 0), tỷ lệ DT dưới giá/tổng DT. Chỉ lấy các năm trong kỳ kiểm tra (K12) - năm ngoài phạm vi không in dòng toàn 0 (Q19).
3. Cột `Đánh giá tuân thủ & Rủi ro`: câu ngắn theo số đã tính (vd "lãi gộp 2023 giảm mạnh so 2022, 14 HĐ bán dưới giá tổng 320tr - cần đối chứng từ gốc"), không gán mức rủi ro/điểm (việc đó vẫn ở KQ3).
4. Nếu cần giữ bản chi tiết từng chứng từ để tham khảo: lưu riêng vào `ket-qua/cau-thanh/Doi_chieu_511_632_<MST>_<ky>.xlsx` (hoặc `.xlsm`), KHÔNG phải sản phẩm giao chính.
5. Sao lưu bản chạy vào `ket-qua/sao-luu/` và `7. Sao luu tong hop/` theo quy chuẩn hồ sơ.

## Định dạng số tiền (bắt buộc)

- **Hiển thị** (chat, Markdown, báo cáo, ô text): chuẩn Việt Nam - dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (openpyxl `number_format`): dùng mã invariant `#,##0` (nguyên) hoặc `#,##0.00` (có phần lẻ). Excel locale VN sẽ hiện thành dấu `.` / `,`. Không mô tả nhầm là "dấu phẩy ngăn cách hàng nghìn".
- **Tỷ lệ %**: hiển thị `12,34%` (dấu `,` phần lẻ).
- Cấm dạng Mỹ `1,234,567.89` trong câu trả lời / báo cáo tiếng Việt.

## Bẫy nghiệp vụ cần tránh

- **Sai tên sheet / sai cột G:** Sheet phải đúng `511`/`632`; cột G bên 511 là PS Có, bên 632 là PS Nợ. Đổi chỗ là lãi/lỗ đảo dấu.
- **Xóa nhầm tiêu đề dòng 1:** Macro đọc từ dòng 2; mất header thì map cột sai.
- **Dòng Tổng cộng:** Phải lọc/không dán dòng tổng; `BaoCao` đã loại prefix `TỔNG`/`TONG` nhưng vẫn nên sạch dữ liệu gốc.
- **Lệch ngày CT là bình thường ở một số DN:** Khóa 2 chấp nhận; không kết luận ngay là sai sót nếu chỉ lệch ngày mà HĐ + mặt hàng khớp.
- **Bán dưới giá P ≤ 0:** Có thể do khuyến mại, xuất trả, sai diễn giải ghép nhầm, hoặc cố ý. Phải chọn mẫu chứng từ gốc trước khi kết luận.
- **Macro tắt / chưa import:** Không thấy `ChayDoiChieu` trong Alt+F8 thì làm theo sheet `HuongDan` và `BangDieuKhien`.
- **Ghi đè kỳ cũ:** Mỗi lần chạy ghi đè `DCDTCP`; muốn giữ kỳ trước phải xuất Yes ra file riêng trước khi dán kỳ mới.

## Kiểm chứng kết quả

- Tổng PS Có trên `DCDTCP` (loại dòng tổng) khớp tổng PS Có sheet `511` đã dán.
- Tổng PS Nợ khớp tổng sheet `632` đã dán.
- `BaoCao!B11` (tổng số dòng tình trạng) = tổng các nhóm R; không có dòng R trống bất thường.
- Tỷ lệ lãi gộp và tỷ lệ DT dưới giá có số, không lỗi `#DIV/0!` khi có doanh thu.
- File kết quả đã nằm trong `ket-qua/cau-thanh/` và được dẫn trong báo cáo phân tích hồ sơ.

## Liên kết quy trình hồ sơ

Playbook lệnh: `playbook-lenh.md`.

Skill này là phân hệ chuyên sâu của `phan-tich-ho-so-thue` tại **Bước 4 - đối chiếu 511/632**. Đầu vào thuộc **Bước 2** (sổ chi tiết TK 511, 632 trong `2.1. So ke toan tong hop va chi tiet/` hoặc `2.2. Kho - Chi phi - Gia thanh/`).

## Lịch sử (tự học)
- [2026-10-19 vòng 2] Đổi sheet giao `KQ2_511_632`: không còn liệt kê từng chứng từ, chuyển sang `BaoCao` tổng hợp theo năm + cột `Đánh giá tuân thủ & Rủi ro` (ngoại lệ 3.4). `DCDTCP` chi tiết giữ nội bộ/tham khảo, không phải sheet giao cuối.
- [2026-10-09] Chốt định dạng số tiền chuẩn VN đồng loạt: hiển thị dấu `.` hàng nghìn và `,` phần lẻ; Excel `#,##0` / `#,##0.00`.
