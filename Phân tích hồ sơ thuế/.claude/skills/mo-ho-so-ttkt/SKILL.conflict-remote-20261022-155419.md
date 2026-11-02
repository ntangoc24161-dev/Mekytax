---
name: Mở hồ sơ TTKT
description: "Checklist mở hồ sơ TTKT: chọn đối tượng, đối chiếu BCTC-tờ khai-HĐĐT, GDLK, phân nhánh thiếu thuế/trốn thuế."
group: Tài chính
updated: '2026-10-09'
---

# Mở hồ sơ thanh tra - kiểm tra thuế

Skill phương pháp cho công chức / trợ lý AI khi **mở và rà hồ sơ** trước hoặc trong thanh tra - kiểm tra thuế. Không thay engine XML `phan-tich-ho-so-thue`; bổ sung lớp đọc nghiệp vụ từ chuyên đề BCTC TTKT. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "mở hồ sơ TTKT", "checklist kiểm tra thuế", "rà BCTC trước kiểm tra", "đối chiếu tờ khai hóa đơn", "lập kế hoạch nội dung kiểm tra".
- Có hoặc sắp có bộ hồ sơ: BCTC, tờ khai GTGT/TNDN/TNCN, dữ liệu HĐĐT, phụ lục GDLK.
- Không dùng khi chỉ cần chạy tool XML/Excel tự động → dùng `phan-tich-ho-so-thue`.
- Không dùng khi sắp xếp thư mục 4 bước / checklist đủ thiếu file → dùng `to-chuc-ho-so-kiem-tra-thue`.
- Không dùng khi hỏi xử phạt mức tiền cụ thể → dùng `xu-phat-vphc-thue-hoa-don`.
- Không dùng khi hỏi khung QLRR / bộ chỉ số chọn NNT → dùng `quan-ly-rui-ro-tuan-thu-thue`.

## Chuẩn bị

1. Đọc checklist 1 trang: `wiki/Checklist-mo-ho-so-TTKT.md` (trong brain này).
2. Thu thập đường dẫn file thật trên đĩa; liệt kê cái **đang thiếu**.
3. Xác định kỳ kiểm tra và loại NNT (DN / hộ / cá nhân; MST 03/06 nếu có).
4. Nếu user đã chạy `phan-tich-ho-so-thue`: lấy sheet Tóm tắt làm đầu vào, rồi đào sâu thủ công các red flag.

## Cách chạy

Làm tuần tự mục A→F trong checklist. Mỗi mục ghi kết luận ngắn + số liệu (nếu có). Thiếu chứng từ thì ghi "thiếu" và đề xuất bước lấy thêm; không bịa số.

Chi tiết kỹ thuật đối chiếu xem `references/doi-chieu-bctc-to-khai.md`.

## Quy trình

1. **Khung pháp lý nhanh**: tách 3 lớp Kế toán (VAS/TT 200-133) / GTGT-hóa đơn / TNDN. Điểm lệch giữa 3 lớp là đầu mối rủi ro.
2. **GTGT + HĐĐT**: map chỉ tiêu 01/GTGT với HĐĐT bán ra/mua vào theo suất; soi bỏ sót, điều chỉnh/thay thế, thời điểm lập HĐ.
3. **BCTC**: đọc ý kiến kiểm toán; soi BCĐKT (tiền, phải thu, tồn kho, TSCĐ, nợ phải trả, vốn); soi BCKQKD (DT, giảm trừ, GVHB, CP BH/QLDN); soi LCTT vs LN trước thuế.
4. **TNDN**: từ LN trước thuế lần B1→B12 trên 03/TNDN; ưu tiên B4, B6, lãi vay vượt khống chế, BĐS, miễn thuế, chuyển lỗ.
5. **GDLK** (nếu có): quan hệ liên kết NĐ 132/2020; khoảng giá độc lập 35-50-75; hồ sơ Local/Master/CbCR.
6. **Phân nhánh**: thủ tục / khai sai 20% / trốn thuế 1-3 lần / cân nhắc hình sự (Điều 200, 203, 204 BLHS).
7. **Đầu ra**: bảng lệch + 5-10 vấn đề Cao/Trung/Thấp + phạm vi đề xuất kiểm tra + danh mục còn thiếu.

## Định dạng số tiền (bắt buộc)

- **Hiển thị** (chat, Markdown, báo cáo, ô text): chuẩn Việt Nam - dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (openpyxl `number_format`): dùng mã invariant `#,##0` (nguyên) hoặc `#,##0.00` (có phần lẻ). Excel locale VN sẽ hiện thành dấu `.` / `,`. Không mô tả nhầm là "dấu phẩy ngăn cách hàng nghìn".
- **Tỷ lệ %**: hiển thị `12,34%` (dấu `,` phần lẻ).
- Cấm dạng Mỹ `1,234,567.89` trong câu trả lời / báo cáo tiếng Việt.

## Bẫy

- Coi BCTC = tờ khai TNDN: sai; luôn có lớp điều chỉnh B1-B12.
- Phạt chồng HĐ không hợp pháp khi đã xử khai sai/trốn thuế về cùng bản chất.
- Bỏ thời hiệu rồi vẫn đề xuất phạt tiền thủ tục.
- Slide/PDF đào tạo có nhiều link file phụ ngoài máy: chỉ dùng nội dung có trong hồ sơ user đưa.
- Không nhầm skill này với engine `phan-tich-ho-so-thue` (chạy file), `to-chuc-ho-so-kiem-tra-thue` (thư mục + checklist file), hay skill xử phạt.

## Kiểm chứng

- Đã đi hết checklist A→F hoặc nêu rõ mục nào bỏ vì thiếu hồ sơ.
- Mọi số liệu trích từ file/tool; chỗ thiếu ghi rõ.
- Có phân nhánh xử lý (thủ tục / thiếu thuế / trốn thuế) và không dùng em dash.

## Lịch sử (tự học)
- [2026-10-09] Chốt định dạng số tiền chuẩn VN đồng loạt: hiển thị dấu `.` hàng nghìn và `,` phần lẻ; Excel `#,##0` / `#,##0.00`.
- [2026-10-14] Phân biệt với skill `to-chuc-ho-so-kiem-tra-thue` (thư mục 4 bước + checklist đủ/thiếu file).
