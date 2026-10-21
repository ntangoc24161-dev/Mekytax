---
name: "Đối chiếu tờ khai GTGT và TNDN chuẩn CBRR"
description: "Đối chiếu Bảng kê - Tờ khai GTGT - Quyết toán TNDN - HĐĐT chuẩn 2 sheet CBRR, bóc tách nguyên nhân trọng yếu và xuất Excel."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-10-01
updated: 2026-10-01
---
## Khi nào dùng

- Đối chiếu toàn diện nghĩa vụ thuế **GTGT** (từng tháng/quý) và **Quyết toán TNDN** (theo năm) giữa Bảng kê mua vào/bán ra, Tờ khai thuế (Mẫu 01/GTGT, Mẫu 03/TNDN) và Hóa đơn điện tử (HĐĐT) qua nhiều năm (ví dụ 10 năm: 2016 - 2025).
- Xuất file Excel báo cáo đối chiếu chuẩn **2 Sheet toàn diện** theo quy chuẩn kỹ thuật số liệu công văn CBRR (`BANG_SO_LIEU_CONG_VAN.md`):
  + **Sheet 1 (`Tong_hop_doi_chieu`)**: Đối chiếu GTGT (Bảng 1, 2, 3), Doanh thu Quyết toán TNDN (Mẫu 4 CBRR), Bảng tổng hợp 23 chỉ tiêu nguyên nhân trọng yếu qua các năm kiểm tra và đề xuất trọng tâm kiểm tra.
  + **Sheet 2 (`Doi_chieu_NCC_CVCBRR`)**: Đối chiếu Người nộp thuế (NNT) & toàn bộ Nhà cung cấp mua vào với Sổ CVCBRR F0/F1, Sổ Công an, Danh bạ NNT và Bảng tổng hợp số liệu chi tiết các NCC bất thường theo từng năm.
- Tự động gom và đồng bộ toàn bộ file báo cáo chính thức, file sao lưu trung gian, file thuyết minh Markdown và dữ liệu JSON vào thư mục sao lưu chuẩn (`7. Sao luu tong hop` trên máy / `sao-luu/` trong Second Brain).

## Nguồn dữ liệu & Chuẩn bị

1. **Tờ khai thuế / BCTC**: File `TONG HOP_TO KHAI.xlsx` (sheet `GTGT 01`, `QT TNDN 03`) hoặc XML tờ khai đã giải mã.
2. **Bảng kê mua vào / bán ra**: File `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx` (sheet `ban ra`, `mua vao` 10 năm đã gộp).
3. **Dữ liệu HĐĐT portal**: Thư mục `BAN RA`, `MUA VAO` chứa file Excel HĐĐT chi tiết.
4. **Dữ liệu rủi ro & Danh bạ**: Sổ CVCBRR (`SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` - Sheet `F0`, `F1`), Sổ Công an (`Bao_cao_cong_van.xlsx`), và Danh bạ NNT (trạng thái ĐKKD).
5. Xác định đúng MST người nộp thuế và thời kỳ kiểm tra chính thức (loại bỏ các năm ngoài kỳ kiểm tra, ví dụ 2015).

## Quy chuẩn Báo cáo Đối chiếu 2 Sheet chuẩn CBRR

Toàn bộ báo cáo được tích hợp trên 2 sheet đồng bộ:

### SHEET 1: `Tong_hop_doi_chieu` (GTGT, TNDN & 23 Chỉ tiêu Nguyên nhân)

#### PHẦN I. ĐỐI CHIẾU NGHĨA VỤ THUẾ GTGT THEO KỲ (40 QUÝ / 10 NĂM)
- **Bảng 1 (Mẫu 3 CBRR - 7 cột)**: Kê khai thuế GTGT theo kỳ (Kỳ, GT mua vào [23], Thuế mua vào [24], Tổng DT [34], Thuế đầu ra [35], PS dương [40], Còn khấu trừ [43]/[41]). Dòng **CỘNG** chỉ tính tổng cột 2–5; cột 6–7 để trống.
- **Bảng 2 (Mẫu 5 CBRR mở rộng - 7 cột)**: Đối chiếu Bán ra theo kỳ (Kỳ, Tờ khai [34]/[35], HĐĐT hợp lệ DT/Thuế, Chênh lệch DT/Thuế = Tờ khai − HĐĐT).
- **Bảng 3 (Mẫu 6 CBRR mở rộng - 7 cột)**: Đối chiếu Mua vào theo kỳ (Kỳ, Tờ khai [23]/[24], Bảng kê mua vào DT/Thuế, Chênh lệch DT/Thuế = Tờ khai − Bảng kê).

#### PHẦN II. ĐỐI CHIẾU DOANH THU & QUYẾT TOÁN TNDN THEO NĂM (Mẫu 4 CBRR)
- Cột: STT, Năm, Doanh thu Bảng kê (1), HĐĐT hợp lệ (2) *(ghi chú tham khảo trước 2023)*, TNDN Tổng DT [01] (3), Giảm trừ DT [02], TNDN DT thuần [03], Chênh lệch BK vs TK [01] (1)-(3), Thu nhập chịu thuế, Thuế TNDN phải nộp, Đánh giá & Ghi chú.
- Dòng CỘNG: Tính tổng các cột số liệu; đánh giá độ khớp 100% giữa Bảng kê vs Chỉ tiêu [01].

#### PHẦN III. BẢNG SỐ LIỆU TỔNG HỢP NGUYÊN NHÂN TRỌNG YẾU THEO NĂM (23 CHỈ TIÊU)
- **Nhóm A (Doanh thu & Quyết toán TNDN)**:
  1. Tổng Doanh thu Bán ra (Bảng kê)
  2. Doanh thu Quyết toán TNDN [01]
  3. Giảm trừ doanh thu [02]
  4. Doanh thu thuần TNDN [03]
  5. Chênh lệch BK vs TNDN [01]
  6. Thu nhập chịu thuế TNDN
  7. Số thuế TNDN phải nộp
- **Nhóm B (Hóa đơn điện tử HĐĐT)**:
  8. Doanh thu HĐĐT hợp lệ
  9. Lệch do chưa áp dụng NĐ 123 (trước T7/2022)
  10. Giá trị HĐ thay thế / hủy bỏ (bóc tách theo từ khóa)
- **Nhóm C (Mua vào & Thuế khấu trừ)**:
  11. Tổng Mua vào theo Bảng kê
  12. Mua vào Tờ khai GTGT [23]
  13. Lệch do sai sót hình thức / gõ phím
  14. Thuế GTGT mua vào Bảng kê
  15. Thuế GTGT mua vào [24]
  16. Lệch Thuế BK vs [24]
- **Nhóm D (Rủi ro Nhà cung cấp)**:
  17. Mua vào từ NCC CVCBRR F0
  18. Mua từ các NCC bất thường (đóng MST, tạm ngừng KD, giải thể...)
  19. Tỷ trọng NCC bất thường (%)
  20. Mua từ NCC bình thường
  21. Tỷ trọng NCC bình thường (%)

#### PHẦN IV. ĐỀ XUẤT TRỌNG TÂM VÀ PHƯƠNG PHÁP KIỂM TRA TẠI DOANH NGHIỆP
- Trọng tâm kiểm tra các NCC có tỷ trọng lớn hoặc có dấu hiệu rủi ro (yêu cầu hợp đồng, biên bản giao nhận, phiếu nhập kho, định mức, ủy nhiệm chi ngân hàng).
- Xác minh nhóm NCC tạm ngừng kinh doanh hoặc đóng MST trước/sau thời điểm xuất hóa đơn.
- Giải trình và xử lý biên bản đối với các lỗi gõ phím / sai sót hình thức trên tờ khai.
- Đôn đốc nộp bổ sung tờ khai chính thức đối với các kỳ còn thiếu XML.

---

### SHEET 2: `Doi_chieu_NCC_CVCBRR` (Đối chiếu NNT & NCC với CVCBRR, Sổ CA, Danh bạ)

#### PHẦN 1. ĐỐI CHIẾU NGƯỜI NỘP THUẾ ĐANG XÉT
- Bảng đối chiếu NNT với Sheet `F1` (CVCBRR Bên mua) và Sổ Công an.
- Kết luận rõ ràng NNT có hay không có trong danh sách cảnh báo hoặc hồ sơ Công an.

#### PHẦN 2. ĐỐI CHIẾU DANH SÁCH NHÀ CUNG CẤP MUA VÀO (10 NĂM)
- **Bảng 2.1: Đối chiếu với Sổ CVCBRR F0 (Bên bán)**: STT, MST, Tên NCC, Số HĐ, Doanh số, Thuế GTGT, Kỳ mua, Mặt hàng chính, Nội dung cảnh báo & Số CV, Trạng thái ĐKKD & Đánh giá rủi ro.
- **Bảng 2.2: Đối chiếu với Danh bạ NNT (Trạng thái bất thường)**: STT, MST, Tên NCC, Số HĐ, Doanh số, Thuế GTGT, Kỳ mua, Trạng thái ĐKKD hiện tại.
- **Bảng 2.3: Bảng tổng hợp số liệu chi tiết các NCC bất thường theo từng năm**: Bóc tách doanh số mua vào từng năm của từng NCC bất thường, kèm dòng Tổng cộng 7 NCC bất thường, Toàn bộ mua vào 10 năm, và Tỷ trọng %.

#### PHẦN 3. ĐỀ XUẤT TRỌNG TÂM KIỂM TRA NHÀ CUNG CẤP TẠI NNT
- Đề xuất chi tiết biện pháp xác minh giao dịch thực tế đối với từng nhà cung cấp có rủi ro.

## Quy chuẩn Thư mục Sao lưu & Lưu trữ Hồ sơ

- **Thư mục trên máy (Hồ sơ kiểm tra)**: `7. Sao luu tong hop` (theo hệ thống đánh số chuẩn: `0. BIEN BAN`, `1. Tong quan chung`, `2. Phan tich doc ngang`, `3. HDDT`, `4. TO KHAI XML`, `5. Bang ke mua vao ban ra`, `6. TNCN`, `7. Sao luu tong hop`).
- **Thư mục đồng bộ trong Second Brain**: `sao-luu/`
- **Danh mục file cần gom vào thư mục sao lưu**:
  1. **File Excel tổng hợp chính thức**: `Tong_hop_doi_chieu_GTGT_TNDN_*.xlsx` (2 sheet) và `Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx`.
  2. **File sao lưu dữ liệu trung gian**: `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx`, `Bang tinh_backup_*.xlsx`, `Tong_hop_doi_chieu_GTGT_TNDN_1sheet.xlsx`, `doi-chieu-hddt-*.xlsx`, `phan-tich-*.xlsx`.
  3. **File tài liệu thuyết minh & cấu trúc**: Các file `.md` thuyết minh số liệu và các file `.json` cấu trúc máy.

## Kỹ thuật Excel & Tránh bẫy XML

1. **Tuyệt đối tránh lỗi Formula XML**: Không bao giờ để nội dung ô văn bản bắt đầu bằng dấu `=`, `+`, `-` (ví dụ `> KẾT LUẬN...` thay vì `=> KẾT LUẬN...` hoặc set tường minh `cell.data_type = 's'`). Nếu bắt đầu bằng `=`, openpyxl sẽ lưu thành thẻ formula `<f>` và Excel sẽ báo lỗi *Removed Records: Formula from /xl/worksheets/sheet1.xml part*.
2. **Định dạng**: Font **Times New Roman** 100% file; tiêu đề cột in đậm; số tiền định dạng phân cách hàng nghìn chuẩn VN (`#,##0`); dòng CỘNG dùng hàm `SUM(...)`.
3. **Hiển thị đầy đủ**: Bật `wrap_text = True` cho các cột diễn giải/ghi chú, tự căn chỉnh chiều cao hàng và độ rộng cột để không bị khuất chữ.

## Bẫy thường gặp & Bài học rút ra

- **Năm ngoài kỳ kiểm tra**: Lọc bỏ các năm nằm ngoài phạm vi quyết định kiểm tra (như 2015) khi lập bảng tổng hợp 10 năm (2016 - 2025).
- **Lọc hóa đơn hủy / thay thế**: Luôn áp dụng bộ lọc từ khóa (`CANCEL_KEYWORDS = ['thay thế', 'xóa bỏ', 'hủy', 'bị thay thế']`) để loại trừ hóa đơn không hợp lệ khỏi số liệu HĐĐT đối chiếu.
- **Lỗi gõ phím trên tờ khai**: Kiểm tra chéo giữa ct23 và ct24 (nếu ct24 khớp đúng số thuế nhưng ct23 lệch do gõ sót số thì đây là lỗi hình thức, số thuế khấu trừ vẫn chính xác).
- **Kỳ thiếu tờ khai XML**: Sử dụng Bảng kê bán ra và HĐĐT hợp lệ để làm căn cứ đối chiếu số liệu, ghi chú rõ để yêu cầu NNT bổ sung tờ khai chính thức.
- **Chuẩn hóa MST**: Loại bỏ ký tự thừa, so khớp cả MST 10 số và MST chi nhánh 13 số khi đối chiếu Sổ CVCBRR và Danh bạ NNT.
