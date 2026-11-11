---
name: Theo dõi công văn cảnh báo rủi ro
description: Bóc tách CV cảnh báo thuế và nạp độc lập sheet F0 (19 cột) và F1 (25 cột) sổ SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.
group: Tài chính
updated: '2026-11-09'
learned_updated: '2026-11-09'
---
## Khi nào dùng
- Nạp dữ liệu doanh nghiệp rủi ro bán ra vào Sheet F0 hoặc doanh nghiệp mua vào Sheet F1 của Sổ theo dõi `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`.
- Nhận hồ sơ công văn cảnh báo từ thư mục `HO SO MOI` hoặc bảng kê Excel đi kèm.

## Chuẩn bị
- Đường dẫn sổ theo dõi: `...\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`.
- Sao lưu sổ sang thư mục `Sao lưu` trước khi thực hiện ghi đè hoặc chèn dòng.
- Xác định đúng cấu trúc cột thực tế: F0 có 19 cột, F1 có 25 cột; tiêu đề nằm ở dòng 5-6 dạng gộp ô.

## Quy trình
1. Kiểm tra và mở file sổ theo dõi bằng openpyxl (chế độ đọc dữ liệu để đếm dòng trước).
2. Tạo bản sao lưu file Excel vào thư mục `Sao lưu` với tiền tố ngày giờ YYYYMMDD_HHMMSS.
3. Trích xuất dữ liệu công văn / tệp Excel đính kèm (ưu tiên đọc bảng Excel trực tiếp thay vì OCR nếu có sẵn).
4. Khử trùng dữ liệu theo khoá chốt 09/11/2026: F0 theo `SoCV + MSTBan`; F1 theo `SoCV + MSTBan + MSTMua`.
5. Nạp dòng mới vào F0 (19 cột) hoặc F1 (25 cột) bắt đầu từ dòng kế tiếp dòng dữ liệu cuối cùng. **F1 bỏ qua hóa đơn không có MST bên mua** (chốt 09/11/2026), không nạp dòng trống `MSTMua`, phải đếm và báo số dòng đã bỏ.
6. Đếm lại số dòng sau khi ghi và xác nhận biến động (F0 trước/sau, F1 trước/sau, số bỏ trùng).

## Bẫy
- Bẫy cấu trúc cột: F0 có 19 cột và F1 có 25 cột, không áp dụng quy chuẩn 25 cột cho cả hai sheet kẻo làm lệch cột dữ liệu.
- Bẫy tiêu đề: Tiêu đề không nằm ở dòng 1 mà ở dòng 5-6 dạng gộp ô 2 tầng.
- Bẫy OCR: Văn bản scan mờ/mất dấu tiếng Việt hoặc mất ký tự số công văn/MST phải đối chiếu tệp Excel đính kèm hoặc nguồn số hóa, không nạp chuỗi OCR rác vào sổ.
- Bẫy sao lưu: Bắt buộc tạo file backup trước khi ghi sổ để tránh mất mát dữ liệu khi tiến trình lỗi.
- Bẫy nguồn chuẩn (chốt 09/11/2026): lô nào có tệp `.xlsx` danh sách doanh nghiệp thì **Excel là nguồn chuẩn** cho MST, tên công ty, địa chỉ, ngành nghề; OCR chỉ dùng lấy `SoCV`, `NgayCV`, `DonViBH`, `NoiDung`.
- Bẫy mốc số dòng: mốc đúng tại 09/11/2026 là F0 **17.075** dòng, F1 **3.761** dòng. Các số 17.333 / 16.952 / 19.638 / 20.050 / 7.282 trong lịch sử là số cũ, không dùng làm mốc đối chiếu.

## Kiểm chứng
- Kiểm tra số dòng F0, F1 tăng chính xác bằng số bản ghi hợp lệ sau khử trùng.
- Kiểm tra cột MST và Tên công ty điền đúng vị trí quy định.

## Lịch sử (tự học)
- [2026-09-22] Xử lý hồ sơ scan 31 trang `[1370_TCS7-QLDN1] 1370TCS7-QLDN1 1298.pdf`, bóc tách 02 công văn cảnh báo (CV 1370 của Thuế cơ sở 7 TP.HCM về Cty Hưng Thành; CV 1432 về Cty Gervina), bổ sung 03 dòng dữ liệu vào sổ.
- [2026-09-23] Đóng gói toàn diện quy trình thành skill chuẩn `theo-doi-cv-canh-bao-rui-ro` lưu tại brain và kho kỹ năng chung `shared-skills`, kèm module tự động hóa `scripts/update_so_theo_doi.py`.
- [2026-09-26] Nâng cấp pipeline tự động hóa cho 57 CV scan tháng 08/2026 và công văn số hóa tháng 06-07/2026. Khắc phục nghẽn XML openpyxl bằng `iter_rows(values_only=True)`, tích hợp cache checkpoint và giám sát Watchdog.
- [2026-09-28] Lô iOffice 22-26/09: từ 36 tệp giữ 24 tệp độc nhất, chuyển đổi OCR sang `fast_batch_ocr`, hạ cấp `python-bidi` về 0.4.2 tránh lỗi DLL.
- [2026-09-30] Hoàn tất đồng bộ dữ liệu bên bán 524 DN từ sheet `CV cong an-ben ban`, khắc phục 499 dòng lỗi font cụm CV 1798, hoàn thiện 108 cty vụ Đào Minh Thọ, nạp 72 cty mới từ Công an Hà Tĩnh và Thanh Hóa.
- [2026-09-30] Tái cấu trúc toàn diện Sổ theo dõi: phân tách thành 2 sheet độc lập `F0` (20.581 dòng bên bán) và `F1` (7.845 dòng bên mua), xóa bỏ các sheet cũ `THEO_DOI`, `TONG_HOP`, `DANH_MUC`, `HUONG_DAN`, đồng bộ toàn bộ font Times New Roman, giảm dung lượng tệp về 2,12 MB và quy hoạch toàn bộ bản sao lưu vào thư mục `00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\Sao lưu`. Đồng thời nâng cấp toàn bộ skill, script và agent theo chuẩn 2 sheet mới.
- [2026-10-01] Làm sạch lại dữ liệu nguồn xlsb: thay 3.351 dòng CBRR cũ trên F0 bằng 2.939 dòng map đúng A/B/E/F; thay 3.649 dòng TB DO KT1 MOI cũ trên F1 bằng 3.087 dòng map đúng B/C/M/N (không gán bên mua vào cột bên bán). F0 còn 19.638 dòng, F1 còn 7.282 dòng.
- [2026-10-01] Tiếp tục không làm lại toàn sổ: xóa 1 MST ngoài `524.xlsx` (`0314123789` trên CV 1798/TCT-TTKT), gộp 2.115 dòng 524 trùng (1770 / 1798.0 / 1798/TCT-TTKT / HDRR trống số CV) thành 637 dòng chuẩn. F0 còn 17.333 dòng, F1 giữ 3.704 dòng.
- [2026-10-01] Bổ sung F0 từ CBRR/HDRR-cuc/CV cong an: điền 521 dòng trống SoCV từ nguồn (CBRR dùng E/F đúng header, không dùng C/D), thêm 3 dòng mới, gộp 407 trùng; còn 372 dòng nguồn không có số CV (chủ yếu CV cong an chỉ ghi tên cơ quan). F0 còn 16.931 dòng.
- [2026-11-09] Sửa lại đặc tả cấu trúc cột thực tế (F0 có 19 cột, F1 có 25 cột, tiêu đề dòng 5-6) sau khi phát hiện lỗi lệch cột theo phản ánh thực tế từ hội thoại 09/11/2026.
