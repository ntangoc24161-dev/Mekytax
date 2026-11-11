---
type: business
provenance: assistant
origin: javis-learned
created: 2026-11-10
updated: 2026-11-10
---
Thực tế tính đến 10/11/2026, ngoài 4 cụm Excel, `HO SO MOI` còn 28 tệp chưa nạp sổ:
- 19 hồ sơ PDF scan/văn bản đơn lẻ chưa chạy OCR (không kèm bảng kê .xlsx): `5645.pdf`, `cv 9955.pdf`, `8835-CV-05112026.pdf`, `6210.pdf`, `12215.pdf`, `8746-CV-04112026.pdf`, `6760TCS15-QLDN1 14445.pdf`, `6136.pdf`, `Cong_van_CBRR_3600930041.pdf`, `CV cảnh báo rủi ro - QLDN THÀNH CÔNG VN.pdf`, `CV cảnh báo rủi ro Cty VĂN THƯƠNG.pdf`, `3. CV cảnh báo rủi ro.pdf`, `VB CANH BAO RUI RO_VAN TAI BA NA.pdf`, `3801351735_Cty Tong Hop Pham Cong_CBRR.pdf`, `CBRR_VPK BAVI.pdf`, `CV canh bao rui ro-CT IN AN TRONG HUY HOANG.pdf`, `CV canh bao rui ro_Trung Dung.pdf`, `VB CANH BAO RUI RO_CTY MEDONG.pdf`, `CV canh bao rui ro_Chinh Hung.pdf` (kèm `Bang ke hoa don mua vao ban ra Chinh Hung.pdf`).
- 8 tệp lọc bỏ: 4-5 bản sao hậu tố `(1)`/`(2)`; 3 văn bản đi kèm cụm 7268 (`CV 7268.pdf`, `TT CV 7268.pdf`, `TRIEN KHAI CV 7268 ve 28DN RUI RO .signed.signed.pdf`) vì 28 DN đã lấy từ Excel; `Trang trắng.pdf` (11 KB, trống).
- 1 tệp loại nghiệp vụ: `DANH SACH BAN RA KHÔNG CO MST.xlsx` (không MST bên mua, không nạp F1).

**Vì sao:** ưu tiên cụm Excel trước nên chưa kích hoạt luồng OCR cho nhóm PDF.
**Áp dụng:** việc tồn tiếp theo là OCR 19 hồ sơ PDF (paddleocr/pdf-sang-md) rồi nạp F0/F1.
