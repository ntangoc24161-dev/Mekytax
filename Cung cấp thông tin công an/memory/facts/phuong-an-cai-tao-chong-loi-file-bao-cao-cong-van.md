---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-22
updated: 2026-11-22
---
Người dùng đã thống nhất phương án cải tạo tối ưu chống hỏng file `Bao_cao_cong_van.xlsx`:
1. **Phân định vai trò:** VPS Linux chỉ bóc tách OCR và đẩy dữ liệu thô vào hàng đợi trung gian (JSON/CSV tạm), tuyệt đối không dùng openpyxl ghi thẳng file chính trên Drive để tránh hỏng Dashboard và mất cache công thức. Javis Desktop Windows dùng Excel COM (`win32com.client`) đọc hàng đợi và nạp vào sổ chính.
2. **Chuyển đổi Excel Table (ListObject):** Chuyển 3 sheet dữ liệu chính (`NHAP_HO_SO`, `TAM_NGUNG`, `XL KET LUAN`) thành dạng Excel Table độc lập để tự động tràn công thức (`Calculated Column`) và định dạng cho dòng mới.
3. **Giữ file nhẹ và chống gãy liên kết:** Thay thế các hàm biến động nặng (`OFFSET`, `INDIRECT`, `TODAY()` lặp) bằng tham chiếu cấu trúc `[@Cot]` và `XLOOKUP`/`INDEX-MATCH`; dán cứng giá trị tĩnh (Paste Values) cho hồ sơ cũ/đã hoàn tất; không format tràn cả cột; khi chuyển sheet chỉ dùng `ClearContents`, không dùng `Delete Row` tránh lỗi gãy tham chiếu `#REF!`.
