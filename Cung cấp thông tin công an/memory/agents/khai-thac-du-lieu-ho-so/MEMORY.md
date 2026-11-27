## Bài học (tự học)
- File live `DB NNT.xls` hiện là SpreadsheetML XML, chỉ có sheet NNT (cột 9 Tổ KT, cột 11 CBQL); sheet KTR1 nằm ở `DB NNT 03.06.2026.xlsx`.
- PDF PaperStream/CamScanner không có lớp chữ, phải OCR rồi đối chiếu ảnh để chốt số hiệu (filename 21143, OCR từng đọc 24143).
- vietqr.io/v2/business/{MST} là API tra cứu MST thay thế khi masothue.com trả 403 - trả JSON chuẩn gồm tên DN, địa chỉ, tình trạng pháp lý, nguồn gdt.gov.vn.
- Đọc sổ XLSX lớn trên Drive: dùng COMPOSIO_MULTI_EXECUTE_TOOL GOOGLEDRIVE_DOWNLOAD_FILE (account=alias) → s3url → requests.get + openpyxl.load_workbook trong COMPOSIO_REMOTE_WORKBENCH; sandbox giữ state giữa các cell.
- Trong COMPOSIO_MULTI_EXECUTE_TOOL, `account` phải là sibling của `tool_slug` (cùng cấp tools[]), không nhét trong `arguments` nếu không sẽ bị lỗi Multiple googledrive accounts.
- Một file PDF có thể là bó nhiều công văn (44321 gồm 10 DN). Đếm trang và đọc số từng trang trước khi tính số dòng sổ.
- `ghi-so-bao-cao-an-toan.py` chỉ ghi chuỗi, không ghi được ô ngày dạng số. Cần đường ghi ngày riêng trước khi ghi cột D, G, Q, R.
- Ghi ngày vào sổ phải dùng số serial Excel (lệnh set-dates), không ghi chuỗi dd/mm/yyyy; upload Drive qua upload_local_file lấy s3key rồi GOOGLEDRIVE_UPLOAD_UPDATE_FILE, và kiểm md5 Drive trước và sau khi ghi.
