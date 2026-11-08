---
type: reference
provenance: assistant
origin: javis-learned
created: 2026-11-05
updated: 2026-11-05
---
Javis ĐỌC TRỰC TIẾP được file trên Google Drive, không cần người dùng tải xuống hay chia sẻ link public.

**Vì sao:** Kho kết nối đã có sẵn Composio (cổng đa app), trong đó toolkit `googledrive` đang ACTIVE với 4 tài khoản. Đã xác minh thực tế ngày 2026-10-26 bằng chuỗi metadata -> download -> đọc sheet.

**Áp dụng (đọc):**
- Gọi `javis_run_tool(name="composio__COMPOSIO_MULTI_EXECUTE_TOOL", args={tools:[{tool_slug, arguments, account}], sync_response_to_workbench:false})`.
- Tool hay dùng: `GOOGLEDRIVE_GET_FILE_METADATA` (xác minh quyền + mimeType), `GOOGLEDRIVE_DOWNLOAD_FILE` (trả `downloaded_file_content.s3url`, hết hạn sau 1 giờ -> `curl` ngay về vault), `GOOGLEDRIVE_FIND_FILE` (tìm theo tên khi chưa có id).
- **PHẢI truyền `account`** vì có nhiều tài khoản. Tài khoản của người dùng: `googledrive_amala-ensign` và `googledrive_ceile-idly` (cả hai = ntangoc2416.1@gmail.com). Mặc định của hệ thống là `googledrive_ralph-urbane` (cungcapthongtin.tcs3tdn@gmail.com), KHÁC drive người dùng, truyền sai sẽ ra notFound.
- `fileId` lấy từ đoạn sau `/d/` trong link Drive.
- Giới hạn: toolkit `googlesheets` CHƯA đấu, các tool `GOOGLESHEETS_*` không chạy. File `.xlsx` thì tải nguyên bản qua Drive rồi đọc bằng `openpyxl`.
- `openpyxl` và `python-docx` nằm ở user site-packages: KHÔNG chạy `python3 -I` (cờ này chặn mất), chỉ cần chạy từ thư mục khác thư mục chứa file tải về.

**Giới hạn ghi (thực tế tính đến 05/11/2026):**
- Kết nối Composio `perm: full`, có tool `GOOGLEDRIVE_UPLOAD_UPDATE_FILE` và `GOOGLEDRIVE_RESUMABLE_UPLOAD`, nhưng cả hai bắt buộc `file_to_upload.s3key` (file phải nằm sẵn trong kho file Composio), không nhận đường dẫn file trong vault; truyền sai ra lỗi 403 "file key does not belong to this project".
- `GOOGLEDRIVE_EDIT_FILE` chỉ nhận chữ UTF-8, đưa docx vào là hỏng file.
- Đường vòng gõ base64 qua sandbox Composio đã thử và hỏng (md5 lệch), không dùng cho file nhị phân.
- Đường vòng đẩy file qua dịch vụ chia sẻ công khai rồi cho sandbox Composio tải về: khả thi nhưng đưa dữ liệu NNT ra ngoài, KHÔNG tự làm, chỉ làm khi user duyệt rõ (05/11/2026 chưa duyệt).
- Việc user đóng Excel KHÔNG mở được đường ghi; ngược lại, khi user lưu sổ trong Excel, bản Drive sẽ đè mất các ô Javis đã sửa ở bản vault.
- Lối mở đề xuất (chưa làm): user cấp Composio API key làm biến môi trường để curl POST file lấy s3key; hoặc bổ sung tool hub kiểu upload file vault. Trong lúc chưa mở, user tự tải file từ vault lên Drive.

Liên quan: [[quy-uoc-soan-bien-ban-ho-so-ve]], [[ho-so-kiem-tra-mst-3603859012]], [[ho-so-kiem-tra-mst-3603721832]]
