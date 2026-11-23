## Bài học (tự học)
- Trước khi `set-status`, tải sổ sống từ Drive và tìm đúng dòng/folder theo số CV; nhật ký cũ ghi "đã ghi dòng N" không đảm bảo dòng đó còn trong sổ sống (sổ từng bị khôi phục, mất dòng).
- PDF ký số có lớp chữ trống ô Số/ngày vẫn có thể là bản ký thật; dùng `pypdf` đọc `/AcroForm` `/Sig` (tên người ký, `/M` ngày ký) để xác nhận thay vì chỉ dựa OCR.
- Khi không thấy CV trong sheet `NHAP_HO_SO`, phải quét cả sheet `XL KET LUAN` (hồ sơ đã hoàn thành được chuyển sang đó, cột A là STT, cột B là folder) trước khi kết luận sổ thiếu dòng.
- Dùng Composio workbench với `account='googledrive_ralph-urbane'` để tải sổ sống (`GOOGLEDRIVE_DOWNLOAD_FILE` trả `s3url`), tính md5 và quét mọi sheet bằng openpyxl; không có account thì lỗi "Multiple googledrive accounts".
