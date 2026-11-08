---
type: preference
provenance: user
origin: user-instruction
created: 2026-10-22
updated: 2026-10-22
---
# Quy tắc mặc định xử lý file PDF và ảnh scan (OCR)
Khi gặp file PDF, ảnh chụp tài liệu hoặc file scan cần OCR:
- Chỉ đọc chữ ra chat: Mặc định dùng skill paddleocr (MCP tool ocr).
- Ghi file Markdown giữ heading/bảng vào sources/: Mặc định dùng skill pdf-sang-md (	ools.doc2md).
- Xuất Office (Word .docx, Excel .xlsx, TXT, searchable PDF), xoay/ghép/tách/so sánh file: Mặc định dùng skill 	ram-tai-lieu (	ools.doctram).
- Tuyệt đối không bịa thông tin khi file scan mờ hoặc không trích xuất được text layer.
