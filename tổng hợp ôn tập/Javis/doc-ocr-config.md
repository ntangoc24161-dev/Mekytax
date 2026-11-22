---
type: config
---
# Cấu hình skill doc-ocr (Đọc PDF và OCR)

Mỗi máy một mục. Skill dùng đường dẫn đầu tiên tồn tại trên máy đang chạy, bỏ qua mục của máy khác.
File này đồng bộ giữa các máy qua sao lưu GitHub của brain, nên giữ đủ mục cho mọi máy.

## Máy Windows (máy chạy Tool OCR)
- ket_qua: D:\03. Ung dung\Tool OCR\ket_qua_thu
- inbox: D:\03. Ung dung\Tool OCR\inbox
- lenh_ocr: "D:\03. Ung dung\Tool OCR\.venv\Scripts\python.exe" "D:\03. Ung dung\Tool OCR\hot_folder_cli.py" --once --mirror "D:\SECOND BRAIN\javis-os\brains\Brain Default\OCR"
- lenh_tram: "D:\SECOND BRAIN\javis-os\.venv-paddleocr\Scripts\python.exe" -m tools.doctram convert "<file>" -f txt -o "D:\03. Ung dung\Tool OCR\ket_qua_thu" (chạy từ D:\SECOND BRAIN\javis-os)

## VPS (Linux, Docker)
- ket_qua: /brains/Brain Default/OCR
- (không có lenh_ocr, lenh_tram: VPS chưa có công cụ OCR)
