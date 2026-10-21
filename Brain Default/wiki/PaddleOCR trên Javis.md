---
type: wiki
created: 2026-09-16
updated: 2026-09-16
aliases: [PaddleOCR, OCR, PP-OCRv6, paddleocr-mcp]
---

PaddleOCR là bộ OCR mã mở của PaddlePaddle, đọc chữ từ ảnh và PDF (thực tế tính đến 2026-09, repo [PaddlePaddle/PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR)). Trên Javis nó chạy qua MCP chính thức `paddleocr-mcp`, không nhét vào `.venv` của app. [[PaddleOCR GitHub]]

## Dùng trên Javis

1. Cài venv riêng `.venv-paddleocr` (một lần), xem skill `paddleocr`.
2. Trang **Kết nối**: thẻ **PaddleOCR**, bật. Mức Chỉ đọc đủ.
3. Chat: gửi ảnh/PDF, nói "OCR" hoặc "đọc chữ". Javis nạp skill `paddleocr` rồi gọi tool MCP.
4. Model mặc định `PP-OCRv6` mở tool `ocr` (chữ thường, gồm tiếng Việt vì PP-OCRv6 gộp Latin). (thực tế theo tài liệu PaddleOCR 3.7)

Muốn bảng/layout ra Markdown: đổi env `PADDLEOCR_MCP_MODEL=PP-StructureV3` trên thẻ, tool lúc đó là `pp_structurev3`. Một thẻ chỉ một tool. [[PaddleOCR GitHub]]

## Không làm

- Không `pip install paddleocr` vào `.venv` Javis.
- Không bịa chữ khi OCR trống.
- Không dùng PaddleOCR-VL trên CPU. [[PaddleOCR GitHub]]

Chi tiết thao tác cho engine: skill `skills/paddleocr/SKILL.md`.
