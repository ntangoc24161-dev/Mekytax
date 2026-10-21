---
type: business
provenance: source
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
---
Zip `ho-so-cv-moi-2026-10/Goi-mang-may-khac-md-skill-paddleocr.zip` không chứa model PaddleOCR, `.venv-paddleocr` hay mã `tools/doc2md`. VPS đang chạy không có venv OCR. Gói `PdfSangMd-ma-nguon.zip` cũng không có trong vault.

Máy Windows đã có Javis OS `D:\SECOND BRAIN\javis-os`, giải nén zip rồi (cần mạng lần đầu):

```
cd "D:\SECOND BRAIN\javis-os"
powershell -ExecutionPolicy Bypass -File "<thư mục giải nén>\skills\paddleocr\scripts\install.ps1"
```

Nếu repo có `tools\doc2md\requirements.txt` thì cài thêm bằng `\.venv-paddleocr\Scripts\python.exe`. Xong: bật thẻ PaddleOCR ở trang Kết nối. Chữ đã bóc thì chép `md/` vào `D:\HSO CONG AN FILE MD`, không OCR lại đợt 01-12/10.

Không cài paddleocr vào `.venv` của Javis. Thiếu venv thì chỉ tạo được skill rồi dừng, không giả vờ đã OCR.
