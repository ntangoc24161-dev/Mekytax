# Gói mang máy khác: MD + skill + PaddleOCR

Xuất 15/10/2026 từ brain Cung cấp thông tin công an.

## Trong zip

| Thư mục / file | Việc |
|---|---|
| `md/` | Chữ PDF/OCR đã bóc đợt iOffice 01-12/10 và bảng 50 MST CV 1625 |
| `skills/boc-tach-pdf-ocr-ca/SKILL.md` | Skill bóc PDF công văn đầy đủ, chép thẳng vào brain máy kia |
| `skills/pdf-sang-md/SKILL.md` | Skill chuyển PDF sang MD (doc2md) |
| `skills/paddleocr/SKILL.md` và `scripts/install.ps1` | Skill OCR chat + script cài venv PaddleOCR |
| `JAVIS-TAO-SKILL-pdf-sang-md.md` | Dán vào chat máy kia, nói tạo skill theo file này |

**Không có trong zip:** model PaddleOCR, thư mục `.venv-paddleocr`, mã nguồn `tools/doc2md`. VPS đang chạy không có các thứ đó.

## Trên máy Windows kia

### 1. Chép skill vào brain

Chép ba thư mục trong `skills/` vào `skills/` của brain. Hoặc gửi `JAVIS-TAO-SKILL-pdf-sang-md.md` vào chat rồi nói tạo skill theo file này.

### 2. Cài PaddleOCR một lần (cần mạng)

Máy đã có Javis OS tại `D:\SECOND BRAIN\javis-os`:

```
cd "D:\SECOND BRAIN\javis-os"
powershell -ExecutionPolicy Bypass -File "<thư mục giải nén>\skills\paddleocr\scripts\install.ps1"
```

Nếu repo Javis OS có `tools\doc2md\requirements.txt` thì chạy thêm:

```
.\.venv-paddleocr\Scripts\python.exe -m pip install -r tools\doc2md\requirements.txt
```

(lệnh trên: `.\.venv-paddleocr\Scripts\python.exe`, không có dấu cách sau `\`)

Xong: bật thẻ PaddleOCR ở trang Kết nối (đọc chữ ra chat). Ghi file `.md` thì dùng skill `pdf-sang-md` hoặc `boc-tach-pdf-ocr-ca`.

### 3. Dùng chữ đã bóc (không cần OCR lại)

Chép `md/` vào `D:\HSO CONG AN FILE MD` hoặc vào brain. File `XUAT-MANG-MAY-KHAC-boc-tach-ocr.md` là bản gộp.

### 4. OCR PDF mới

Từ gốc Javis OS, dùng `.\.venv-paddleocr\Scripts\python.exe -m tools.doc2md`:

- `--list-backends` để xem pymupdf / paddle đã sẵn chưa
- `-o "D:\HSO CONG AN FILE MD" -r --lang vi` cho thư mục PDF
- PDF scan: thêm `--backend paddle`

## Không làm

- Không cài paddleocr vào `.venv` của Javis.
- Không giả vờ đã OCR khi `--list-backends` chưa có `paddle OK`.
