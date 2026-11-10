---
name: PDF sang Markdown
description: "Chuyển PDF/ảnh scan thành file Markdown (.md) cho Second Brain, giữ heading/bảng, lọc header/footer."
group: Năng suất
---

# PDF sang Markdown - ghi source sạch cho Javis

Chạy module `tools/doc2md` (venv `.venv-paddleocr`). Không gọi MCP PaddleOCR. Không `pip install` vào `.venv` của Javis.

## Khi nào dùng

- "chuyển PDF thành markdown", "OCR xong ghi vào sources", "scan sang .md", "batch PDF trong thư mục này".
- Hồ sơ scan, công văn, bảng kê cần **file .md** trong brain (heading `# ## ###`, bảng Markdown, hết header/footer lặp).
- Không dùng khi chỉ cần **đọc chữ ra chat** (một ảnh/PDF, không ghi file) → skill `paddleocr`.
- Không dùng khi file đã là `.md`/`.txt`. Không dịch, không bịa chữ OCR.

## Chuẩn bị

1. Gốc repo Javis (thư mục có `server/main.py`). Máy này: `D:\SECOND BRAIN\javis-os`.
2. Đã có `.venv-paddleocr`. Thiếu thì chạy `brains/Brain Default/skills/paddleocr/scripts/install.ps1` rồi:
   `.\.venv-paddleocr\Scripts\python.exe -m pip install -r tools\doc2md\requirements.txt`
3. Python chạy lệnh: `.\.venv-paddleocr\Scripts\python.exe` - **không** `.venv\Scripts\python.exe`.

Engine **không có shell** (OpenRouter, OpenAI/Anthropic/Gemini API): không chạy được lệnh máy. Đưa nguyên lệnh cho người dùng, nói rõ chưa chạy. Đừng giả vờ đã OCR.

## Cách chạy

Từ gốc repo:

```
.\.venv-paddleocr\Scripts\python.exe -m tools.doc2md "<đường dẫn PDF/ảnh hoặc thư mục>" [cờ]
```

Cờ thật (đừng bịa cờ khác):

- `-o OUT` file `.md` hoặc thư mục đích
- `--brain "TÊN BRAIN"` ghi vào `brains/<tên>/sources/`
- `-r` quét thư mục con
- `--force` ghi đè `.md` đã có
- `--backend auto|paddle|pymupdf|docling|marker` (mặc định `auto`)
- `--lang vi` (mặc định; Latin = Việt có dấu + Anh)
- `--no-orientation` tắt bước tự nhận hướng xoay trang và dòng chữ. Nhanh hơn khi scan đã thẳng. Ảnh lệch hoặc ngược sẽ đọc sai. Không thêm cờ thì bước này vẫn bật.
- `--dry-run` chỉ liệt kê file
- `--list-backends` xem backend đã cài

Lần Paddle đầu tiên tải model, có thể vài phút. Không gọi song song nhiều file nặng.

## Quy trình

1. Lấy đường dẫn file/thư mục thật trên đĩa. Ảnh chat thường ở `inbox/` hoặc `.staging` của brain.
2. Chọn đích: người dùng nói brain nào thì `--brain "..."`. Im lặng thì `-o` cạnh file nguồn, hoặc `sources/` của brain đang mở.
3. Engine có shell: chạy lệnh, đọc stdout (`OK` / `SKIP` / `FAIL`).
4. Báo: bao nhiêu file thành công, đường dẫn `.md` vừa ghi. Không dán cả file dài trừ khi họ hỏi.
5. Họ nói "tiêu hoá / ingest" sau khi có `.md` → mới dùng skill `ingest-source`. Không ingest nếu họ chỉ bảo chuyển file.
6. Lỗi `chưa cài` / traceback: nói đúng thiếu gì (venv OCR, pymupdf, paddle). Không bịa nội dung tài liệu.

Ví dụ ghi vào brain kiểm tra thuế:

```
.\.venv-paddleocr\Scripts\python.exe -m tools.doc2md "D:\hs\scan" --brain "KIỂM TRA THUẾ" -r
```

## Bẫy

- Trùng với `paddleocr`: skill kia trả **chữ trong chat** qua MCP. Skill này **ghi file .md**.
- Cài paddleocr/pymupdf vào `.venv` Javis sẽ đụng FastAPI.
- `--brain` sai tên thì lệnh thất bại; liệt kê thư mục `brains/` chứ đừng đoán.
- PDF có lớp chữ: `auto` dùng PyMuPDF (nhanh). Scan mờ: `--backend paddle`.
- Scan máy photocopy đã thẳng hướng: thêm `--no-orientation` để bớt một bước suy luận mỗi trang. Không dùng cờ này khi ảnh bị xoay lệch hoặc chụp nghiêng.
- Không đọc được thì nói thất bại, không bịa số trên công văn/hóa đơn.

## Kiểm chứng

- `.\.venv-paddleocr\Scripts\python.exe -m tools.doc2md --list-backends` hiện `pymupdf OK` và `paddle OK`.
- `--dry-run` trên một thư mục in đúng danh sách PDF/ảnh.
- Chat: "chuyển file này sang markdown vào sources" → Javis chạy lệnh trên (nếu có shell) hoặc đưa lệnh (nếu không).
