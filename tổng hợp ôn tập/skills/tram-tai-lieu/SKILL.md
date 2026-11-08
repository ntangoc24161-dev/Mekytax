---
name: Trạm tài liệu
description: "Chuyển PDF/ảnh scan ra Word, Excel, TXT, PDF tìm kiếm được; xoay/ghép/tách trang; so chữ hai file; theo dõi thư mục."
group: Năng suất
---

# Trạm tài liệu

Ứng dụng local `tools/doctram` trong repo Javis. OCR tiếng Việt và Anh đi qua PyMuPDF (PDF đã có lớp chữ) hoặc PaddleOCR (scan, ảnh). Không gọi ABBYY FineReader, không copy model trong thư mục cài FineReader.

## Khi nào dùng

- "chuyển PDF sang Word / Excel / TXT", "xuất PDF tìm kiếm được", "xoay trang", "tách PDF", "ghép PDF", "xóa trang", "sắp lại trang", "so hai công văn", "theo dõi thư mục scan".
- Chỉ đọc chữ ra chat, không ghi file: skill `paddleocr`.
- Chỉ ghi Markdown vào `sources/` của brain: skill `pdf-sang-md` (cùng engine `tools/doc2md`). Skill này vẫn làm được `-f md`, dùng khi người dùng đang ở luồng Trạm tài liệu.

## Chuẩn bị

1. Gốc repo: `D:\SECOND BRAIN\javis-os` (nơi có `server/main.py`).
2. Python: `.\.venv-paddleocr\Scripts\python.exe`. Thiếu venv thì chạy `brains\Brain Default\skills\paddleocr\scripts\install.ps1`.
3. Thư viện xuất file, một lần:

```
.\.venv-paddleocr\Scripts\python.exe -m pip install -r tools\doctram\requirements.txt
```

Không `pip` vào `.venv` của Javis. Engine không có shell: đưa nguyên lệnh, nói rõ chưa chạy.

## Cách chạy

Mọi lệnh chạy từ gốc repo.

```
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram convert INPUT -f md|docx|xlsx|txt|pdf -o OUT [--lang vi] [--backend auto|paddle|pymupdf] [--brain "TÊN"] [-r] [--force] [--dry-run]
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram pages INPUT --rotate 90 --out OUT [--force]
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram pages INPUT --delete 1,3-5 --out OUT
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram pages INPUT --order 3,1,2 --out OUT
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram pages --merge A.pdf B.pdf --out OUT
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram pages INPUT --split --out THUMUC
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram compare A B -o chenh.md [--backend auto] [--lang vi] [--force]
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram watch IN --out OUT -f docx [--force]
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram ui
```

`pages` chỉ nhận đúng một việc trong `--rotate`, `--delete`, `--order`, `--merge`, `--split`. Góc xoay: 90, 180, 270. Số trang đánh từ 1. `ui` mở http://127.0.0.1:8765 . Giao diện bấm được qua `tools\doctram\Chay.bat`.

`--backend auto`: PDF có lớp chữ dùng PyMuPDF; scan và ảnh dùng PaddleOCR. `--brain` ghi vào `brains/<tên>/sources/`.

## Quy trình

1. Lấy đường dẫn thật trên đĩa. Ảnh chat thường ở `inbox/` hoặc `.staging` của brain.
2. Chọn lệnh theo việc, không ghép hai việc `pages` một lúc. Nhiều file thì `-o` là thư mục.
3. Có shell: chạy lệnh, đọc dòng `OK` / `SKIP` / `FAIL` / `DRY`.
4. Báo số file, đường dẫn đích, và dòng `!` nếu có. Không dán cả file Word/Excel. Không bịa chữ khi `FAIL`.
5. Họ nói tiêu hoá sau khi đã có `.md` thì mới dùng skill `ingest-source`.

Ví dụ:

```
.\.venv-paddleocr\Scripts\python.exe -m tools.doctram convert "D:\hs\scan.pdf" -f docx -o "D:\hs\ra"
```

## Bẫy

- PDF đã có lớp chữ khi xuất `-f pdf` được copy nguyên, không OCR lại. Scan và ảnh được dựng PDF mới: ảnh trang cộng lớp chữ trong suốt để tìm được, không khớp từng dòng như bản Word.
- So sánh là so đoạn chữ đã đọc. Không báo lệch hình, dấu đỏ, hay con dấu.
- `watch` chỉ một tiến trình, chỉ đuôi pdf/png/jpg/jpeg/tif/tiff/bmp/webp, bỏ file đang copy dở. Thư mục đích không được nằm trong thư mục theo dõi.
- Một file lỗi không dừng cả thư mục. `SKIP` nghĩa là đích đã có; thêm `--force` khi họ muốn ghi đè.
- Lần Paddle đầu tải model, có thể vài phút. Không gọi song song nhiều file nặng.
- Không đọc được chữ thì nói thất bại. Không điền số trên công văn hay hóa đơn.

## Kiểm chứng

- `--help` in đủ `convert`, `pages`, `compare`, `watch`, `ui`.
- `--dry-run` liệt kê file, không tạo file đích.
- Một PDF có chữ, `-f txt --backend pymupdf`, file txt chứa câu đã biết trên PDF.
- `pages --split` tạo đúng số file với số trang.
