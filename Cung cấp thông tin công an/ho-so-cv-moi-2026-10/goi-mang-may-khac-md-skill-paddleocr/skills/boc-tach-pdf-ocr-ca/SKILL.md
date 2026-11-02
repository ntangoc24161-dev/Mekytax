---
name: Bóc tách PDF/OCR công văn
description: "Bóc PDF/ảnh scan công văn thành Markdown đầy đủ (PyMuPDF hoặc OCR Paddle), tách đa CV, không bịa chữ."
group: Tài chính
---

# Bóc tách PDF/OCR công văn

Chuyển PDF/ảnh scan thành file Markdown **đủ nội dung** (số hiệu, ngày, cơ quan, yêu cầu, bảng MST, phụ lục). Chạy `tools.doc2md` trên venv OCR. Không gọi MCP PaddleOCR. Không `pip install` vào `.venv` của Javis. Không dịch. Không bịa chữ OCR, MST, số tiền.

## Khi nào dùng

- "bóc PDF", "OCR công văn", "chuyển scan sang md", "trích xuất đầy đủ nội dung file OCR/PDF", "mang skill sang máy khác chạy doc2md".
- Hồ sơ công văn Công an / Thuế TP: PDF ký số có lớp chữ, PDF scan, ảnh trang, bảng kê MST.
- Không dùng khi chỉ cần đọc lướt một trang ra chat (không ghi file).
- Không dùng khi file đã là `.md`/`.txt`.
- Không dùng để phân Tổ KT, ghi Excel, soạn dự thảo (skill `theo-doi-cung-cap-ca` / `tra-loi-cong-van`).

## Chuẩn bị

Thứ tự tìm công cụ (dừng khi chạy được `--list-backends`):

1. Gốc Javis OS (thư mục có `server/main.py`): `D:\SECOND BRAIN\javis-os`. Python: `.\.venv-paddleocr\Scripts\python.exe` (không dùng `.venv\Scripts\python.exe`).
2. Gói portable: thư mục có `tools\doc2md` và `.\.venv-ocr\Scripts\python.exe`.
3. Thiếu venv: máy Javis OS chạy `skills/paddleocr/scripts/install.ps1` (trong gói mang theo hoặc `brains\Brain Default\skills\paddleocr\scripts\install.ps1`). Nếu repo có `tools\doc2md\requirements.txt` thì cài thêm bằng cùng python venv đó.

Engine không có shell (OpenRouter, API): đưa nguyên lệnh, nói rõ chưa chạy. Đừng giả vờ đã OCR.

Đích lưu `.md` công văn Công an (thử lần lượt):

1. `D:\HSO CONG AN FILE MD` (đĩa Windows; tập trung, không bỏ lẫn thư mục PDF).
2. Không có ổ `D:`: `ho-so-cv-moi-YYYY-MM/md/` trong brain đang mở.

PDF gốc giữ nguyên chỗ nhận (`CV MOI` hoặc `ho-so-cv-moi-YYYY-MM/pdf/`).

## Cách chạy

Từ gốc Javis OS, python venv OCR, module `tools.doc2md`:

- `--list-backends` xem pymupdf / paddle
- `-o "D:\HSO CONG AN FILE MD"` đích `.md`
- `-r` quét thư mục con
- `--force` ghi đè
- `--backend auto|paddle|pymupdf|docling|marker` (mặc định `auto`)
- `--lang vi`
- `--dry-run` chỉ liệt kê

PDF ký số / lớp chữ: `auto` dùng PyMuPDF. PDF scan / ảnh mờ: `--backend paddle`. Lần Paddle đầu tải model, có thể vài phút. Không gọi song song nhiều file nặng.

`--brain "TÊN BRAIN"` chỉ khi user bảo ghi `sources/` của Javis. Hồ sơ CA dùng `-o "D:\HSO CONG AN FILE MD"`.

## Quy trình

1. Lấy đường dẫn file/thư mục PDF thật trên đĩa. `--dry-run` trước khi chạy hàng loạt.
2. Với từng PDF:
   - Có lớp chữ: `--backend pymupdf` hoặc `auto`.
   - Không lớp chữ / chữ rác: `--backend paddle`.
   - Vẫn trống sau OCR: xuất ảnh từng trang rồi đọc ảnh; ghi `.md` những gì đọc được; chỗ không đọc được ghi `[không đọc được]`. Không bỏ hồ sơ. Không bịa MST.
3. Một PDF nhiều công văn đính kèm: tách thành nhiều file `.md`, mỗi công văn một file, đặt tên theo số hiệu. Không gộp thành một khối.
4. File `.md` phải có đủ, nếu văn bản có:
   - Số hiệu, ngày ban hành, nơi ban hành, dấu đến (số đến, ngày đến)
   - Kính gửi, trích yếu, căn cứ
   - Từng điểm yêu cầu (không rút thành một câu)
   - Điều tra viên / số điện thoại / hạn trả lời / nơi nhận
   - Bảng kê: giữ dạng bảng Markdown, đủ cột STT, MST, tên (và cột khác nếu có). STT in trên giấy giữ nguyên, kể cả nhảy số.
   - Chữ ký / chức danh (không cần vẽ dấu)
5. Đặt tên file `.md`: số hiệu ASCII, ví dụ `1625-YC-ANDT-D2.md`, `10212-CSKT-D3.md`. Không lưu `.md` trong thư mục PDF lưu trữ `STT. CV ...`.
6. Đọc stdout (`OK` / `SKIP` / `FAIL`). Báo số file thành công và đường dẫn. Không dán cả file dài trừ khi user hỏi.
7. Lỗi `chưa cài` / traceback: nói đúng thiếu gì (venv OCR, pymupdf, paddle). Không bịa nội dung tài liệu.
8. Không ingest wiki, không ghi Excel, không phân Tổ KT trong skill này.

## Bẫy

- Trùng `pdf-sang-md`: skill kia là lệnh doc2md chung. Skill này thêm chuẩn đủ nội dung công văn và đích `D:\HSO CONG AN FILE MD`.
- Trùng `paddleocr` MCP: skill kia trả chữ trong chat, không ghi file.
- Cài paddleocr/pymupdf vào `.venv` Javis sẽ đụng FastAPI.
- PDF chỉ đạo Thuế TP (`DON-KTr1`) thường không kèm bảng MST: ghi rõ thiếu bảng kê, không bịa danh sách.
- Bảng scan nhảy STT (trùng 15, thiếu 6-7): giữ STT trên giấy, không tự đánh lại.
- Không đọc được mà vẫn điền MST/số tiền: bịa. Để `[không đọc được]`.

## Kiểm chứng

- `--list-backends` hiện `pymupdf OK`; PDF scan cần thêm `paddle OK`.
- `--dry-run` in đúng danh sách PDF/ảnh sẽ xử lý.
- Mỗi công văn một `.md` trong `D:\HSO CONG AN FILE MD` (hoặc `ho-so-cv-moi-YYYY-MM/md/` khi không có ổ D:).
- File `.md` có heading, bảng Markdown nếu nguồn có bảng; không còn PAGE BREAK rỗng.
- PDF gốc vẫn ở thư mục nhận, không bị xóa.
