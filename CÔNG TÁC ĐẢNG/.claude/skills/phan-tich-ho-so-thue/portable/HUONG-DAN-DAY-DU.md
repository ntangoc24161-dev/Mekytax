# Hướng dẫn sử dụng

Hai công cụ mang USB, dùng trên máy Windows 64-bit. **Không đăng nhập HDDT / TMS / thuế điện tử.**

| Việc cần làm | File mang đi | Cài trên máy kia |
|---|---|---|
| Phân tích hồ sơ kiểm tra thuế (XML tờ khai + Excel hóa đơn) | `PhanTichHoSoThue-portable.zip` | Không cần. Giải nén rồi chạy. |
| Chuyển PDF / ảnh scan sang Markdown | `PdfSangMd-ma-nguon.zip` | Cần Python 3.12, cài **một lần** (có mạng). |

Đọc mục A hoặc mục B tùy việc. Mục C: đưa kết quả vào Javis.

---

# A. Phân tích hồ sơ kiểm tra thuế

Không cần Python. Không cần mạng.

## A1. Giải nén

1. Copy `PhanTichHoSoThue-portable.zip` sang USB.
2. Giải nén ra ổ đĩa, ví dụ `D:\PhanTichHoSoThue\`. **Không chạy trong file zip.**
3. Thấy `Chay.bat` và thư mục `PhanTichHoSoThue\`.

Lần đầu Windows có thể hỏi SmartScreen: **More info** → **Run anyway**.

## A2. Chuẩn bị hồ sơ

Gom file **đã tải sẵn** vào một thư mục (được phép có thư mục con):

- XML tờ khai GTGT, quyết toán TNDN, quyết toán TNCN, BCTC
- Excel bảng kê hóa đơn: sheet **2.BR** (bán ra), **4.MV** (mua vào)

Không cần đặt tên file đúng mẫu. Chương trình đọc thẻ XML `<maTKhai>`:

- 842 = GTGT  
- 892 = TNDN  
- 953 = TNCN  
- 402 = BCTC  

Tờ khai bổ sung được ưu tiên hơn tờ chính thức cùng kỳ.

Bỏ qua: thư mục `BIEN BAN`, file `PTDN01` / `PTDN02`.

## A3. Chạy

- **Chuột:** double-click `Chay.bat` → chọn thư mục hồ sơ → chọn chỗ ghi kết quả (Hủy = ghi vào `<hồ sơ>\ket-qua`).
- **Kéo thả:** kéo thư mục hồ sơ thả vào `Chay.bat`. Kết quả: `<hồ sơ>\ket-qua\`.

Cửa sổ hiện `OK MST ... findings` là xong.

## A4. Đọc kết quả

- `phan-tich-<MST>-<nam>.xlsx` - mở Excel (sheet Tóm tắt, GTGT, TNDN, TNCN, Hóa đơn; dòng Cao tô đỏ)  
- `phan-tich-<MST>-<nam>.md` - mở Notepad / Word / Javis  
- `phan-tich-<MST>-<nam>.json` - danh sách phát hiện  

Trong file Excel: tóm tắt rủi ro (cao / trung bình / thấp), mỗi loại thuế một sheet. Mỗi mục có số lệch, căn cứ, câu hỏi chất vấn.

Chương trình **không sửa** file gốc hồ sơ. Tiêu chí cần TMS (TC24) không chạy.

## A5. Lỗi hay gặp

| Hiện tượng | Cách xử lý |
|---|---|
| SmartScreen chặn | More info → Run anyway |
| Không đọc được XML / bảng kê | Sai thư mục, hoặc Excel không có sheet BR/MV |
| Antivirus xóa exe | Thêm thư mục giải nén vào ngoại lệ |

---

# B. Chuyển PDF / ảnh scan sang Markdown

Không phải exe một nút. OCR tiếng Việt cần PaddleOCR: máy đích cài Python **một lần**, có mạng lúc cài và lúc tải model lần đầu.

## B1. Giải nén

Copy `PdfSangMd-ma-nguon.zip`, giải nén ra ví dụ `D:\PdfSangMd\`. Thấy `cai-dat-may-moi.ps1`, `Chay.bat`, thư mục `tools\`.

## B2. Cài một lần

Cần Python 3.12 (python.org, tick **Add to PATH**).

Mở PowerShell **trong thư mục vừa giải nén**:

```
powershell -ExecutionPolicy Bypass -File .\cai-dat-may-moi.ps1
```

Đợi pip xong. Không cài các gói này vào `.venv` của Javis.

## B3. Chạy

Kéo thả file PDF, ảnh, hoặc cả thư mục vào `Chay.bat`.

Hoặc:

```
.\.venv-ocr\Scripts\python.exe -m tools.doc2md "D:\hs\cong-van.pdf"
.\.venv-ocr\Scripts\python.exe -m tools.doc2md "D:\hs\scan" -o "D:\out" -r
```

| Cờ | Ý nghĩa |
|---|---|
| `-o D:\out` | Thư mục (hoặc file) đích |
| `-r` | Quét cả thư mục con |
| `--force` | Ghi đè file `.md` đã có |
| `--backend paddle` | Bắt buộc OCR (PDF scan) |
| `--backend pymupdf` | Chỉ PDF có lớp chữ, không OCR |
| `--dry-run` | Chỉ liệt kê file, chưa chuyển |

PDF có chữ: nhanh (PyMuPDF). PDF scan / ảnh: chậm hơn, lần đầu tải model.

Máy **đã có Javis OS**:

```
cd "D:\SECOND BRAIN\javis-os"
.\.venv-paddleocr\Scripts\python.exe -m tools.doc2md "D:\hs\scan" --brain "KIỂM TRA THUẾ" -r
```

## B4. Kết quả

File `.md` cạnh PDF hoặc trong thư mục `-o`. Có tiêu đề `# ## ###`, bảng Markdown, đã lọc header/footer/số trang. Frontmatter `type: source` để thả vào Javis.

## B5. Lỗi hay gặp

| Hiện tượng | Cách xử lý |
|---|---|
| `py -3.12` không chạy | Cài Python 3.12, tick Add to PATH |
| Lần OCR đầu rất chậm | Đang tải model, chờ |
| Chữ trong PDF ra rác | Đó là file scan → `--backend paddle` |
| `Chưa cai venv` | Chạy lại `cai-dat-may-moi.ps1` |

---

# C. Đưa vào Javis

1. Copy file `.md` kết quả vào `sources/` của brain (ví dụ **KIỂM TRA THUẾ**).
2. Chat: "tiêu hoá source vừa thêm" nếu muốn lên wiki.
3. Để Javis tự biết cách chạy tool:
   - Phân tích hồ sơ: gửi `JAVIS-TAO-SKILL-phan-tich-ho-so-thue.md` → "tạo skill theo file này"
   - PDF sang MD: gửi `JAVIS-TAO-SKILL-pdf-sang-md.md` → "tạo skill theo file này"
4. Muốn Javis **viết lại app PDF sang MD** trên máy có mã nguồn: gửi `JAVIS-TAO-UNG-DUNG-pdf-sang-md.md`.

Brain KIỂM TRA THUẾ trên máy gốc **đã có** hai skill. Máy / brain khác mới cần bước 3.
