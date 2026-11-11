---
name: PaddleOCR
description: "OCR ảnh và PDF bằng PaddleOCR qua MCP (tool ocr). Bảng/layout: đổi model PP-StructureV3."
group: Năng suất
---

# PaddleOCR - đọc chữ từ ảnh và PDF

Dùng MCP chính thức của [PaddlePaddle/PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) đã đấu vào trang Kết nối. Một kết nối chỉ mở **một** tool, tùy biến `PADDLEOCR_MCP_MODEL`.

## Khi nào dùng

- Ảnh hóa đơn, CMND, biển số, tem, bảng, trang scan, PDF cần **đọc chữ ra chat**.
- Người dùng nói "đọc chữ ảnh này", "OCR giúp", "lấy chữ trong ảnh".
- Muốn **ghi file .md vào sources** (heading, bảng, batch) → skill `pdf-sang-md`, không gọi MCP.
- Không dùng khi file đã là chữ (`.md`, `.txt`, `.docx` có text). Không dịch, không sửa ảnh.

## Chuẩn bị

1. Trang **Kết nối** có thẻ **PaddleOCR**, đang **bật**. Thẻ stdio mặc định tắt; bật sau khi venv OCR đã cài.
2. Venv riêng (không được cài vào `.venv` của Javis): `.venv-paddleocr`. Cài một lần:

```
cd "D:\SECOND BRAIN\javis-os"
py -3.12 -m venv .venv-paddleocr
.\.venv-paddleocr\Scripts\python.exe -m pip install -U "paddleocr-mcp[local-cpu]"
```

3. File ảnh/PDF nằm trong brain (`attachments/`, `inbox/`) hoặc đường dẫn máy mà Javis đọc được. Lần chạy local đầu tiên tải model, chậm vài phút.

Nguồn chạy mặc định: `PADDLEOCR_MCP_PPOCR_SOURCE=local`, model `PP-OCRv6` (tool `ocr`). Đổi sang API Baidu thì cần token AI Studio, không đoán token.

## Cách chạy

Gọi đúng tool MCP mà kết nối đang mở. Tham số chung (từ server `paddleocr-mcp`):

- `input_data` (bắt buộc): đường dẫn file, URL, hoặc (nguồn HTTP) dữ liệu file. Local: đưa **đường dẫn file**, không nhét PDF base64 (giới hạn chính thức khi chạy local).
- `output_mode`: `simple` (chữ thường, mặc định) hoặc `detailed` (JSON).
- `file_type`: `image` hoặc `pdf` khi nguồn HTTP không đoán được loại.
- `return_images`: với parse tài liệu; tắt nếu không cần ảnh trong kết quả (tiết kiệm token).
- `runtime_params`: object tùy chọn, chỉ key pipeline hợp lệ.

| Việc người dùng muốn | Model trên kết nối | Tool |
|---|---|---|
| Đọc chữ (hóa đơn, biển, chữ in/tay) | `PP-OCRv6` (mặc định) | `ocr` |
| Latin chuyên (không CJK) | `PP-OCRv5-latin` | `ocr` |
| PDF/ảnh phức tạp ra Markdown (bảng, tiêu đề) | `PP-StructureV3` | `pp_structurev3` |
| Parse bằng VLM | `PaddleOCR-VL-1.6` | `paddleocr_vl` |

Muốn tool khác: sửa env `PADDLEOCR_MCP_MODEL` của thẻ Kết nối rồi bật lại. Không giả vờ gọi `pp_structurev3` khi thẻ đang là `PP-OCRv6`.

## Quy trình

1. Xác định file. Ảnh vừa gửi chat thường ở `inbox/` hoặc `.staging`. Đường dẫn tương đối brain hoặc tuyệt đối đều được; local OCR cần file thật trên đĩa.
2. Kiểm tra Kết nối PaddleOCR đang bật. Tắt thì nói rõ: bật thẻ, hoặc chưa cài `.venv-paddleocr`.
3. Gọi tool. OCR thường: `ocr` + `input_data` + `output_mode=simple`.
4. Trả lời người dùng bằng **chữ đã đọc**, gọn. Chỉ đưa JSON khi họ hỏi tọa độ/độ tin cậy.
5. Nếu kết quả trống: nói không đọc được chữ, gợi ý ảnh rõ hơn, không bịa nội dung.

Ví dụ lời gọi (OCR thường):

```
ocr
input_data: <đường dẫn file ảnh hoặc PDF>
output_mode: simple
```

## Bẫy

- Cài `paddleocr` vào `.venv` Javis sẽ đụng FastAPI/starlette. Chỉ cài `.venv-paddleocr`.
- Một tiến trình MCP = một model = một tool. Sai tool là kết nối đang mở model khác.
- Local không nhận PDF base64. Gửi file trên đĩa.
- `PaddleOCR-VL` không nên chạy CPU. Giữ `PP-OCRv6` cho máy không GPU.
- Lần đầu tải model; timeout MCP có thể 10 phút. Đừng gọi song song nhiều file nặng.
- Không đọc được thì nói thất bại, không bịa số trên hóa đơn.

## Kiểm chứng

- Thẻ Kết nối **PaddleOCR** bật, mức Chỉ đọc là đủ (OCR không ghi ra ngoài).
- Chat: gửi một ảnh chữ rõ, bảo "OCR giúp". Javis gọi `ocr` và trả đúng chữ trên ảnh.
- `paddleocr_mcp --help` chạy được bằng `.venv-paddleocr\Scripts\paddleocr_mcp.exe`.
