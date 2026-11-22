---
name: Đọc PDF và OCR
description: "Đọc nội dung PDF hoặc ảnh scan, kể cả file chưa OCR: dùng kết quả có sẵn, hoặc OCR file mới rồi đọc."
group: Năng suất
---

# Đọc PDF và OCR

Đây là điểm vào duy nhất khi người dùng muốn đọc, tìm, hoặc bóc số liệu từ PDF hoặc ảnh scan. Skill tự chọn: dùng kết quả OCR có sẵn trước, OCR mới khi chưa có. Việc OCR mới đi qua skill `tram-tai-lieu` khi máy không có công cụ OCR chính.

## Khi nào dùng

- "đọc file PDF này", "file scan này nói gì", "tìm trong công văn...", "bóc số tiền trong bảng...".
- Người dùng nhắc tới một tệp PDF hoặc ảnh đã hoặc chưa OCR.
- Không dùng khi chỉ cần chuyển file sang Word, Excel hay PDF tìm kiếm được (dùng `tram-tai-lieu`), hoặc xoay, tách, ghép trang.

## Bước 1: đọc cấu hình máy

Đọc `Javis/doc-ocr-config.md` của brain. Mỗi máy có một mục. Dùng mục của máy đang chạy:
- `ket_qua`: thư mục kết quả. Dùng đường dẫn đầu tiên tồn tại trên máy này.
- `inbox`, `lenh_ocr`: công cụ OCR chính (Tool OCR, Tesseract). Có thì dùng.
- `lenh_tram`: công cụ dự phòng (`tram-tai-lieu`). Có thì dùng khi thiếu `lenh_ocr`.

Không có file cấu hình, hoặc máy không có mục nào: nói rõ "máy này chưa cấu hình OCR", không đoán đường dẫn.

## Bước 2: tìm kết quả có sẵn

Trong `ket_qua`, tìm theo tên gốc của tệp (không phân biệt hoa thường):
- `<tên>_OCR.txt` (Tool OCR) hoặc `<tên>.txt` (đầu ra của `tram`): đọc trước.
- `<tên>_OCR_sua.txt`: danh sách chỗ đã tự sửa. Đọc khi cần kiểm tra số, MST, ngày tháng.
- `<tên>_OCR.xlsx`: bảng. Chỉ mở khi `.txt` không đủ cho số liệu trong bảng. Chỉ có trên máy chạy Tool OCR.

Có kết quả thì sang bước 4.

## Bước 3: OCR file mới

Chỉ làm khi người dùng đưa file mới mà chưa có kết quả:

1. Nếu có `lenh_ocr` (ưu tiên, Tesseract đọc dấu tiếng Việt tốt): chép file vào `inbox`, chạy `lenh_ocr`. Lệnh mất vài phút mỗi trang, nên chạy nền nếu engine hỗ trợ. Kết quả sẽ nằm trong `ket_qua`.
2. Nếu không có `lenh_ocr` nhưng có `lenh_tram`: chạy `tram-tai-lieu convert` với `-f txt` và `-o` là `ket_qua`. Cảnh báo người dùng: lệnh này có thể làm mất dấu tiếng Việt, nên mọi số liệu phải đối chiếu với bản gốc.
3. Không có cả hai (thường là VPS): không OCR được. Báo người dùng thả file vào `inbox` trên máy Windows, kết quả sẽ tự lên VPS ở lần đồng bộ brain sau.

Không hứa sẽ tự báo lại khi xong. Nói rõ đã đưa vào OCR và kết quả nằm ở đâu.

## Bước 4: đọc và trả lời

- Đọc đúng đoạn cần, không dán cả file vào chat. Tìm theo từ khóa, rồi mới đọc quanh chỗ tìm thấy.
- Số tiền, MST, ngày tháng: đối chiếu với `_OCR_sua.txt` hoặc bản gốc trước khi kết luận. Nếu còn nghi ngờ, nói rõ đây là kết quả OCR và chỉ ra chỗ không chắc.
- Trích dẫn: ghi tên file kết quả và trang nếu có.

## Lưu ý

- Chữ OCR có thể sai dấu hoặc sai số. Không khẳng định số liệu quan trọng khi chưa đối chiếu.
- Engine không có shell (API): không tự chạy được lệnh OCR. Chỉ đọc kết quả có sẵn, và nói rõ nếu cần OCR thêm.
- Không sửa, không xóa file trong `ket_qua`.
