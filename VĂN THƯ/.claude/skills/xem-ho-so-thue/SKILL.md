---
name: Xem hồ sơ thuế điện tử
description: "Mở XML hồ sơ thuế, dựng tờ khai bằng mẫu iTax Viewer, kiểm chữ ký số, xuất HTML."
group: Tài chính
---

# Xem hồ sơ thuế điện tử

Đọc một tệp XML hồ sơ thuế (tờ khai, thông báo, chứng từ, hóa đơn, đăng ký). Dựng HTML bằng mẫu XSLT của iTax Viewer đã cài trên máy, kiểm chữ ký số, trả tóm tắt từ JSON. Không chấm rủi ro. Không ký, không sửa tệp, không nộp hồ sơ.

## Khi nào dùng

- "mở tờ khai xml", "xem hồ sơ thuế điện tử", "đọc file HSoThueDTu", "kiểm chữ ký tờ khai", "in tờ khai như iTax Viewer".
- User đưa đường dẫn một tệp `.xml` hồ sơ thuế.
- Không dùng khi user muốn chấm 28 tiêu chí, đối chiếu hóa đơn với tờ khai, hoặc chạy PhanTichHoSoThue. Việc đó là skill `phan-tich-ho-so-thue`.
- Không dùng cho file Excel bảng kê mua vào, bán ra.

## Chuẩn bị

- Windows, có iTax Viewer (mặc định `C:\Program Files (x86)\iTax Viewer`).
- Python của app: `D:\SECOND BRAIN\xem-ho-so-thue\.venv\Scripts\python.exe`
- Thiếu venv: trong thư mục đó chạy `python -m venv .venv` rồi `.venv\Scripts\python.exe -m pip install -r requirements.txt`. Không cài vào `.venv` của Javis.

Engine không có shell: đưa nguyên lệnh, nói rõ chưa chạy. Đừng bịa MST, kỳ, hay kết quả chữ ký.

## Cách chạy

```
D:\SECOND BRAIN\xem-ho-so-thue\.venv\Scripts\python.exe "D:\SECOND BRAIN\xem-ho-so-thue\app.py" --file "<xml>" --html "<thu muc ra>" --json
```

`<thu muc ra>` mặc định hợp lý là thư mục chứa tệp XML, thư mục con `xem-html`. Xem trên trình duyệt: chạy `app.py` không tham số, mở `http://127.0.0.1:8765`.

## Quy trình

1. Lấy đường dẫn tệp XML thật trên đĩa. Không nhận nội dung dán nếu không ghi ra tệp.
2. Có shell: chạy lệnh trên. Đọc JSON in ra stdout.
3. Tóm tắt ngắn: loại hồ sơ, mã, phiên bản, MST, tên người nộp, kỳ, số phụ lục, trạng thái chữ ký (`hop_le` / `canh_bao` / `loi` / `khong_co`) và một câu `ghi_chu` của từng chữ ký. Nếu `canh_bao_mau` khác rỗng thì nói rõ vì sao không dựng được mẫu.
4. Đưa đường dẫn file HTML. Không dán cả trang tờ khai trừ khi user hỏi.
5. Không ingest wiki trừ khi user nói "tiêu hoá".

## Bẫy

- Số liệu chỉ lấy từ JSON và từ chính tệp. Thiếu trường thì nói không có, không điền hộ.
- `canh_bao` là chứng thư hết hạn hoặc không nằm trong `certstore` của iTax Viewer. `loi` là chữ ký vỡ hoặc chứng thư bị thu hồi. "Chưa kiểm tra thu hồi" không có nghĩa là chữ ký sai.
- Mẫu XSLT thuộc bản cài iTax Viewer. Máy không cài thì app chỉ hiện cây XML.
- App chỉ chạy trên Windows vì mẫu dùng MSXML.

## Kiểm chứng

- Lệnh thoát mã 0.
- JSON có `mst`, `ma`, `ky`, `chu_ky`, `html`.
- File tại `html` tồn tại và không rỗng.
- Câu trả lời không thêm số tiền nào ngoài những gì JSON hoặc user đã nêu.
