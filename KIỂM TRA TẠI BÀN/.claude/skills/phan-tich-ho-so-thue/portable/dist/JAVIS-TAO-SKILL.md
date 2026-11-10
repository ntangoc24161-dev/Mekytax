# Lệnh cho Javis: tạo skill Phân tích hồ sơ kiểm tra thuế

Người dùng dán file này vào chat Javis (hoặc để Javis đọc file) và nói: **tạo skill theo file này**.

Javis ghi đúng một file:

`skills/phan-tich-ho-so-thue/SKILL.md`

trong **brain đang mở**. Nếu skill `phan-tich-ho-so-thue` đã có thì **cập nhật** file cũ, không tạo bản sao. `description` dưới đây đã đếm **109** ký tự, không viết dài hơn 150. `group: Tài chính`. Không dùng em dash.

---

Sao chép nguyên khối dưới đây thành nội dung `SKILL.md`:

```
---
name: Phân tích hồ sơ kiểm tra thuế
description: "Phân tích hồ sơ kiểm tra thuế từ XML tờ khai/BCTC và Excel hóa đơn, đối chiếu HĐ-tờ khai, 28 tiêu chí rủi ro."
group: Tài chính
---

# Phân tích hồ sơ kiểm tra thuế

Đọc file đã xuất sẵn (XML tờ khai/BCTC, Excel bảng kê HĐ). Không đăng nhập HDDT, TMS, thuế điện tử, DVC. Kết quả: file Excel (mỗi loại thuế một sheet) + Markdown + JSON.

## Khi nào dùng

- "phân tích hồ sơ thuế", "đối chiếu hóa đơn tờ khai", "chấm rủi ro BCTC", "28 tiêu chí", "chạy PhanTichHoSoThue", "Chay.bat".
- User đã có thư mục XML GTGT/TNDN/TNCN/BCTC và/hoặc Excel bảng kê mua-bán.
- Không dùng khi chỉ OCR PDF (`pdf-sang-md`) hoặc hỏi thời hạn QĐ 1161 (`quy-trinh-kiem-tra-thue`).
- Không dùng để cào web hay xin tài khoản hệ thống thuế.

## Chuẩn bị

Ưu tiên bản portable (không cần Python trên máy):

- `Chay.bat` và thư mục `PhanTichHoSoThue\` (giải nén từ `PhanTichHoSoThue-portable.zip`).
- Thư mục hồ sơ NNT trên đĩa.

Nếu đang ở repo Javis OS, có thể chạy engine Python:

- `.venv-paddleocr\Scripts\python.exe`
- `brains/KIỂM TRA THUẾ/skills/phan-tich-ho-so-thue/engine/cli.py`

Không `pip install` vào `.venv` của Javis.

Engine không có shell (OpenRouter, API): đưa nguyên lệnh, nói rõ chưa chạy. Đừng giả vờ đã phân tích.

## Cách chạy

Portable:

```
Chay.bat
```

Kéo thả thư mục hồ sơ vào `Chay.bat`, hoặc:

```
PhanTichHoSoThue\PhanTichHoSoThue.exe --input-dir "<thu muc ho so>" --out "<thu muc ket qua>"
```

Mặc định `--out` = `<hồ sơ>\ket-qua` nếu không chọn.

Trong repo Javis:

```
.\.venv-paddleocr\Scripts\python.exe "brains\KIỂM TRA THUẾ\skills\phan-tich-ho-so-thue\engine\cli.py" --input-dir "<thu muc ho so>" --out "<brain>\sources"
```

XML: `<maTKhai>` 842 GTGT, 892 TNDN, 953 TNCN, 402 BCTC. `loaiTKhai` B đè C cùng kỳ. Excel sheet `2.BR` / `4.MV`. Bỏ `BIEN BAN`, `PTDN01`, `PTDN02`.

## Quy trình

1. Lấy đường dẫn thư mục hồ sơ thật trên đĩa.
2. Engine có shell: chạy lệnh trên. Đọc dòng `OK MST ... findings`.
3. Mở file `.xlsx` vừa ghi (sheet Tóm tắt, rồi GTGT/TNDN/TNCN/Hóa đơn). Tóm tắt số mục cao/trung bình/thấp và 3-5 red flag. File `.md` để nạp Javis, không dán cả báo cáo trừ khi user hỏi.
4. Không ingest wiki trừ khi user nói "tiêu hoá".
5. Không bịa số. Thiếu BCTC XML thì nói M2 ít tiêu chí. TC24 (cần TMS) không chạy.

## Bẫy

- `TO_KHAI_*.xml` trùng nội dung GTGT: hash-dedup.
- File `*.xml.xml` vẫn parse.
- Ngưỡng TC2 = 5% DT, TC3 = 30% DT (sheet TUDIEN tool 4.4). Thang điểm tổng hợp ghi (cần xác minh) so với QĐ 1109.
- SmartScreen có thể chặn exe lần đầu: More info, Run anyway.

## Kiểm chứng

- Chạy xong có `phan-tich-<MST>-<nam>.xlsx`, `.md` và `.json`.
- File Excel có sheet Tóm tắt / GTGT / TNDN / TNCN / Hóa đơn / Phụ lục; dòng Cao tô đỏ.
- File md có mục Tóm tắt rủi ro, không có em dash.
```

Sau khi ghi file, báo ngắn: đã tạo/cập nhật skill `phan-tich-ho-so-thue`, đường dẫn file, user gõ `/phan-tich-ho-so-thue` hoặc "phân tích hồ sơ thuế thư mục ...".
