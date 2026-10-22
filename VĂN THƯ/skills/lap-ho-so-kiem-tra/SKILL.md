---
name: Lập hồ sơ kiểm tra thuế
description: "Vận hành app hồ sơ kiểm tra thuế: nhập vụ, import Excel, xuất Word đi, biên bản, phát sinh và về."
group: Tài chính
---

# Lập hồ sơ kiểm tra thuế

App web nội bộ của Thuế cơ sở 3: nhập một vụ kiểm tra hoặc nạp Excel, rồi xuất bộ Word hồ sơ đi, biên bản, phát sinh và hồ sơ về. App không ký số, không đóng dấu, không cấp số quyết định, không xuất PDF.

## Khi nào dùng

- "lập hồ sơ kiểm tra", "xuất biên bản kiểm tra", "import Excel vụ kiểm tra", "chạy app hồ sơ", "xuất DI-03", "xuất BB-05", "hoso.zip".
- User đưa MST, số quyết định, file Excel mẫu app, hoặc bảo xuất một mã mẫu.
- Không dùng để chấm 28 tiêu chí XML/Excel hóa đơn. Việc đó là `phan-tich-ho-so-thue`.
- Không dùng để mở tờ khai XML trên iTax Viewer. Việc đó là `xem-ho-so-thue`.
- Không dùng khi chỉ hỏi thời hạn QĐ 1161. Việc đó là `quy-trinh-kiem-tra-thue`.
- Không dùng cho checklist mở hồ sơ TTKT. Việc đó là `mo-ho-so-ttkt`.

## Chuẩn bị

- App: `D:\01. Cong tac kiem tra\HO SO KIEM TRA\app`
- Mở: `chay.bat` trong thư mục đó. Trang: `http://127.0.0.1:8000`
- Script (không cài thêm gói, không đụng `.venv` của Javis):

```
python "D:\SECOND BRAIN\javis-os\shared-skills\lap-ho-so-kiem-tra\scripts\chay_vu.py" status
```

Thiếu lệnh `python` thì dùng `C:\Users\NGOC\AppData\Local\Programs\Python\Python312\python.exe`.

Script đăng nhập tài khoản nhập liệu `canbo` đã có trong app. Đổi bằng biến `HOSO_USER` và `HOSO_PASS` nếu mật khẩu đã đổi. Không dùng tài khoản admin cho việc nhập và xuất.

Danh mục mẫu và cột Excel: `references/mau-van-ban.md`.

Engine không có shell: đưa nguyên lệnh, nói rõ chưa chạy. Không bịa đường dẫn file xuất.

## Cách chạy

```
python "...\lap-ho-so-kiem-tra\scripts\chay_vu.py" status
python "...\lap-ho-so-kiem-tra\scripts\chay_vu.py" import "D:\duong\dan\vu.xlsx"
python "...\lap-ho-so-kiem-tra\scripts\chay_vu.py" xuat --mst 3603132097 --qd "4988/QĐ-TCS3" --docs DI-03,BB-05
```

Bỏ `--docs` thì script lấy đúng các ô đang tick trên trang xuất của vụ đó (nhóm luôn có, cộng mẫu theo hướng xử lý và phát sinh đã bật).

`status` in `dang chay` hoặc `chua chay`. Lệnh này không tự mở app.

## Quy trình

1. Lấy MST, số quyết định, đường dẫn Excel nếu có, và mã mẫu cần xuất. Thiếu MST thì hỏi một câu rồi dừng. Thiếu Số QĐKT: hỏi rõ; nếu user chọn **xuất trình ký trước** thì dùng mã tạm `DU_THAO-TRINH-KY` (không bịa số chính thức).
2. Chạy `status`. In `chua chay` thì khởi động `chay.bat` nền, đợi rồi gọi lại `status`. Giữ cổng 8000. Cửa sổ đen cũ còn chiếm cổng thì đóng cửa sổ đó trước, không mở cổng khác.
3. Có file `.xlsx` đúng mẫu app thì chạy `import`. Đọc số tạo mới, số cập nhật và mọi cảnh báo. Cùng MST + số quyết định là cập nhật vụ cũ, không tạo vụ thứ hai.
4. Chạy `xuat`. Script dừng vì lỗi (thiếu tên, MST, trưởng đoàn, thành viên) hoặc vì cảnh báo cần xác nhận thì đưa nguyên chữ cảnh báo cho user. Không tự xác nhận, không xuất lần hai. `pick_vu` lệch chuỗi QĐ thì liệt kê `id=` và xuất đúng `vu_id`, hoặc sửa số QĐ trên app cho khớp.
5. Trả thư mục `output\<MST>\<số quyết định, dấu / đổi thành gạch ngang>\<mốc thời gian>\` và file `hoso.zip`. Nhắc mở Word để điền chỗ ký. Có Số QĐKT thật sau thì cập nhật vụ và xuất lại.

Xem trước trên web lấy form đang gõ. Xuất Word lấy dữ liệu đã lưu. Import ghi thẳng vào cơ sở dữ liệu nên xuất được ngay sau import.

Sổ theo dõi Excel JAVIS PRO + thêm vụ từ DB NNT: làm theo skill cùng slug trong brain `Hồ sơ kiểm tra` (`skills/lap-ho-so-kiem-tra/`).

## Bẫy

- `HUONG_DAN.md` và `TIEN_TRINH.md` còn ghi `D:\HO SO KIEM TRA\app`. Thư mục đó không còn. Dùng đường dẫn ở mục Chuẩn bị.
- Không sửa file trong `HO SO NGUON`. App xuất từ `templates\clean`.
- BB-03 chỉ thành viên (`vai_tro` = tv). BB-04 chỉ trưởng đoàn (`td`). BB-05 đủ chỉ tiêu cộng 3 dòng thuế (`tong_hop`).
- Số âm trên Excel ghi dạng `(40.829.663)`.
- Thời kỳ kiểm tra ghi `Từ năm 2022 đến năm 2025` thì app mới tách bảng từng năm.
- Thiếu giải trình khi có chênh lệch vẫn xuất được; ô trống trên Word tô đỏ. Script coi đó là cảnh báo và dừng để user quyết.
- VE-KL chỉ xuất được khi đã có mẫu kết luận trong `templates\clean`.
- Không ký số, không xuất PDF, không bịa số tiền.

## Kiểm chứng

- `status` thoát 0 và in `dang chay` khi app mở; thoát khác 0 và in `chua chay` khi cổng 8000 tắt.
- Sau `import`, stdout có dòng `tao=` và `cap_nhat=`.
- Sau `xuat` thành công, stdout có `OK` và ít nhất một đường dẫn dưới `output\`, gồm `hoso.zip` khi có file Word.
- Câu trả lời không thêm số tiền ngoài file Excel, trang app, hoặc lời user đã đưa.
