# Hướng dẫn sử dụng - Phân tích hồ sơ kiểm tra thuế

Bản mang USB. Máy Windows 64-bit. **Không cài Python, không cần mạng, không đăng nhập HDDT/TMS.**

Bản đủ hai công cụ (kèm PDF sang Markdown) nằm tại `D:\SOFT\TEST TOOL\HUONG-DAN-SU-DUNG.md`.

## 1. Cài ra máy (giải nén, không phải cài đặt)

1. Copy file `PhanTichHoSoThue-portable.zip` sang USB hoặc ổ đĩa.
2. **Giải nén** ra một thư mục (ví dụ `D:\PhanTichHoSoThue\`). Không chạy trực tiếp trong file zip.
3. Trong thư mục giải nén có:
   - `Chay.bat` - bấm cái này
   - `PhanTichHoSoThue\PhanTichHoSoThue.exe` - chương trình
   - `HUONG-DAN-SU-DUNG.md` - file này
   - `JAVIS-TAO-SKILL.md` - đưa cho Javis để tạo skill
   - `DOC.txt` - bản rút gọn

Windows SmartScreen có thể hỏi lần đầu: **More info** rồi **Run anyway**.

## 2. Chuẩn bị thư mục hồ sơ

Gom file **đã xuất sẵn** vào một thư mục (có thể có thư mục con). Chương trình quét hết.

| Loại | Định dạng | Ghi chú |
|---|---|---|
| Tờ khai GTGT (01/GTGT) | XML | Cổng thuế điện tử / HTKK |
| Quyết toán TNDN (03/TNDN) | XML | |
| Quyết toán TNCN | XML | |
| Báo cáo tài chính | XML | TT200/TT133 |
| Bảng kê hóa đơn mua/bán | Excel `.xlsx` | Sheet `2.BR` (bán ra), `4.MV` (mua vào) |

Không cần đặt tên file đúng mẫu. Chương trình nhận XML theo thẻ `<maTKhai>`:

- `842` = GTGT
- `892` = TNDN
- `953` = TNCN
- `402` = BCTC

Tờ khai **bổ sung** (`loaiTKhai` = B) được ưu tiên hơn tờ **chính thức** (C) cùng kỳ.

Bỏ qua: thư mục tên `BIEN BAN`, file `PTDN01` / `PTDN02` (báo cáo tool cũ).

## 3. Chạy phân tích

**Cách 1 - chuột:** double-click `Chay.bat` → chọn thư mục hồ sơ → chọn thư mục kết quả (bấm Hủy thì ghi vào `<hồ sơ>\ket-qua`).

**Cách 2 - kéo thả:** kéo thư mục hồ sơ thả vào `Chay.bat`. Kết quả: `<hồ sơ>\ket-qua\`.

**Cách 3 - dòng lệnh:**

```
PhanTichHoSoThue\PhanTichHoSoThue.exe --input-dir "D:\hs\3603859012" --out "D:\hs\3603859012\ket-qua"
```

Cửa sổ in dòng `OK MST ... findings` là xong.

## 4. Đọc kết quả

Trong thư mục đích:

- `phan-tich-<MST>-<nam>.xlsx` - báo cáo làm việc (mở Excel). Mỗi loại thuế một sheet, dòng **Cao** tô đỏ
- `phan-tich-<MST>-<nam>.md` - báo cáo chữ (nạp Javis)
- `phan-tich-<MST>-<nam>.json` - danh sách phát hiện có cấu trúc

Bố cục file Excel:

1. Sheet **Tóm tắt** - điểm rủi ro, số mục cao / trung bình / thấp, ưu tiên giải trình
2. Sheet **GTGT**
3. Sheet **TNDN**
4. Sheet **TNCN**
5. Sheet **Hóa đơn**
6. Sheet **Phụ lục** - cảnh báo đọc file

Mỗi sheet thuế: bảng nhóm phát hiện, rồi bảng chi tiết (nội dung bất thường, căn cứ, câu hỏi chất vấn). Màu đỏ = Cao, cam = Trung bình, xanh = Thấp.

## 5. Việc chương trình làm / không làm

**Làm:** đọc XML + Excel trên đĩa, đối chiếu HĐ với tờ khai, chạy các tiêu chí rủi ro BCTC/TNDN lấy từ tool 4.4 (sheet TUDIEN), ghi báo cáo.

**Không làm:** đăng nhập TMS, HDDT, thuế điện tử, DVC. Không sửa file gốc của hồ sơ.

Tiêu chí cần dữ liệu TMS (ví dụ TC24) **không chạy**. Thiếu XML BCTC thì phần chỉ tiêu BCTC sẽ ít hơn - không bịa số.

## 6. Lỗi thường gặp

| Hiện tượng | Việc làm |
|---|---|
| SmartScreen chặn | More info → Run anyway |
| `không đọc được XML tờ khai hay bảng kê` | Sai thư mục, hoặc XML không phải `HSoThueDTu`, Excel không có sheet BR/MV |
| Excel không ra hóa đơn | Cần sheet tên có `BR` / `MV` / `ban ra` / `mua` |
| Antivirus xóa exe | Thêm thư mục giải nén vào ngoại lệ |

## 7. Đưa báo cáo vào Javis

1. Copy file `.md` vào `sources/` của brain (ví dụ brain KIỂM TRA THUẾ).
2. Trong chat Javis: bảo "tiêu hoá source vừa thêm" nếu muốn lên wiki.
3. Để Javis **tự biết cách chạy tool** trên máy có Javis: đưa file `JAVIS-TAO-SKILL.md` (cùng zip) và nói "tạo skill theo file này".
