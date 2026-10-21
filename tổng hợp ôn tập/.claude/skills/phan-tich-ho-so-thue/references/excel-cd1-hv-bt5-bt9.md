# Kỹ thuật Excel kiểm tra thuế - Chuyên đề 1 (CD1)

Nguồn: Hội Kế toán Thuế TP.HCM, Excel phân tích số liệu kế toán phục vụ kiểm tra thuế (6/2026).  
Dữ liệu luyện: `Du lieu_HV` (bản học viên). Đáp án BT5-BT9 tự giải 2026-10-07 (không phải file GV chính thức).  
Bản chạy mẫu: `hoc-tap/chuyen-de-1-excel-kt/ket-qua/` trong brain Phân tích hồ sơ thuế.

Dùng reference này khi cần công thức / khóa đối chiếu chi tiết. Thân skill chỉ giữ checklist nghiệp vụ.

## Bản đồ bài tập → nghiệp vụ hồ sơ

| BT | Mục tiêu | Áp vào bước phân tích |
|---|---|---|
| Pivot NKC | Tổng hợp PS theo TK / tháng từ Nhật ký chung | Bước 3 làm sạch; Bước 4 đối chiếu tổng |
| BT4 | BKBR ↔ NKC ↔ HĐĐT bán ra | Đối chiếu 3 chiều bán ra |
| BT5 | Trùng HĐ: COUNTIFS(Số HĐ + MST + giá trị chưa thuế) | Làm sạch HĐĐT / bảng kê trước đối chiếu |
| BT6 | Chuẩn hóa MST (10 / 13 số, thiếu 0 đầu, thiếu dấu `-`) | Mọi khớp NCC, CVCBRR, danh bạ |
| BT7A | Cụm HĐ cùng ngày + cùng NCC, tổng > 20 triệu | Chi tiền mặt đến 06/2025 |
| BT7B | Cùng logic, ngưỡng > 5 triệu | Chi tiền mặt từ 07/2025 (Luật GTGT 48/2024) |
| BT8 | HĐĐT mua vào ↔ BKMV ↔ PS Nợ TK 133* trên NKC | Đối chiếu thuế GTGT được khấu trừ |
| BT9 | DN thương mại: BKMV ↔ NKC (mẫu 1 Nợ 1 Có) | Khớp thuế mua vào khi sổ ghi gọn |

## Công thức chốt

### BT5 - Hóa đơn trùng
```
=COUNTIFS($F$9:$F$last;F9;$I$9:$I$last;I9;$K$9:$K$last;K9)
```
Khóa: Số HĐ + MST + Giá trị chưa thuế. Flag khi COUNTIFS > 1.

### BT6 - Chuẩn hóa MST
```
=LEN(A2)
=IF(OR(B2=10;B2=14);A2;IF(B2=9;"0"&A2;IF(B2=12;"0"&LEFT(A2;9)&"-"&RIGHT(A2;3);IF(MID(A2;10;1)="-";"0"&A2;LEFT(A2;10)&"-"&RIGHT(A2;3)))))
```
Quy tắc: độ dài 10 hoặc 14 (đã có `-`) giữ nguyên; 9 số → thêm `0` đầu; 12 số → `0` + 9 số + `-` + 3 số nhánh; thiếu `0` trước dấu `-` thì thêm; còn lại cắt/ghép dạng 10-3.

### BT7A / BT7B - Cụm cùng ngày cùng NCC
```
J = G + I   (doanh số + thuế, hoặc cột tổng thanh toán tương đương)
=SUMIFS($J$...;$F$...;F4;$D$...;D4)   # BT7A: khóa ngày + MST/NCC, ngưỡng 20.000.000
=SUMIFS($J$...;$E$...;E4;$C$...;C4)   # BT7B: khóa ngày + MST/NCC, ngưỡng 5.000.000
```
Flag dòng thuộc nhóm có tổng > ngưỡng. Ngưỡng theo kỳ phát sinh HĐ, không theo ngày chạy file.

### BT8 - Đối chiếu mua vào HĐĐT / BKMV / NKC 133
- Khóa dòng: `SoHD(bỏ 0 đầu) & "|" & MST_chuan`
- Cột `KTRA`: KHOP / #N/A khi map HĐĐT ↔ BKMV
- Đối chiếu tháng: tổng tiền thuế BKMV vs tổng PS Nợ các TK bắt đầu `133` trên NKC

### BT9 - DN thương mại (BKMV ↔ NKC)
- Khóa ưu tiên: `SoHD|GiaTri`; thử thêm `SoHD|TienThue` với dòng Nợ 133
- Đối chiếu tháng/năm: PS Nợ 133 vs tiền thuế BKMV
- Phù hợp sổ kiểu 1 Nợ 1 Có (ít phân bổ nhiều TK trên một chứng từ)

## Kết quả tự giải trên Du lieu_HV (2026-10-07)

- BT5: 1240 dòng; **9** dòng trùng; **4** nhóm
- BT6: 751 dòng; **417** MST phải chỉnh
- BT7A: 62 dòng / 34 nhóm > 20tr
- BT7B: 182 dòng / 144 nhóm > 5tr
- BT8: matched 407; BKMV_NA 529; HDDT_NA 17; thuế BK 3.153.428.532 vs Nợ 133 3.153.909.044 (lệch nhỏ theo tháng)
- BT9: matched **119/119**; thuế BK = Nợ 133 = **178.228.430**

## Lưu ý vận hành

- Excel tiếng Việt thường dùng `;` thay `,` trong công thức.
- File HV có thể không có sẵn VLOOKUP; khóa `SoHD|MST` / `SoHD|GiaTri` là logic tương đương slide.
- Số HĐ bỏ số 0 đầu trước khi ghép khóa; MST phải qua BT6 trước khi COUNTIFS / SUMIFS / VLOOKUP.
- Không kết luận “khớp tuyệt đối” khi còn #N/A: ghi rõ số dòng NA và nguyên nhân (thiếu HĐĐT portal, HĐ hủy, lệch kỳ).

## Định dạng số tiền

Số tiền trên báo cáo/chat: chuẩn VN `1.234.567,89` (`.` hàng nghìn, `,` phần lẻ). Excel `number_format`: `#,##0` / `#,##0.00`.
