---
name: Làm sạch MST và hóa đơn CD1
description: "Chuẩn hóa MST, quét HĐ trùng, cụm cùng ngày 20tr/5tr, đối chiếu BKMV-HĐĐT-NKC 133 theo CD1 BT5-BT9."
group: Tài chính
updated: '2026-10-15'
---
# Làm sạch MST và hóa đơn CD1

Chuyên đề 1 Excel kiểm tra thuế (BT5-BT9): chuẩn hóa khóa MST/Số HĐ trước mọi VLOOKUP, COUNTIFS, đối chiếu NCC/CVCBRR/GTGT. **Không** xuất file mẫu 03 tách cho DN; gắn kết quả vào bảng kê, đối chiếu HĐĐT↔BK, GTGT/133, công nợ/tiền mặt.

Công thức chi tiết: `references/excel-cd1-hv-bt5-bt9.md` (cùng folder skill này) hoặc bản gốc `skills/phan-tich-ho-so-thue/references/excel-cd1-hv-bt5-bt9.md`.

Skill trục: `phan-tich-ho-so-thue` playbook Bước 1b. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "chuẩn hóa MST", "hóa đơn trùng", "CD1", "BT5 BT6 BT7 BT8 BT9", "đối chiếu BKMV NKC 133", "cụm hóa đơn cùng ngày".
- Đã có bảng kê hoặc HĐĐT gộp (`tong-hop-dau-vao-hddt-bang-ke-nld`).

Không dùng khi user chỉ muốn chính sách thuế hoặc sắp folder.

## Chuẩn bị

1. Bảng kê và/hoặc HĐĐT đã gộp.
2. Tốt hơn: Sổ Nhật ký chung để Pivot TK/tháng và BT8/BT9 (Nợ `133*`).
3. Excel tiếng Việt: công thức dùng `;` phân tham số.

## Quy trình

Gắn kỹ thuật vào bảng biểu nghiệp vụ, không tạo `03_Lam_sach_CD1` làm kết quả DN.

1. **BT6 - MST_chuan** trên mọi sheet có MST (bảng kê, HĐĐT, NCC). 10 hoặc 14 ký tự (đã có `-`) giữ; 9 số thêm `0` đầu; 12 số → `0` + 9 số + `-` + 3 số nhánh.
2. **BT5 - HĐ trùng:** COUNTIFS khóa Số HĐ + MST_chuan + giá trị chưa thuế. Flag > 1; xuất nhóm trùng để loại hoặc giải trình **trước** khi cộng doanh số.
3. **BT5 gắn HĐĐT ↔ bảng kê:** cột `KHOP` / `CHI_CO_BANG_KE` / `CHI_CO_HDDT`. Số HĐ bỏ 0 đầu rồi ghép MST_chuan.
4. **BT7A/BT7B - cụm cùng ngày cùng NCC:** SUMIFS (doanh số + thuế) theo ngày + MST. Ngưỡng **20.000.000** đến 06/2025; **5.000.000** từ 07/2025 (Luật GTGT 48/2024). Ngưỡng theo kỳ phát sinh HĐ. Gắn sheet công nợ/tiền mặt (`phan-tich-cong-no-tien-mat`).
5. **BT8:** khóa `SoHD(bỏ 0 đầu)|MST_chuan`; thuế BKMV theo tháng vs PS Nợ TK `133*` trên NKC. Ghi số dòng #N/A từng phía.
6. **BT9 (DN thương mại, sổ 1 Nợ 1 Có):** khóa `SoHD|GiaTri`, thử `SoHD|TienThue` với dòng 133. Khớp tuyệt đối tiền thuế tháng/năm trước khi kết luận lệch kê khai. Tóm tắt nhúng `doi-chieu-gtgt-tndn-cbrr` Phần III.
7. Pivot NKC theo TK / tháng nếu có sổ.

Lưu `Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx` vào `cau-thanh/` (file chuyên đề, không phải mẫu 03). Nhúng tóm tắt vào Master.

## Định dạng số tiền

Hiển thị chuẩn VN: `.` hàng nghìn, `,` phần lẻ. Excel `#,##0` / `#,##0.00`. Cấm ô text bắt đầu bằng `=`.

## Bẫy

- Khớp MST 9/12 số với CVCBRR/danh bạ khi chưa BT6.
- Số HĐ `00044` vs `44` coi là khác nhau.
- Kết luận khớp tuyệt đối khi còn #N/A.
- Dùng `,` trong công thức Excel tiếng Việt.

## Kiểm chứng

- Cột `MST_chuan` đủ 10 hoặc 13 số có `-`.
- Nhóm trùng BT5 đã liệt kê, chưa cộng mù.
- BT8/BT9 ghi số dòng KHOP và #N/A, thuế theo tháng.
- Bước tiếp: `doi-chieu-gtgt-tndn-cbrr`, `doi-chieu-ncc-mua-vao-rui-ro`.
