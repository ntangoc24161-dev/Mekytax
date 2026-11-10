# Cấu trúc BBXN / BBKT theo năm (chuẩn phôi + VPT)

Nguồn chuẩn mắt: `templates/TONG HOP CHUAN DINH DANG.docx` và fact `chuan-dinh-dang-tong-hop-vpt-2026-10-08`.

## BBXN - mỗi năm

1. Dòng năm + ĐVT **cùng một dòng**, có tab right ~8931 twips:
   - Thành viên: `+ NĂM 20xx:` đậm + tab + `Đơn vị tính: Đồng` nghiêng cỡ nhỏ hơn.
   - Trưởng đoàn: `+Năm 20xx` đậm + tab + ĐVT nghiêng.
2. Bảng từ DTCHEP: `TT | CHỈ TIÊU | SỐ BÁO CÁO | SỐ KIỂM TRA | CHÊNH LỆCH`.
   - Thành viên (CPTC): 1 chỉ tiêu/năm, TT = 1.
   - Trưởng đoàn (QLDN + bán hàng): **2 chỉ tiêu trong 1 bảng/năm**, TT = 1 cho QLDN, TT = 2 cho bán hàng (không tách 2 bảng).
3. Có chênh: thêm dòng tiêu đề `Giải trình số liệu chênh lệch qua kiểm tra:`, xuống hàng mới `*/` + nội dung giải trình + căn cứ (TT 96/2015 trước 2025; NĐ 320/2025 từ 2025).
4. Không chênh: chỉ `*/ Không chênh lệch.` (viết thường, không có dòng tiêu đề "Giải trình...").
5. Không ghi TNTT / thuế truy thu ở BBXN.

Mục `I. Nội dung kiểm tra` / `II. Kết quả kiểm tra`: đậm, canh đều, thụt ~1.0 cm. Tên chỉ tiêu/chuyên đề chi phí ghi thẳng, **không gạch đầu dòng `-`**, thống nhất dùng **`Chi phí quản lý doanh nghiệp`** (không dùng tên cũ `Chi phí quản lý kinh doanh`).

## Checkbox Báo cáo giám sát (mục I, 9 câu)

- Chuẩn từ 2026-10-27: ô trống **☐** (U+2610), ô đã chọn **☑** (U+2611), đặt **ngay sau nhãn chữ** đúng vị trí phôi gốc (không Wingdings, không đặt đầu chữ).
- Phôi `templates/MAUTONG HOP.docx` đã có sẵn 23 ô (10 ☑ mặc định + 13 ☐) theo đúng vị trí này; chỉ cần giữ nguyên khi copy khối hoặc đổi lựa chọn theo thực tế vụ, không tự ý reset.
- Mục 9 có 3 lựa chọn cùng hàng ("Có thực hiện kiểm tra" / "Không thực hiện kiểm tra" / "Không có thông tin thu thập từ bên ngoài") - cả 3 đều có ô riêng.

## BBKT - mỗi năm

- Dòng năm: `+Năm 20xx` + tab + `Đơn vị tính: Đồng` (cùng dòng).
- Có chênh: Giải trình (đậm) → Công ty báo TNTT... → `*/` nguyên nhân → `=>Thuế TNDN truy thu...` → Giảm lỗ nếu có.
- Không chênh: `*/ Không chênh lệch.`
- Đủ mục III (nghĩa vụ) và IV (kết luận/kiến nghị) theo phôi. Với vụ kiểm tra giải thể: Mục III Nghĩa vụ NSNN bắt buộc điền đủ số phát sinh phải nộp các năm theo công thức `(3)=0`, `(6)=(4)`, `(7)=(3)+(4)+(5)-(6)`, không để trống hay gạch ngang.
- Đoạn mở "Hôm nay, vào lúc..." canh trái, tránh giãn chữ.

## CPTC

Câu thống nhất: `Công ty hạch toán chi phí lãi vay không tương ứng doanh thu`.