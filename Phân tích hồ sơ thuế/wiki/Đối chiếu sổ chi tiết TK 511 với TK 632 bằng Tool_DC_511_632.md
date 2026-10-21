---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-04
updated: 2026-10-04
source: [[conversations/2026-10-04]]
---
Tool_DC_511_632.xlsm là mini-app Excel (macro `ChayDoiChieu`, module `DoiChieu511632`) đối chiếu từng dòng sổ chi tiết doanh thu TK 511 (Phát sinh Có) với sổ chi tiết giá vốn TK 632 (Phát sinh Nợ). [[Tool_DC_511_632.xlsm]]

File gồm 7 sheet: `BangDieuKhien`, `511`, `632`, `DCDTCP`, `BaoCao`, `CauHinh`, `HuongDan`. [[Tool_DC_511_632.xlsm]]

Cột bắt buộc A→G đúng thứ tự: Ngày hạch toán, Ngày chứng từ, Số chứng từ, Số hóa đơn, Diễn giải, Tài khoản, Số tiền (511 = PS Có; 632 = PS Nợ); không đổi tên sheet `511`/`632`. [[Tool_DC_511_632.xlsm]]

Sheet kết quả `DCDTCP` có 3 khối: A–G dòng 511, I–O dòng 632 ghép, P–S kết quả (P lãi/lỗ = G−O; Q khớp diễn giải; R tình trạng; S năm chứng từ). [[Tool_DC_511_632.xlsm]]

Ba khóa khớp theo thứ tự ưu tiên (thực tế theo cấu hình tool): (1) Số HĐ đã bỏ 0 đầu + ngày CT + diễn giải mặt hàng; (2) Số HĐ + diễn giải, chấp nhận lệch ngày CT; (3) Cùng ngày CT khi thiếu HĐ hoặc tên hàng lệch. [[Tool_DC_511_632.xlsm]]

Năm nhãn tình trạng R: khớp đủ 3 khóa; khớp HĐ+mặt hàng khác ngày CT; ghép theo ngày CT khác HĐ/mặt hàng; Không khớp 632; Có giá vốn không có doanh thu. [[Tool_DC_511_632.xlsm]]

Sheet `BaoCao` tổng hợp theo năm chứng từ (2018–2030): số dòng theo R, DT/GV/lãi gộp/tỷ lệ lãi gộp, PS Nợ 632 theo tình trạng, nhóm bán dưới giá P≤0 và tỷ lệ DT dưới giá; công thức loại dòng bắt đầu bằng TỔNG/TONG. [[Tool_DC_511_632.xlsm]]

Vận hành 4 bước: dán 511 → dán 632 → Alt+F8 `ChayDoiChieu` → lọc `DCDTCP` cột R/P và đọc `BaoCao`; xuất Yes ra file riêng trước khi chạy kỳ mới vì macro ghi đè `DCDTCP`. [[Tool_DC_511_632.xlsm]]
