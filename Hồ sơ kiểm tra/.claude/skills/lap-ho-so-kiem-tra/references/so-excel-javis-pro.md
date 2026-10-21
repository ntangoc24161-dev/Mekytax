# Sổ Excel JAVIS PRO - cấu trúc

## Sheet

- `00_DASHBOARD` - KPI, SLA theo QĐ 1161 / Luật 108.
- `01_DA_BAN_HANH` - đã ban hành (có Số QĐKT, hoặc user xếp vào đây để trình ký).
- `02_DANG_THUC_HIEN` - đang thực hiện (chưa có Số QĐKT).
- `DANH_MUC` - danh mục (CQT, cán bộ, trạng thái...).

Đã bỏ: `LICH_SU_2024`, `LICH_SU_2025`, `DU_LIEU_IN_AN`.

## Chuyển trạng thái (macro)

Mở `.xlsm`, bật Enable Content. Đổi trạng thái:

- `1. Đã ban hành` trên `02_DANG_THUC_HIEN` → chuyển sang `01_DA_BAN_HANH`.
- `2. Đang thực hiện` trên `01_DA_BAN_HANH` → chuyển sang `02_DANG_THUC_HIEN`.

Script gắn macro: `scripts/inject_status_move_vba.py`.

## Cột thông tin công ty cần giữ

Thời kỳ kiểm tra, Ngành nghề KD, Giám đốc, Giới tính, Số CCCD, Ngày cấp CCCD, Nơi cấp CCCD, Kế toán trưởng, CCCD Kế toán, Số GCN ĐKKD, Ngày cấp ĐKKD, Chi nhánh, Tài khoản ngân hàng.

## Cột đã bỏ

Hạn công bố QĐ, Hạn ký biên bản KT, Hạn nộp tiền, Tình trạng nộp, và phần theo dõi nộp NSNN.

## Cảnh báo / định dạng

- Cột `Thông báo quá hạn` báo khi quá hạn mà chưa thực hiện.
- Font Times New Roman toàn file.
- Layout hiện tại khoảng **50 cột** (nhóm thông tin công ty đã bổ sung lại).

## Script tái tạo

- `scripts/create_redesigned_excel.py` - dựng cấu trúc / đổ lại dữ liệu.
- `scripts/javis_tax_tracker.py` - đọc và báo cáo sổ.

## Ví dụ hồ sơ phiên 2026-10-01

- MST `3601409441` - CÔNG TY TNHH MTV BAO BÌ GIẤY VẠN PHÁT THÀNH.
- Thời kỳ: Từ năm 2016 đến năm 2025.
- Đoàn: Nguyễn Thị Ánh Ngọc (TĐ), Nguyễn Thành Đồng (TV).
- Xuất DI trình ký: `app\output\3601409441\DU_THAO-TRINH-KY\20261001-012121\`.
