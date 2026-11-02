# Wiki Log

Nhật ký thời gian (append-only). Mỗi entry: `## [YYYY-MM-DD] loại | tiêu đề` (loại: init/ingest/update/query/lint/migrate).

## [2026-10-23] update | Cắt cấu hình token Brain Default
- Tạo: [[Cat-cau-hinh-token-Brain-Default]]
- Việc đã làm trên brain Brain Default: orchestration off, huỷ reminder r_24e0350726, tắt 23 skill chuyên môn (giữ 8 skill inbox/OCR).

## [2026-10-23] update | Cắt cấu hình token Phân tích hồ sơ thuế
- Tạo: [[Cat-cau-hinh-token-Phan-tich-ho-so-thue]]
- orchestration off; huỷ task blocked t_62548b57a135; tắt 8 skill ngoài phạm vi; giữ 24 skill + workflow 3 giai đoạn + 6 agent.

## [2026-10-23] update | Cắt cấu hình token Hồ sơ kiểm tra
- Tạo: [[Cat-cau-hinh-token-Ho-so-kiem-tra]]
- orchestration/workflow đã off; tắt 11 skill lệch phòng; còn 15 skill lập/soạn hồ sơ + OCR/hệ thống.

## [2026-10-23] update | Cắt token đợt 3 Công an / Cảnh báo / Ôn tập
- Tạo: [[Cat-cau-hinh-token-dot-3-cong-an-canh-bao-on-tap]]
- 3 brain orchestration off; huỷ task đề thi nhầm trên Cảnh báo; tắt skill lệch từng phòng.

## [2026-10-23] query | Ước tiết kiệm sau cắt
- Tạo: [[Uoc-tinh-tiet-kiem-token-sau-cat]]
- Cơ sở: background 86.3% tháng 10; orch off không xoá hết nhãn nền; loop 180 phút ~2% tổng.
