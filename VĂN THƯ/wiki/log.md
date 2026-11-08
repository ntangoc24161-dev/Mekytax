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

## [2026-10-23] update | Liên kết phòng/brain
- Tạo: [[Ban-do-lien-ket-phong-brain]]
- Skill mới: `skills/ban-giao-lien-phong/SKILL.md`
- Sửa loop `van-thu-dieu-phoi-don-doc` + agent `phong-van-thu`: thêm Brain đích / hướng chạy / dòng bàn giao.

## [2026-10-23] query | Ưu nhược phương pháp liên kết
- Tạo: [[Phuong-phap-lien-ket-phong-uu-nhuoc]]

## [2026-11-06] update | Backlog vận hành liên phòng
- Tạo: [[Backlog-van-hanh-lien-phong]]
- Skill `ban-giao-lien-phong` + agent `phong-van-thu`: bước hỏi xác nhận trước khi hướng dẫn đổi brain.
- Ghim model hub: `phong-van-thu`/`so-hoa-ocr-cong-van` = haiku; `soan-thao-van-ban`/`tri-thuc-second-brain` = sonnet.
- Kiểm kê mirror (22 ứng viên gỡ) + đề xuất archive 4 brain ít dùng; chưa xoá (chờ chủ).
- Hướng dẫn `/brain` Telegram/Zalo tránh Default.

## [2026-11-06] update | Duyệt backlog 1+2+3
- Archive 22 agent mirror + 5 workflow chuyên khỏi VĂN THƯ → `_archive-mirror-2026-11-06/` (còn 4 agent hub + 4 workflow văn thư).
- Ghim model 20 agent trên Phân tích / Hồ sơ kiểm tra / Công an / Cảnh báo (haiku nhẹ, sonnet chốt); bỏ qua 2 conflict-remote.
- Tạo: [[Checklist-archive-brain-it-dung]] (chủ tự xoá brain trên UI).

## [2026-11-06] update | Open question tự đổi brain
- Chủ muốn: sau xác nhận bàn giao, Javis tự đổi brain + mở agent đích.
- Ghi `_open-questions` + cập nhật mục tiêu [[Backlog-van-hanh-lien-phong]] mục 1 (vẫn là kế hoạch sản phẩm, chưa có hiện trạng).
