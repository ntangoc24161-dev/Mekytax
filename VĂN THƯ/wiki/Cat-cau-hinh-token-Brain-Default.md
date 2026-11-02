---
type: wiki
status: active
tags: [wiki, van-hanh, token]
created: 2026-10-23
updated: 2026-10-23
source: usage_index + cấu hình Brain Default
---
# Cắt cấu hình token - Brain Default

## Hiện trạng đo được (tháng 10/2026)

Theo chỉ mục `usage_index` (trang Mức dùng Javis):

- Brain Default xếp **hạng 2** token tháng này: ~**41.8M** token, ~**422** lượt, quy đổi ~**$11.8**.
- Phần lớn là **background** (362 lượt), model chủ đạo **claude-haiku-4-5**.
- Phòng **không có agent / workflow / loop**, nhưng vẫn là nơi Telegram/Zalo hay mở phiên đa việc (phân tích hồ sơ, công an, ôn thi...).
- Trước khi cắt: Kanban **orchestration=auto**, 1 nhắc cron **07:00 mỗi ngày** (`muc_quyen: full`, kể chuyện Telegram), ~**31-33 skill** chuyên sâu chồng lên các phòng khác.

Nguồn đo: `server/usage_index.db` project=`Brain Default`; cấu hình: `brains/Brain Default/Javis/kanban.json`, `reminders.json`, `skills/`.

## Quyết định đã áp (2026-10-23)

1. **Tắt AI tự vận hành** Kanban của Brain Default: `orchestration: off` (API `POST /kanban/orchestration`).
2. **Huỷ nhắc hẹn** `r_24e0350726` (cron `0 7 * * *`, kể chuyện Telegram mỗi sáng) qua `POST /reminders/cancel`.
3. **Tắt (không xoá)** 23 skill chuyên môn trùng phòng khác, chuyển vào `skills/.disabled/`. Giữ 8 skill nhẹ cho inbox:
   - AI hệ thống: `ingest-source`, `javis-builder`, `lint-wiki`, `notes`, `query-wiki`
   - OCR/tài liệu: `paddleocr`, `pdf-sang-md`, `tram-tai-lieu`
4. Dung lượng mô tả skill đang bật còn ~**36 KB** (trước cắt ~**321 KB** chỉ tính `SKILL.md`).

Skill đã tắt (có thể bật lại bằng `POST /skills/toggle`): bộ phân tích hồ sơ thuế / GTGT-TNDN / NCC / chính sách thuế / soạn thảo / iOffice / trả lời công an / ôn thi / webcake...

## Mục tiêu vận hành sau cắt

- Brain Default = **hộp thư / điều hướng**, không phải phòng chuyên môn.
- Việc phân tích thuế, công an, văn thư, ôn thi chạy ở brain chuyên tương ứng.
- Việc nền trên Brain Default chỉ chạy khi anh bật lại orchestration hoặc bấm chạy tay.

## Việc còn mở

- Gắn kênh Telegram/Zalo vào brain đúng việc (tránh mọi chat rơi vào Default).
- Quyết định có archive hẳn Brain Default hay giữ làm inbox nhẹ.
- Xem [[Đánh giá token theo phòng]] khi trang đó được lập đầy đủ.

## Mâu thuẫn

Chưa có. Số liệu tháng này phản ánh cả trước và sau cắt; hiệu quả tiết kiệm chỉ đo rõ từ các ngày sau 2026-10-23.
