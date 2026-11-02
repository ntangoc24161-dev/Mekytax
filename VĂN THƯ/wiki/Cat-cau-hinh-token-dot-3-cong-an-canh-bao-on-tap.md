---
type: wiki
status: active
tags: [wiki, van-hanh, token]
created: 2026-10-23
updated: 2026-10-23
source: usage_index + cấu hình 3 brain
---
# Cắt cấu hình token đợt 3 - Công an / Cảnh báo / Ôn tập

Liên quan: [[Cat-cau-hinh-token-Brain-Default]], [[Cat-cau-hinh-token-Phan-tich-ho-so-thue]], [[Cat-cau-hinh-token-Ho-so-kiem-tra]].

## Hiện trạng đo được trước cắt

| Brain | Token tháng (ước từ usage) | Orch trước | Điểm nóng |
|---|---:|---|---|
| Cung cấp thông tin công an | ~12.3M | auto | 4 agent, workflow active, 38 skill (~485 KB), nền Opus/Haiku |
| Cập nhật công văn cảnh báo | ~1.2M | auto | workflow active, việc blocked nhầm (thu thập đề thi), 22 skill |
| Ôn tập thi | ít trong by_project | auto | 0 agent, 33 skill (~397 KB), nhiều task archived giải đề |

## Quyết định đã áp (2026-10-23)

### Cung cấp thông tin công an
1. `orchestration: off`
2. Tắt 3 skill ngoài phạm vi: `html-to-webcake`, `phan-tich-cau-hoi-sinh-doi-cheat-sheet`, `tien-trinh-supervisor-tu-dong`
3. **Giữ** workflow `quản-lý-công-văn-công-an` active (chạy tay) và skill gắn 4 agent (điều phối, soạn thảo, khai thác, pháp chế)

### Cập nhật công văn cảnh báo
1. `orchestration: off`
2. Huỷ việc blocked sai phòng `t_50db6f95c23a` (thu thập đề thi P3-39917)
3. Tắt 9 skill lệch: GĐ1-3, phân tích hồ sơ/công nợ, xử phạt, webcake, cheat-sheet, supervisor
4. **Giữ** skill chuỗi cảnh báo: theo dõi CV, đối chiếu NCC, iOffice, OCR, soạn thảo/trả lời

### Ôn tập thi
1. `orchestration: off`
2. Tắt 10 skill vận hành lệch phòng (GĐ1-3, phân tích hồ sơ/NCC/công nợ, soạn thảo, iOffice, trả lời CV, webcake)
3. **Giữ** skill ôn/giải đề + chính sách phục vụ thi + OCR/hệ thống

## Tổng kết đợt cắt cả chuỗi (2026-10-23)

Các brain đã `orchestration=off`: Brain Default, Phân tích hồ sơ thuế, Hồ sơ kiểm tra, Cung cấp thông tin công an, Cập nhật CV cảnh báo, Ôn tập thi. (VĂN THƯ vốn đã off.)

## Việc còn mở / đã xử lý thêm

- Loop VĂN THƯ `van-thu-dieu-phoi-don-doc`: đã giãn từ **60 → 180 phút** (2026-10-23), vẫn bật, `max_runs_per_day: 10`, quiet hours 18-07.
- Ghim model nhẹ theo agent.
- Đo lại token sau 7 ngày trên trang Mức dùng.
- Các brain gần như không dùng (IFOFFIC, KIỂM TRA THUẾ, Brain tính cup, tổng hợp ôn tập): archive hoặc gộp nếu anh muốn dọn tiếp.
