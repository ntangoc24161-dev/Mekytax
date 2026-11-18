---
type: decision
created: 2026-11-18
updated: 2026-11-18
source: chat 2026-11-18 (mwlp8ddh)
---
Ngày 18/11/2026 chủ chọn **Gói A**, rồi chốt dọn mâu thuẫn **1A + 2A + 3A + 4A**. Cùng ngày bổ sung **nhánh B**: bản trả lời đã ký trong `CV MOI` xử lý luôn lúc quét trưa.

**Vì sao:** Muốn gọn, đúng 1 kênh quét lúc 12h, soạn thảo chỉ khi có lệnh, ghi sổ đúng cả Linux và Windows; bản ký rơi vào `CV MOI` không chờ lệnh riêng.

**Áp dụng:**
- Reminder cron `r_7c63a0c056` (`0 12 * * *`) là **kênh quét duy nhất**; `chat_id` phiên dashboard hiện tại.
- Loop `ho-so-cung-cap-cong-an`: giữ mô tả quy trình, **`enabled: false`** (2A).
- Phân loại file theo **checklist cứng** (nhánh A / nhánh B / không chắc → báo chat). Chi tiết trong loop và text cron.
- **Nhánh A - CV Công an mới:** workflow `quản-lý-công-văn-công-an` chỉ 2 bước bóc tách → phân công (xem trước); **dừng**; soạn thảo chỉ khi có lệnh (3A).
- **Nhánh B - CV trả lời đã ký trong CV MOI:** workflow `theo-doi-cv-tra-loi-da-ky` chạy ngay trong lần quét trưa (`Đang xử lý` → `Đã hoàn thành`). Quét rộng ngoài `CV MOI` vẫn cần lệnh theo dõi.
- Ghi sổ: Linux/VPS = Drive Composio `googledrive_ralph-urbane` + vá XML; Windows = Excel COM (4A).
- Memory `mau-thuan-hai-loop…` và `quy-tac-in-phan-cong-kt1` đã cập nhật khớp (1A).
