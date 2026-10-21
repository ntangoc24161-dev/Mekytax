---
superseded_by: decision-chot-kho-skill-chung-shared-skills
type: decision
created: 2026-10-04
updated: 2026-10-04
---
**Chốt:** đường chuẩn của skill chung là `D:\SECOND BRAIN\javis-os\shared-skills/<slug>/`.

**Vì sao:** Đây là kho chung đúng theo router Javis (`skill_router`), gắn theo brain qua `.javis/shared-skills.json`. `Brain Default/skills/` chỉ là skill local của não mặc định, không phải kho chung toàn hệ.
**Áp dụng:** Từ 2026-10-04, khi user yêu cầu “skill chung” thì chỉ ghi/cập nhật `shared-skills/`. Không bắt buộc dọn bản trùng đang nằm ở `Brain Default/skills/` (giữ nguyên nếu đã có). Bản local trong brain đang làm việc (`<brain>/skills/`) vẫn thắng bản chung khi cùng slug.


## Lịch sử
- [2026-10-04] Bị thay thế bởi [[decision-chot-kho-skill-chung-shared-skills]] (user đổi thông tin).
