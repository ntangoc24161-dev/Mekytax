---
type: preference
provenance: user
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
---
Phiên Telegram làm việc hồ sơ Công an đi qua bot **Trợ lý quản lý công văn CA** (`Kiem_soat_dieu_phoi_bot`), agent `kiem-soat-dieu-phoi-quy-trinh`.

Chat chủ Telegram `8799384849` đã map (file `tg_brain.json` / `tg_agent.json`) sang brain `Cung cấp thông tin công an` + agent Điều phối; cần Restart Telegram hoặc `/brain` + `/agent` để RAM nạp lại.

**Vì sao:** Chat chủ mặc định Brain Default (giọng nàng/ta) không có tool Zalo/ngữ cảnh CA; lệnh "Chuyển zalo cho kt2" bị Default trả lời sai.

**Áp dụng:** Lệnh phân công/Zalo/KT gửi vào bot CA. Không dùng Brain Default cho việc CA. Đổi model bot = đổi model agent Điều phối (không có ô model riêng trên bot).
