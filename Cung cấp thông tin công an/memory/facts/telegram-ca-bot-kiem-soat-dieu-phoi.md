---
type: preference
provenance: user
origin: javis-learned
created: 2026-11-18
updated: 2026-11-20
source: chat 2026-11-20 (mwlp8ddh); chốt lại 2026-11-20 (mwnu8fh5tv6n99vy)
---
Phiên Telegram làm việc hồ sơ Công an đi qua bot **Trợ lý quản lý công văn CA** (`@Kiem_soat_dieu_phoi_bot`, account `acc_656f8aaa4e`, chatbot `bot_4e7cc7806a`), agent `kiem-soat-dieu-phoi-quy-trinh`.

**Kênh nhận file xem trước / dự thảo (chốt 20/11/2026):** cùng bot CA trên. Gửi bằng script `scripts/gui-telegram-bot-ca.py` tới chat chủ `8799384849`. **Không** dùng `POST /telegram/send-file` (bot Javis chính ở Cài đặt → Telegram) cho việc CA.

Chat chủ Telegram `8799384849` đã map (file `tg_brain.json` / `tg_agent.json`) sang brain `Cung cấp thông tin công an` + agent Điều phối khi nhắn bot Javis chính; việc CA vẫn ưu tiên nhắn thẳng `@Kiem_soat_dieu_phoi_bot`.

**Vì sao:** Tách việc CA khỏi Brain Default / bot Javis chính; file xem trước và lệnh phân công cùng một khung bot CA.

**Áp dụng:** Lệnh phân công/Zalo/KT, xem trước `[XEM TRƯỚC]`, dự thảo Word gửi vào bot CA. Đổi model bot = đổi model agent Điều phối.

**Telegram mặc định của brain này:** bot Trợ lý quản lý công văn CA (`@Kiem_soat_dieu_phoi_bot`). Mọi gửi file/xem trước/phân công từ brain này đi bot này, không dùng bot Javis chính.
