---
type: preference
created: 2026-11-18
updated: 2026-11-20
source: chat 2026-11-18 15:27 (mwlp8ddh) - lựa chọn "OK quy trình lần sau"; chat 2026-11-20 chuyển bot CA
---
Khi phân công hồ sơ Công an, **bản xem trước gửi qua Telegram bot CA** `@Kiem_soat_dieu_phoi_bot` (script `scripts/gui-telegram-bot-ca.py`) trước; người dùng OK rồi mới bàn giao thật.

**Vì sao:** Người dùng muốn kiểm tra trên Telegram (tiện xem file) trước khi gửi Zalo KT2/KT3 hoặc chốt in KT1; từ 20/11/2026 gom lệnh + file CA về cùng bot CA, không dùng bot Javis chính.

**Áp dụng:**
- KT1: skill `in-phan-cong-kt1` → bot CA xem trước → chốt in → cập nhật `Đang xử lý`.
- KT2/KT3: skill `gui-phan-cong-zalo-kt23` → bot CA xem trước (tin nháp + file) → chốt gửi → Zalo đầu mối → cập nhật `Đang xử lý`.
- Không gửi lại Zalo cho CV đã bàn giao (1806, 2126, 9372) trừ khi người dùng lệnh rõ "gửi lại".
- Không dùng `POST /telegram/send-file` cho việc CA.
