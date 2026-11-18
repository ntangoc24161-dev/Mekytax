---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
source: chat 2026-11-18 (mwlp8ddh)
---
**Đã giải quyết (18/11/2026, Gói A rồi chốt 1A+2A+3A+4A).**

Trước đó hai loop cùng tên từng chạy song song (stub 720 phút và conflict-remote 1440 phút). Hiện chỉ còn **1 file** `Javis/loops/ho-so-cung-cap-cong-an.md` (quy trình đầy đủ, `enabled: false` sau 2A), mốc quét cố định là reminder cron `r_7c63a0c056` lúc **12:00** mỗi ngày.

**Vì sao:** Chủ chọn Gói A gộp loop, rồi chọn 2A để còn đúng 1 kênh quét lúc 12h, tránh nhiễu 2 lần/ngày.

**Áp dụng:** Không khẳng định còn 2 loop. Fact này chỉ còn là lịch sử đã đóng; trạng thái sống lấy từ file loop + reminder API.
