---
type: preference
provenance: user
origin: javis-learned
created: 2026-11-25
updated: 2026-11-25
source: chat 2026-11-25 (43c376b1)
---
Khi có yêu cầu lập hồ sơ đi, biên bản hoặc hồ sơ về cho các công ty, mặc định sử dụng phôi Mail Merge chuẩn:
- Hồ sơ kiểm tra theo kế hoạch: `templates/MAUTONG HOP.docx` (kết nối trực tiếp sheet `01_DA_BAN_HANH$` qua Dòng 1 ẩn).
- Hồ sơ kiểm tra giải thể: `templates/MAUTONG HOP GIAI THE.docx`.

**Vì sao:** Thống nhất quy trình Hybrid 2 tầng, tránh sai sót nguồn dữ liệu và không dùng các script chèn text tĩnh cũ.
**Áp dụng:** Tự động áp dụng cho các trợ lý `soan-ho-so-di`, `soan-bien-ban-ho-so-ve` và các skill lập hồ sơ liên quan.
