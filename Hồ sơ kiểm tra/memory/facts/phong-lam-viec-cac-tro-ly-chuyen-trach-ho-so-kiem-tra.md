---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
---
Người dùng thiết lập phòng làm việc chuyên trách cho brain Hồ sơ kiểm tra gồm các vai:
- Trợ lý khởi tạo hồ sơ (`agents/khoi-tao-ho-so.md`): cập nhật thông tin dữ liệu công ty, gắn các skill `lap-ho-so-kiem-tra`, `nhap-bang-tinh-tu-bao-cao-hdkd`, `boc-tach-bang-tinh-phan-tich`, `xem-ho-so-thue`.
- Trợ lý Soạn thảo hồ sơ đi (`agents/soan-ho-so-di.md`): tạo văn bản chuẩn thể thức hồ sơ đi trình ký, gắn `ho-so-kiem-tra-ke-hoach`, `soan-thao-van-ban`.
- Trợ lý Soạn thảo biên bản và hồ sơ về (`agents/soan-bien-ban-ho-so-ve.md`): tạo văn bản thể thức biên bản và hồ sơ về, trình ký QĐ xử phạt VPHC/kết luận, gắn `soan-thao-van-ban`, `dien-bien-ban-tu-bang-tinh`, `xu-phat-vphc-thue-hoa-don`.
- Trợ lý Tri thức (`agents/tri-thuc-second-brain.md`): quản trị wiki, trả lời có trích dẫn, gắn `ingest-source`, `query-wiki`, `lint-wiki`.
