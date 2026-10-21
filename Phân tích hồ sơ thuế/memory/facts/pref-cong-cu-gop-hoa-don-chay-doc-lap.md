---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-06
updated: 2026-10-06
---
Khi tổng hợp hóa đơn mua vào/bán ra để mang sang máy khác, user muốn công cụ chạy độc lập: không cài Python/Excel, mỗi công ty một file Excel 2 sheet Mua vào / Bán ra có tiêu đề. Sau bản HTML còn đòi exe: bấm vào, thêm file hoặc thư mục rồi chạy, không đòi quyền admin.

Bản HTML tại `apps/gop-hoa-don-mua-ban/` (copy cả thư mục kèm `vendor/` và `mau/`). Bản exe portable đã đóng tại `apps/gop-hoa-don-mua-ban/GopHoaDon.exe` (Windows 64-bit) kèm `GopHoaDon-portable.zip`.

**Vì sao:** User giao mini app chạy mọi máy không cần cài; sau khi có HTML còn đòi exe một nút không admin.
**Áp dụng:** Đưa `GopHoaDon.exe` (hoặc folder HTML). Không bảo cài Python. Không đòi quyền admin. Máy đích cần Edge hoặc Chrome. Lần đầu SmartScreen: Thông tin thêm -> Vẫn chạy. Chưa bấm thử trên Windows.
