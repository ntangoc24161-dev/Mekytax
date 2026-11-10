---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-11
updated: 2026-11-11
---
Ngày 11/11/2026, user chỉ đạo agent **Thu thập công văn cảnh báo** (slug `tra-cuu-tai-cv-ioffice`) thôi đăng nhập VNPT iOffice. Nguồn tiếp nhận gồm 2 luồng riêng:
- Thư mục `HO SO MOI` trên Drive (hộp thư hiện tại).
- Tệp gửi qua Telegram/Zalo: user sẽ liên kết Javis Core với Telegram để chuyển dữ liệu; agent **tự lấy, xử lý và bàn giao trực tiếp**, user KHÔNG phải chuyển tay tệp vào HO SO MOI.

Tệp đầu vào: mọi loại, thường là PDF/ảnh scan.

Việc bước 1: đổi tên chuẩn `[Số đến]_<Số ký hiệu>_<tên gốc>`, báo danh sách (số CV, ngày, trích yếu, đường dẫn), rồi bàn giao danh sách PDF/ảnh cho bước 2 (OCR). Các bước sau giữ nguyên.

**Vì sao:** user bỏ khâu iOffice; hồ sơ đến qua thư mục và kênh chat, không muốn làm trung gian thủ công.
**Áp dụng:** không chạy `tra-cuu-ioffice-dong-nai` ở bước thu thập; tệp Telegram/Zalo xử lý thẳng, HO SO MOI vẫn là nguồn song song (user chưa bỏ nguồn này).
