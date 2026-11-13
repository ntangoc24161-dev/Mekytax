---
type: loop
name: Nhắc nhở thông báo hàng ngày
slug: nham-nho-thong-bao-hang-ngay
enabled: false
mode: full
goal: custom
interval_min: 1440
owner_chat: ""
updated: 2026-11-13
---

## Mục tiêu
Hàng ngày lúc 7h sáng, kiểm tra bảng quản lý thông báo (`QL_Thong_bao_Dong_Mo_HD_2026.xlsx`) và gửi nhắc nhở Telegram cho cán bộ về các deadline sắp đến hoặc quá hạn.

## Quy trình mỗi vòng (mỗi 24 giờ)

1. **Mở bảng**: `QL_Thong_bao_Dong_Mo_HD_2026.xlsx`, sheet "Quản lý thông báo"
2. **Phân tích 4 danh sách**:
   - **Chưa gửi L1**: Ngày gửi L1 = trống, Deadline L1 < hôm nay
   - **Sắp hết hạn L1**: Ngày gửi L1 ≠ trống, Deadline L1 sắp tới (≤ 3 ngày)
   - **Quá hạn L1**: Deadline L1 < hôm nay - 3 ngày (RED)
   - **Chưa gửi L2**: Giải trình L1 = "Không", Ngày gửi L2 = trống
3. **Nhóm theo cán bộ**: Tổng hợp số DN chưa làm của từng cán bộ
4. **Gửi Telegram**:
   - Nếu có chưa gửi: `⏰ [Cán bộ A]: Còn 3 DN chưa gửi L1. MST: 360314..., 360365..., 360385...`
   - Nếu có sắp hết hạn: `⏱️ [Cán bộ B]: Còn 2 ngày hết hạn cho 5 DN`
   - Nếu có quá hạn 3+ ngày: `🔴 [Cán bộ C]: ĐỎ - 2 DN quá hạn 3+ ngày! Báo Trưởng phòng ngay`
5. **Ghi log**: Lưu vào file `Javis/loops/logs/nham-nho-[YYYY-MM-DD].log` (ai được nhắn, bao nhiêu DN, kết quả)

## Điểm chú ý
- Chỉ gửi Telegram nếu có DN cần nhắc nhở (không gửi nếu toàn bộ đúng deadline)
- Nếu Telegram gửi thất bại → Ghi log + đề xuất liên hệ khác (email, gọi)
- Quá hạn 3+ ngày là ưu tiên cao (RED) - gửi ngay lập tức, đừng chờ
- Báo cáo phải có con số cụ thể (bao nhiêu DN, bao nhiêu cán bộ)
