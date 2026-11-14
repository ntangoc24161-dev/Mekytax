---
type: workflow
name: Quy trình xử lý thông báo đóng/mở HĐ
slug: quy-trinh-xu-ly-thong-bao
status: 'off'
group: Vận hành
description: Tiếp nhận → Phân công → Gửi L1/L2 → Giải trình → Quyết định (Đóng HĐ / Phiếu chuyển)
steps:
  - agent: phong-van-thu
    task: "Tiếp nhận danh sách thông báo mới từ {{input}}, phân công cán bộ xử lý, cập nhật bảng QL_Thong_bao_Dong_Mo_HD_2026.xlsx"
    verify_agent: phong-van-thu
    max_retries: 1
  
  - agent: theo-doi-tien-do
    task: "Hàng ngày kiểm tra deadline, gửi nhắc nhở Telegram cho cán bộ quá hạn, cập nhật trạng thái bảng. Lần 1: Gửi L1 → Chờ 7 ngày (Deadline L1). Lần 2: Gửi L2 → Chờ 7 ngày (Deadline L2)"
    max_retries: 1
  
  - agent: xu-ly-cuoi
    task: "Sau 7 ngày từ Deadline L2, kiểm tra DN không giải trình → Tạo phiếu chuyển hoặc ghi chép đóng HĐ → Lưu vào exports/ → Cập nhật bảng cột Hành động"
    verify_agent: xu-ly-cuoi
    max_retries: 1

updated: 2026-11-13
---

## Mô tả quy trình

**Giai đoạn 1: Tiếp nhận & Phân công (Agent Phòng Văn Thư)**
- Input: Danh sách thông báo mới (file Excel hoặc Zalo)
- Output: Bảng cập nhật, cán bộ được phân công, ngày L1 ghi nhận

**Giai đoạn 2: Theo dõi & Nhắc nhở (Agent Theo dõi) - Chạy hàng ngày**
- Kiểm tra chưa gửi L1 (quá 0 ngày)
- Kiểm tra sắp hết hạn L1 (quá hạn 3 ngày)
- Kiểm tra chưa nhận GT L1
- Gửi nhắc nhở Telegram → Cánhội xử lý gửi L2
- Lặp lại cho lần 2

**Giai đoạn 3: Xử lý cuối (Agent Xử lý cuối) - Mỗi thứ 2 tuần**
- Kiểm tra DN không GT sau 2 lần gửi (Deadline L2 ≤ 7 ngày)
- Tạo phiếu chuyển (mặc định) hoặc ghi đóng HĐ (nếu có lý do)
- Lưu tài liệu vào `exports/`
- Cập nhật bảng (cột Hành động)

## Lưu ý
- Workflow này để `status: 'off'` cho người dùng xem trước rồi bật
- Giai đoạn 2 chạy tự động hàng ngày (sẽ tạo loop riêng)
- Giai đoạn 3 chạy hàng tuần (thứ 2 - sẽ tạo loop riêng)
