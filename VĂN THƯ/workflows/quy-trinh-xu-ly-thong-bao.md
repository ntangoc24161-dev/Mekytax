---
type: workflow
name: Quy trình xử lý thông báo đóng/mở HĐ
slug: quy-trinh-xu-ly-thong-bao
status: 'off'
group: Vận hành
description: Tiếp nhận → Phân công trên Master → Theo dõi L1/L2 → Quyết định Đóng HĐ / Phiếu chuyển
steps:
  - agent: phong-van-thu
    task: "Tiếp nhận danh sách từ {{input}}, phân công cán bộ, cập nhật QL_CONG_TAC_KIEM_TRA_MASTER_2026.xlsx (sheet DANH_SACH_THEO_DOI_CHUNG)"
    verify_agent: phong-van-thu
    max_retries: 1

  - agent: theo-doi-tien-do
    task: "Đọc Master, kiểm tra deadline L1/L2 (7 ngày làm việc), gửi nhắc Telegram, báo cấp đỏ. Nguồn: G:\\Other computers\\My Laptop\\1. CONG TAC KIEM TRA\\QL_CONG_TAC_KIEM_TRA_MASTER_2026.xlsx"
    max_retries: 1

  - agent: xu-ly-cuoi
    task: "Sau hạn L2, lọc DN không giải trình trên Master → cán bộ chốt Đóng HĐ hoặc Phiếu chuyển (mặc định phiếu chuyển) → lưu exports/ → cập nhật cột Hành động"
    verify_agent: xu-ly-cuoi
    max_retries: 1

updated: 2026-11-20
---

## Mô tả quy trình

**Giai đoạn 1: Tiếp nhận & Phân công (Agent Phòng Văn Thư)**
- Input: kế hoạch / danh sách thông báo / thư mục Word
- Output: dòng mới hoặc cập nhật trên file Master

**Giai đoạn 2: Theo dõi & Nhắc nhở (Agent Theo dõi + loop hàng ngày)**
- Nhắc Telegram theo tồn L1/L2
- Cấp đỏ khi quá hạn ≥ 3 ngày làm việc

**Giai đoạn 3: Xử lý cuối (Agent Xử lý cuối + loop thứ 2)**
- DN không GT sau L2 → phiếu chuyển hoặc đóng HĐ
- Cập nhật Master + lưu phiếu ở `exports/`

## Lưu ý
- Workflow để `status: 'off'` cho user xem rồi bật
- File Master chính nằm ở ổ `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\`
- Loop nhắc nhở và xử lý cuối tạo riêng, mặc định tắt đến khi user bật
