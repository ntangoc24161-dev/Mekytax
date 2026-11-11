---
type: workflow
name: Bóc tách & Phân công công văn Công an
slug: boc-tach-phan-cong-cong-van
status: 'off'
group: Vận hành
description: Bóc tách danh sách doanh nghiệp từ CV Công an, tra cứu DB NNT, phân công cho Tổ KT tương ứng.
steps:
  - agent: khai-thac-du-lieu-ho-so
    task: "Bóc tách công văn Công an: trích danh sách doanh nghiệp, MST, yêu cầu cung cấp tài liệu. Tra cứu DB NNT.xls để xác định tình trạng, cơ quan quản lý. Đầu vào: {{input}}"
    verify_agent: Kiểm soát & Điều phối quy trình
    max_retries: 1
  - agent: Theo dõi & Phân công công văn
    task: "Phân công doanh nghiệp cho Tổ KT1/KT2/KT3 phù hợp dựa trên kết quả tra cứu. Xác định CBQL KT1. Cập nhật danh sách phân công. Dữ liệu từ bước trước: {{prev}}"
    verify_agent: Kiểm soát & Điều phối quy trình
    max_retries: 1
updated: '2026-11-08'
---

## Mục tiêu
Khi có công văn Công an mới, tự động bóc tách danh sách doanh nghiệp, tra cứu thông tin pháp lý, phân công cho các Tổ kiểm tra (KT1/KT2/KT3) và gán cán bộ quản lý tương ứng. Mỗi bước đều được `Kiểm soát & Điều phối quy trình` kiểm tra chất lượng trước khi chuyển sang bước tiếp theo.

## Trigger
- **Loop hàng ngày** lúc 12:00 giờ Việt Nam (quy trình tự động)
- **Hoặc** bạn giao CV thêm qua Telegram/Zalo, tôi sẽ trigger workflow này ngay

## Đầu ra
- Danh sách doanh nghiệp được phân công (tên DN, MST, Tổ KT, CBQL)
- File phân công cập nhật lưu trên Drive
- Báo cáo về chat này để xác nhận

## Ghi chú
Sau khi bước phân công xong, workflow **tạm dừng**. Chờ yêu cầu tiếp theo (soạn thảo công văn trả lời) từ bạn hoặc Công an, khi đó sẽ trigger Workflow "Soạn thảo & Quản lý hồ sơ".
