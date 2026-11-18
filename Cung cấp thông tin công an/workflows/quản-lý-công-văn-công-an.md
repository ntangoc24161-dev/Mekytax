---
type: workflow
name: Quản lý công văn công an
slug: quản-lý-công-văn-công-an
status: active
group: Tài chính
description: Tiếp nhận, bóc tách, phân công KT1 in / KT2-KT3 Zalo (Telegram xem trước); dừng sau phân công. Soạn thảo khi có lệnh riêng.
steps:
- agent: khai-thac-du-lieu-ho-so
  task: Tiếp nhận công văn từ thư mục CV MOI; bóc tách file .md lưu D:\HSO CONG AN FILE MD; tra cứu CSDL DB NNT.xls và cổng thông tin; phân Tổ KT; cập nhật Bao_cao_cong_van.xlsx (sheet NHAP_HO_SO); ghi tên CBQL Tổ KT1 nếu công văn có doanh nghiệp thuộc TCS3. Ghi sổ theo nền tảng - Linux/VPS Drive Composio ralph-urbane vá XML; Windows Excel COM.
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
- agent: theo-doi-phan-cong
  task: "Rà soát công văn Mới; tra cứu CBQL. KT1: skill in-phan-cong-kt1 - gửi PDF xem trước qua Telegram, chỉ in khi user chốt, rồi cập nhật cột O sang Đang xử lý. KT2/KT3: skill gui-phan-cong-zalo-kt23 - tin ngắn (tên DN, số CV/ngày, hạn), gửi nháp + file xem trước qua Telegram, chỉ gửi Zalo khi user chốt, rồi cập nhật cột O sang Đang xử lý. Ghi sổ: Linux/VPS = Drive Composio ralph-urbane + vá XML (giữ Dashboard); Windows = Excel COM. Không đổi trạng thái trước khi bàn giao thành công. Sau bước này WORKFLOW DỪNG - không sang soạn thảo."
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
updated: '2026-11-18'
---

## Phạm vi tự động (cron 12h / khi giao CV mới)
Chỉ **2 bước**: bóc tách → phân công (Telegram xem trước → chốt). Xong thì **dừng** và báo chat.
Workflow này chỉ cho **CV Công an mới**. Bản trả lời đã ký trong CV MOI do cron phân nhánh sang `theo-doi-cv-tra-loi-da-ky`.

## Việc làm khi có lệnh riêng hoặc nhánh trưa
- **Soạn thảo** công văn trả lời / tờ trình: gọi agent `soan-thao-tra-loi-cong-an` hoặc bật workflow `soan-thao-quan-ly-ho-so` (vẫn chỉ khi có lệnh).
- **CV trả lời đã ký:** workflow `theo-doi-cv-tra-loi-da-ky` (`Đang xử lý` → `Đã hoàn thành`) - cron 12h tự kích khi file nằm trong CV MOI; ngoài CV MOI thì cần lệnh theo dõi.
- **Lưu trữ / xuất danh sách / hậu tố -R:** agent `quản-lý-hồ-sơ-công-an`.
