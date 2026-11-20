---
type: workflow
name: Quản lý công văn công an
slug: quản-lý-công-văn-công-an
status: active
group: Tài chính
description: Tiếp nhận, bóc tách, phân công KT1 in / KT2-KT3 Zalo (xem trước qua bot CA @Kiem_soat_dieu_phoi_bot); dừng sau phân công. Soạn thảo khi có lệnh riêng.
steps:
- agent: khai-thac-du-lieu-ho-so
  task: "BƯỚC 0 BẮT BUỘC trước bóc tách/ghi sổ: kiểm tra Bao_cao_cong_van.xlsx đang mở không. Nếu nghi đang mở: chạy scripts/cua-kiem-dong-bao-cao.py (nhắc bot CA, best-effort đóng Excel, chờ sync ~60s, KHÔNG chờ chủ trả lời), rồi tự tiếp. Tiếp nhận công văn từ thư mục CV MOI; bóc tách file .md lưu D:\\HSO CONG AN FILE MD; tra cứu CSDL DB NNT.xls và cổng thông tin; phân Tổ KT; cập nhật đúng 1 file sống Bao_cao_cong_van.xlsx sheet NHAP_HO_SO (Drive fileId 1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph, không tạo bản dated mới); ghi tên CBQL Tổ KT1 nếu công văn có doanh nghiệp thuộc TCS3. Ghi sổ theo nền tảng - Linux/VPS: scripts/ghi-so-bao-cao-an-toan.py (validate/set-status/set-cells) rồi Drive Composio ralph-urbane UPLOAD_UPDATE; Windows Excel COM."
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
- agent: theo-doi-phan-cong
  task: "Trước mọi lần cập nhật cột O: xác nhận Bao_cao_cong_van.xlsx đã đóng và Drive đã sync. Rà soát công văn Mới; tra cứu CBQL. KT1: skill in-phan-cong-kt1 - gửi PDF xem trước qua bot CA @Kiem_soat_dieu_phoi_bot (scripts/gui-telegram-bot-ca.py), chỉ in khi user chốt, rồi cập nhật cột O sang Đang xử lý. KT2/KT3: skill gui-phan-cong-zalo-kt23 - tin ngắn (tên DN, số CV/ngày, hạn), gửi nháp + file xem trước qua cùng bot CA, chỉ gửi Zalo khi user chốt, rồi cập nhật cột O sang Đang xử lý. Ghi sổ: Linux/VPS = scripts/ghi-so-bao-cao-an-toan.py + Drive Composio ralph-urbane UPLOAD_UPDATE; Windows = Excel COM. Không đổi trạng thái trước khi bàn giao thành công. Sau bước này WORKFLOW DỪNG - không sang soạn thảo."
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
updated: '2026-11-20'
---

## Phạm vi tự động (cron 12h / khi giao CV mới)
Chỉ **2 bước**: bóc tách → phân công (xem trước qua bot CA `@Kiem_soat_dieu_phoi_bot` → chốt). Xong thì **dừng** và báo chat.
Workflow này chỉ cho **CV Công an mới**. Bản trả lời đã ký trong CV MOI do cron phân nhánh sang `theo-doi-cv-tra-loi-da-ky`.

## Cửa kiểm sổ trước bóc tách (chốt 20/11/2026, cập nhật cùng ngày)
Trước khi bóc tách dữ liệu và ghi `Bao_cao_cong_van.xlsx`:

1. Nghi / kiểm tra file sổ đang mở.
2. Chạy `python scripts/cua-kiem-dong-bao-cao.py`:
   - Nhắc qua bot CA `@Kiem_soat_dieu_phoi_bot`
   - Best-effort đóng Excel trên máy đang chạy Javis (`javis_app_close`)
   - **Không chờ** anh trả lời chat; chờ ngắn Drive sync (~60s)
3. Tự tiếp tục bóc tách và ghi đè đúng **một** file sống trên Drive (fileId `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`; không tạo bản dated mới).

VPS Docker thường **không tắt được** cửa sổ Excel trên máy Windows của chủ; bước tự động vẫn chạy tiếp sau nhắc + chờ sync. Bỏ qua cửa này dễ làm Excel Windows báo file hỏng khi Drive sync đè lên bản đang mở.

## Việc làm khi có lệnh riêng hoặc nhánh trưa
- **Soạn thảo** công văn trả lời / tờ trình: gọi agent `soan-thao-tra-loi-cong-an` hoặc bật workflow `soan-thao-quan-ly-ho-so` (vẫn chỉ khi có lệnh).
- **CV trả lời đã ký:** workflow `theo-doi-cv-tra-loi-da-ky` (`Đang xử lý` → `Đã hoàn thành`) - cron 12h tự kích khi file nằm trong CV MOI; ngoài CV MOI thì cần lệnh theo dõi.
- **Lưu trữ / xuất danh sách / hậu tố -R:** agent `quản-lý-hồ-sơ-công-an`.
