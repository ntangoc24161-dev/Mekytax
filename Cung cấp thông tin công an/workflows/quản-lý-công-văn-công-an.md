---
type: workflow
name: Quản lý công văn công an
slug: quản-lý-công-văn-công-an
status: active
group: Tài chính
description: "Tiếp nhận CV MOI: bóc tách, ghi sổ, tạo folder+MOVE PDF; phân công KT1 in / KT2-KT3 Zalo qua bot CA; dừng sau phân công."
steps:
- agent: khai-thac-du-lieu-ho-so
  task: "B1 BẮT BUỘC đủ checklist rồi mới báo xong (không dừng chờ chốt chủ giữa chừng). (0) Cửa kiểm: chạy scripts/cua-kiem-dong-bao-cao.py (nhắc bot CA, best-effort đóng Excel, chờ sync ~60s, KHÔNG chờ chat) rồi tự tiếp. VPS thường không tắt được Excel Windows - ghi nhận cảnh báo, KHÔNG coi đó là lý do dừng B1 nếu UPLOAD_UPDATE sau này md5 khớp + còn chart. (1) Tiếp nhận CV MOI; bóc .md (vault ho-so-cv-moi-.../md và/hoặc D:\\HSO CONG AN FILE MD); PDF nhiều CV thì tách dòng riêng. (2) Tra DB NNT + phân Tổ KT; CBQL KT1 chỉ khi thuộc TCS3. (3) Ghi sổ sống fileId 1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph bằng scripts/ghi-so-bao-cao-an-toan.py rồi Composio ralph-urbane UPLOAD_UPDATE (cấm openpyxl.save cả workbook). Ô chưa chắc: ghi [CHƯA XÁC ĐỊNH]/cần đối chiếu rồi VẪN GHI DÒNG - không treo B1 chờ anh chốt. Điểm cần anh quyết: liệt kê NEEDS_INPUT trong báo cáo chat (web:mwlp8ddhcfj9yn08) để xử lý SONG SONG, không chặn verify. (4) CÙNG B1 - tạo thư mục Drive dưới 1. CUNG CAP CONG AN NAM 2026 theo STT. CV SO CV-CQBAN HANH-TO KT XU LY (STT tiếp sau folder max hiện có), MOVE/COPY PDF từ CV MOI vào đúng folder, cập nhật cột B (folder) trên sổ; Excel theo dõi không phải CV thì để lại CV MOI. (5) Số CV mờ: OCR + gửi ảnh trang Số qua scripts/gui-telegram-bot-ca.py để anh đối chiếu SONG SONG; folder tạm dùng số đọc được hoặc 4432X/4433X+MST; KHÔNG fail B1 chỉ vì còn số mờ nếu đã có dòng sổ + folder + PDF đã chuyển. Đầu vào: {{input}}"
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
- agent: theo-doi-phan-cong
  task: "B2 chỉ chạy khi B1 đã pass (đã có dòng sổ + folder + PDF đã chuyển). Trước cập nhật cột O: xác nhận sổ đóng/sync tốt nhất có thể. Rà Mới; KT1 skill in-phan-cong-kt1 - gửi PDF xem trước qua bot CA @Kiem_soat_dieu_phoi_bot (scripts/gui-telegram-bot-ca.py), chỉ in khi user chốt, rồi O=Đang xử lý. KT2/KT3 skill gui-phan-cong-zalo-kt23 - nháp + xem trước cùng bot CA, chỉ Zalo khi user chốt, rồi O=Đang xử lý. Ghi sổ: ghi-so-bao-cao-an-toan.py + UPLOAD_UPDATE. Không đổi O trước bàn giao thành công. Sau bước này WORKFLOW DỪNG - không soạn thảo. Dữ liệu B1: {{prev}}"
  verify_agent: kiem-soat-dieu-phoi-quy-trinh
  max_retries: 2
updated: '2026-11-27'
---

## Phạm vi tự động (cron 12h / khi giao CV mới)
Chỉ **2 bước**: B1 bóc tách + ghi sổ + **tạo folder/MOVE PDF** → B2 phân công (xem trước bot CA `@Kiem_soat_dieu_phoi_bot` → chốt). Xong thì **dừng** và báo chat.
Workflow này chỉ cho **CV Công an mới**. Bản trả lời đã ký trong CV MOI do cron phân nhánh sang `theo-doi-cv-tra-loi-da-ky`.

## Bài học 26/11/2026 (chỉnh 27/11)
Lần 12:07 dừng vì chờ chốt chủ nên chưa ghi sổ → verify fail. Lần 14:38 đã ghi sổ nhưng verify vẫn fail vì Cổng 1 / số mờ, và **không tạo folder/MOVE trong B1**. Từ bản này:

1. **Chờ chốt ≠ điều kiện verify B1.** Điểm cần anh quyết ghi marker + báo NEEDS_INPUT; vẫn ghi dòng và hoàn tất lưu trữ.
2. **Tạo folder + MOVE/COPY PDF + cột B là checklist BẮT BUỘC của B1**, không để sang B2 hay làm tay sau.
3. **Cổng 1 dư / số mờ xử lý song song** (Telegram đối chiếu). Sau UPLOAD_UPDATE md5 khớp + còn chart DASHBOARD thì verify B1 được pass dù còn cảnh báo Cổng 1 trên VPS hoặc còn số đang đối chiếu.

### Tiêu chí verify B1 (pass)
- Có dòng `NHAP_HO_SO` tương ứng CV mới (đủ DN theo quy tắc tách PDF).
- UPLOAD_UPDATE đúng fileId sống; validate/charts ok (md5 báo cáo khớp bản tải lại nếu kiểm).
- Mỗi hồ sơ mới có folder STT đúng quy tắc; PDF đã ra khỏi CV MOI (trừ file không phải CV nhánh A).
- Cột B đã gắn STT folder.
- Phần còn mờ / cần anh chốt: đã gắn marker + đã gửi Telegram đối chiếu (nếu có số mờ) + liệt kê NEEDS_INPUT trong báo cáo.

### Không được fail B1 chỉ vì
- Đang chờ anh chốt giả định nghiệp vụ (đã ghi marker).
- Cổng 1 trên VPS không tắt được Excel Windows (đã chạy cua-kiem + upload md5 khớp).
- Còn 1-vài số CV đang đối chiếu ảnh (đã gửi Telegram, folder tạm đã tạo).

## Cửa kiểm sổ trước bóc tách (chốt 20/11/2026)
1. Nghi / kiểm tra file sổ đang mở.
2. Chạy `python scripts/cua-kiem-dong-bao-cao.py` (nhắc bot CA → best-effort đóng → chờ sync ~60s → tự tiếp, không chờ chat).
3. Ghi đè đúng một file sống Drive fileId `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`.

## Việc làm khi có lệnh riêng
- **Soạn thảo** trả lời / tờ trình: agent `soan-thao-tra-loi-cong-an` hoặc workflow `soan-thao-quan-ly-ho-so`.
- **CV trả lời đã ký:** workflow `theo-doi-cv-tra-loi-da-ky`.
- **Hậu tố -R / xuất danh sách:** agent `quản-lý-hồ-sơ-công-an`.
