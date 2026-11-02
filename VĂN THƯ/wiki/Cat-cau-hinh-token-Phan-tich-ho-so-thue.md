---
type: wiki
status: active
tags: [wiki, van-hanh, token]
created: 2026-10-23
updated: 2026-10-23
source: usage_index + cấu hình brain Phân tích hồ sơ thuế
---
# Cắt cấu hình token - Phân tích hồ sơ thuế

## Hiện trạng đo được (tháng 10/2026)

Theo chỉ mục `usage_index`:

- Brain **Phân tích hồ sơ thuế** xếp **hạng 1** token tháng này: ~**82.8M** token, ~**656** lượt, quy đổi ~**$31.15**.
- Chủ yếu **background** (551 lượt) + **subagent** (70 lượt ngày 19/10); model nặng **claude-sonnet-5** (330 lượt nền ngày 19/10) và **haiku** (198 lượt nền ngày 23/10).
- Có **6 agent** đúng vai GĐ1-GĐ3, **1 workflow** `phan-tich-ho-so-thue-3-giai-doan` (`status: active`, 7 bước), **0 loop**, **0 reminder**.
- Trước cắt: Kanban **orchestration=auto**; 1 việc `blocked` treo tổng hợp KQ TT_BIEN_HOA; **32 skill** (~260 KB `SKILL.md`), trong đó nhiều skill ngoài phạm vi phòng (ôn thi, iOffice, soạn thảo, webcake...).

Liên quan: [[Cat-cau-hinh-token-Brain-Default]].

## Quyết định đã áp (2026-10-23)

1. **Tắt AI tự vận hành** Kanban: `orchestration: off`.
2. **Huỷ việc blocked** `t_62548b57a135` (Tổng hợp kết quả cuối 6 tệp KQ TT_BIEN_HOA) để khỏi bị dispatcher/worker quấy.
3. **Tắt (không xoá)** 8 skill ngoài phạm vi phân tích hồ sơ:
   - `html-to-webcake`, `phan-tich-cau-hoi-sinh-doi-cheat-sheet`, `tien-trinh-supervisor-tu-dong`
   - `soan-thao-van-ban`, `tra-cuu-ioffice-dong-nai`, `tra-loi-cong-van`
   - `xu-phat-vphc-thue-hoa-don`, `quan-ly-rui-ro-tuan-thu-thue`
4. **Giữ** workflow 3 giai đoạn và toàn bộ skill gắn 6 agent (GĐ1-3, đối chiếu 511/632, NCC, TNCN, kho, giá thành, công nợ, tổ chức hồ sơ...) cộng skill hệ thống/OCR.
5. **Không tắt** workflow: chỉ chạy khi anh chủ động khởi động; orchestration off đã chặn tự chạy nền qua Kanban.

## Mục tiêu vận hành sau cắt

- Phòng chỉ chạy khi có DN cụ thể (folder + MST + kỳ) và anh bật chạy workflow/việc tay.
- GĐ1/chuẩn hóa ưu tiên model nhẹ (Haiku); Sonnet/Opus chỉ khi chốt KQ3 / nhận xét trọng yếu (hướng dẫn dùng, chưa ghim cứng trong file agent).
- Skill ôn thi / công an / soạn thảo nằm ở brain chuyên tương ứng.

## Việc còn mở

- Ghim model theo agent (nếu Javis/engine hỗ trợ ổn định trên bản đang chạy).
- Đo lại token 7 ngày sau cắt để xác nhận giảm nền.
- Cắt tiếp phòng **Hồ sơ kiểm tra** (hạng 3 token, chi phí quy đổi cao vì Opus/Sonnet).
