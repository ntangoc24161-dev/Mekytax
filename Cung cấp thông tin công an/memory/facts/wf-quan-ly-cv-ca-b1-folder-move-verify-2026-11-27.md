---
type: decision
created: 2026-11-27
updated: 2026-11-27
source: chat 2026-11-26 (mwlp8ddh)
---
Sau lỗi workflow trưa 26/11 (B1 fail vì chờ chốt / Cổng 1 / số mờ; chưa tạo folder+MOVE trong B1), đã chỉnh `workflows/quản-lý-công-văn-công-an.md` cùng agent `khai-thac-du-lieu-ho-so` và `kiem-soat-dieu-phoi-quy-trinh`.

**Vì sao:** Verify B1 đang trộn “chờ chủ” với “đã ghi sổ”, và lưu trữ Drive không nằm trong checklist B1 nên PDF nằm mãi ở CV MOI.

**Áp dụng:** (1) Chờ chốt ≠ điều kiện fail B1 - ghi marker + NEEDS_INPUT. (2) Tạo folder + MOVE/COPY PDF + cột B bắt buộc trong B1. (3) Cổng 1 VPS / số mờ xử lý song song; pass khi md5 sổ khớp, charts còn, folder+PDF xong.
