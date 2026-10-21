---
type: wiki
title: Playbook lệnh - Làm sạch MST và hóa đơn CD1
updated: 2026-10-15
---

# Playbook `lam-sach-cd1-mst-hoa-don`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A1b). Công thức: `references/excel-cd1-hv-bt5-bt9.md`.

## Cần

Bảng kê hoặc HĐĐT đã gộp. Tốt hơn có NKC cho BT8/BT9.

## Lệnh

```
Chạy CD1 BT5-BT9 tại <FOLDER> kỳ <KY> và gắn vào bảng biểu (không tạo mẫu 03 tách):
1) BT6 chuẩn hóa MST trên bảng kê + HĐĐT.
2) BT5 đối chiếu HĐĐT ↔ bảng kê: cột trùng khớp / không khớp + thống kê.
3) BT7 cụm cùng ngày cùng NCC (ngưỡng 20tr đến 06/2025; 5tr từ 07/2025).
4) BT8/BT9 thuế theo tháng BKMV vs HĐĐT vs Nợ 133 NKC; khóa SoHD|MST / SoHD.
Lưu Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx vào cau-thanh; nhúng tóm tắt vào Master.
```

**Siêu ngắn:** `Chuẩn hóa MST + quét HĐ trùng <FOLDER>` hoặc `Đối chiếu BKMV vs HĐĐT vs Nợ 133 <FOLDER>`

## File ra

`Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx` tại `cau-thanh/`. Không xuất mẫu 03 tách.

## Thiếu thì

Không có NKC thì vẫn BT5/BT6; BT8/BT9 ghi thiếu sổ. Bước tiếp: `doi-chieu-gtgt-tndn-cbrr`, `doi-chieu-ncc-mua-vao-rui-ro`.
