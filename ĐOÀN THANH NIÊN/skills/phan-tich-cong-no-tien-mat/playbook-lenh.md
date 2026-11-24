---
type: wiki
title: Playbook lệnh - Phân tích công nợ và tiền mặt
updated: 2026-10-15
---

# Playbook `phan-tich-cong-no-tien-mat`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A8). Playbook trục Bước 8. BT7: `lam-sach-cd1-mst-hoa-don`.

## Cần

Sổ 111/112/131/331. Sao kê NH: Excel tốt hơn PDF.

## Lệnh

```
Phân tích công nợ 131/331 và quét chi tiền mặt TK 111 tại <FOLDER> kỳ <KY>.
Ngưỡng: >= 20 triệu đến 06/2025; >= 5 triệu từ 07/2025 (K5 đã chốt >=).
Quét cụm HĐ mua cùng ngày cùng NCC bằng SUMIFS (CD1 BT7A/BT7B); flag nhóm vượt ngưỡng.
Phân loại mục đích chi; nêu rõ sao kê đang PDF thì chưa kết luận thiếu UNC.
Cập nhật sheet Cong_no_va_Ngan_hang bằng số liệu thật.
```

**Siêu ngắn:** `Quét 111, cụm cùng ngày NCC và công nợ 131/331 <FOLDER>`

## File ra

`Phan_tich_Tien_mat_va_Ngan_hang_<MST>_<KY>.xlsx`. Nhúng Master `Cong_no_va_Ngan_hang`.

## Thiếu thì

Sao kê PDF: ghi PDF, không kết luận thiếu UNC. Bước tiếp: Master 8 sheet trong `phan-tich-ho-so-thue`.
