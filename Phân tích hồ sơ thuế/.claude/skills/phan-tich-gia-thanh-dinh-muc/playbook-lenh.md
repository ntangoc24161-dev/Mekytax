---
type: wiki
title: Playbook lệnh - Phân tích giá thành và định mức
updated: 2026-10-15
---

# Playbook `phan-tich-gia-thanh-dinh-muc`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A7). Playbook trục Bước 7.

## Cần

Sản xuất: sổ 154/621/622/627 + bảng giá thành + định mức + 632 + BCTC [02].
Thương mại: GV 632 + CP BH/QL từ TNDN/BCTC (+ kết quả 511-632).

## Lệnh sản xuất

```
Phân tích giá thành và định mức tại <FOLDER> kỳ <KY>.
Bóc 621/622/627, khớp GV với BCTC [02] và TK 632.
Cập nhật sheet Gia_thanh_va_Dinh_muc.
```

## Lệnh thương mại

```
DN thương mại: lập sheet Gia_thanh_va_Dinh_muc theo cơ cấu
Doanh thu - GV 632 - lãi gộp - chi phí bán hàng/QLDN (từ TNDN),
không bắt buộc 621/622/627.
```

**Siêu ngắn:** `Phân tích giá thành/định mức <FOLDER>` hoặc `Cơ cấu GV thương mại <FOLDER>`

## File ra

Master sheet `Gia_thanh_va_Dinh_muc`.

## Thiếu thì

Ghi hiện trạng, không 0 giả. Không ép DN thương mại phải có 621/622/627.
