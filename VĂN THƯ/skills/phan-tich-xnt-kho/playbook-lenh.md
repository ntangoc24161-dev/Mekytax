---
type: wiki
title: Playbook lệnh - Phân tích XNT kho
updated: 2026-10-15
---

# Playbook `phan-tich-xnt-kho`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A6). Playbook trục Bước 6. TC3: `skills/phan-tich-ho-so-thue/references/tieu-chi-28.md`.

## Cần

Sổ XNT 152/155 (sản xuất) hoặc 156 (thương mại) + BCTC.

## Lệnh

```
Phân tích XNT kho tại <FOLDER> kỳ <KY>.
DN thương mại: TK 156. DN sản xuất: 152+155 (+156 nếu có).
Đối chiếu tồn CK với BCTC: DNN dùng ct140; đầy đủ dùng [141] = 152+155+156.
Đối chiếu xuất kho với GV 632.
Xuất/cập nhật sheet Tong_hop_XNT_va_Kho, không để giá trị 0 nếu đã có sổ.
```

**Siêu ngắn:** `Phân tích XNT kho <FOLDER>`

## File ra

Sheet Master `Tong_hop_XNT_va_Kho` hoặc file `Phan_tich_XNT_va_Gia_thanh_<MST>_<KY>.xlsx`.

## Thiếu thì

Ghi thiếu nguồn. Không kết luận HTK=0 vì ct141=0 trên BCTC DNN. Bước tiếp: `phan-tich-gia-thanh-dinh-muc`.
