---
type: wiki
title: Playbook lệnh - Đối chiếu GTGT TNDN HĐĐT CBRR
updated: 2026-10-15
---

# Playbook `doi-chieu-gtgt-tndn-cbrr`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A2). Playbook trục Bước 2.

## Cần

XML GTGT + TNDN. Bảng kê hoặc HĐĐT nếu có.

## Lệnh

```
Đối chiếu tờ khai GTGT - QTT TNDN - bảng kê - HĐĐT chuẩn CBRR
cho <TEN> MST <MST> tại <FOLDER> kỳ <KY>.
Xuất Tong_hop_doi_chieu_GTGT_TNDN_<MST>.xlsx
(2 sheet chuẩn CBRR nếu đủ dữ liệu).
Bóc tách kỳ BS, HĐ hủy/thay thế, nguyên nhân lệch.
Nếu có HĐĐT mua + BKMV + NKC: đối chiếu thêm khóa SoHD|MST và thuế BKMV vs PS Nợ 133* theo tháng (CD1 BT8/BT9).
```

**Siêu ngắn:** `Đối chiếu CBRR GTGT-TNDN-HĐĐT <FOLDER>`

## File ra

`Tong_hop_doi_chieu_GTGT_TNDN_<MST>.xlsx`. Nhúng Master sheet `Tong_hop_GTGT_TNDN`.

## Thiếu thì

Thiếu XML quý: chạy năm có, liệt kê quý thiếu. Bước tiếp: `doi-chieu-ncc-mua-vao-rui-ro`.
