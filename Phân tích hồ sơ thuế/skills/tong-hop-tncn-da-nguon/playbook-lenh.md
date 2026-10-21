---
type: wiki
title: Playbook lệnh - Tổng hợp TNCN đa nguồn
updated: 2026-10-15
---

# Playbook `tong-hop-tncn-da-nguon`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A4). Chính sách: skill chung `thue-tncn`.

## Cần

XML QTT TNCN tại `1.1`. Tốt hơn: 05/KK quý, phụ lục NLĐ, file đa nguồn `3.3`.

## Lệnh

```
Tổng hợp QTT TNCN và danh sách NLĐ MST <MST> kỳ <KY>
từ XML tại <FOLDER>\1. DU LIEU THO NGANH THUE\1.1. To khai XML
xuất vào 1.4 và cau-thanh.
Nếu có file tra cứu đa nguồn ở 3.3 thì đối chiếu TC22/TC23, trùng CCCD.
Xuất Tong_hop_TNCN_va_Da_nguon_<MST>_<KY>.xlsx.
```

**Siêu ngắn:** `Tổng hợp TNCN <MST> <FOLDER>`

## File ra

`Tong_hop_TNCN_va_Da_nguon_<MST>_<KY>.xlsx` + danh sách NLĐ tại `1.4`. Nhúng Master `Tong_hop_TNCN_va_Da_nguon`.

## Thiếu thì

Thiếu `3.3`: vẫn lập QTT, ghi thiếu đa nguồn. Không điền 0 giả cho quý thiếu XML.
