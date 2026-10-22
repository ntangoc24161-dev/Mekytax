---
type: wiki
title: Playbook lệnh - Đối chiếu 511 và 632
updated: 2026-10-15
---

# Playbook `doi-chieu-511-632`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A5). Tool: `assets/Tool_DC_511_632.xlsm`. Cấu trúc: `references/cau-truc-tool-dc-511-632.md`.

## Cần

Sổ chi tiết TK 511 (PS Có) và TK 632 (PS Nợ), cột A→G, sheet tên `511` / `632`.

## Lệnh

```
Đối chiếu tự động doanh thu - giá vốn theo skill doi-chieu-511-632
cho công ty <TEN> MST <MST> tại <FOLDER> kỳ <KY>.
Xuất Doi_chieu_511_632_<MST>_<KY>.xlsx/.xlsm vào cau-thanh
và nhúng sheet Doi_chieu_511_632 vào Master.
Báo lãi gộp theo năm, dòng P<=0, dòng không khớp 632.
```

**Siêu ngắn:** `Đối chiếu 511-632 <FOLDER>`

## File ra

`Doi_chieu_511_632_<MST>_<KY>.xlsx` tại `cau-thanh/`. Nhúng Master `Doi_chieu_511_632`.

## Thiếu thì

Chưa có sổ 511/632: treo bước này. Vẫn làm đối chiếu GTGT từ tờ khai.
