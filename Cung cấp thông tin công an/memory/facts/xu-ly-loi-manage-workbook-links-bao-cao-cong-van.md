---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-09-30
updated: 2026-09-30
---
Trên file Bao_cao_cong_van.xlsx (sheet NHAP_HO_SO), lỗi Manage Workbook Links phát sinh do 20 ô tại cột Cán bộ quản lý chứa công thức `=VLOOKUP(..., [1]KTR1!..., 6, 0)` trỏ file ngoài bị thiếu, đồng thời gói OpenXML còn lưu thẻ `<externalReferences>`, quan hệ `externalLink1.xml` và thư mục `xl/externalLinks/`. Cách xử lý chuẩn hóa:
1. Tra cứu MST qua `DB NNT.xls` để điền trực tiếp giá trị text tên cán bộ quản lý vào cột N.
2. Bóc gói zip `.xlsx`, gỡ bỏ triệt để phần tử `<externalReferences>` trong `xl/workbook.xml`, xóa Relationship trong `xl/_rels/workbook.xml.rels`, làm sạch `[Content_Types].xml` và xóa thư mục `xl/externalLinks/`.
