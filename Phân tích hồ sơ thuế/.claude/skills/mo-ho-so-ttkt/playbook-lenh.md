---
type: wiki
title: Playbook lệnh - Mở hồ sơ TTKT
updated: 2026-10-15
---

# Playbook `mo-ho-so-ttkt`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (B1). Checklist: `wiki/Checklist-mo-ho-so-TTKT.md`.

Không sắp folder (`to-chuc-ho-so-kiem-tra-thue`). Không chạy engine XML / Master 8 sheet (`phan-tich-ho-so-thue`).

## Cần

Kỳ kiểm tra, loại NNT, đường dẫn hồ sơ hoặc sheet tóm tắt đã chạy phân tích.

## Lệnh

```
Mở hồ sơ TTKT cho <TEN> MST <MST> kỳ <KY> tại <FOLDER>
theo skill mo-ho-so-ttkt, checklist A→F.
Đối chiếu BCTC - tờ khai - HĐĐT, GDLK nếu có.
Phân nhánh thủ tục / thiếu thuế / trốn thuế.
Liệt kê 5-10 vấn đề Cao/Trung/Thấp và hồ sơ còn thiếu.
Không sắp thư mục, không xuất Master 8 sheet.
```

**Siêu ngắn:** `Checklist mở hồ sơ TTKT A-F <FOLDER>`

## File ra

Bảng lệch + danh mục vấn đề + hồ sơ thiếu (Markdown hoặc Excel tại `cau-thanh/` nếu user yêu cầu xuất file).

## Thiếu thì

Mục nào thiếu chứng từ: ghi thiếu, đề xuất bước lấy thêm. Không bịa số.
