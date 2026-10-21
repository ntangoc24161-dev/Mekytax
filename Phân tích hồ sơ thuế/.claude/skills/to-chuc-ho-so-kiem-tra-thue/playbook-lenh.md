---
type: wiki
title: Playbook lệnh - Tổ chức hồ sơ kiểm tra thuế
updated: 2026-10-15
---

# Playbook `to-chuc-ho-so-kiem-tra-thue`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A0). Cây thư mục: `references/cay-thu-muc-4-buoc.md`.

## Cần

Đường dẫn tuyệt đối. Suy được MST + tên ngắn ASCII thì làm luôn.

## Lệnh

```
Tổ chức hồ sơ 4 bước tại <FOLDER>
MST <MST>, tên <TEN>, kỳ <KY> theo skill to-chuc-ho-so-kiem-tra-thue.
Tạo thư mục <TEN>_<MST>, phân loại file, xuất 1 checklist A4
tại 4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx.
Xóa nguồn cũ sau khi đối soát. Không phân tích số liệu.
Nêu module phân tích nào chạy được ngay.
```

**Siêu ngắn:** `Tổ chức / kiểm kê checklist <FOLDER>`

## File ra

Folder `<TEN>_<MST>` + `4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx`.

## Thiếu thì

Chưa có folder tuyệt đối thì dừng. Bước tiếp: `tong-hop-dau-vao-hddt-bang-ke-nld`.
