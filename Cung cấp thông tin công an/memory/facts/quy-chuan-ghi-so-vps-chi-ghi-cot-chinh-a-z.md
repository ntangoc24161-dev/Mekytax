---
type: decision
created: 2026-11-20
updated: 2026-11-20
source: chat 2026-11-20 (38b144d6)
---
Khi cập nhật sổ `Bao_cao_cong_van.xlsx` trên Google Drive, VPS chỉ ghi dữ liệu thô vào các cột chính từ A đến Z trên sheet `NHAP_HO_SO`. Tuyệt đối không tính toán, không ghi đè dải công thức phụ từ cột AE đến AZ, không can thiệp sheet DASHBOARD hay nén lại biểu đồ. Việc kéo công thức phụ và làm mới Dashboard do phần mềm Microsoft Excel trên máy Windows phụ trách khi mở file/máy.

**Vì sao:** Ngăn chặn triệt để nguy cơ làm lệch công thức XML, mất biểu đồ DASHBOARD và vỡ cấu trúc ZIP của Microsoft Excel khi ghi từ môi trường Linux VPS.

**Áp dụng:** Áp dụng cho mọi tác vụ bóc tách, cập nhật số liệu công văn mới của các agent và skill thao tác trên sổ `Bao_cao_cong_van.xlsx`.
