---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-14
updated: 2026-10-14
---
Khi tra cứu NNT để nhập sổ cung cấp Công an:
- Trước khi gán Tổ KT, bắt buộc đọc **cơ quan thuế quản lý** (có trong `DB NNT` thì thuộc TCS3; không có thì lấy trường Quản lý bởi trên masothue.com).
- NNT do Thuế thành phố hoặc cơ quan thuế khác quản lý: loại khỏi danh sách TCS3, không gán KT1/KT2/KT3, không gán CBQL, không soạn dự thảo.
- MST không có trên sheet `KTR1` thì để trống cột cán bộ xử lý, không suy tên.

**Vì sao:** Phường trùng địa bàn KT1 (ví dụ Long Bình / KCN Amata) vẫn có NNT do Thuế thành phố quản lý; skill cũ gán KT1 theo phường rồi bịa CBQL khi không khớp `KTR1`.

**Áp dụng:** Mọi lượt tra cứu MST / phân công Tổ KT trên sổ `Bao_cao_cong_van.xlsx` và khi lập dự thảo trả lời.
