---
type: decision
created: 2026-10-16
updated: 2026-10-16
---
Khi phân chia Tổ KT trên sổ `Bao_cao_cong_van.xlsx` (ô Tổ KT):

1. **Tra cứu cơ quan thuế quản lý trước.** Nguồn: cột `Tên cơ quan thuế` trên sheet `NNT` của `I:\My Drive\DB NNT.xls` (ví dụ `TP.Biên Hòa-TCS3 Đồng Nai`); thiếu thì trường Cơ quan thuế quản lý trên masothue.com.
2. **Chỉ phân KT1/KT2/KT3 và gán CBQL** khi CQ là Thuế cơ sở 3 thành phố Đồng Nai (Biên Hòa - Vĩnh Cửu): `TP.Biên Hòa-TCS3 Đồng Nai`, `TCS3`, `Thuế cơ sở 3 thành phố Đồng Nai`, hoặc tên cũ Chi cục Thuế khu vực Biên Hòa - Vĩnh Cửu.
3. **Vĩnh Cửu** (kể cả `Huyện Vĩnh Cửu(Hết HL)`) thuộc TCS3 thì gán **KT2**.
4. **CQ khác:** ghi nguyên văn tên cơ quan thuế quản lý vào ô Tổ KT; để trống CBQL; không đưa vào dự thảo trả lời.

**Vì sao:** Địa chỉ phường không đủ để kết luận DN thuộc TCS3. Phân tổ khi CQ khác quản lý làm sai sổ và soạn thảo nhầm DN ngoài địa bàn.

**Áp dụng:** Skill `tra-cuu-nnt-tu-dong`, `theo-doi-cung-cap-ca`, agent Tiếp nhận & Điều phối công văn.
