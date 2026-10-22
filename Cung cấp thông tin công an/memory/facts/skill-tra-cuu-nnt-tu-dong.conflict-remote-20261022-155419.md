---
type: decision
created: '2026-09-26'
updated: '2026-10-16'
---
Quy chuẩn kỹ năng tra cứu NNT và phân Tổ Kiểm tra tự động:
- **Tên kỹ năng:** `tra-cuu-nnt-tu-dong`.
- **Cơ chế tra cứu 2 tầng:** Ưu tiên `I:\My Drive\DB NNT.xls` sheet `NNT` (gồm cột `Tên cơ quan thuế`, `Tên Phường Xã`, `Tên Quận/Huyện`); thiếu thì masothue.com / dangkykinhdoanh.gov.vn / tracuunnt.gdt.gov.vn. Không bịa.
- **Cửa kiểm CQ quản lý trước khi phân tổ:**
  - Thuộc TCS3 (`TP.Biên Hòa-TCS3 Đồng Nai` / Thuế cơ sở 3 TP Đồng Nai / Chi cục Thuế KV Biên Hòa - Vĩnh Cửu): mới chia KT1-KT3. Vĩnh Cửu -> KT2. KT1 mới gán CBQL từ sheet `KTR1`.
  - CQ khác: ghi nguyên văn tên CQ vào ô Tổ KT, không gán CBQL.
- **Bảng địa bàn TCS3:** KT1 Long Bình, Hố Nai, Long Hưng, Phước Tân, Tam Phước; KT2 Tam Hiệp, Trảng Dài và toàn bộ Vĩnh Cửu (Trị An, Tân An, Tân Triều, Phú Lý...); KT3 Biên Hòa, Trấn Biên.

**Vì sao:** Sổ theo dõi và dự thảo Word chỉ xử lý DN do TCS3 quản lý.

**Áp dụng:** Mọi lượt tiếp nhận công văn từ CV MOI hoặc Zalo, và khi chạy workflow quản lý công văn công an.
