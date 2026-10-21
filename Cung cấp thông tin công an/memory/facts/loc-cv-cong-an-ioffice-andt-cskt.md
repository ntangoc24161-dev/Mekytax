---
type: preference
provenance: assistant
origin: javis-learned
created: 2026-10-12
updated: 2026-10-12
---
Khi rà hệ thống VNPT iOffice các công văn cung cấp thông tin Công an chưa có trên sổ theo dõi, không dùng từ khóa trích yếu "công an" vì khớp quá rộng.

Lọc đúng văn bản có số ký hiệu hoặc trích yếu thuộc **ANĐT**, **CSKT** và công văn chuyển **KTr1**.

**Vì sao:** Lần quét từ 01/10/2026 lấy mặc định 10 dòng theo "công an" ra tập không đúng luồng cung cấp hồ sơ.

**Áp dụng:** Mở tìm kiếm nâng cao, nới số bản ghi, lọc ANĐT/CSKT/KTr1, rồi đối chiếu `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`) trước khi tải và bóc tách.
