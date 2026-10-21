---
type: decision
created: 2026-10-01
updated: 2026-10-01
---
Đã làm sạch sổ `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` theo map cột nguồn `Tong hop danh sach moi rui ro (1).xlsb`.

**Vì sao:** Các lần nạp cũ lấy lệch cột (fill-down Số CV, ghép MST/tên mua-bán), nên CBRR trên F0 và TB DO KT1 MOI trên F1 không khớp file nguồn.

**Áp dụng:**
- F0 từ sheet CBRR: A=MST, B=Tên, E=Số CV, F=Ngày CV (C=Ngày CV đến, D=Nơi gửi). Thay 3.351 dòng CBRR cũ bằng 2.939 dòng chuẩn.
- F1 từ sheet TB DO KT1 MOI: B=MST, C=Tên, M=Số CV, N=Ngày CV. Thay 3.649 dòng cũ bằng 3.087 dòng chuẩn; để trống TenBan/MSTBan.
- Giữ nguyên HDRR, 524/Công an, F1 nguồn 2026.09.16. F0 còn 19.638 dòng, F1 còn 7.282 dòng.
- Backup: `00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\Sao lưu\SoTheoDoi_CV_CanhBaoRuiRo_Thue.bak_20261001_010038.xlsx`.
