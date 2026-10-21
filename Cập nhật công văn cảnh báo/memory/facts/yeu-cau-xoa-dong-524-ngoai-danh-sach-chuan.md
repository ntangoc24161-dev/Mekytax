---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-01
updated: 2026-10-01
---
Trong tiến trình làm sạch Sổ theo dõi `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` (sheet `F0`), người dùng chỉ đạo tiếp tục, không chạy lại từ đầu, và xóa dòng 524 nằm ngoài `INPUT/524.xlsx`.

**Vì sao:** HDRR và CV cong an nạp lại cụm 524 thành 4 bản (1770, 1798.0, 1798/TCT-TTKT, trống số CV), lẫn 1 MST không có trong danh sách chuẩn.

**Áp dụng:**
- Chuẩn 524.xlsx: 524 MST theo `1770/CAT-CSKT` ngày 29/03/2023; 113 MST theo `05/BKL-CSKT(Đ4)` ngày 16/08/2023.
- Đã xóa MST `0314123789` trên `1798/TCT-TTKT`; gộp 2.115 dòng trùng thành 637 dòng nguồn `524.xlsx`.
- F0 còn 17.333 dòng; F1 giữ 3.704 dòng. Backup `Sao lưu\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.bak_20261001_011911`.
