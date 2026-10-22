---
type: decision
created: 2026-10-09
updated: 2026-10-09
---
Đã đổi cấu trúc cột Sổ theo dõi `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`: đưa **MST bên bán** và **Tên công ty bán** lên trước khối thông tin công văn cảnh báo.

**Thứ tự đọc mới:** `STT → MSTBan → TenBan → SoCV → NgayCV → DonViBH → ...`

- F0: 19 cột, **17.018** dòng dữ liệu.
- F1: 25 cột, **3.704** dòng; cùng nguyên tắc MST/tên bên bán trước công văn.
- Đã bỏ cột công thức hack `=G` từng chèn ở F0.
- Backup: `Sao lưu/SoTheoDoi_CV_CanhBaoRuiRo_Thue.bak_reorder_mst_20261009_160711.xlsx`
- Script: `skills/theo-doi-cv-canh-bao-rui-ro/scripts/reorder_mst_ten_truoc_cv.py`
- Đã đồng bộ skill/script nạp dữ liệu và 2 bản sao lưu tại `D:\SECOND BRAIN\KIEM SOAT RUI RO\QUAN LY HD RUI RO\` và `D:\Theo doi CVCBRR-ok\`.

**Vì sao:** Tra cứu theo doanh nghiệp/MST tiện hơn khi hai cột này đứng trước số công văn.
**Áp dụng:** Mọi lần nạp mới phải theo layout cột mới trong `update_so_theo_doi.py` và skill `theo-doi-cv-canh-bao-rui-ro`.
