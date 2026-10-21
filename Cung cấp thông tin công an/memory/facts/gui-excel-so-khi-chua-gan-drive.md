---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
---
Khi Javis trên VPS chưa đọc được `/mnt/i` và không có chữ `I:`, người dùng gửi file `Bao_cao_cong_van.xlsx` vào khung chat để rà CQT và phân công TCS3.

Bản đính kèm là bản làm việc của lượt đó. Kết quả sửa lưu trong vault (ví dụ `ho-so-cv-moi-2026-10/Bao_cao_cong_van_ra-soat-761-810.xlsx`). Không khẳng định đã ghi sổ chính trên Google Drive.

**Vì sao:** 15/10/2026 rclone chưa có token OAuth; người dùng gửi Excel để rà dòng 761-810 ngay.

**Áp dụng:** Có file đính kèm thì đọc file đó, không dừng vì hết ổ `I:`. Người dùng copy bản vault đè lên sổ Drive khi sẵn sàng.
