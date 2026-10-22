---
type: decision
created: 2026-10-07
updated: 2026-10-07
---
Đã cập nhật nội dung Chuyên đề 1 Excel kiểm tra thuế (CD1, Hội Kế toán Thuế TP.HCM 6/2026 + đáp án tự giải BT5-BT9 trên Du lieu_HV) vào các skill hiện có, không tạo skill mới trùng.

**Áp dụng:**
- Brain + skill chung `phan-tich-ho-so-thue`: thêm `references/excel-cd1-hv-bt5-bt9.md`; bổ sung Bước 3 làm sạch MST/HĐ trùng + Pivot NKC; Bước 4 đối chiếu HĐĐT-BKMV-NKC 133, DN thương mại 1Nợ1Có, SUMIFS cụm cùng ngày ngưỡng 20tr/5tr; bẫy MST/Số HĐ/Excel VN; cập nhật playbook (bước 1b + lệnh siêu ngắn).
- `doi-chieu-gtgt-tndn-cbrr` và `doi-chieu-ncc-mua-vao-rui-ro` (brain): gắn chuẩn hóa MST CD1 và nhắc khóa đối chiếu / cụm cùng ngày.
- Skill chung lưu tại `shared-skills/`; mirror `.claude/skills` đã đồng bộ.

**Vì sao:** Hoàn thiện kỹ thuật Excel dùng lại cho mọi hồ sơ kiểm tra sau này, gắn đúng chỗ trong quy trình 4 bước thay vì skill Excel tách rời.

**Áp dụng:** Khi phân tích hồ sơ mới, chạy làm sạch khóa (MST, HĐ trùng) trước đối chiếu; quét cụm cùng ngày khi rà tiền mặt; dùng reference CD1 khi cần công thức chi tiết.
