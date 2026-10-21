---
type: decision
provenance: source
origin: javis-learned
created: 2026-10-12
updated: 2026-10-12
---
Khi rà iOffice để cập nhật sổ cung cấp hồ sơ Công an:
- Không ghi vào sổ các công văn `DON-NVDTPC` (kiến nghị phòng ngừa tội phạm, chỉ đạo nội bộ).
- Không ghi vào sổ các công văn `DON-KTr1` về xác minh đối tượng truy nã (không phải yêu cầu cung cấp hồ sơ doanh nghiệp theo mẫu thường).

**Vì sao:** Lần rà 01/10-12/10/2026, `18979/DON-NVDTPC` và `19022/DON-KTr1` (theo 9650/PC02-Đ2) lọt tập ANĐT/CSKT/KTr1 nhưng không thuộc luồng cung cấp hồ sơ DN.

**Áp dụng:** Sau khi lọc ANĐT/CSKT/KTr1, đọc trích yếu và số ký hiệu; loại hai nhóm này khỏi sheet `NHAP_HO_SO`. Công văn `YC-CSKT`/`CSKT` gửi thẳng Thuế cơ sở 3 và `DON-KTr1` chỉ đạo cung cấp hồ sơ thì vẫn nhập; thiếu bảng kê MST thì ghi chú thiếu, không bịa MST.
