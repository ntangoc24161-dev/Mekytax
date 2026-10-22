---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
---
Người dùng quy định rõ 4 nguồn dữ liệu độc lập khi phân tích đối chiếu hồ sơ thuế:
1. HĐĐT: Dữ liệu tổng hợp hóa đơn lấy từ hệ thống của cơ quan thuế.
2. Bảng kê mua vào/bán ra: Do doanh nghiệp cung cấp, kê khai theo kỳ.
3. Tờ khai thuế (GTGT, TNDN): Do doanh nghiệp kê khai nộp trên hệ thống cơ quan thuế (chỉ tiêu [23], [24], [34], [35]...).
4. Sổ kế toán (NKC, Sổ cái các TK 133, 511, 632...): Dữ liệu kế toán nội bộ do doanh nghiệp cung cấp.
Khi thiết kế bảng đối chiếu (như PHẦN II bán ra, PHẦN III mua vào trên sheet Tong_hop_GTGT_TNDN), tuyệt đối không gộp chung các nguồn (không gộp HĐĐT với Bảng kê chung cột) mà phải tách riêng từng cột nguồn để tính chênh lệch bắt sai sót giữa các bên.
