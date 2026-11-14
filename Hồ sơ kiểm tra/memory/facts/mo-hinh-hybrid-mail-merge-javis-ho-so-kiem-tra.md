---
superseded_by: chot-phuong-an-hybrid-2-tang-va-cap-nhat-3-skill-ho-so-thue
type: decision
provenance: user
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
source: chat 2026-11-12 (ce012a55)
---
Người dùng đã chốt mô hình kết hợp 2 tầng (Hybrid) để tự động hóa bộ hồ sơ kiểm tra quyết toán thuế (chốt ngày 12/11/2026, trao đổi với Javis VPS).

- **Tầng 1 - Word Mail Merge (người dùng chủ động):** sổ `Danh sach kiem tra quyet toan pro.xlsm` thêm sheet đệm phẳng `DuLieuMerge`, chỉ chứa 1 dòng của vụ đang làm. Phôi Word gắn Mail Merge cố định vào sheet này để điền thông tin hành chính (tên NNT, MST, địa chỉ, đoàn kiểm tra, số QĐ...) và các chỉ tiêu NSNN tổng hợp (truy thu, phạt 20%, chậm nộp, giảm lỗ, giảm khấu trừ). Số tiền bằng chữ sinh bằng hàm VBA `=DocSo(...)` ngay trên `DuLieuMerge` (Phương án 2).
- **Tầng 2 - Javis:** điền bảng ma trận nhiều dòng từ sheet DTCHEP vào BBXN và BBKT; đối chiếu số tổng trên `DuLieuMerge` với dòng tổng sheet TINH PHAT + CHAM NOP của Bang tinh.xlsx; chuẩn thể thức NĐ 30 (Tờ trình/QĐKT 1 trang A4, checkbox ☐/☑, khoảng cách chữ ký).
- Có 5 nhóm mẫu Javis phải điền bảng chi tiết trực tiếp: DI-01 (Mẫu 08), BBXN (Mẫu 11), BBKT (Mẫu 14), Báo cáo kết quả (Mẫu 12), QĐ xử phạt VPHC.

**Vì sao:** Mail Merge không tự làm mới khi sổ Excel đổi, dễ vỡ form bảng biểu và không đọc được dữ liệu nhiều năm từ Bảng tính.
**Áp dụng:** khi lập hồ sơ, không dùng Mail Merge cho bảng số liệu nhiều năm. Trước khi xuất, luôn kiểm tra chéo số trên DuLieuMerge với TINH PHAT. Đang chờ: tạo sheet DuLieuMerge trong .xlsm và kiểm tra xem sổ đã có cột bằng chữ dùng công thức VBA DocSo chưa.


## Lịch sử
- [2026-11-13] Bị thay thế bởi [[chot-phuong-an-hybrid-2-tang-va-cap-nhat-3-skill-ho-so-thue]] (user đổi thông tin).
