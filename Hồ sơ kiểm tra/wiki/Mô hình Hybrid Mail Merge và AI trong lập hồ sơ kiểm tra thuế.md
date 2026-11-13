---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
source: [[conversations/2026-11-12]]
---
## Ý chính
- Word Mail Merge không tự làm mới khi sửa sổ Excel nguồn, dễ vỡ form bảng biểu và không đọc được dữ liệu dạng ma trận nhiều năm [[conversations/2026-11-12]].
- Phương án đã chốt là phân 2 tầng (thực tế tính đến 12/11/2026): Mail Merge lo trường định danh hành chính và số liệu đơn lẻ, AI lo bảng chi tiết nhiều dòng, kiểm tra chéo và thể thức [[conversations/2026-11-12]].

## Tầng 1 - Mail Merge
- Nguồn cấp dữ liệu đã được tinh gọn kết nối trực tiếp từ sheet `01_DA_BAN_HANH` thông qua Dòng 1 ẩn chứa 58 mã trường chuẩn hóa, xóa bỏ hoàn toàn sheet đệm trung gian `DuLieuMerge` [[conversations/2026-11-13]].
- Trường gồm thông tin NNT, đoàn kiểm tra, số và ngày quyết định, các chỉ tiêu NSNN tổng hợp (truy thu, phạt 20%, chậm nộp, giảm lỗ, giảm khấu trừ) [[conversations/2026-11-12]].
- 6 cột số tiền bằng chữ (`Tong_truy_thu_chu`, `Tong_cham_nop_chu`, `Tong_tien_phat_chu`, `Tong_phat_sinh_chu`, `Giam_lo_chu`, `Giam_khau_tru_chu`) được sinh trực tiếp ở cuối bảng `01_DA_BAN_HANH` qua hàm `=DocSo(...)` [[conversations/2026-11-13]].

## Tầng 2 - AI
- Điền bảng chi tiết từ sheet DTCHEP vào BBXN và BBKT [[conversations/2026-11-12]].
- Đối chiếu số tổng trên sheet đệm với dòng tổng TINH PHAT + CHAM NOP [[conversations/2026-11-12]].
- Chuẩn hóa thể thức theo NĐ 30/2020/NĐ-CP [[conversations/2026-11-12]].
- 5 nhóm mẫu cần điền bảng chi tiết: DI-01 (Mẫu 08), BBXN (Mẫu 11), BBKT (Mẫu 14), Báo cáo kết quả (Mẫu 12), QĐ xử phạt VPHC [[conversations/2026-11-12]].

Liên quan: [[Quy trình hoàn thiện hồ sơ kiểm tra thuế từ Bảng tính]], [[Kỹ thuật điền hồ sơ Word kiểm tra thuế từ phôi tổng hợp]].
