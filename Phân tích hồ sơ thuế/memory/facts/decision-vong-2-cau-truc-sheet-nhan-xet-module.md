---
superseded_by: pref-giu-hd04-top-5-ncc-va-mo-rong-nhan-xet-tat-ca-sheet-kq2
type: decision
provenance: user
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
---
Ngày 19/10/2026 (vòng 2, sau khi cán bộ xem bản chạy thật TT_BIEN_HOA MST 3603883505), chốt thêm các thay đổi cấu trúc trên bộ mẫu v1.4 (`thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/`), đã sửa trực tiếp vào cả mẫu .xlsx và skill .md:

- `KQ3_Ket_luan_danh_gia.xlsx` thêm sheet `Danh_gia_28_tieu_chi_rui_ro` (narrative, 7 cột, lấy lại từ mẫu 06 cũ) bên cạnh sheet `KQ3` (80 quy tắc, 39 cột) - không thay cho nhau, là 2 lớp đọc khác nhau.
- `KQ2_511_632` (trong KQ2.2) đổi từ liệt kê từng chứng từ sang bảng tổng hợp theo năm tài chính + cột **Đánh giá tuân thủ & Rủi ro** (lấy từ `BaoCao` của skill `doi-chieu-511-632`).
- `KQ2_NCC_mua_vao` (trong KQ2.1) đổi từ bảng tổng hợp theo NCC (gây lẫn 2 bảng không phân cách ở dòng 147/148 TT_BIEN_HOA) sang MỘT bảng chi tiết 100% hóa đơn mua vào, là bảng khai triển của `IN07_NCC_CanhBao`.
- Ngoại lệ 3.4 (sửa trong `skills/phan-tich-ho-so-thue/SKILL.md`): mọi sheet KQ2 được phép có cột/khối nhận xét module ngắn (`KQ2_511_632`, `KQ2_Doi_chieu_GTGT`, `KQ2_NCC_mua_vao`, `KQ2_TNCN_LD2`, `KQ2_Kho_Gia_thanh`, `KQ2_Gia_thanh_va_Dinh_muc`, `KQ2_Cong_no_Tien_mat`). Mức rủi ro / điểm / đề cương toàn hồ sơ vẫn chỉ ở KQ3.
- HD04 (tỷ trọng Top-5 NCC / P14) bắt buộc giữ: tính bằng gộp theo MST người bán, không dùng LARGE trên dòng hóa đơn.
- Thêm Q14-Q19 vào `phan-tich-ho-so-thue/SKILL.md`: ngày dd/mm/yyyy không giờ, số kiểu VN, font Unicode tự sửa mã lỗi, auto-width cột, xóa dòng ví dụ khi chạy thật, xóa dòng/kỳ không phát sinh.

**Vì sao:** cán bộ xem bản chạy thật phát hiện nhiều lệch giữa quy trình/skill đã viết và kết quả thực tế (bug thực thi, không phải lỗi mẫu) cộng với 3 điểm mẫu thật sự thiếu cấu trúc (thiếu sheet 28 tiêu chí, KQ2_511_632 và KQ2_NCC_mua_vao chưa đúng ý).

**Áp dụng:** mọi hồ sơ chạy sau 19/10/2026 vòng 2 dùng cấu trúc mẫu mới này. Hồ sơ đã chạy trước đó (TT_BIEN_HOA) cần chạy lại GĐ2/GĐ3 để nhận cấu trúc mới - KHÔNG tự coi bản cũ là đạt.

**Đã chốt thêm 19/10/2026 vòng 2b:** giữ HD04; mở rộng nhận xét sang TNCN/kho/giá thành/công nợ; chạy lại 6 tệp TT_BIEN_HOA theo form mới (v2).

**Nguồn:** hội thoại 19/10/2026 sau khi xem TT_BIEN_HOA; chi tiết ở `skills/phan-tich-ho-so-thue/SKILL.md`, `skills/giai-doan-1-kiem-ke-danh-gia/SKILL.md`, `skills/giai-doan-2-phan-tich-du-lieu-co-quan-thue/SKILL.md`, `skills/giai-doan-3-hoan-chinh-ho-so/SKILL.md`, `skills/doi-chieu-511-632/SKILL.md`, `skills/doi-chieu-ncc-mua-vao-rui-ro/SKILL.md`. Xem thêm [[decision-quy-trinh-1-4-sau-tep-p7]].


## Lịch sử
- [2026-10-20] Bị thay thế bởi [[pref-giu-hd04-top-5-ncc-va-mo-rong-nhan-xet-tat-ca-sheet-kq2]] (user đổi thông tin).
