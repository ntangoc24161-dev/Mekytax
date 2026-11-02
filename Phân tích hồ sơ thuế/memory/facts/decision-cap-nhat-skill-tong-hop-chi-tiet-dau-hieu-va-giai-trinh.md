---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-23
updated: 2026-10-23
---

# Đã cập nhật chuẩn tổng hợp dấu hiệu bất thường và chi tiết giải trình NNT vào bộ Skill

User yêu cầu cập nhật toàn bộ quy trình và kết quả làm việc đợt kiểm tra thuế (tổng hợp 9 dấu hiệu bất thường, phân tích EBITDA, chi tiết bán dưới giá, bộ file giải trình NNT 10 sheet và quy chuẩn định dạng Excel) thành skill dùng chung và dùng lại cho các hồ sơ sau.

**Vì sao:**
Tránh tình trạng các lần phân tích hồ sơ tiếp theo bị lặp lại lỗi (quên phân tích EBITDA, không bóc tách chi tiết bán dưới giá, xuất file thô thiếu định dạng hoặc để file mẫu trống cho NNT). Cần quy chuẩn hóa thành năng lực thường trực của quy trình phân tích hồ sơ thuế 1.4.

**Áp dụng:**
1. **Skill `giai-doan-3-hoan-chinh-ho-so`:**
   - Mở rộng đầu ra: bên cạnh `KQ3_Ket_luan_danh_gia` (2 sheet), chuẩn hóa việc xuất 2 file:
     + `KQ3_Danh_gia_Dau_hieu_Bat_thuong_<MST>_<KY>.xlsx` (tổng quan 9 dấu hiệu, hướng kiểm tra chi tiết, hồ sơ cần xem).
     + `KQ3_NNT_Giai_trinh_Chi_tiet_<MST>_<KY>.xlsx` (10 sheet: gồm sheet `00_Tong_hop_chi_tiet` tổng hợp chung và 9 sheet chuyên đề chi tiết với số liệu thật bóc tách từ KQ2/KQ3 để đối chiếu trực tiếp).
   - Bắt buộc áp dụng bộ quy chuẩn định dạng Q14-Q20 (phông Times New Roman, tiêu đề Navy `#1F4E78`, viền mỏng, showGridLines=True, ngày dd/mm/yyyy không giờ, số #,##0, tự giãn cột, tô màu cảnh báo đỏ/vàng/xanh theo mức độ rủi ro).

2. **Skill `giai-doan-2-phan-tich-du-lieu-co-quan-thue`:**
   - Bổ sung sheet `KQ2_EBITDA_Phan_tich` vào `KQ2.1`: tính EBITDA 4 thành phần theo Nghị định 132/2020/NĐ-CP; khi không vay nợ thì kết luận an toàn không rủi ro K9.
   - Bổ sung sheet `KQ2_511_632_Chi_tiet` vào `KQ2.2`: bóc tách chi tiết từng dòng, quét phát hiện hóa đơn bán dưới giá vốn `511_DT < 632_GV` và tính số tiền lỗ.
   - Chuẩn hóa đọc `IN04_BCTC_TNDN` theo dòng các năm (KQKD 50, CDKT 141, TNDN ct16); quy tắc đo lớp chữ PDF sao kê IN11 (text thì trích NH+TK, scan 0 text thì ghi "PDF - chưa kết luận" theo Q7).

3. **Skill trục `phan-tich-ho-so-thue`:**
   - Cập nhật mục 2 (Cấu trúc bộ tệp), mục 4.6 (Module ghi vào đâu), bổ sung quy tắc Q20 và cập nhật nhật ký Lịch sử phiên bản.
   - Đồng bộ toàn bộ cập nhật sang kho skill chung `shared-skills/` và `.claude/skills/`.
