---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-07
updated: 2026-10-09
---
Trước khi chạy phân tích hồ sơ, tạo **1 file Excel tổng hợp đầu vào riêng** (không nhét vào Master 8 sheet):

1. **HĐĐT CQ thuế:** 2 sheet chi tiết `HDDT_Mua_vao`, `HDDT_Ban_ra`.
2. **Bảng kê do công ty cung cấp:** 2 sheet chi tiết `Bang_ke_Mua_vao`, `Bang_ke_Ban_ra` (không trộn chung sheet với HĐĐT).
3. **NLĐ qua các năm:** 1 sheet `Tong_hop_NLD`; bắt buộc ghi nguồn `05-1` hoặc `05-2`. Nếu nguồn = **05-2** mà công ty **không trích 10%** → cột **Đề xuất kiểm tra cam kết = Có**.

**Khung cột `Tong_hop_NLD` (2026-10-09):** giữ nguyên chỉ tiêu trên tờ khai QTT TNCN / phụ lục. 05-1 và 05-2 khác biểu mẫu → chỉ tiêu khác nhau tách thành cột riêng (cột không thuộc nguồn để trống, không ghi 0 giả). Nhóm cột: định danh & nguồn; chỉ tiêu chung (TNCT, thuế đã khấu trừ); chỉ tiêu riêng 05-1 (HĐLĐ, ủy quyền, giảm trừ, TN tính thuế, thuế phải nộp); chỉ tiêu riêng 05-2 (không cư trú, TNCT/thuế phí BH, miễn hiệp định/dầu khí, trích 10%, cam kết); đối soát & nguồn file. **Sắp xếp bắt buộc:** Năm QT tăng dần → nguồn 05-1 trước 05-2 → Họ tên.

Mẫu: `02_Tong_hop_dau_vao_HDDT_BangKe_NLD_MAU.xlsx` (kèm `Huong_dan`, `Tom_tat_kiem_tra`).

**Vì sao:** Tách lớp bóc tách dữ liệu thô khỏi lớp phân tích/đối chiếu; giữ dấu vết nguồn HĐĐT vs Bảng kê và kiểm soát rủi ro TNCN 05-2 không khấu trừ; không làm mất chỉ tiêu khi gộp 2 biểu mẫu khác nhau.

**Áp dụng:** Playbook Bước 0b trong `phan-tich-ho-so-thue`; xuất file thực vào `cau-thanh/`.
