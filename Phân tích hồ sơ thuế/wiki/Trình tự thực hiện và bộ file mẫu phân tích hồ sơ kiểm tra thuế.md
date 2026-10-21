---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-08
updated: 2026-10-08
source: [[conversations/2026-10-08]]
---
Bộ bảng mẫu chuẩn phục vụ phân tích hồ sơ kiểm tra thuế được chuẩn hóa đánh số thứ tự từ 01 đến 07 theo đúng trình tự thực hiện tại thư mục `MAU_FILE_TONG_HOP` [[conversations/2026-10-08]]. Luồng xử lý một doanh nghiệp gồm 6 bước liên hoàn: (1) `01_CHECKLIST_TONG_HOP_HO_SO_MAU.xlsx` kiểm kê tính đầy đủ của hồ sơ (Bước 0); (2) `02_Tong_hop_dau_vao_HDDT_BangKe_NLD_MAU.xlsx` bóc tách dữ liệu đầu vào gồm HĐĐT, Bảng kê mua vào/bán ra và danh sách NLĐ từ bảng kê 05-1/05-2 (Bước 0b); (3) `03_Lam_sach_CD1_BT5_BT9_MAU.xlsx` làm sạch dữ liệu MST, lọc hóa đơn trùng và cụm chi tiền mặt (Bước 1b); (4) `04_Tong_hop_danh_sach_doi_tac_HDDT_Bang_ke_MAU.xlsx` tổng hợp danh bạ đối tác phục vụ tra cứu trạng thái ĐKKD; (5) `05_Tong_hop_NCC_mua_vao_CV_trang_thai_MAU.xlsx` đối chiếu nhà cung cấp mua vào với danh sách rủi ro CVCBRR F0/F1 và trạng thái ĐKKD; (6) `06_Phan_tich_ho_so_cong_ty_MAU.xlsx` là file Master 8 sheet tổng hợp phân tích toàn diện (mục tiêu) [[conversations/2026-10-08]]. Riêng file `07_TONG_HOP_TIEN_DO_DOAN_KIEM_TRA_MAU.xlsx` được sử dụng riêng ở cấp đoàn kiểm tra để theo dõi tiến độ tổng hợp nhiều doanh nghiệp [[conversations/2026-10-08]].
