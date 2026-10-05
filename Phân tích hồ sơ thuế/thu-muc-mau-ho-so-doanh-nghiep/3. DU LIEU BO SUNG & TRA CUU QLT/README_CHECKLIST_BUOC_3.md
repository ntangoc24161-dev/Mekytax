# HƯỚNG DẪN & CHECKLIST BƯỚC 3: DỮ LIỆU BỔ SUNG & TRA CỨU QLT
(Dữ liệu tra cứu ngành Thuế, cảnh báo rủi ro & Hồ sơ giải trình bổ sung)

Thư mục này chứa các cơ sở dữ liệu làm giàu từ hệ thống Quản lý thuế, dữ liệu liên ngành và các hồ sơ NNT nộp bổ sung sau khi phát hiện khoảng trống dữ liệu ở Bước 1 & 2.

---

## 1. Cấu trúc các thư mục con

```text
3. DU LIEU BO SUNG & TRA CUU QLT/
│
├── 3.1. Canh bao rui ro CVCBRR & Cong an/
│   ├── So_CVCBRR_F0_Nguoi_ban_rui_ro.xlsx   (Danh sach DN phat hanh hoa don bat hop phap)
│   ├── So_CVCBRR_F1_Doanh_nghiep_mua.xlsx   (Danh sach DN mua vao dang bi theo doi/dieu tra)
│   └── So_theo_doi_Cong_van_Cong_an.xlsx    (Danh sach vu an, doanh nghiep do Cong an dieu tra)
│
├── 3.2. Danh ba trang thai NNT/
│   └── Danh_ba_trang_thai_DKKD_NCC.xlsx     (Tra cuu trang thai 100% NCC: dang HĐ, dong MST, bo tron...)
│
├── 3.3. Tra cuu TNCN da nguon & Nhan su/
│   ├── Ket_qua_tra_cuu_TNCN_toan_quoc.xlsx  (Tra cuu 100% nguoi lao dong xem co thu nhap noi khac khong)
│   └── Danh_sach_lao_dong_trung_CCCD.xlsx   (Boc tach cac truong hop trung ma so thue/CCCD tren XML)
│
└── 3.4. Giai trinh & Ho so bo sung/
    ├── Phieu_yeu_cau_cung_cap_bo_sung.pdf   (Phieu yeu cau do Doan kiem tra phat hanh)
    ├── Cong_van_giai_trinh_cua_NNT.pdf      (Giai trinh cua NNT kem chung tu chung minh)
    └── Chung_tu_bo_sung_theo_yeu_cau/       (UNC bo sung, hop dong bo sung...)
```

---

## 2. Checklist kiểm tra tính đầy đủ của Bước 3

- [ ] **Đối chiếu Cảnh báo rủi ro & Công an:**
  - [ ] Nạp đầy đủ Sổ CVCBRR F0 (người bán), F1 (người mua) và Sổ Công an mới nhất.
  - [ ] Đã quét 100% Mã số thuế của các nhà cung cấp mua vào trong kỳ.
- [ ] **Trạng thái ĐKKD của Nhà cung cấp:**
  - [ ] Đã tra cứu tình trạng hoạt động của toàn bộ NCC (phát hiện NCC ngừng hoạt động nhưng vẫn xuất HĐ hoặc đóng MST trước ngày xuất HĐ).
- [ ] **Thu nhập nhiều nguồn của Người lao động:**
  - [ ] Đã chạy tra cứu trên CSDL Quản lý thuế ngành Thuế cho 100% nhân sự có tên trên bảng quyết toán.
  - [ ] Xác định số lao động có từ 2 nguồn thu nhập trở lên để kiểm tra điều kiện ủy quyền quyết toán thay.
  - [ ] Rà soát phát hiện các lỗi trùng số CCCD/CMND giữa các cá nhân trên dữ liệu XML.
- [ ] **Giải trình & Bổ sung:**
  - [ ] Đã phát hành Phiếu yêu cầu giải trình đối với các điểm bất thường (khoảng trống hóa đơn, NCC rủi ro F0, mượn tiền cá nhân...).
  - [ ] Đã tiếp nhận và lưu trữ đầy đủ hồ sơ phản hồi của NNT.
