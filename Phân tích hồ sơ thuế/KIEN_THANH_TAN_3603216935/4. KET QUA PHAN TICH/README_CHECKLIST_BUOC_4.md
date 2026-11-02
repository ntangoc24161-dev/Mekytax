# HƯỚNG DẪN & CHECKLIST BƯỚC 4: KẾT QUẢ PHÂN TÍCH
(Sản phẩm phân tích chuyên sâu, Master Excel 7 Sheet và Đề cương kiểm tra)

Thư mục này chứa toàn bộ sản phẩm kết xuất sau khi hệ thống tự động chạy phân tích đa chiều và tích hợp dữ liệu từ Bước 1, Bước 2 và Bước 3.

---

## 1. Cấu trúc các thư mục con và tệp tin chuẩn

```text
4. KET QUA PHAN TICH/
│
├── Phan_tich_ho_so_cong_ty_<TenDN>.xlsx   --> FILE MASTER EXCEL 7 SHEET DUY NHẤT
│   ├── Sheet 1: Danh_gia_28_tieu_chi_rui_ro (Cham diem theo QD 1109/QD-TCT & Tool 4.4)
│   ├── Sheet 2: Tong_hop_GTGT_TNDN (Doi chieu 3 chieu GTGT 40 quy vs TNDN & HDDT)
│   ├── Sheet 3: Tong_hop_NCC_CVCBRR (Doi chieu 100% NCC mua vao voi So F0, F1, Cong an, DKKD)
│   ├── Sheet 4: Tong_hop_TNCN_va_Da_nguon (So lieu TNCN 11 nam, 25+ ban ghi tra cuu da nguon)
│   ├── Sheet 5: Tong_hop_XNT_va_Kho (Doi chieu XNT 152/155/156 voi BCTC [141] - Giai toa TC3)
│   ├── Sheet 6: Gia_thanh_va_Dinh_muc (Boc tach 621/622/627, Z SP, doi chieu BCTC [02] - TC28)
│   └── Sheet 7: Cong_no_va_Ngan_hang (Cong no 131/331, Quet chi tien mat >20tr/5tr, Cum HD trong ngay)
│
├── cau-thanh/                             --> Thư mục lưu các file phân tích chuyên đề độc lập
│   ├── Bang_ke_mua_vao_ban_ra_tong_hop.xlsx (Bang ke gop da loc dong Tong cong)
│   ├── Tong_hop_doi_chieu_GTGT_TNDN.xlsx    (Bang doi chieu GTGT 10 nam chuan 5 bang CVCBRR)
│   ├── Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx (Xet 58+ NCC theo F0, F1, DKKD & chi tiet HD)
│   ├── Tong_hop_TNCN_va_28_tieu_chi_rui_ro.xlsx (Chi tiet QTT TNCN & 28 tieu chi)
│   ├── Phan_tich_XNT_va_Gia_thanh.xlsx      (Boc tach XNT kho, gia thanh va dinh muc)
│   └── Phan_tich_Tien_mat_va_Ngan_hang.xlsx (Danh sach chi tien mat vuot nguong & sao ke)
│
├── sao-luu/                               --> Thư mục lưu trữ kỹ thuật
│   ├── Dữ liệu json trích xuất từ XML
│   └── Các bản nháp thuyết minh Markdown
│
└── De_cuong_kiem_tra_tai_tru_so.docx      --> Bản Kế hoạch / Đề cương kiểm tra chi tiết tại trụ sở NNT
```

---

## 2. Checklist kiểm tra sản phẩm đầu ra Bước 4

- [ ] **Khớp đúng 100% số liệu giữa các phân hệ:**
  - [ ] Doanh thu bán ra HĐĐT khớp Chỉ tiêu [34] GTGT và khớp Chỉ tiêu [01] BCTC (sau khi bóc tách HĐ hủy).
  - [ ] Tổng tồn kho cuối kỳ (TK 152 + 155 + 156) khớp đúng 100% Chỉ tiêu [141] Hàng tồn kho BCTC.
  - [ ] Tổng giá vốn hàng bán xuất kho (TK 632) khớp đúng 100% Chỉ tiêu [02] Giá vốn BCTC.
  - [ ] Chi phí lương trên Bảng tính giá thành (TK 622) khớp với Chi phí lương QTT TNCN (Chỉ tiêu [23]).
- [ ] **Khoanh vùng rủi ro rõ ràng:**
  - [ ] Đã chỉ rõ danh sách và giá trị mua vào của các NCC thuộc diện F0 / Công an / đóng MST.
  - [ ] Đã quét và lập danh sách chi tiết các khoản chi tiền mặt > 20 triệu (trước 07/2025) và >= 5 triệu (từ 07/2025).
  - [ ] Đã bóc tách các ngày có cụm hóa đơn của 1 NCC vượt ngưỡng thanh toán không dùng tiền mặt.
  - [ ] Đã xác định rõ các điểm bất thường nhân sự (lao động đa nguồn, trùng CCCD).
- [ ] **Đồng bộ lưu trữ 3 cấp:**
  - [ ] Gốc `ket-qua/`: Duy nhất 01 file `Phan_tich_ho_so_cong_ty_<TenDN>.xlsx`.
  - [ ] Thư mục `cau-thanh/`: Đầy đủ 6 file chuyên đề độc lập.
  - [ ] Thư mục `sao-luu/`: Lưu trữ file kỹ thuật `.json`, `.md`.
