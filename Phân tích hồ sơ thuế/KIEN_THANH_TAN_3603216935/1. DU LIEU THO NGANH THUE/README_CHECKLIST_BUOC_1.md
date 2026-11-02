# HƯỚNG DẪN & CHECKLIST BƯỚC 1: DỮ LIỆU THÔ NGÀNH THUẾ
(Dữ liệu do Cơ quan thuế chuẩn bị sẵn trước khi mở hồ sơ kiểm tra)

Thư mục này chứa toàn bộ các dữ liệu điện tử gốc trích xuất từ hệ thống công nghệ thông tin ngành Thuế.

---

## 1. Cấu trúc các thư mục con

```text
1. DU LIEU THO NGANH THUE/
│
├── 1.1. To khai XML/                  --> Toàn bộ file tờ khai XML kết xuất từ hệ thống Quản lý thuế
│   ├── To khai GTGT (Mau 842 - thang hoac quy)
│   ├── To khai Bo sung GTGT (KHBS)
│   ├── To khai Quyet toan TNDN (Mau 892 kem Phu luc 03-1A)
│   ├── To khai Quyet toan TNCN (Mau 953 kem Phu luc 05-1/BK)
│   ├── To khai Khau tru TNCN quy (Mau 05/KK-TNCN)
│   ├── Bao cao tai chinh nam (Mau 402: CDKT, KQKD, LCTT, Thuyet minh)
│   └── Phu luc Giao dich lien ket (Nghi dinh 132/2020/ND-CP neu co)
│
├── 1.2. Hoa don dien tu & Bang ke/    --> Dữ liệu hóa đơn và bảng kê chi tiết
│   ├── HDDT.xlsx                      (File trich xuat tu Cong thong tin Hoa don dien tu)
│   ├── Bang ke ban ra cac ky.xlsx     (File bang ke hoa don ban ra chi tiet)
│   └── Bang ke mua vao cac ky.xlsx    (File bang ke hoa don mua vao chi tiet)
│
└── 1.3. Nghia vu nop thue & Thong bao/--> Tình hình thực hiện nghĩa vụ thuế
    ├── Lich su nop thue vao NSNN.xlsx
    ├── Thong bao tien thue no, tien phat, tien cham nop.pdf
    └── Thong tin dang ky thue ban dau (Mau 08-MST neu co).pdf
```

---

## 2. Checklist kiểm tra tính đầy đủ của Bước 1

- [ ] **Đủ số lượng tờ khai XML:** Đủ 4 quý/năm (hoặc 12 tháng/năm) cho toàn bộ các năm trong niên độ kiểm tra đối với:
  - [ ] Thuế GTGT (Mẫu 842)
  - [ ] Quyết toán TNDN (Mẫu 892)
  - [ ] Quyết toán TNCN (Mẫu 953)
  - [ ] Báo cáo tài chính (Mẫu 402)
- [ ] **Đã thu thập Tờ khai bổ sung (KHBS):** Kiểm tra có tờ khai bổ sung nào làm thay đổi số thuế phải nộp hoặc số thuế được khấu trừ chuyển kỳ sau không.
- [ ] **Đủ Bảng kê Hóa đơn:** Đủ bảng kê mua vào và bán ra của từng quý, định dạng file Excel rõ ràng.
- [ ] **Dữ liệu HDDT:** Đã tải file trích xuất từ hệ thống HĐĐT tập trung ngành Thuế để đối chiếu ngược với bảng kê do NNT nộp.

---

## 3. Quy chuẩn đặt tên file khuyến nghị
- Tờ khai GTGT: `GTGT_Q1_2024_842.xml`, `GTGT_Q2_2024_KHBS_842.xml`
- Tờ khai TNDN: `QTT_TNDN_2024_892.xml`
- Tờ khai TNCN: `QTT_TNCN_2024_953.xml`, `TNCN_Q1_2024_05KK.xml`
- Báo cáo tài chính: `BCTC_2024_402.xml`
- Bảng kê: `Bang_ke_ban_ra_Q1_2024.xlsx`, `Bang_ke_mua_vao_Q1_2024.xlsx`
