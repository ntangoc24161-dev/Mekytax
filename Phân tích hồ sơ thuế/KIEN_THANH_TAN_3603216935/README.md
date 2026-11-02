# HƯỚNG DẪN SỬ DỤNG BỘ THƯ MỤC MẪU HỒ SƠ KIỂM TRA THUẾ

Bộ thư mục này được chuẩn hóa theo **Quy trình phân tích hồ sơ kiểm tra thuế 4 bước khép kín** dành cho công chức và đoàn kiểm tra thuế. Khi bắt đầu kiểm tra một doanh nghiệp mới, bạn chỉ cần sao chép toàn bộ thư mục này và đổi tên thành:
`<TEN_DOANH_NGHIEP>_<MST>` (Ví dụ: `VAN_PHAT_THANH_3601409441`).

---

## 1. Cấu trúc tổng quan 4 cấp thư mục chuẩn hóa

```text
<TEN_DOANH_NGHIEP>_<MST>/
│
├── 1. DU LIEU THO NGANH THUE/             --> BƯỚC 1: Hồ sơ cơ quan thuế chuẩn bị sẵn
│   ├── 1.1. To khai XML/                  --> Tờ khai GTGT (842), QTT TNDN (892), QTT TNCN (953), BCTC (402), KHBS
│   ├── 1.2. Hoa don dien tu/              --> File HDDT.xlsx tập hợp danh sách hóa đơn điện tử portal (mua vào, bán ra)
│   ├── 1.3. Nghia vu nop thue & Thong bao/--> Lịch sử nộp NSNN, thông báo nợ thuế, đăng ký thuế ban đầu, PTDN01, PTDN02
│   └── 1.4. Tong hop danh sach NLD theo QTT TNCN/ --> File tổng hợp danh sách người lao động theo quyết toán thuế TNCN hàng năm
│
├── 2. HO SO NNT CUNG CAP/                 --> BƯỚC 2: Hồ sơ do Doanh nghiệp nộp theo Phiếu yêu cầu
│   ├── 2.1. So ke toan tong hop va chi tiet/ --> Sổ Nhật ký chung, Sổ Cái, Sổ chi tiết TK 111, 112, 131, 331, 3388, 341
│   ├── 2.2. Kho - Chi phi - Gia thanh/    --> Sổ XNT 152/155/156, Sổ 154/621/622/627, Bảng tính giá thành SP, Bảng định mức kỹ thuật
│   ├── 2.3. Ngan hang & Dong tien/        --> File sao kê Excel Internet Banking, Sao kê PDF có dấu, UNC cho HĐ >= 20tr (hoặc 5tr)
│   ├── 2.4. Lao dong - Tien luong - TSCD/ --> Bảng thanh toán lương 12 tháng, Bảng chấm công, HĐLĐ, Bảng trích khấu hao 214, phân bổ 242
│   └── 2.5. Bang ke mua vao ban ra/       --> Tập hợp dữ liệu bảng kê mua vào, bán ra do công ty kê khai và cung cấp
│
├── 3. DU LIEU BO SUNG & TRA CUU QLT/      --> BƯỚC 3: Dữ liệu tra cứu quản lý thuế & Giải trình bổ sung
│   ├── 3.1. Canh bao rui ro CVCBRR & Cong an/ --> Sổ F0 người bán rủi ro, Sổ F1 người mua, Sổ Công an theo dõi hóa đơn
│   ├── 3.2. Danh ba trang thai NNT/       --> 01 file Excel tập hợp danh sách đối tác mua vào bán ra theo danh sách HĐĐT & Bảng kê + Danh bạ ĐKKD
│   ├── 3.3. Tra cuu TNCN da nguon & Nhan su/ --> Dữ liệu tra cứu thu nhập nhiều nơi ngành Thuế, xác minh CCCD/CMND trùng lặp
│   └── 3.4. Giai trinh & Ho so bo sung/   --> Quyết định kiểm tra, Biên bản làm việc, Văn bản giải trình & Hồ sơ bổ sung của NNT
│
└── 4. KET QUA PHAN TICH/                  --> BƯỚC 4: Kết xuất phân tích chuyên sâu & Hồ sơ kiểm tra
    ├── CHECKLIST_TONG_HOP_HO_SO.xlsx      --> File Master Checklist duy nhất (In chuẩn 01 trang A4 đứng, giải thích rõ thiếu/đủ)
    ├── Phan_tich_ho_so_cong_ty_<TenDN>.xlsx --> File Master Excel 7 Sheet chuẩn hóa duy nhất
    ├── cau-thanh/                         --> Các file Excel phân tích chuyên đề độc lập (gốc bảng kê, đối chiếu 3 chiều, XNT, Giá thành, NCC F0, TNCN)
    └── sao-luu/                           --> File json, md, bản nháp kỹ thuật trung gian
```

---

## 2. Cách kiểm tra tình trạng thiếu / đủ hồ sơ

1. Nạp các file tài liệu thu thập được vào đúng các thư mục con từ mục `1.` đến mục `3.`.
2. Chạy kịch bản tự động kiểm tra và xuất phiếu kết quả rà soát ra file Master Excel Checklist in ấn tại thư mục `4. KET QUA PHAN TICH/`:
   ```bash
   python kiem_tra_tinh_day_du_ho_so.py
   ```
3. Mở file `4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx` để in trực tiếp phiếu kiểm soát phục vụ hồ sơ kiểm tra lưu trữ.
