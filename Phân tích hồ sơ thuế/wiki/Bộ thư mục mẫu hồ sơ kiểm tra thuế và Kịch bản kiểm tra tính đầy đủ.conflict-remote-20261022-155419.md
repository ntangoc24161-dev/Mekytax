---
type: wiki
status: active
tags: [wiki, phan-tich-ho-so-thue, quy-trinh, thu-muc-mau, checklist, excel-a4]
created: '2026-10-02'
updated: '2026-10-02'
source: [[wiki/Quy trình phân tích toàn diện hồ sơ kiểm tra thuế 4 bước và Bộ checklist hồ sơ.md]]
---

# Bộ thư mục mẫu hồ sơ kiểm tra thuế và Kịch bản kiểm tra tính đầy đủ

Bộ thư mục mẫu chuẩn hóa theo **Quy trình 4 bước khép kín** giúp Đoàn kiểm tra dễ dàng tiếp nhận, lưu trữ, tự động rà soát tính đầy đủ của hồ sơ kiểm tra thuế và in ấn các biểu mẫu checklist phục vụ lưu trữ hồ sơ.

Khi làm việc trên hồ sơ thật (dựng folder, phân loại file, xuất checklist, xóa nguồn cũ): skill `to-chuc-ho-so-kiem-tra-thue`. Trang này giữ cây mẫu; cây sống mới nhất cũng nằm tại `skills/to-chuc-ho-so-kiem-tra-thue/references/cay-thu-muc-4-buoc.md` (`1.2` = HĐĐT portal, `2.5` = bảng kê NNT).

## 1. Cấu trúc cây thư mục mẫu chuẩn

Thư mục mẫu được lưu tại `thu-muc-mau-ho-so-doanh-nghiep/` trong Second Brain:

```text
<TEN_DOANH_NGHIEP>_<MST>/
├── 1. DU LIEU THO NGANH THUE/
│   ├── 1.1. To khai XML/                  (Tờ khai GTGT 842, QTT TNDN 892, QTT TNCN 953, BCTC 402, KHBS)
│   ├── 1.2. Hoa don dien tu & Bang ke/    (HDDT.xlsx portal, Bảng kê mua vào, Bảng kê bán ra)
│   ├── 1.3. Nghia vu nop thue & Thong bao/(Lịch sử nộp NSNN, Thông báo nợ thuế, Đăng ký thuế)
│   ├── CHECKLIST_BUOC_1_DU_LIEU_THO_CQT.xlsx (File Excel Checklist Bước 1 chuẩn in 01 trang A4)
│   └── README_CHECKLIST_BUOC_1.md
├── 2. HO SO NNT CUNG CAP/
│   ├── 2.1. So ke toan tong hop va chi tiet/ (Sổ Nhật ký chung, Sổ Cái, TK 1111, 1121, 131, 331, 336/3388)
│   ├── 2.2. Kho - Chi phi - Gia thanh/    (Sổ XNT 152/155/156, Sổ 154/621/622/627, Bảng Z SP, Định mức, BB kiểm kê)
│   ├── 2.3. Ngan hang & Dong tien/        (Sao kê Excel Internet Banking, PDF có dấu, UNC cho HĐ >= 20tr / 5tr)
│   ├── 2.4. Lao dong - Tien luong - TSCD/ (Bảng lương 12 tháng, Bảng chấm công, HĐLĐ, Khấu hao 214, Phân bổ 242)
│   ├── CHECKLIST_BUOC_2_HO_SO_NNT_CUNG_CAP.xlsx (File Excel Checklist Bước 2 chuẩn in 01 trang A4)
│   └── README_CHECKLIST_BUOC_2.md
├── 3. DU LIEU BO SUNG & TRA CUU QLT/
│   ├── 3.1. Canh bao rui ro CVCBRR & Cong an/ (Sổ F0 người bán, Sổ F1 người mua, Sổ Công an)
│   ├── 3.2. Danh ba trang thai NNT/       (Trạng thái ĐKKD của 100% NCC mua vào)
│   ├── 3.3. Tra cuu TNCN da nguon & Nhan su/ (Tra cứu TNCN đa nguồn ngành Thuế, xác minh CCCD trùng)
│   ├── 3.4. Giai trinh & Ho so bo sung/   (Phiếu yêu cầu, Công văn giải trình, Chứng từ bổ sung)
│   ├── CHECKLIST_BUOC_3_DU_LIEU_BO_SUNG_QLT.xlsx (File Excel Checklist Bước 3 chuẩn in 01 trang A4)
│   └── README_CHECKLIST_BUOC_3.md
├── 4. KET QUA PHAN TICH/
│   ├── Phan_tich_ho_so_cong_ty_<TenDN>.xlsx (Master Excel 7 Sheet duy nhất)
│   ├── cau-thanh/                         (6 file phân tích chuyên đề độc lập)
│   ├── sao-luu/                           (File json, md kỹ thuật)
│   ├── CHECKLIST_BUOC_4_KET_QUA_PHAN_TICH.xlsx (File Excel Checklist Bước 4 chuẩn in 01 trang A4)
│   └── README_CHECKLIST_BUOC_4.md
├── CHECKLIST_TONG_HOP_HO_SO.xlsx          (Master Excel Checklist 6 Sheet chuẩn in 01 trang A4)
├── CHECKLIST_TONG_HOP_HO_SO.md            (Checklist tổng hợp 35 tiêu chí kiểm soát dạng Markdown)
└── kiem_tra_tinh_day_du_ho_so.py          (Kịch bản Python tự động quét và xuất kết quả Excel A4)
```

## 2. Hệ thống Checklist Excel định dạng chuẩn in ấn 01 trang A4

- **Master Workbook (`CHECKLIST_TONG_HOP_HO_SO.xlsx`):** Gồm 6 Sheet, thiết lập sẵn `Page Setup: A4`, `Fit to 1 page wide x 1 page tall`, Quốc hiệu, Tiêu ngữ, Tên cơ quan thuế, Thông tin doanh nghiệp, Bảng checklist kẻ viền, Hộp kết luận đánh giá và Ô ký duyệt 2 cấp (Cán bộ phân tích / Trưởng đoàn kiểm tra):
  1. `Phieu_kiem_soat_tong_hop`: Phiếu kiểm soát nhanh 13 nhóm hồ sơ theo chiều dọc (A4 Portrait).
  2. `Checklist_toan_bo_35_muc`: Toàn bộ 35 hạng mục hồ sơ chi tiết theo 4 bước (A4 Landscape).
  3. `Checklist_Buoc_1`: Chi tiết Bước 1 - Dữ liệu thô ngành Thuế (A4 Landscape).
  4. `Checklist_Buoc_2`: Chi tiết Bước 2 - Hồ sơ NNT cung cấp (A4 Landscape).
  5. `Checklist_Buoc_3`: Chi tiết Bước 3 - Dữ liệu bổ sung & Tra cứu QLT (A4 Landscape).
  6. `Checklist_Buoc_4`: Chi tiết Bước 4 - Kết quả phân tích & Báo cáo đầu ra (A4 Landscape).
- **Các file Excel chuyên đề riêng từng bước:** Đặt trực tiếp tại thư mục con tương ứng (`1.`, `2.`, `3.`, `4.`) giúp công chức kiểm tra mở và in ấn ngay tại bước nghiệp vụ đang thực hiện.

## 3. Kịch bản quét tự động tính đầy đủ hồ sơ (`kiem_tra_tinh_day_du_ho_so.py`)

Kịch bản Python tự động kiểm tra thư mục hồ sơ của doanh nghiệp và xuất báo cáo Markdown / Console hoặc xuất trực tiếp phiếu kết quả kiểm soát ra file Excel định dạng in 01 trang A4:
```powershell
python "thu-muc-mau-ho-so-doanh-nghiep/kiem_tra_tinh_day_du_ho_so.py" "<Thu_muc_doanh_nghiep>" --excel "Phieu_ket_qua_ra_soat.xlsx"
```
- Tự động đánh dấu `[X] ĐẠT` hoặc `[X] THIẾU` tương ứng với số lượng file thực tế quét được.
- Đưa ra kết luận rõ ràng: **Đủ điều kiện phân tích chuyên sâu Bước 4** hay **Cần phát hành Phiếu yêu cầu bổ sung tài liệu**.

## 4. Liên kết liên quan
- [[wiki/Quy trình phân tích toàn diện hồ sơ kiểm tra thuế 4 bước và Bộ checklist hồ sơ.md]]
- [[wiki/Quy định ngưỡng thanh toán không dùng tiền mặt trong khấu trừ thuế GTGT và chi phí TNDN.md]]
- [[wiki/Checklist mở hồ sơ thanh tra kiểm tra thuế (TTKT).md]]
