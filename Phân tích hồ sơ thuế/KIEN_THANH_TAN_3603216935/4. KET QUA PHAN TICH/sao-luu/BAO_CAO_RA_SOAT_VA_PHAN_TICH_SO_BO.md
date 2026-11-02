# Báo cáo rà soát & phân tích sơ bộ - Kiến Thành Tân (3603216935)

- Thời điểm: 2026-10-16 16:16 Asia/Ho_Chi_Minh
- Loại việc gốc: giải thể tại bàn (`GIAI THE`); đồng thời sắp theo danh mục mẫu 4 bước để phân tích.
- Vault: `KIEN_THANH_TAN_3603216935/`
- Drive: https://drive.google.com/drive/folders/1HEBh6PfW6thceRsqfRmzLnMiI3skbtVh

## Kết luận ngắn

**Đã sắp xếp theo thư mục mẫu 4 bước** trên vault và trên Drive (20 mục đã chuyển đúng slot). Checklist tự động vault: **5/13 đạt**; còn thiếu XML/HDDT/CVCBRR/NH và Master 8 sheet nên **chưa đủ điều kiện phân tích chuyên sâu đầy đủ**.

Lưu ý: Drive đã có sổ NKC, sổ chi tiết TK, công nợ, tồn kho, lương trong `2.1`/`2.2`/`2.4`, nhưng vault chưa tải bulk nên script checklist vẫn báo 0 file ở các mục đó.

## Checklist tóm tắt

| Nhóm | Đã có | Thiếu bắt buộc |
|---|---|---|
| B1 Dữ liệu thô CQT | 1.3 Bang tinh, PTDN, nghĩa vụ, tổng quan | 1.1 XML; 1.2 HDDT portal |
| B2 Hồ sơ NNT | 2.5 bảng kê GTGT; 2.4 CCDC; Drive có NKC/SCT/công nợ/tồn kho/lương | 2.3 Sao kê NH/UNC; sync vault từ Drive 2.1-2.2 |
| B3 Bổ sung QLT | 3.2 Danh bạ; 3.4 hồ sơ giải thể Word | 3.1 CVCBRR F0/F1 & Công an; 3.3 TNCN đa nguồn |
| B4 Kết quả | Báo cáo sơ bộ + Bang tinh trong cau-thanh | Master Excel 8 sheet |

## Phân tích sơ bộ từ nguồn đã có

### Doanh thu thuần (Bang tinh/PTDN, đồng)

| Năm | Doanh thu thuần |
|---|---:|
| 2014 | 0 |
| 2015 | 688,482,148 |
| 2016 | 1,336,008,048 |
| 2017 | 562,685,000 |
| 2018 | 1,754,982,791 |
| 2019 | 1,573,134,000 |
| 2020 | 1,424,825,750 |
| 2021 | 1,041,955,000 |
| 2022 | 640,013,500 |
| 2023 | 182,350,000 |
| 2024 | 1,690,000,000 |
| 2025 | 1,140,000,000 |
| 2026 | 0 |

Nguồn: Bang tinh/PTDN đã điền hồ sơ giải thể; 2023-2025 lấy theo bộ số đã chốt khi lập Word (làm tròn tỷ với 2024-2025).

### Bảng kê GTGT (file chính 2014-2025)

- Sheet `MV` / `BR` có trong `2.5. Bang ke mua vao ban ra/`.
- Đã nhận diện đúng MST 3603216935, địa chỉ Long Bình.
- Cấu trúc theo từng năm/kỳ; còn dòng nhãn `TỔNG CỘNG` cần lọc khi gộp phân tích chi tiết.

### Hồ sơ giải thể

- Word: `3.4. Giai trinh & Ho so bo sung/HO_SO_GIAI_THE_TAI_BAN_Kien_Thanh_Tan_3603216935.docx`
- Drive: https://drive.google.com/file/d/1Pp_wqI7W2GFbc6GRelTbKx-Ebqd7OrGC/view

## Việc tiếp theo (khi anh chốt)

1. Bổ sung XML tờ khai + HDDT portal vào `1.1` / `1.2`.
2. Đổ sổ CVCBRR/Công an vào `3.1`; sao kê NH vào `2.3`.
3. Sync bulk Drive `2.1`/`2.2`/`2.4` xuống vault rồi chạy module 511-632, XNT, tiền mặt.
4. Xuất Master 8 sheet khi đủ nguồn.
