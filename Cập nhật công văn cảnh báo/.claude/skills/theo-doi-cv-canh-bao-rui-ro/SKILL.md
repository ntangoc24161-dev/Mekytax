---
name: Theo dõi công văn cảnh báo rủi ro
description: Bóc tách công văn cảnh báo thuế/HĐĐT, cập nhật độc lập sheet F0 bên bán và F1 bên mua sổ SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.
group: Tài chính
updated: '2026-10-01'
learned_updated: '2026-10-01'
---

# Theo dõi công văn cảnh báo rủi ro

Kỹ năng tự động bóc tách thông tin từ các công văn cảnh báo rủi ro thuế, hóa đơn điện tử (PDF scan, văn bản số hóa, trích xuất iOffice) và cập nhật dữ liệu chuẩn hóa vào Sổ theo dõi tập trung ngành Thuế.

## Khi nào dùng
- Khi tiếp nhận các công văn cảnh báo rủi ro hóa đơn từ Tổng cục Thuế, Cục Thuế các tỉnh hoặc Công an các địa phương chuyển đến.
- Khi cần bóc tách danh sách doanh nghiệp bán hàng rủi ro (F0) hoặc doanh nghiệp mua hàng hóa dịch vụ liên quan (F1).
- Khi cần nạp hàng loạt dữ liệu mới vào bảng tính Excel theo dõi mà không làm đứt gãy cấu trúc hay hỏng định dạng.

## Vị trí tệp & Cấu trúc Sổ theo dõi

### 1. Đường dẫn tệp cốt lõi
- **Sổ theo dõi chính:** `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`
- **Thư mục sao lưu (Backup):** `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\Sao lưu`
- **Kịch bản thực thi:** `skills/theo-doi-cv-canh-bao-rui-ro/scripts/update_so_theo_doi.py`

### 2. Cấu trúc 2 sheet dữ liệu độc lập
Sổ theo dõi đã được tối ưu hóa dung lượng (giảm từ 2,76 MB xuống 2,12 MB) và lược bỏ hoàn toàn các sheet cũ (`THEO_DOI`, `TONG_HOP`, `DANH_MUC`, `HUONG_DAN`), chia thành 2 sheet độc lập:
- **Sheet `F0` (Danh sách bên bán):** Quản lý toàn bộ danh sách người nộp thuế bán hàng có dấu hiệu rủi ro, xuất hóa đơn bất hợp pháp do Thuế và Công an cảnh báo (hiện có 17.333 dòng, 19 cột, đã bỏ nhóm mua/hóa đơn).
- **Sheet `F1` (Danh sách bên mua):** Quản lý toàn bộ danh sách người nộp thuế mua hàng hóa, dịch vụ liên quan đến hóa đơn rủi ro cần rà soát, giải trình (hiện có 7.282 dòng nguồn cũ; sau làm sạch TB DO KT1 MOI + CV 33 còn 3.704 dòng, 25 cột).

### 3. Quy chuẩn 25 cột dữ liệu thống nhất (áp dụng cho cả F0 và F1)
Mỗi dòng dữ liệu được chuẩn hóa đúng 25 cột:
1. `STT`: Số thứ tự tăng dần theo STT cuối của từng sheet.
2. `SoCV`: Số ký hiệu công văn cảnh báo (ví dụ: `1370/TCS7-QLDN1`).
3. `NgayCV`: Ngày ban hành văn bản (kiểu Date).
4. `DonViBH`: Cơ quan ban hành (chuẩn hóa tên đầy đủ).
5. `TenBan`: Tên doanh nghiệp bên bán.
6. `MSTBan`: Mã số thuế bên bán.
7. `NoiDung`: Tóm tắt hành vi/dấu hiệu rủi ro cốt lõi.
8. `TenMua`: Tên doanh nghiệp bên mua.
9. `MSTMua`: Mã số thuế bên mua.
10. `TamNgung`: Trạng thái NNT (Tạm ngưng, Bỏ trốn, Đóng MST...).
11. `KyKK`: Kỳ kê khai (Tháng/Quý/Năm).
12. `SoHD`: Ký hiệu và số hóa đơn liên quan.
13. `DoanhThu`: Doanh thu chưa thuế (số nguyên/thực).
14. `ThueGTGT`: Tiền thuế GTGT (số nguyên/thực).
15. `NhomND`: Nhóm nội dung rủi ro.
16. `DiaChi`: Địa chỉ trụ sở doanh nghiệp.
17. `NgayNhan`: Ngày cơ quan thuế tiếp nhận văn bản.
18. `TinhTrang`: Tình trạng xử lý/giải trình hồ sơ.
19. `GhiChu`: Ghi chú nghiệp vụ bổ sung.
20. `File`: Tên tệp hồ sơ đính kèm hoặc văn bản scan gốc.
21. `NamCV`: Năm ban hành công văn.
22. `Trung`: Cờ đánh dấu trùng lặp.
23. `KhoaCV`: Khóa liên kết công văn (`SoCV_Nam`).
24. `CVDau`: Đánh dấu `1` tại dòng đầu tiên của công văn, các dòng sau để trống.
25. `MSTDau`: Đánh dấu `1` tại dòng đầu tiên của MST, các dòng sau để trống.

## Quy trình bóc tách & Nạp dữ liệu chuẩn

```text
[Công văn PDF/Scan/iOffice] 
       │
       ▼ (PaddleOCR / doc2md)
[Nội dung số hóa Markdown]
       │
       ▼ (Phân loại nghiệp vụ)
┌────────────────────────────────────────┐
│ Phân loại đối tượng:                   │
│ - Bên bán vi phạm / xuất HĐ -> Sheet F0│
│ - Bên mua sử dụng HĐ / giải trình -> F1│
└────────────────────────────────────────┘
       │
       ▼ (Đối chiếu khóa chống trùng: SoCV + MST + Sheet)
[Tạo bản sao lưu tại thư mục Sao lưu/]
       │
       ▼ (openpyxl iter_rows & append)
[Nạp vào sheet F0 / F1 + Định dạng Times New Roman]
```

### Các bước thực hiện chi tiết:
1. **Bóc tách văn bản:** Sử dụng module OCR hoặc đọc file `.md` đã chuyển đổi để trích xuất đầy đủ các trường thông tin.
2. **Phân loại sheet mục tiêu:**
   - Nếu công văn gửi danh sách đơn vị bán hóa đơn rủi ro: chuẩn bị bản ghi nạp vào sheet `F0`.
   - Nếu công văn yêu cầu rà soát đơn vị mua hàng hóa dịch vụ: chuẩn bị bản ghi nạp vào sheet `F1`.
3. **Kiểm tra chống trùng:** Kiểm tra trùng lặp dựa trên bộ ba `(Số CV + MST + Sheet mục tiêu)` trước khi nạp.
4. **Tạo bản sao lưu tự động:** Hệ thống tự động tạo tệp backup định dạng `.bak_YYYYMMDD_HHMMSS` trong thư mục `Sao lưu/`.
5. **Ghi dữ liệu & chuẩn hóa phông chữ:** Sử dụng `openpyxl` nạp dữ liệu nối tiếp, tự động tính STT và đặt phông chữ chuẩn **Times New Roman** (size 10) cho toàn bộ ô mới.

## Cách chạy kịch bản tự động

### 1. Cập nhật từ file JSON trích xuất:
```powershell
python "skills/theo-doi-cv-canh-bao-rui-ro/scripts/update_so_theo_doi.py" "records_new.json"
```

### 2. Chỉ định nạp cụ thể vào sheet F0 hoặc F1:
```powershell
# Nạp vào sheet F0 (Bên bán)
python "skills/theo-doi-cv-canh-bao-rui-ro/scripts/update_so_theo_doi.py" "records_seller.json" --sheet F0

# Nạp vào sheet F1 (Bên mua)
python "skills/theo-doi-cv-canh-bao-rui-ro/scripts/update_so_theo_doi.py" "records_buyer.json" --sheet F1
```

## Bẫy nghiệp vụ & Lưu ý kỹ thuật
- **Map cột nguồn `Tong hop danh sach moi rui ro (1).xlsb` (bắt buộc, 2026-10-01):**
  - Sheet `CBRR` → `F0` (bên bán): cột **A = MST**, **B = Tên công ty**, **E = Số CV**, **F = Ngày CV**. Cột C là Ngày CV đến, cột D là Nơi gửi. Không lấy C/D làm Số CV/Ngày CV.
  - Sheet `TB DO KT1 MOI` → `F1` (bên mua): cột **B = MST**, **C = Tên công ty**, **M = Số CV**, **N = Ngày CV**. Cột N đôi khi chứa tên cơ quan Công an thay vì ngày: đưa vào `DonViBH`. Không copy MST/tên bên mua sang cột bên bán.
  - Kịch bản làm sạch: `scripts/clean_f0_cbrr_f1_tbkt1.py`, `scripts/lam_sach_so_theo_doi.py`.
- **Danh sách 524/637 DN lấy chuẩn từ `INPUT/524.xlsx`, không lấy lại từ HDRR hay CV cong an:** Đợt 1 = 524 MST theo `1770/CAT-CSKT` ngày 29/03/2023; đợt 2 = 113 MST theo `05/BKL-CSKT(Đ4)` ngày 16/08/2023. Xóa dòng 524/1798/1770/BKL ngoài danh sách này. Không gộp trùng với dòng CBRR chỉ *nhắc* 524 trong dấu hiệu. Kịch bản: `scripts/xoa_dong_524_ngoai_danh_sach.py`.
- **Tuyệt đối không ghi đè vào sheet THEO_DOI cũ:** Sổ theo dõi hiện chỉ gồm 2 sheet `F0` và `F1`. Ghi vào sheet cũ sẽ gây lỗi bảng tính.
- **Không duyệt cell ngẫu nhiên khi mở file lớn:** Luôn sử dụng `iter_rows(values_only=True)` để đọc dữ liệu, tránh tràn bộ nhớ RAM khi mở file hơn 28.000 dòng.
- **Bảo toàn phông chữ:** Toàn bộ dữ liệu mới nạp bắt buộc phải gán phông **Times New Roman** để đồng bộ với toàn bộ sổ theo dõi.
- **Không tự ý tính toán công thức ngoài:** Do sheet `TONG_HOP` đã được lược bỏ, việc tổng hợp báo cáo sẽ được thực hiện trực tiếp thông qua truy vấn dữ liệu từ sheet `F0` và `F1`.

## Kiểm chứng sau cập nhật
1. Mở file Excel kiểm tra số dòng tăng thêm tại đúng sheet `F0` hoặc `F1`.
2. Kiểm tra cột STT nối tiếp chính xác, không bị ngắt quãng hoặc trùng lặp.
3. Kiểm tra phông chữ hiển thị đúng Times New Roman.
4. Kiểm tra file backup mới nhất đã được tạo trong thư mục `Sao lưu/`.

## Lịch sử (tự học)
- [2026-09-22] Xử lý hồ sơ scan 31 trang `[1370_TCS7-QLDN1] 1370TCS7-QLDN1 1298.pdf`, bóc tách 02 công văn cảnh báo (CV 1370 của Thuế cơ sở 7 TP.HCM về Cty Hưng Thành; CV 1432 về Cty Gervina), bổ sung 03 dòng dữ liệu vào sổ.
- [2026-09-23] Đóng gói toàn diện quy trình thành skill chuẩn `theo-doi-cv-canh-bao-rui-ro` lưu tại brain và kho kỹ năng chung `shared-skills`, kèm module tự động hóa `scripts/update_so_theo_doi.py`.
- [2026-09-26] Nâng cấp pipeline tự động hóa cho 57 CV scan tháng 08/2026 và công văn số hóa tháng 06-07/2026. Khắc phục nghẽn XML openpyxl bằng `iter_rows(values_only=True)`, tích hợp cache checkpoint và giám sát Watchdog.
- [2026-09-28] Lô iOffice 22-26/09: từ 36 tệp giữ 24 tệp độc nhất, chuyển đổi OCR sang `fast_batch_ocr`, hạ cấp `python-bidi` về 0.4.2 tránh lỗi DLL.
- [2026-09-30] Hoàn tất đồng bộ dữ liệu bên bán 524 DN từ sheet `CV cong an-ben ban`, khắc phục 499 dòng lỗi font cụm CV 1798, hoàn thiện 108 cty vụ Đào Minh Thọ, nạp 72 cty mới từ Công an Hà Tĩnh và Thanh Hóa.
- [2026-09-30] Tái cấu trúc toàn diện Sổ theo dõi: phân tách thành 2 sheet độc lập `F0` (20.581 dòng bên bán) và `F1` (7.845 dòng bên mua), xóa bỏ các sheet cũ `THEO_DOI`, `TONG_HOP`, `DANH_MUC`, `HUONG_DAN`, đồng bộ toàn bộ font Times New Roman, giảm dung lượng tệp về 2,12 MB và quy hoạch toàn bộ bản sao lưu vào thư mục `00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\Sao lưu`. Đồng thời nâng cấp toàn bộ skill, script và agent theo chuẩn 2 sheet mới.
- [2026-10-01] Làm sạch lại dữ liệu nguồn xlsb: thay 3.351 dòng CBRR cũ trên F0 bằng 2.939 dòng map đúng A/B/E/F; thay 3.649 dòng TB DO KT1 MOI cũ trên F1 bằng 3.087 dòng map đúng B/C/M/N (không gán bên mua vào cột bên bán). F0 còn 19.638 dòng, F1 còn 7.282 dòng.
- [2026-10-01] Tiếp tục không làm lại toàn sổ: xóa 1 MST ngoài `524.xlsx` (`0314123789` trên CV 1798/TCT-TTKT), gộp 2.115 dòng 524 trùng (1770 / 1798.0 / 1798/TCT-TTKT / HDRR trống số CV) thành 637 dòng chuẩn. F0 còn 17.333 dòng, F1 giữ 3.704 dòng.
- [2026-10-01] Bổ sung F0 từ CBRR/HDRR-cuc/CV cong an: điền 521 dòng trống SoCV từ nguồn (CBRR dùng E/F đúng header, không dùng C/D), thêm 3 dòng mới, gộp 407 trùng; còn 372 dòng nguồn không có số CV (chủ yếu CV cong an chỉ ghi tên cơ quan). F0 còn 16.931 dòng.
