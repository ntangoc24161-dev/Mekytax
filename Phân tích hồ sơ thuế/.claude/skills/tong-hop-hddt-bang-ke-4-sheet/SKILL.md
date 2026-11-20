---
name: Tổng hợp hóa đơn và bảng kê 4 sheet
description: "Tổng hợp HĐĐT cổng CQT và bảng kê NNT thành file 5 sheet (mua/bán riêng + tổng hợp rủi ro NCC), gắn cờ F0/CA/F1, chuẩn hóa MST BT6."
group: Tài chính
updated: '2026-11-18'
---

# Tổng hợp hóa đơn và bảng kê kèm đối chiếu rủi ro NCC

Skill này tổng hợp toàn diện hóa đơn từ hai nguồn độc lập: **Hóa đơn điện tử (HĐĐT)** từ cổng CQT và **Bảng kê** do người nộp thuế (NNT) cung cấp. Quy trình chạy theo **Hướng 2 (đóng gói 2 file chuẩn liên kết dữ liệu)**:

1. **File 1**: `Tong_hop_HDDT_BangKe_<MST>_<KY>.xlsx` gồm 4 sheet dữ liệu mua/bán riêng biệt và 1 sheet tổng hợp rủi ro NCC (tổng cộng 5 sheet).
2. **File 2**: `Tong_hop_NCC_mua_vao_CV_trang_thai_<MST>_<KY>.xlsx` gồm 5 sheet chuyên đề kiểm tra rủi ro NCC mua vào (chuẩn skill `doi-chieu-ncc-mua-vao-rui-ro`).

## Khi nào dùng

- Cần tổng hợp dữ liệu hóa đơn mua vào và bán ra tách bạch giữa HĐĐT (cổng CQT) và Bảng kê NNT.
- Gắn cờ rủi ro trực tiếp vào từng dòng hóa đơn mua vào: đối soát với Sổ CVCBRR (F0, F1) và Sổ Công an (`Bao_cao_cong_van.xlsx`).
- Bổ sung tra cứu Danh bạ NNT (Tình trạng hoạt động ĐKKD & Cơ quan thuế quản lý).
- Tự động xuất sheet tổng hợp theo trạng thái hoạt động ĐKKD và cảnh báo rủi ro.
- Xuất file chuyên đề đối chiếu 100% NCC mua vào với CVCBRR và Công an.

## Cấu trúc File 1: Tong_hop_HDDT_BangKe_<MST>_<KY>.xlsx (5 sheet)

1. **`HDDT_Mua_vao`** (27 cột): 100% dòng hóa đơn mua vào từ cổng CQT (loại bỏ HĐ hủy/thay thế nếu có cờ, giữ đầy đủ thông tin).
   - Thêm cột `MST_chuan` (BT6).
   - 4 cột cuối đối soát: `Tình trạng hoạt động`, `Cơ quan thuế quản lý`, `Cảnh báo rủi ro` (F0/Công an/F1), `Thông tin công văn CA / CVCBRR`.
2. **`HDDT_Ban_ra`** (25 cột): 100% dòng hóa đơn bán ra từ cổng CQT, thêm cột `MST_chuan` và 2 cột Danh bạ NNT.
3. **`Bang_ke_Mua_vao`** (22 cột): 100% dòng bảng kê mua vào NNT nộp (lọc bỏ các dòng nhãn "Tổng cộng", giữ hóa đơn đối tác "TỔNG CÔNG TY...").
   - Thêm cột `MST_chuan` (BT6).
   - 4 cột cuối đối soát: `Tình trạng hoạt động`, `Cơ quan thuế quản lý`, `Cảnh báo rủi ro` (F0/Công an/F1), `Thông tin công văn CA / CVCBRR`.
4. **`Bang_ke_Ban_ra`** (19 cột): 100% dòng bảng kê bán ra NNT nộp, thêm cột `MST_chuan` và 2 cột Danh bạ NNT.
5. **`Tong_hop_trang_thai_rui_ro`**: Sheet tổng hợp tóm tắt toàn bộ NCC mua vào gồm 3 phần:
   - **Phần I**: Bảng thống kê theo Tình trạng hoạt động (ĐKKD - Danh bạ NNT).
   - **Phần II**: Bảng thống kê theo Cảnh báo rủi ro (Công an, F0, F1, Bình thường).
   - **Phần III**: Danh sách chi tiết các NCC có cảnh báo rủi ro hoặc trạng thái ĐKKD bất thường kèm kiến nghị kiểm tra.

## Cấu trúc File 2: Tong_hop_NCC_mua_vao_CV_trang_thai_<MST>_<KY>.xlsx (5 sheet)

File chuyên đề đối chiếu NCC rủi ro theo chuẩn Vạn Phát Thành:
1. **`Phan_1_F1_NNT`**: Rà soát rủi ro tự thân của NNT (kiểm F0, F1, Công an đối với chính NNT).
2. **`Bang_2_1_NCC_F0`**: Danh sách NCC thuộc diện CVCBRR (F0, F1) và Sổ Công an.
3. **`Bang_2_2_Toan_bo_NCC`**: Bảng tổng hợp 100% NCC mua vào (Union Bảng kê ∪ HĐĐT) có cột Nguồn (`Bảng kê`, `HĐĐT`, `Chung`), đối chiếu chênh lệch số lượng HĐ, doanh số, tiền thuế.
4. **`Bang_2_3_Ky_ke_khai_CT`**: Bóc tách chi tiết từng hóa đơn theo kỳ kê khai của các NCC dính F0/CA/F1 hoặc ĐKKD bất thường.
5. **`Phan_3_Nhan_xet`**: Đánh giá rủi ro và đề xuất trọng tâm kiểm tra hồ sơ, tài khoản ngân hàng thực tế.

## Quy tắc chuẩn hóa Mã số thuế (BT6 - MST_chuan)

1. **10 chữ số**: Giữ nguyên (VD: `3602996707` → `3602996707`).
2. **9 chữ số**: Thêm 1 số 0 ở đầu để đủ 10 số (VD: `602996707` → `0602996707`).
3. **12 chữ số** (bắt đầu từ 0): Thêm 1 số 0 ở đầu và thêm dấu gạch ngang `-` trước 3 chữ số chi nhánh cuối (VD: `602996707001` → `0602996707-001`).
4. **13 chữ số** (bắt đầu từ 3): Giữ nguyên và thêm dấu gạch ngang `-` trước 3 chữ số chi nhánh cuối (VD: `3602996707001` → `3602996707-001`).
5. **14 ký tự** (đã có sẵn dấu `-`): Giữ nguyên.

## Thứ tự ưu tiên gắn cờ rủi ro NCC

Đối soát MST người bán với các nguồn theo thứ tự ưu tiên:
1. **Sổ Công an (`Bao_cao_cong_van.xlsx`)**: Gắn nhãn `Công an`, ghi rõ số CV, ngày CV, cơ quan điều tra.
2. **CVCBRR F0 (`SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` sheet `F0`)**: Gắn nhãn `F0`, ghi số CV cảnh báo và cơ quan thuế ban hành.
3. **CVCBRR F1 (`SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` sheet `F1`)**: Gắn nhãn `F1`, ghi số CV.
4. Nếu vừa dính Công an vừa F0/F1 thì ưu tiên gắn nhãn `Công an`.

## Định dạng & Lưu trữ

- Font: Times New Roman thống nhất toàn bộ các sheet.
- Số tiền định dạng `#,##0`. Dòng tổng cộng dùng công thức `SUM`.
- Các dòng dính Công an/F0 tô nền hồng nhạt (`FCE4D6`), chữ đậm đỏ (`C00000`); F1 tô nền vàng nhạt (`FFF2CC`), chữ đậm cam (`B25900`).
- Lưu đồng thời vào `ket-qua/cau-thanh/` trong Second Brain và thư mục hồ sơ `4. KET QUA PHAN TICH/cau-thanh/` trên máy kiểm tra.
