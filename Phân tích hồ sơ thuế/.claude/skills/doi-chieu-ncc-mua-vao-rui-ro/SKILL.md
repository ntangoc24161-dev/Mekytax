---
name: Đối chiếu NCC mua vào với CVCBRR và danh bạ
description: Đối chiếu NCC bảng kê mua vào với sổ CVCBRR (F0, F1), sổ Công an và danh bạ NNT gom 1 sheet kèm chi tiết hóa đơn.
group: Tài chính
origin: javis-learned
status: active
created: 2026-09-30
updated: '2026-10-05'
learned_updated: '2026-10-05'
---
## Khi nào dùng

- Đối chiếu nhà cung cấp / bên bán trên bảng kê **mua vào** với Sổ CVCBRR (F0, F1), Sổ Công an (`Bao_cao_cong_van.xlsx`) và Danh bạ NNT.
- Xuất file Excel chuẩn 2 sheet: `Tong_hop_doi_chieu` + `Chi_tiet_hoa_don`.
- So khớp với bản chuẩn Vạn Phát Thành khi nghi file bị rút gọn.

## Chuẩn bị

1. **Bảng kê mua vào**: `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx` (sheet `mua vao` đã gộp, đã lọc dòng Tổng cộng).
2. **Sổ CVCBRR**: `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` (F0 bên bán, F1 bên mua).
3. **Sổ Công an**: `Bao_cao_cong_van.xlsx`.
4. **Danh bạ NNT**: trạng thái ĐKKD.

## Cách chạy / Quy trình xuất

File Excel gồm 2 sheet, font **Times New Roman**:

### 1. Sheet `Tong_hop_doi_chieu`
- **Phần 1:** Đối chiếu NNT đang xét với F1 và Sổ Công an; ghi kết luận an toàn/có cảnh báo.
- **Phần 2:**
  - **Bảng 2.1 (F0):** NCC dính công văn cảnh báo (khoảng 18 cột: Số dòng BK, Năm phát sinh, Từ/Đến ngày HĐ, Mặt hàng chính, Tình trạng F0, Số CV, Ngày CV, Đơn vị ban hành, Nội dung cảnh báo, Trạng thái ĐKKD...).
  - **Bảng 2.2:** Liệt kê **100% NCC** trên bảng kê mua vào (bình thường + bất thường), có dòng **Tổng cộng toàn bộ**; đánh dấu nhóm trạng thái ĐKKD bất thường. **Cấm** chỉ giữ nhóm bất thường.
- **Phần 3 (bắt buộc):** (1) Đánh giá rủi ro NNT, (2) Đánh giá rủi ro NCC theo nhóm, (3) Đề xuất trọng tâm kiểm tra chứng từ.

### 2. Sheet `Chi_tiet_hoa_don`
- Toàn bộ dòng HĐ mua vào trong kỳ, sắp **`Ngay_HD` tăng dần**.
- Đủ khoảng **20 cột**, gồm nhóm chứng từ kế toán (`Loại chứng từ`, `Số chứng từ`, `Ngày chứng từ`), nhóm F0 (`Cảnh báo F0`, `Số CV F0`, `Đơn vị ban hành`, `Nội dung cảnh báo`) và cột `Trạng thái hoạt động ĐKKD` riêng (không gộp nhầm vào ghi chú số tiền).

## Bẫy

- **Bẫy rút gọn:** chỉ xuất NCC bất thường / cắt Phần 3 / thiếu cột chứng từ-F0 -> mất NCC chủ lực (ví dụ T&T từng thiếu 108 NCC ~99,5% doanh số).
- Nhầm F1 (bên mua) với F0 (bên bán).
- Khớp sót MST có/không gạch nhánh (chuẩn hóa 10 và 13 số).
- Sort `Chi_tiet` theo MST thay vì `Ngay_HD`.
- Ô text bắt đầu bằng `=` gây lỗi Removed Records Formula.

## Kiểm chứng

- Tổng doanh số mọi NCC trên Bảng 2.2 = tổng bảng kê mua vào sau lọc.
- Có đủ Phần 1 + 2.1 + 2.2 (100% NCC) + Phần 3.
- `Chi_tiet_hoa_don` đủ dòng HĐ, đủ cột chuẩn, font Times New Roman.

## Lịch sử (tự học)
- [2026-10-05] User chỉ ra file Tong_hop_NCC của T&T thiếu thông tin so với Vạn Phát Thành; skill cũ mô tả Bảng 2.2 như chỉ liệt kê NCC bất thường, dễ dẫn tới bản rút gọn. Cập nhật bắt buộc 100% NCC, Phần 3 và đủ cột Chi_tiet.
