---
name: Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN
description: "KQ3 80 quy tắc, 28 tiêu chí rủi ro, bảng tổng hợp 9 dấu hiệu bất thường và chi tiết giải trình NNT 10 sheet."
group: Tài chính
updated: '2026-10-23'
---

# Giai đoạn 3 - Kết luận đánh giá & Bóc tách dấu hiệu bất thường (v1.4)

Áp dụng HƯỚNG DẪN THƯỜNG TRỰC trong skill `phan-tich-ho-so-thue`. Sao mẫu KQ3 rồi điền. Mọi kết luận sai phạm, mức độ rủi ro và bộ hồ sơ đối chiếu giải trình chi tiết được hoàn chỉnh tại đây.

Tên skill giữ "sổ sách DN" vì slug đã khóa; GĐ3 chạy cả khi chưa có sổ: quy tắc thiếu nguồn ghi KHÔNG_ĐỦ_DỮ_LIỆU / Chưa kết luận, không dừng.

## Khi nào dùng

Sau KQ1.3, KQ2.1, KQ2.2. "phân tích giai đoạn 3", "Lệnh 3", "xuất KQ3", "đề cương", "chuỗi bằng chứng", "tổng hợp dấu hiệu bất thường", "xuất file NNT giải trình chi tiết".

## Đầu ra

### 1. File kết luận chính: `KQ3_Ket_luan_danh_gia_<MST>_<KY>_v<n>.xlsx` (2 sheet)
- `KQ3`: 39 cột A-AM. Dữ liệu từ dòng 10. Đủ 80 quy tắc, mỗi quy tắc ít nhất một dòng.
- `Danh_gia_28_tieu_chi_rui_ro`: 7 cột (`ID`, `Tên tiêu chí`, `Phân loại`, `Ngưỡng rủi ro`, `Kết quả phân tích trên hồ sơ`, `Mức độ rủi ro`, `Kiến nghị kiểm tra`). 28 dòng narrative độc lập với bảng 80 quy tắc.

### 2. File tổng quan dấu hiệu: `KQ3_Danh_gia_Dau_hieu_Bat_thuong_<MST>_<KY>.xlsx`
Bảng tổng hợp các dấu hiệu bất thường trọng yếu phát hiện qua phân tích (mức độ Cao / Trung bình / Thấp, số liệu cụ thể, vị trí sheet KQ, hướng kiểm tra chi tiết 5-6 bước, tài liệu cần chuẩn bị và trạng thái kiểm tra).

### 3. File chi tiết số liệu đối chiếu giải trình: `KQ3_NNT_Giai_trinh_Chi_tiet_<MST>_<KY>.xlsx`
Bảng tổng hợp chi tiết toàn bộ các dấu hiệu bất thường đã bóc tách từ `KQ2.1` và `KQ2.2` thành dữ liệu số liệu thật (không phải file mẫu để trống cho NNT tự điền), gồm **10 Sheet**:
- `00_Tong_hop_chi_tiet`: Bảng tổng quan + 9 Phần chi tiết (Phần I đến IX) gom trên 1 sheet duy nhất phục vụ đối chiếu trực tiếp.
- `1_NCC_F0_CanhBao`: Chi tiết danh mục NCC F0/Công an từ IN07 (tô đỏ nổi bật `#FCE4D6` / chữ `#C00000`).
- `2_NCC_DKKD_bat_thuong`: Chi tiết hóa đơn mua vào từ 24 NCC tạm ngừng KD, đóng MST, bỏ địa chỉ (tô vàng cam `#FFF2CC` / chữ `#B25900`).
- `3_Ty_le_gia_von`: Phân tích tập hợp chi phí và tỷ lệ giá vốn theo năm, biên lãi gộp.
- `4_HD_ban_duoi_gia`: Chi tiết các dòng đối chiếu 511-632 bán dưới giá vốn (tô đỏ nổi bật dòng bán lỗ).
- `5_Khong_vay_no_EBITDA`: Phân tích EBITDA theo NĐ 132/2020 4 năm, chi phí lãi vay = 0 (tô xanh lá `#E2EFDA` / chữ `#385723`).
- `6_Top5_NCC_95%`: Tỷ trọng mua vào Top-5 nhà cung cấp (MST, tên NCC, doanh số, % tỷ trọng, hợp đồng).
- `7_Lech_HDDT_BangKe`: Đối chiếu nghĩa vụ thuế GTGT qua các năm giữa HĐĐT, Bảng kê và Tờ khai.
- `8_Cong_no_TK131_TK331`: Biến động nợ 131, 331, chi tiền mặt 111 & danh mục sao kê NH 112 (phân loại PDF scan/trích xuất).
- `9_TNCN_NLD_nhieu_nguon`: Quyết toán thuế TNCN, rà soát lao động đa nguồn, kiểm tra trùng CCCD.

## Quy chuẩn định dạng chuyên nghiệp bắt buộc (Q14-Q20)

Mọi file Excel xuất ra từ Giai đoạn 3 (kể cả file thô và file tổng hợp chi tiết) BẮT BUỘC áp dụng bộ quy chuẩn:
- **Font chữ:** Toàn bộ workbook dùng font **Times New Roman**, Unicode chuẩn (tự sửa lỗi mã VNI/TCVN3); cỡ chữ 10pt cho dữ liệu, 10-11pt in đậm cho tiêu đề bảng, 14pt in đậm cho tiêu đề chính.
- **Màu sắc nhận diện:** Tiêu đề bảng nền xanh Navy đậm (`#1F4E78`), chữ trắng in đậm, căn giữa. Tiêu đề phân đoạn nền xanh nhạt (`#D9E1F2`), chữ xanh đậm.
- **Định dạng số và ngày tháng:**
  - Ngày tháng hiển thị dạng **`dd/mm/yyyy`**, không hiển thị giờ, căn giữa.
  - Số tiền và giá trị hiển thị có dấu chấm phân cách hàng nghìn (`#,##0`), căn phải.
  - Tỷ lệ phần trăm hiển thị dạng **`0.00%`**, căn phải.
- **Đường viền và hiển thị:** Kẻ viền ô mỏng rõ ràng (`#D9D9D9`), bật chế độ hiển thị đường lưới (`showGridLines = True`) trên 100% sheet.
- **Tự động giãn cột:** Tính toán độ rộng cột tối ưu theo nội dung dài nhất (từ 10 đến 50-60 ký tự, `wrap_text` nếu quá dài), không để lỗi `###` hoặc tràn chữ.
- **Tô màu cảnh báo rủi ro:**
  - **Màu đỏ đào (`#FCE4D6` / chữ đỏ `#C00000`):** Nổi bật NCC F0/Công an và dòng hóa đơn bán dưới giá vốn.
  - **Màu vàng cam (`#FFF2CC` / chữ cam `#B25900`):** Nổi bật NCC có trạng thái ĐKKD bất thường.
  - **Màu xanh lá nhạt (`#E2EFDA` / chữ xanh `#385723`):** Nổi bật thông tin tích cực / tuân thủ tốt (không vay nợ, khớp 100%).

## Quy trình (Lệnh 3)

1. Sao KQ3 từ mẫu. Chạy 80 quy tắc. Thiếu nguồn: O = KHÔNG_ĐỦ_DỮ_LIỆU, P = Chưa kết luận, AA nêu IN còn thiếu.
2. Dòng Có phát hiện: Nhãn, Mức (F0/F1 = Cao, K2), hệ số, điểm, sắc thuế, V. Chống trùng Mã_PH, dòng chủ, cột W theo T1-T10. Cột W tính bằng mã. 0 ≤ W ≤ V; W < V thì X bắt buộc.
3. Điền AB, AC, AH cho dòng Cao và Trung bình; AI, AJ cho NCC/HĐ trọng điểm (mắt xích thiếu = "Không đủ dữ liệu", K11). AD, AK, AL, AM để trống cho cán bộ. AL không cộng vào W.
4. Sắp xếp mục 5.4 quy trình. Kiểm lại độc lập 10 phát hiện điểm cao nhất.
5. Báo số dòng theo kết quả; tổng cột W theo sắc thuế; việc công ty còn thiếu (cột AA).
6. Điền `Danh_gia_28_tieu_chi_rui_ro`: narrative 28 dòng dựa trên các dòng KQ3 liên quan, gán Mức độ rủi ro và Kiến nghị kiểm tra nhất quán với kết luận đã chốt.
7. Bóc tách danh mục các dấu hiệu bất thường, xuất file `KQ3_Danh_gia_Dau_hieu_Bat_thuong_<MST>_<KY>.xlsx`.
8. Trích xuất toàn bộ dữ liệu chi tiết từ KQ2.1, KQ2.2 và KQ3, dựng file `KQ3_NNT_Giai_trinh_Chi_tiet_<MST>_<KY>.xlsx` 10 sheet (gồm sheet `00_Tong_hop_chi_tiet` và 9 sheet chuyên đề), chạy script định dạng chuẩn hóa toàn bộ workbook.

## Lệnh 4 (công ty gửi thêm)

Cùng Mã_PH; cột AE = Mới / Giữ nguyên / Nâng lên Chắc / Đổi kết luận / Đổi mức / Giải tỏa; cột AG lý do. Bản cũ giữ lịch sử. Chỉ điền ô trống của quy tắc mở khóa ở KQ2. Cập nhật lại file dấu hiệu bất thường và file giải trình chi tiết tương ứng.

## Cấm

Tự gắn nhãn trốn thuế. Số 0 giả. Ghi đè bản cũ. Tách sheet KQ3 cũ. Cộng số từ KQ2 vào tổng. Để file chi tiết giải trình ở dạng bảng trống không có số liệu.

## Lệnh siêu ngắn

```
Hoàn chỉnh GĐ3 hồ sơ <FOLDER> MST <MST> kỳ <KY> xuất KQ3 và chi tiết giải trình
```
