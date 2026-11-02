---
name: Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế
description: "Phân tích số liệu KQ2.1 (CQT+IN07+EBITDA) và KQ2.2 (sổ DN+511/632 chi tiết): số, chênh lệch, tỷ lệ, cờ máy; không kết luận."
group: Tài chính
updated: '2026-10-23'
---

# Giai đoạn 2 - Phân tích hồ sơ, chỉ số liệu (v1.4)

Áp dụng HƯỚNG DẪN THƯỜNG TRỰC trong skill `phan-tich-ho-so-thue`. Sao mẫu KQ2.1 và KQ2.2 rồi điền. Không ghi mức rủi ro / điểm / đề cương toàn hồ sơ vào KQ2. Nhận xét module ngắn theo mục 3.4 được phép.

Tên skill giữ "cơ quan thuế" vì slug đã khóa; phạm vi thật là cả KQ2.1 (CQT) và KQ2.2 (sổ công ty, ô trống nếu chưa có).

## Khi nào dùng

Sau KQ1.3. "phân tích giai đoạn 2", "Lệnh 2", "xuất KQ2.1", "đối chiếu GTGT", "đối chiếu 511-632", "phân tích EBITDA". Mức A-E không chặn việc dựng hai tệp.

## Đầu ra

- `KQ2.1_Phan_tich_CQT_<MST>_<KY>_v<n>.xlsx`:
  - `KQ2_Tong_hop_co`: đếm cờ máy bằng công thức.
  - `KQ2_EBITDA_Phan_tich`: phân tích EBITDA 4 năm theo Nghị định 132/2020/NĐ-CP (Lợi nhuận BDKD, Khấu hao, Chi phí lãi vay, Lãi tiền gửi, EBITDA Gross, EBITDA 30%, Lãi vay không hợp lệ).
  - `KQ2_Doi_chieu_GTGT`: đối chiếu nghĩa vụ thuế GTGT theo kỳ kê khai thực tế.
  - `KQ2_NCC_mua_vao`: bảng chi tiết 100% hóa đơn mua vào.
  - `KQ2_TNCN_LD2`: đối chiếu thuế TNCN và lao động đa nguồn.
  - `IN07_NCC_CanhBao`: danh mục tổng hợp NCC cảnh báo rủi ro từ nguồn chỉ định.
- `KQ2.2_Phan_tich_ho_so_cty_<MST>_<KY>_v<n>.xlsx`:
  - `KQ2_Tong_hop_co`: đếm cờ máy sổ công ty.
  - `KQ2_511_632`: bảng tổng hợp đối chiếu doanh thu giá vốn theo năm tài chính kèm cột Đánh giá tuân thủ & Rủi ro.
  - `KQ2_511_632_Chi_tiet`: bảng chi tiết từng dòng đối chiếu 511 và 632 (ghép theo HĐ/ngày CT/mặt hàng), công thức phát hiện bán dưới giá `IF(511_DT < 632_GV, "Có", "Không")` và tính số tiền lỗ.
  - `KQ2_Kho_Gia_thanh`: phân tích XNT kho 152/155/156 và giá vốn.
  - `KQ2_Gia_thanh_va_Dinh_muc`: phân tích cơ cấu giá thành và biên lãi gộp.
  - `KQ2_Cong_no_Tien_mat`: phân tích công nợ 131/331, chi tiền mặt 111 và ngân hàng 112.

Cờ máy: Khớp | Lệch | Chưa đủ nguồn; Vượt ngưỡng | Không vượt | Chưa đủ nguồn. Tong_hop_co đếm bằng công thức, không điền tay, không cộng vào KQ3.

## Quy trình (Lệnh 2)

1. Sao KQ2.1 và KQ2.2 từ mẫu. Lập IN07 trong KQ2.1 từ danh sách đối tác IN02 và NGUỒN CHỈ ĐỊNH; chưa có IN02 thì để trống.
2. Điền bảng chỉ bằng số liệu, chênh lệch, tỷ lệ, cờ máy. Quét 100% NCC và 100% người lao động. Ô phụ thuộc dữ liệu công ty chưa có thì ĐỂ TRỐNG (T10).
3. Bóc tách EBITDA: quét `IN04_BCTC_TNDN` theo dòng các năm (KQKD 50, CDKT 141, TNDN ct16, ct10), tính EBITDA 30% và so sánh với chi phí lãi vay phát sinh. Nếu chi phí lãi vay = 0 đ, ghi rõ "Không vay nợ - Không rủi ro K9".
4. Đối chiếu 511-632: điền sheet tổng hợp năm trên `KQ2_511_632` và lập sheet `KQ2_511_632_Chi_tiet` để phát hiện các hóa đơn bán dưới giá vốn.
5. Ánh xạ năm BCTC với năm kê khai kèm độ tin cậy. Kiểm tổng khớp chi tiết.
6. HD07 TREO (K13). TT01 ngưỡng ≥ 20 triệu đến 30/06/2025 và ≥ 5 triệu từ 01/07/2025 (K5). P23 = 10% trên sheet công nợ (K14).

## Bài học kỹ thuật & Xử lý tình huống

- **Cấu trúc `IN04_BCTC_TNDN`:** IN04 có cấu trúc **dòng xếp chồng** theo các năm liên tiếp (không phải cột F). Quét cột A (năm) + B (loại báo cáo CDKT/KQKD/TNDN) + C (mã chỉ tiêu) để bóc đúng số liệu 4 năm.
- **Xử lý PDF sao kê ngân hàng IN11:** Đo lớp chữ trước bằng `extract_text()`. Nếu có text → trích tên NH + số TK để đối chiếu sổ 112/131/331. Nếu scan 0 text (như bộ 5 sổ phụ scan) → ghi "PDF - chưa kết luận" (Q7), không tự bịa dòng UNC; chuyển sang skill `tram-tai-lieu` / OCR khi người dùng yêu cầu.
- **`KQ2_Cong_no_Tien_mat`:** Bắt buộc chạy phần 131/331 và chi tiền mặt 111 khi đã có IN08/IN10; chỉ các dòng phụ thuộc sao kê IN11 mới ghi "PDF - chưa kết luận".
- **HD04 (Top-5 NCC):** Tính gộp theo MST người bán trên toàn kỳ rồi chia tổng mua kỳ; không dùng hàm `LARGE` trên từng dòng hóa đơn.
- **Nhận xét module (3.4):** Các sheet KQ2 được phép có cột/khối nhận xét chuyên đề ngắn gắn liền số liệu, không kết luận mức rủi ro toàn hồ sơ (việc đó ở KQ3).

## Cấm

Kết luận ở KQ2. Số 0 giả. Đổi phông. Xóa định dạng có điều kiện. Xuất Tong_hop_*, Doi_chieu_*, sheet PL, workbook trọn gói.

## Lệnh siêu ngắn

```
Phân tích GĐ2 <FOLDER> MST <MST> kỳ <KY> xuất KQ2.1 KQ2.2
```
