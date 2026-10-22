---
name: Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế
description: "Phân tích số liệu KQ2.1 (CQT+IN07) và KQ2.2 (sổ DN): số, chênh lệch, tỷ lệ, cờ máy; không kết luận."
group: Tài chính
updated: '2026-10-19'
---

# Giai đoạn 2 - Phân tích hồ sơ, chỉ số liệu (v1.4)

Áp dụng HƯỚNG DẪN THƯỜNG TRỰC trong skill `phan-tich-ho-so-thue`. Sao mẫu KQ2.1 và KQ2.2 rồi điền. Không ghi mức rủi ro, nhận xét, kiến nghị, tổng số tiền vào KQ2.

Tên skill giữ "cơ quan thuế" vì slug đã khóa; phạm vi thật là cả KQ2.1 (CQT) và KQ2.2 (sổ công ty, ô trống nếu chưa có).

## Khi nào dùng

Sau KQ1.3. "phân tích giai đoạn 2", "Lệnh 2", "xuất KQ2.1", "đối chiếu GTGT". Mức A-E không chặn việc dựng hai tệp.

## Đầu ra

- `KQ2.1_Phan_tich_CQT_<MST>_<KY>_v<n>.xlsx`: KQ2_Tong_hop_co (16 chỉ tiêu), KQ2_Doi_chieu_GTGT, KQ2_NCC_mua_vao, KQ2_TNCN_LD2, IN07_NCC_CanhBao.
- `KQ2.2_Phan_tich_ho_so_cty_<MST>_<KY>_v<n>.xlsx`: KQ2_Tong_hop_co (11 chỉ tiêu), KQ2_511_632, KQ2_Kho_Gia_thanh, KQ2_Gia_thanh_va_Dinh_muc, KQ2_Cong_no_Tien_mat.

Cờ máy: Khớp | Lệch | Chưa đủ nguồn; Vượt ngưỡng | Không vượt | Chưa đủ nguồn. Tong_hop_co đếm bằng công thức, không điền tay, không cộng vào KQ3.

## Quy trình (Lệnh 2)

1. Sao KQ2.1 và KQ2.2 từ mẫu. Lập IN07 trong KQ2.1 từ danh sách đối tác IN02 và NGUỒN CHỈ ĐỊNH; chưa có IN02 thì để trống.
2. Điền bảng chỉ bằng số liệu, chênh lệch, tỷ lệ, cờ máy. Quét 100% NCC và 100% người lao động. Ô phụ thuộc dữ liệu công ty chưa có thì ĐỂ TRỐNG (T10).
3. Ánh xạ năm BCTC với năm kê khai kèm độ tin cậy. Kiểm tổng khớp chi tiết.
4. HD07 TREO (K13). TT01 ngưỡng ≥ 20 triệu đến 30/06/2025 và ≥ 5 triệu từ 01/07/2025 (K5). P23 = 10% trên sheet công nợ (K14).

Module bước 2/3/4/5/6/7/8 ghi đúng sheet đích, tạo đúng một phiên bản mới của tệp đó, quy tắc khác CHƯA_CHẠY.

## Bổ sung 19/10/2026 vòng 2 (sửa lỗi phát hiện trên TT_BIEN_HOA)

**`KQ2_Doi_chieu_GTGT` - Phần I không được để trống:** lấy số liệu từ TỔNG HỢP TỜ KHAI GTGT các kỳ (IN03), gộp theo quý hoặc tháng ĐÚNG theo kỳ kê khai thật của NNT (không tự ý chọn quý khi NNT khai tháng hoặc ngược lại). Chỉ in các kỳ NẰM TRONG phạm vi kiểm tra (K12); xóa hẳn các năm/kỳ ngoài phạm vi thay vì in dòng toàn số 0 (Q19) - lỗi cũ ở TT_BIEN_HOA in cả 2016-2021 toàn số 0 trong khi kỳ kiểm tra là 2022-2025. Theo đúng cách làm đã chốt ở skill `doi-chieu-gtgt-tndn-cbrr`.

**`KQ2_NCC_mua_vao` - đổi thành MỘT bảng chi tiết hóa đơn duy nhất** (không còn 2 bảng lồng nhau kiểu Bảng 2.2 + Bảng 2.3 không phân cách - lỗi cũ dòng 147 lẫn dòng 148 ở TT_BIEN_HOA): mỗi dòng = một hóa đơn mua vào (100% hóa đơn hợp lệ, không chỉ NCC rủi ro), cột gồm MST/tên người bán, số HĐ, ngày HĐ (dd/mm/yyyy), mặt hàng, giá trị chưa thuế, thuế GTGT, nguồn (Bảng kê/HĐĐT/Chung), có công văn cảnh báo hay không, nhóm rủi ro (F0/F1/Không), trạng thái ĐKKD (đang hoạt động/bỏ địa chỉ/ngừng kinh doanh/khác). Đây là bảng CHI TIẾT của `IN07_NCC_CanhBao` (IN07 giữ vai trò tổng hợp theo từng NCC; `KQ2_NCC_mua_vao` là khai triển theo từng hóa đơn của chính các NCC đó). Nếu cần dòng Tổng cộng: đặt Ở CUỐI bảng, có 1 dòng trống ngăn cách rõ, không chèn giữa dữ liệu. Theo skill `doi-chieu-ncc-mua-vao-rui-ro` (mục Bảng 2.2/2.3 đã gộp lại).

**`IN07_NCC_CanhBao`:** tiêu đề cột in đậm, chữ đứng (không nghiêng), Times New Roman - định dạng rõ ràng hơn bản cũ.

**`KQ2_Cong_no_Tien_mat` không được để trống khi IN08 (sổ kế toán) hoặc IN10 (chi tiền mặt) đã có dữ liệu** - chỉ các dòng phụ thuộc IN11 (sao kê ngân hàng) chưa trích mới ghi KHÔNG_ĐỦ_DỮ_LIỆU/"PDF - chưa kết luận" (Q7), phần còn lại (131/331, chi tiền mặt 111, cụm HĐ cùng ngày) vẫn phải chạy. Theo skill `phan-tich-cong-no-tien-mat`.

**`KQ2_511_632` đổi cấu trúc** (không còn liệt kê từng chứng từ 1 dòng/HĐ): dùng bảng tổng hợp theo năm tài chính "I. BẢNG TỔNG HỢP ĐỐI CHIẾU THEO NĂM TÀI CHÍNH" (cột Chỉ tiêu | Tổng cộng kỳ | từng năm | **Đánh giá tuân thủ & Rủi ro**) theo skill `doi-chieu-511-632` đã cập nhật.

**HD04 (tỷ trọng Top-5 NCC, P14 mặc định 50%) bắt buộc giữ** trên `KQ2_NCC_mua_vao`: vì sheet này là chi tiết hóa đơn, tính Top-5 bằng cách gộp theo MST người bán rồi chia cho tổng mua kỳ; ghi số vào ô tỷ trọng và cờ `Vượt ngưỡng` / `Không vượt` / `Chưa đủ nguồn` (Tong_hop_co đọc H5). Không dùng `LARGE` trên từng dòng HĐ.

**Nhận xét module (3.4, đã mở rộng):** ngoài 511-632 / GTGT / NCC, các sheet `KQ2_TNCN_LD2`, `KQ2_Kho_Gia_thanh`, `KQ2_Gia_thanh_va_Dinh_muc`, `KQ2_Cong_no_Tien_mat` cũng có khối nhận xét ngắn cuối sheet. Không ghi mức rủi ro / điểm / đề cương hồ sơ (việc đó ở KQ3).

## Cấm

Kết luận ở KQ2. Số 0 giả. Đổi phông. Xóa định dạng có điều kiện. Xuất Tong_hop_*, Doi_chieu_*, sheet PL, workbook trọn gói.

## Lệnh siêu ngắn

```
Phân tích GĐ2 <FOLDER> MST <MST> kỳ <KY> xuất KQ2.1 KQ2.2
```
