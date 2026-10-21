---
name: Giai đoạn 1 - Kiểm kê & Đánh giá hồ sơ thuế
description: "Kiểm kê IN00-IN14 thành KQ1.1 (CQT), KQ1.2 (công ty), KQ1.3 (mức A-E, 80 quy tắc, đề nghị bổ sung)."
group: Tài chính
updated: '2026-10-19'
---

# Giai đoạn 1 - Kiểm kê hồ sơ & Đánh giá mức độ dữ liệu (v1.4)

Áp dụng HƯỚNG DẪN THƯỜNG TRỰC trong skill `phan-tich-ho-so-thue` và quy trình 1.4. Sao mẫu rồi điền. Chưa phân tích rủi ro.

## Khi nào dùng

"kiểm kê hồ sơ", "chạy giai đoạn 1", "Lệnh 1", "mức A-E", "xuất KQ1.1".

## Đầu ra

Ba tệp, phiên bản `v<n>`:

- `KQ1.1_Kiem_ke_CQT_<MST>_<KY>_v<n>.xlsx`: IN00, IN01, IN03, IN04, IN05 (cột Loai_dong = NLD|QTT), IN14. Không có sheet KQ1_*. Không có IN07.
- `KQ1.2_Kiem_ke_ho_so_NNT_<MST>_<KY>_v<n>.xlsx`: IN00, IN02, IN08-IN13. IN00 chỉ liệt kê tệp nhóm công ty. Chưa có tệp công ty thì vẫn xuất, IN00 trống (trừ dòng hướng dẫn).
- `KQ1.3_Danh_gia_ho_so_<MST>_<KY>_v<n>.xlsx`: KQ1_Kiem_ke (13 dòng IN), KQ1_Kha_nang_phan_tich (80 quy tắc), KQ1_De_nghi_bo_sung.

Mẫu: `thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/KQ1.1_Kiem_ke_CQT.xlsx` (và KQ1.2, KQ1.3).

## Quy trình (Lệnh 1)

1. Quét toàn bộ thư mục (cả tệp ngoài cấu trúc). Gán nguồn CQT hoặc công ty theo thư mục cấp 1 (K1). Xếp mỗi tệp vào một IN; tệp không xếp được ghi riêng ở IN00 với Mã IN "-".
2. Chuẩn hóa vào IN: nhóm CQT vào KQ1.1; nhóm công ty vào KQ1.2. Mỗi tệp có IN00 riêng. Tệp mẫu cũ là nguồn nhập (T8), ghi ở IN00, không cộng.
3. Điền KQ1.3. Xác định mức A-E (cổng: mức cho biết quy tắc nào có số; không chặn dựng KQ2/KQ3).
4. Báo mức hồ sơ, số quy tắc Chạy ngay / Chạy một phần / Bị chặn, tối đa 3 điều cần xin công ty.

IN03 và IN04 chỉ bản CQT (K8/Q3). IN05 gộp QTT, không sheet IN06. IN07 không nằm ở KQ1.1.

### IN02_BangKe (KQ1.2) - bổ sung 19/10/2026 vòng 2

- Cột `Chieu` = `MUA` hoặc `BAN`. BẮT BUỘC lọc và điền CẢ HAI chiều từ bảng kê gốc, không chỉ mua vào. Thiếu lọc bán ra là lỗi (đã xảy ra ở TT_BIEN_HOA 19/10, đã sửa).
- Hai khối phân đoạn "I. BẢNG KÊ MUA VÀO" và "II. BẢNG KÊ BÁN RA": chữ ĐỨNG (không in nghiêng), in đậm, có thể tô nền nhạt để tách khối rõ hơn thay cho in nghiêng.
- Dòng tiêu đề cột: in đậm, chữ đứng, Times New Roman (Q13/Q16).
- Cột `Ngay_HD`: `number_format = "dd/mm/yyyy"`, không giờ (Q14) - sửa lỗi cũ `yyyy-mm-dd h:mm:ss`.

### IN05_NLD_TNCN (KQ1.1) - bổ sung 19/10/2026 vòng 2

- Nhóm theo năm (`Nam`); trong MỖI năm, dòng `Loai_dong = QTT` (quyết toán tổng năm đó) đứng TRƯỚC các dòng `Loai_dong = NLD` (từng người lao động) của chính năm đó. Thứ tự đúng: QTT 2022, NLD 2022 (nhiều dòng), QTT 2023, NLD 2023 (nhiều dòng)... Không xếp tất cả QTT của mọi năm lên đầu rồi mới tới NLD (lỗi cũ ở TT_BIEN_HOA 19/10: QTT 2022-2025 dồn đầu, NLD xếp sau cùng).

## Cổng dừng

- Folder không đọc được: dừng.
- Thiếu MST hoặc kỳ: hỏi một câu (K12).
- Cấm tính chênh lệch rủi ro, cấm xóa nguồn, cấm số 0 giả, cấm đổi phông Times New Roman.

## Lệnh siêu ngắn

```
Kiểm kê GĐ1 hồ sơ <FOLDER> MST <MST> kỳ <KY> xuất KQ1.1 KQ1.2 KQ1.3
```
