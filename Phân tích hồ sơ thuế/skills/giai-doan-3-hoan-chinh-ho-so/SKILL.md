---
name: Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN
description: "KQ3 một sheet 39 cột, 80 quy tắc, chống trùng T1-T10; kết luận, mức, điểm, đề cương, chuỗi bằng chứng."
group: Tài chính
updated: '2026-10-19'
---

# Giai đoạn 3 - Kết luận đánh giá (v1.4)

Áp dụng HƯỚNG DẪN THƯỜNG TRỰC trong skill `phan-tich-ho-so-thue`. Sao mẫu KQ3 rồi điền một sheet 39 cột A-AM. Mọi kết luận chỉ ở đây.

Tên skill giữ "sổ sách DN" vì slug đã khóa; GĐ3 chạy cả khi chưa có sổ: quy tắc thiếu nguồn ghi KHÔNG_ĐỦ_DỮ_LIỆU / Chưa kết luận, không dừng.

## Khi nào dùng

Sau KQ1.3, KQ2.1, KQ2.2. "phân tích giai đoạn 3", "Lệnh 3", "xuất KQ3", "đề cương", "chuỗi bằng chứng". Không tách sheet KQ3_De_cuong, KQ3_Master_28TC, KQ3_Chuoi_bang_chung, BM06.

## Đầu ra

`KQ3_Ket_luan_danh_gia_<MST>_<KY>_v<n>.xlsx` - hai sheet:

- `KQ3`: 39 cột A-AM. Dữ liệu từ dòng 10. Đủ 80 quy tắc, mỗi quy tắc ít nhất một dòng.
- `Danh_gia_28_tieu_chi_rui_ro` (bổ sung lại 19/10/2026 vòng 2, theo mẫu 06 cũ): 7 cột - `ID`, `Tên tiêu chí`, `Phân loại`, `Ngưỡng rủi ro`, `Kết quả phân tích trên hồ sơ`, `Mức độ rủi ro`, `Kiến nghị kiểm tra`. 28 dòng, narrative (câu chữ đánh giá, không phải bảng mã hóa như sheet `KQ3`). Đây là lớp tổng hợp DỄ ĐỌC cho cán bộ, độc lập với bảng 80 quy tắc - không cần ánh xạ 1-1 giữa 28 tiêu chí và 80 quy tắc, chỉ cần nội dung đúng với kết luận đã có ở `KQ3`.

36 quy tắc lấy số từ bảng KQ2 (KQ2.1 = 21, KQ2.2 = 15), cột Z tham chiếu `tệp!sheet!ô`. 44 quy tắc Trực tiếp: cột B = `Trực tiếp`, cột Z trống, chứng cứ cột Y và AB.

## Quy trình (Lệnh 3)

1. Sao KQ3 từ mẫu. Chạy 80 quy tắc. Thiếu nguồn: O = KHÔNG_ĐỦ_DỮ_LIỆU, P = Chưa kết luận, AA nêu IN còn thiếu.
2. Dòng Có phát hiện: Nhãn, Mức (F0/F1 = Cao, K2), hệ số, điểm, sắc thuế, V. Chống trùng Mã_PH, dòng chủ, cột W theo T1-T10. Cột W tính bằng mã. 0 ≤ W ≤ V; W < V thì X bắt buộc.
3. Điền AB, AC, AH cho dòng Cao và Trung bình; AI, AJ cho NCC/HĐ trọng điểm (mắt xích thiếu = "Không đủ dữ liệu", K11). AD, AK, AL, AM để trống cho cán bộ. AL không cộng vào W.
4. Sắp xếp mục 5.4 quy trình. Kiểm lại độc lập 10 phát hiện điểm cao nhất.
5. Báo số dòng theo kết quả; tổng cột W theo sắc thuế; việc công ty còn thiếu (cột AA).
6. Điền `Danh_gia_28_tieu_chi_rui_ro`: với mỗi tiêu chí, viết "Kết quả phân tích trên hồ sơ" dựa trên các dòng KQ3 liên quan (không suy diễn ngoài số đã có ở KQ3), gán `Mức độ rủi ro` và `Kiến nghị kiểm tra` nhất quán với kết luận đã chốt ở KQ3 cho các quy tắc thuộc tiêu chí đó. Thiếu dữ liệu: ghi "chưa đủ dữ liệu để đánh giá", không tự đoán mức rủi ro.

Master 28 tiêu chí chỉ sinh từ KQ3 khi cán bộ yêu cầu (K3).

## Lệnh 4 (công ty gửi thêm)

Cùng Mã_PH; cột AE = Mới / Giữ nguyên / Nâng lên Chắc / Đổi kết luận / Đổi mức / Giải tỏa; cột AG lý do. Bản cũ giữ lịch sử. Chỉ điền ô trống của quy tắc mở khóa ở KQ2.

## Cấm

Tự gắn nhãn trốn thuế. Số 0 giả. Ghi đè bản cũ. Tách sheet KQ3 cũ. Cộng số từ KQ2 vào tổng.

## Lệnh siêu ngắn

```
Hoàn chỉnh GĐ3 hồ sơ <FOLDER> MST <MST> kỳ <KY> xuất KQ3
```
