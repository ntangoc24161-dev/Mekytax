---
type: wiki
title: Playbook 3 giai đoạn phân tích hồ sơ thuế (v1.4 - P7)
updated: 2026-10-19
---

# Playbook 3 giai đoạn (Chuẩn v1.4)

Căn cứ: `references/Quy_trinh_3_giai_doan_va_mau_bieu_chot_v1.4.md` và `references/Prompt_P7_AI_chi_dung_bo_quy_trinh_1.4.md`. Thay P0-P6. Hướng dẫn thường trực nằm trong skill `phan-tich-ho-so-thue`.

Mẫu: `thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/KQ1.1_Kiem_ke_CQT.xlsx` (và KQ1.2, KQ1.3, KQ2.1, KQ2.2, KQ3). Sao chép rồi điền. Không dựng lại cấu trúc.

## Lệnh 1: Giai đoạn 1

```
LỆNH 1 – Giai đoạn 1 cho hồ sơ <THƯ MỤC>, MST <MST>, kỳ kiểm tra <KỲ>, loại hình <LOẠI HÌNH>. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC.
1. Quét toàn bộ thư mục (cả tệp ngoài cấu trúc). Gán nguồn CQT hoặc công ty theo thư mục cấp 1. Xếp mỗi tệp vào một IN; tệp không xếp được ghi riêng ở IN00 với Mã IN "-".
2. Chuẩn hóa dữ liệu thô vào IN theo mẫu: nhóm CQT vào KQ1.1 (IN01, IN03, IN04, IN05 gồm QTT, IN14); nhóm công ty vào KQ1.2 (IN02, IN08–IN13). Mỗi tệp có IN00 riêng, chỉ liệt kê tệp của nhóm mình. Chưa có tệp nhóm công ty thì vẫn xuất KQ1.2 với IN00 trống.
3. Điền KQ1_Kiem_ke (13 dòng IN), KQ1_Kha_nang_phan_tich (80 quy tắc), KQ1_De_nghi_bo_sung (xếp theo số quy tắc mở khóa). Xác định mức hồ sơ A–E và tiến độ.
4. Báo mức hồ sơ, số quy tắc Chạy ngay / Chạy một phần / Bị chặn, tối đa 3 điều cần xin công ty.
Chưa phân tích rủi ro. Đầu ra: KQ1.1, KQ1.2, KQ1.3 (v1).
```

## Lệnh 2: Giai đoạn 2

```
LỆNH 2 – Giai đoạn 2 cho hồ sơ <THƯ MỤC>, MST <MST>, kỳ <KỲ>. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC. Đã có KQ1.1, KQ1.2, KQ1.3.
1. Sao KQ2.1 và KQ2.2 từ mẫu. Lập IN07 trong KQ2.1 từ danh sách đối tác của IN02 và NGUỒN CHỈ ĐỊNH; chưa có IN02 thì để trống.
2. Điền các bảng chỉ bằng số liệu, chênh lệch, tỷ lệ, cờ máy. Quét 100% NCC và 100% người lao động. Ô phụ thuộc dữ liệu công ty chưa có thì ĐỂ TRỐNG.
3. Ánh xạ năm BCTC với năm kê khai kèm độ tin cậy. Kiểm tổng khớp chi tiết.
4. Không ghi mức rủi ro, nhận xét, kiến nghị, tổng tiền vào KQ2.
Đầu ra: KQ2.1, KQ2.2 (v1).
```

## Lệnh 3: Giai đoạn 3

```
LỆNH 3 – Giai đoạn 3 cho hồ sơ <THƯ MỤC>, MST <MST>, kỳ <KỲ>. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC. Đã có KQ1.3, KQ2.1, KQ2.2.
1. Sao KQ3 từ mẫu. Chạy 80 quy tắc: quy tắc có bảng ở KQ2 lấy số từ bảng đó và ghi tham chiếu ở cột Z; quy tắc "Trực tiếp" tính từ IN và ghi nguồn ở cột Y, AB. Mỗi quy tắc ít nhất một dòng, kể cả Không phát hiện, Chưa kết luận, Không áp dụng.
2. Dòng Có phát hiện: Nhãn, Mức, hệ số trọng yếu, điểm, sắc thuế, V; chống trùng bằng Mã_PH, dòng chủ, cột W theo T1–T10; cột W tính bằng mã.
3. Điền AB, AC, AH cho dòng Cao và Trung bình; AI, AJ (chuỗi bằng chứng 5 mắt xích) cho NCC, hóa đơn trọng điểm, mắt xích thiếu nguồn ghi "Không đủ dữ liệu" (K11); AD, AK, AL, AM để trống cho cán bộ; AL không cộng vào cột W. Sắp xếp theo mục 5.4 quy trình. Kiểm lại độc lập 10 phát hiện điểm cao nhất.
4. Báo: số dòng theo kết quả; tổng cột W theo sắc thuế; việc công ty còn thiếu (cột AA).
Đầu ra: KQ3 (v1).
```

## Lệnh 4: Cập nhật khi công ty gửi thêm hồ sơ

```
LỆNH 4 – Cập nhật hồ sơ <THƯ MỤC>, MST <MST>, kỳ <KỲ>. Công ty đã gửi thêm. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC.
1. Chạy lại Giai đoạn 1: so IN00 mới với IN00 cũ (tệp mới, tệp đổi mã băm, bản bổ sung). Xuất KQ1.1, KQ1.2, KQ1.3 phiên bản kế tiếp.
2. Xác định quy tắc bị ảnh hưởng. Điền lại ĐÚNG các ô trống ở KQ2.1, KQ2.2 của quy tắc được mở khóa; bảng khác giữ nguyên.
3. KQ3 phiên bản kế tiếp: cùng Mã_PH; cột AE ghi Mới / Giữ nguyên / Nâng lên Chắc / Đổi kết luận / Đổi mức / Giải tỏa; cột AG ghi lý do và hồ sơ công ty làm thay đổi. Bản cũ giữ làm lịch sử, không cộng.
4. Báo: số quy tắc được mở khóa, số phát hiện mới, đổi mức, giải tỏa; việc còn thiếu.
```

## Thiếu dữ liệu thì làm gì

| Nguồn thiếu | Hành động |
|---|---|
| IN02, IN08-IN13 | Vẫn tạo KQ2.1, KQ2.2, KQ3. Ô công ty trống. Quy tắc: KHÔNG_ĐỦ_DỮ_LIỆU / Chưa kết luận, cột AA nêu IN. |
| IN07 / sổ F0 F1 | Q11: không tự tìm nguồn khác. HD01-HD05, HD08: KHÔNG_ĐỦ_DỮ_LIỆU, báo cán bộ. |
| IN14 TMS | LS01, LS02: KHÔNG_ĐỦ_DỮ_LIỆU (K10). |
| Sao kê PDF chưa trích | `PDF - chưa kết luận` (K11). Không kết luận thiếu UNC. |
| Số liệu ngành | HD07 TREO; TN03, TN04, TN19, TN22 sơ bộ - chưa so ngành (K13). |

Mức A-E là cổng cho biết quy tắc nào có số, không chặn dựng KQ2/KQ3.
