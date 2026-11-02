---
type: wiki
title: Playbook 3 giai đoạn phân tích hồ sơ thuế (v1.4 - P7)
updated: 2026-10-23
---

# Playbook 3 giai đoạn (Chuẩn v1.4 vòng 2)

Căn cứ: `references/Quy_trinh_3_giai_doan_va_mau_bieu_chot_v1.4.md` và Hướng dẫn thường trực trong skill `phan-tich-ho-so-thue`.

Mẫu: `thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/KQ1.1_Kiem_ke_CQT.xlsx` (và KQ1.2, KQ1.3, KQ2.1, KQ2.2, KQ3). Sao chép rồi điền. Không dựng lại cấu trúc.

**Phòng điều phối (6 vai):** `thu-ho-so-kiem-ke` → `chuyen-vien-du-lieu-dau-vao` → `phan-tich-vien-cqt` (KQ2.1) → `phan-tich-vien-so-cty` (KQ2.2) → `phan-tich-vien-so-sach-dn` (GĐ3) → `truong-phong-phan-tich-ho-so`. Workflow: `workflows/phan-tich-ho-so-thue-3-giai-doan.md` (7 bước).

## Lệnh 1: Giai đoạn 1

```
LỆNH 1 – Giai đoạn 1 cho hồ sơ <THƯ MỤC>, MST <MST>, kỳ kiểm tra <KỲ>, loại hình <LOẠI HÌNH>. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC.
1. Quét toàn bộ thư mục (cả tệp ngoài cấu trúc). Gán nguồn CQT hoặc công ty theo thư mục cấp 1. Xếp mỗi tệp vào một IN; tệp không xếp được ghi riêng ở IN00 với Mã IN "-".
2. Chuẩn hóa dữ liệu thô vào IN theo mẫu: nhóm CQT vào KQ1.1 (IN01, IN03, IN04, IN05 gồm QTT, IN14); nhóm công ty vào KQ1.2 (IN02 đủ mua+bán, IN08–IN13). Mỗi tệp có IN00 riêng. Chưa có tệp nhóm công ty thì vẫn xuất KQ1.2 với IN00 trống.
3. Điền KQ1_Kiem_ke (13 dòng IN), KQ1_Kha_nang_phan_tich (80 quy tắc), KQ1_De_nghi_bo_sung. Xác định mức hồ sơ A–E và tiến độ.
4. Báo mức hồ sơ, số quy tắc Chạy ngay / Chạy một phần / Bị chặn, tối đa 3 điều cần xin công ty.
Chưa phân tích rủi ro. Đầu ra: KQ1.1, KQ1.2, KQ1.3 (v1). Agent: thu-ho-so-kiem-ke (+ chuyen-vien-du-lieu-dau-vao cho IN01/IN02/IN05).
```

## Lệnh 2: Giai đoạn 2

```
LỆNH 2 – Giai đoạn 2 cho hồ sơ <THƯ MỤC>, MST <MST>, kỳ <KỲ>. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC. Đã có KQ1.1, KQ1.2, KQ1.3.
1. Sao KQ2.1 và KQ2.2 từ mẫu. Lập IN07 trong KQ2.1 từ danh sách đối tác của IN02 và NGUỒN CHỈ ĐỊNH; chưa có IN02 thì để trống.
2. KQ2.1 (agent phan-tich-vien-cqt): GTGT, NCC 100% (giữ HD04), TNCN, EBITDA, Tong_hop_co. KQ2.2 (agent phan-tich-vien-so-cty): 511-632 tổng + Chi_tiet, kho, giá thành, công nợ tiền mặt, Tong_hop_co.
3. Điền số liệu, chênh lệch, tỷ lệ, cờ máy. Quét 100% NCC và 100% NLĐ. Ô phụ thuộc dữ liệu công ty chưa có thì ĐỂ TRỐNG.
4. Được phép nhận xét module ngắn theo mục 3.4 skill trục. Cấm mức rủi ro / điểm / đề cương toàn hồ sơ trên KQ2.
Đầu ra: KQ2.1, KQ2.2 (v1).
```

## Lệnh 3: Giai đoạn 3

```
LỆNH 3 – Giai đoạn 3 cho hồ sơ <THƯ MỤC>, MST <MST>, kỳ <KỲ>. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC. Đã có KQ1.3, KQ2.1, KQ2.2. Agent: phan-tich-vien-so-sach-dn.
1. Sao KQ3 từ mẫu. Chạy 80 quy tắc trên sheet KQ3; điền sheet Danh_gia_28_tieu_chi_rui_ro (28 dòng narrative).
2. Dòng Có phát hiện: Nhãn, Mức, hệ số, điểm, sắc thuế, V; chống trùng T1–T10; cột W tính bằng mã.
3. Điền AB, AC, AH cho Cao/Trung bình; AI, AJ cho NCC/HĐ trọng điểm; AD, AK, AL, AM để trống cho cán bộ.
4. Xuất KQ3_Danh_gia_Dau_hieu_Bat_thuong_<MST>_<KY>.xlsx và KQ3_NNT_Giai_trinh_Chi_tiet_<MST>_<KY>.xlsx (10 sheet có số liệu thật). Áp Q14-Q20.
5. Báo: số dòng theo kết quả; tổng cột W theo sắc thuế; việc công ty còn thiếu (cột AA).
Đầu ra: KQ3 (2 sheet) + 2 file phụ GĐ3.
```

## Lệnh 4: Cập nhật khi công ty gửi thêm hồ sơ

```
LỆNH 4 – Cập nhật hồ sơ <THƯ MỤC>, MST <MST>, kỳ <KỲ>. Công ty đã gửi thêm. Áp dụng HƯỚNG DẪN THƯỜNG TRỰC.
1. Chạy lại Giai đoạn 1: so IN00 mới với IN00 cũ. Xuất KQ1.1, KQ1.2, KQ1.3 phiên bản kế tiếp.
2. Điền lại ĐÚNG các ô trống ở KQ2.1, KQ2.2 của quy tắc được mở khóa; bảng khác giữ nguyên.
3. KQ3 phiên bản kế tiếp: cùng Mã_PH; cột AE/AG; bản cũ giữ lịch sử. Cập nhật lại 2 file phụ GĐ3.
4. Báo: số quy tắc mở khóa, phát hiện mới, đổi mức, giải tỏa; việc còn thiếu.
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
