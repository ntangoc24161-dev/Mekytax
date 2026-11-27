---
name: Loại lương chủ DNTN và chủ TNHH MTV
description: "Xét loại lương chủ DNTN và chủ TNHH MTV kiêm GĐ; ghép TMS DIRNA/NAME_FIRS với 05-1; điền LOẠI LƯƠNG."
group: Tài chính
updated: '2026-11-27'
---
> **Hợp đồng đầu ra v1.4:** module nội bộ. Số liệu ghi `KQ2.1!KQ2_TNCN_LD2`. Kết luận mức rủi ro / số truy thu chỉ ở KQ3. Không giao tệp chuyên đề như sản phẩm sáu tệp. Sheet `LOẠI LƯƠNG` trên Bang tinh chỉ là listing; **chưa trừ** cột Số kiểm tra 642/641 khi cán bộ chưa chốt.

# Loại chi phí lương chủ DNTN và chủ TNHH MTV kiêm Giám đốc

Cùng một điều luật: không được trừ tiền lương, tiền công của **chủ doanh nghiệp tư nhân** và **chủ công ty TNHH MTV do một cá nhân làm chủ**. Không loại theo chức danh Giám đốc đơn lẻ.

Wiki: [[Nguyên tắc loại trừ chi phí tiền lương Giám đốc công ty TNHH MTV]].

Skill trục: `phan-tich-ho-so-thue` playbook **Bước 4b** (sau TNCN). Khác `tong-hop-tncn-da-nguon` (QTT, đa nguồn, TC22/TC23) và `giam-thue-tndn` (giảm thuế NĐ).

## Khi nào dùng

- "loại lương giám đốc", "chủ kiêm GĐ", "lương chủ DNTN", "TNHH MTV không trừ lương chủ".
- Loại hình trên GCN/TMS là **DNTN** hoặc **TNHH MTV do cá nhân làm chủ**.
- Đã có IN05 (05-1) và/hoặc lịch sử TMS (`DIRNA`, `NAME_FIRS`).

Không dùng: hộ kinh doanh (không phải TNDN DN); CTCP, TNHH hai thành viên trở lên, MTV do **tổ chức** làm chủ (ghi `KHÔNG_ÁP_DỤNG` điều này).

## Chuẩn bị

1. Loại hình: GCN/ĐKKD hoặc TMS. DNTN và TNHH MTV cá nhân mới mở skill này.
2. Lịch sử TMS (thư mục `1.3`): tách `DIRNA` (GĐ đăng ký) và `NAME_FIRS` (chủ / tên NNT) **từng năm**, kể cả ngày đổi giữa năm.
3. Phụ lục 05-1/TNCN XML (IN05): họ tên, MST CN, chức danh, số tiền từng năm.
4. Bảng lương, chấm công, HĐLĐ, BHXH (thư mục `2.4`) nếu có. Thiếu thì để trống, không điền 0.
5. Bang tinh (nếu hồ sơ giải thể/biên bản có sheet `LOẠI LƯƠNG`).

## Quy trình

### 1. Xác định loại hình

- DNTN: luôn xét lương **chủ DNTN**.
- TNHH MTV cá nhân: xét từng năm GĐ có đồng thời là chủ hay không.
- Khác: dừng, `KHÔNG_ÁP_DỤNG`.

### 2. Xếp trường hợp từng năm

Ghép 3 trục: chủ (`NAME_FIRS`) + GĐ TMS (`DIRNA`) + người lĩnh lương trên 05-1/bảng lương.

| Mã | Đối tượng | TNDN | TNCN |
|---|---|---|---|
| A1 | Chủ DNTN lĩnh lương (kể cả ghi chức danh GĐ/quản lý) | Loại toàn bộ | Không chịu TNCN từ TTLĐ (thu nhập chủ) |
| A2 | Người thuê điều hành DNTN, không phải chủ | Được trừ nếu đủ HĐLĐ, chấm công, bảng lương, chứng từ chi | Chịu TNCN như NLĐ |
| B1 | Chủ TNHH MTV cá nhân kiêm GĐ | Loại toàn bộ, dù có điều hành hay không | Không chịu TNCN từ TTLĐ |
| B2 | GĐ thuê TNHH MTV, không phải chủ | Được trừ nếu đủ 4 loại chứng từ như A2 | Chịu TNCN như NLĐ |
| B3 | TNHH MTV do tổ chức làm chủ | Được trừ nếu đủ quyết định/HĐLĐ và chứng từ | Chịu TNCN bình thường |

Kèm A1/B1 nếu có nguồn: phụ cấp, thưởng, BHXH/BHYT/BHTN phần DN đóng, thuê nhà và phúc lợi chi riêng cho chủ/chủ kiêm GĐ.

Căn cứ: tiết d điểm 2.6 khoản 2 Điều 6 TT 78/2014 (TT 96/2015); CV 727/TCT-CS 03/03/2015. Hiện hành: điểm h khoản 2 Điều 9 Luật TNDN 2025; điểm d khoản 8 Điều 10 NĐ 320/2025/NĐ-CP.

### 3. Lấy số

- Số loại A1/B1: dòng 05-1 (hoặc bảng lương nếu 05-1 trống) của **đúng người chủ**, đúng năm.
- Không loại nhân viên khác, không loại GĐ thuê (A2/B2) theo điều này.
- A2/B2 thiếu HĐLĐ: ghi rủi ro chứng từ riêng, **không gom** vào khoản loại chủ.
- Đổi chủ giữa năm: mặc định điền đủ số 05-1 năm đó, ghi chú ngày TMS; cắt pro-rata chỉ khi cán bộ chốt.
- GĐ TMS khác người lĩnh lương GĐ trên bảng lương: tin 05-1/bảng lương cho số tiền, tin TMS cho tư cách chủ/GĐ. Không lấy `DIRNA` làm người bị loại.

### 4. Ghi kết quả

GĐ2 (có IN05 + IN14, chưa có bảng lương): bảng phân loại năm + số A1/B1; nhãn sơ bộ với năm thiếu bảng lương.

GĐ3 (có `2.4`): hoàn thiện A2/B2 (HĐLĐ/chấm công/chứng từ chi); cộng BHXH/phụ cấp A1/B1 nếu tách được.

Ghi:

- `KQ2.1!KQ2_TNCN_LD2`: bảng năm / người / case / số loại / số không loại / nguồn. Nhận xét module (3.4). Không ghi mức rủi ro.
- KQ3: kết luận, số TNDN dự thảo, đề cương. Tính thuế trên LNTT sau cộng chi phí loại; năm lỗ thì chỉ truy thu phần LNTT dương sau cộng. Chưa phạt / chậm nộp nếu cán bộ chưa lệnh.
- Bang tinh sheet `LOẠI LƯƠNG`: listing năm, họ tên, MST CN, số tiền. **Không** trừ cột Số kiểm tra 642/641 cho đến khi cán bộ chốt.
- File Word/Excel dự thảo tại `0. BIEN BAN` hoặc `ket-qua/<TEN>_<MST>/` chỉ khi cán bộ yêu cầu lập biên bản.

Font Times New Roman, ngày `dd/MM/yyyy`, số Excel `#,##0`.

## Bẫy

- Loại mọi người mang chức danh Giám đốc. `DIRNA` của Kiến Thành Tân giữ Khánh Hằng đến 2024 trong khi lương GĐ 2017-2022 là Kim Ngân (Case B2).
- Coi GĐ thuê thiếu HĐLĐ là cùng một khoản với chủ kiêm GĐ.
- Điền 0 cho năm không có 05-1 hoặc không có bảng lương.
- Trừ 642 Số kiểm tra trước khi chốt.
- Cộng BHXH/phụ cấp khi không có chứng từ tách được.
- Áp cho hộ kinh doanh hoặc CTCP.
- Lẫn skill `giam-thue-tndn` (NĐ 114/92) với điều loại lương chủ.

## Kiểm chứng

- Mỗi năm có đủ 3 cột: chủ, GĐ TMS, người lĩnh lương.
- Tổng loại = cộng các dòng A1/B1; đối chiếu 05-1 gốc.
- Nhân viên khác và Case A2/B2 không nằm trong tổng loại chủ.
- Sheet `LOẠI LƯƠNG` đã điền thì 642 Số kiểm tra vẫn bằng kê khai, trừ khi cán bộ đã chốt trừ.
- Thiếu HĐLĐ/BHXH được ghi thiếu, không bịa.

## Lệnh siêu ngắn

```
Xét loại lương chủ DNTN/TNHH MTV tại <FOLDER> MST <MST> kỳ <KY>
theo skill loai-luong-chu-dntn-tnhh-mtv.
Ghi KQ2.1!KQ2_TNCN_LD2. Chưa trừ 642 Số kiểm tra.
```
