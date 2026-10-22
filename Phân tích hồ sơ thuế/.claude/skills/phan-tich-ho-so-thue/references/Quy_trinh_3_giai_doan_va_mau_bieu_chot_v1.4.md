# Quy trình kiểm tra hồ sơ thuế theo 3 giai đoạn và bộ mẫu biểu chốt (bản 1.4)

Ngày lập: 19/10/2026. Bản 1.4 chốt, dựng theo cấu trúc sáu tệp bạn gửi (KQ1.1, KQ1.2, KQ1.3, KQ2.1, KQ2.2, KQ3) và theo hai câu trả lời của bạn lúc 15:01 ngày 19/10/2026.

**Thay thế hoàn toàn** bản nháp 1.4 gửi buổi sáng, bản 1.3 (`Quy_trinh_3_giai_doan_va_mau_bieu_chot_v1.3.md`), ba tệp mẫu `KQ1_MAU_Ho_so`, `KQ2_MAU_Ket_qua`, `KQ3_MAU_Ket_qua`, sheet 09 của workbook 1.3, sheet phụ lục PL và hợp đồng 19 sheet. **Các tài liệu cũ không dùng để phân tích hồ sơ nữa.**

**Vẫn hiệu lực:** 13 quyết định K1–K13 ngày 15/10/2026 (`Prompt_bo_sung_P5_quyet_dinh_15-10-2026.md`, mục A). Prompt thực thi là `Prompt_P7_AI_chi_dung_bo_quy_trinh_1.4.md`.

**Chưa xác minh (ghi "tạm" trong tệp):** thân các SKILL.md; cột thật của sổ F0/F1/CA và tệp TMS; mã TT99 (Phụ lục IV); số liệu tham chiếu ngành; mốc trọng yếu P03, P04 và các ngưỡng đề xuất.

## 1. Quyết định chốt ngày 19/10/2026

| # | Quyết định | Hệ quả trong bộ mẫu |
|---|---|---|
| Q1 | **IN02 (bảng kê) lấy từ bảng kê do công ty gửi**, thuộc nhóm công ty | IN02 nằm ở KQ1.2. Quy tắc cần IN02 chưa chạy được khi chưa có dữ liệu công ty |
| Q2 | **Chưa có dữ liệu công ty thì vẫn chạy KQ2.1 và KQ2.2**, ô công ty để trống đến khi có lệnh cập nhật | Luật T10. Ô trống không phải số 0. Quy tắc tương ứng ở KQ3 ghi KHÔNG_ĐỦ_DỮ_LIỆU. Cập nhật điền lại đúng ô đó |
| Q3 | **IN07 (danh bạ đối tác, cảnh báo) xếp vào KQ2.1**, vì phải có IN02 mới lập được | KQ2.1 có sheet `IN07_NCC_CanhBao`. `KQ2_NCC_mua_vao` chỉ nhắc lại **một** thuộc tính từ IN07 là Nhóm F (F0 \| F1) bằng công thức tra cứu, và dùng ngày cấp MST để tính số tháng (HD05). Số công văn, trạng thái ĐKKD, ngày hiệu lực ở lại IN07 (T9) |
| Q4 | **IN12 (giao dịch liên kết) ở nhóm công ty** | IN12 nằm ở KQ1.2 |
| Q5 | **Gộp IN05 với IN06** | Một sheet `IN05_NLD_TNCN`, cột `Loai_dong` = `NLD` hoặc `QTT`, 29 cột |
| Q6 | **Tên KQ1.3** thay cho KQ1.0 | Tên tệp `KQ1.3_Danh_gia_ho_so.xlsx` |
| Q7 | **Quy tắc không có bảng ở KQ2 thì ghi thẳng vào KQ3** | Cột B của KQ3 ghi "Trực tiếp"; chứng cứ ở cột Y, AB |
| Q8 | **IN03, IN04 chỉ chứa bản của cơ quan thuế**; không xét bản BCTC, tờ khai do công ty gửi | Không có cột `Nguon_ban`, không có phép so hai bản. Bản công ty gửi (nếu có) không nạp vào IN03, IN04 |
| Q9 | **Ngưỡng TN17 = 10%** (dư Có TK 331 cuối kỳ / doanh thu thuần BCTC [01]) | Tham số mới **P23** ở sheet 05 và ô tham số của `KQ2_Cong_no_Tien_mat`; TN17 so với P23 bằng toán tử > |

Những điều tôi đã chọn thêm, để bạn biết: IN14 (TMS) ở KQ1.1; "13" ở nhóm CQT trong yêu cầu hiểu là IN14; IN08 thêm vào KQ1.2.

## 2. Năm nguyên tắc

1. **Một dữ kiện, một chỗ lưu.** IN lưu dữ liệu chuẩn hóa; KQ2 phân tích (số liệu); KQ3 kết luận. KQ1.3 chỉ đánh giá độ đủ.
2. **Mỗi giai đoạn có tệp đúng vai trò.** Không tệp nào làm việc của tệp khác.
3. **Kết luận chỉ ở KQ3.** KQ2.1 và KQ2.2 chỉ có số liệu, chênh lệch, tỷ lệ và cờ máy (`Khớp` / `Lệch` / `Chưa đủ nguồn`; `Vượt ngưỡng` / `Không vượt` / `Chưa đủ nguồn`). Không mức rủi ro, nhận xét, kiến nghị, tổng số tiền.
4. **Một phát hiện, một dòng, một mã** (Mã_PH). Số tiền tính vào tổng một lần (mục 7).
5. **Mẫu cũ chỉ là nguồn nhập**, không tính lại. Gọi mẫu cũ bằng tên tệp, không bằng số mẫu.

## 3. Cấu trúc sáu tệp

| Giai đoạn | Tệp | Sheet | Vai trò |
|---|---|---|---|
| **GĐ1** Kiểm kê, đánh giá tiến độ, chuẩn hóa dữ liệu thô thành đầu vào | `KQ1.1_Kiem_ke_CQT.xlsx` | `IN00_Manifest`, `IN01_HDDT`, `IN03_ToKhai_GTGT`, `IN04_BCTC_TNDN`, `IN05_NLD_TNCN` (gồm cả quyết toán TNCN), `IN14_TMS_Nghia_vu` | Đầu vào từ cơ quan thuế |
| | `KQ1.2_Kiem_ke_ho_so_NNT.xlsx` | `IN00_Manifest`, `IN02_BangKe`, `IN08_So_KeToan`, `IN09_Kho_GiaThanh`, `IN10_Chi_tien_mat`, `IN11_Sao_ke_NH`, `IN12_GDLK`, `IN13_So_511_632` | Đầu vào do công ty cung cấp |
| | `KQ1.3_Danh_gia_ho_so.xlsx` | `KQ1_Kiem_ke`, `KQ1_Kha_nang_phan_tich`, `KQ1_De_nghi_bo_sung` | Kiểm kê, đánh giá, ghi nhận chuẩn hóa. Cổng sang GĐ2 |
| **GĐ2** Phân tích hồ sơ | `KQ2.1_Phan_tich_CQT.xlsx` | `KQ2_Tong_hop_co`, `KQ2_Doi_chieu_GTGT`, `KQ2_NCC_mua_vao`, `KQ2_TNCN_LD2`, `IN07_NCC_CanhBao` | Phân tích trên dữ liệu CQT; phần cần IN02 để trống nếu chưa có |
| | `KQ2.2_Phan_tich_ho_so_cty.xlsx` | `KQ2_Tong_hop_co`, `KQ2_511_632`, `KQ2_Kho_Gia_thanh`, `KQ2_Gia_thanh_va_Dinh_muc`, `KQ2_Cong_no_Tien_mat` | Phân tích trên dữ liệu công ty; để trống nếu chưa có |
| **GĐ3** Kết luận, sai phạm, gợi ý kiểm tra trọng điểm | `KQ3_Ket_luan_danh_gia.xlsx` | `KQ3` (39 cột A–AM) | Sổ kết quả duy nhất, tích lũy |

Không còn sheet PL. Workbook chuẩn (sheet 00–09, BM01–BM06) ở lại nơi làm việc nội bộ, không giao theo từng hồ sơ.

**Nguồn dữ liệu và phụ thuộc công ty.** Nhóm CQT: IN01, IN03, IN04, IN05, IN14. Nhóm công ty: IN02, IN08–IN13. IN07 do cán bộ tra cứu (sổ F0/F1, công văn, danh bạ ĐKKD) nhưng chỉ lập được khi có danh sách đối tác từ IN02, nên đặt ở KQ2.1.

Trong 80 quy tắc: **50 cần dữ liệu công ty, 30 không cần** (cột T sheet 04 của workbook chuẩn; cột "Cần dữ liệu công ty" ở `KQ1_Kha_nang_phan_tich`).

## 4. Quy trình

```
Thư mục hồ sơ ─► GĐ1 ─► KQ1.1 (dữ liệu CQT)
                  │      KQ1.2 (dữ liệu công ty; có thể chưa có)
                  └────► KQ1.3 (kiểm kê, khả năng phân tích, đề nghị bổ sung)
                              │ cổng: mức hồ sơ
                              ▼
                  GĐ2 ─► KQ2.1  ─ ô công ty để trống nếu chưa có IN02, IN08…
                         KQ2.2  ─ ô trống nếu chưa có IN08–IN13
                              │
                              ▼
                  GĐ3 ─► KQ3 (80 quy tắc; quy tắc thiếu nguồn: KHÔNG_ĐỦ_DỮ_LIỆU)
                              ▲
   Công ty gửi thêm: chạy lại GĐ1, rồi cập nhật đúng ô, đúng quy tắc được mở khóa (4.5)
```

Giai đoạn là trạng thái dữ liệu. Không có dữ liệu công ty thì **không dừng**: KQ2.1 và KQ2.2 vẫn được tạo từ mẫu, ô công ty để trống (T10).

### 4.1 Giai đoạn 1: kiểm kê, đánh giá tiến độ, xử lý dữ liệu thô thành đầu vào

**Đầu vào.** Thư mục hồ sơ, MST, kỳ kiểm tra (K12; thiếu thì hỏi một câu), chế độ kế toán từng năm, loại hình.

**Bước 1a. Quét thư mục.**
1. Kỳ kiểm tra do cán bộ nêu, không đặt mặc định.
2. Quét toàn bộ thư mục, kể cả tệp ngoài cấu trúc. Nguồn gốc theo thư mục cấp 1 (K1): `1. DU LIEU THO NGANH THUE` = CQT; `2. HO SO NNT CUNG CAP` = công ty. Không sửa, không xóa tệp nguồn.
3. Xếp mỗi tệp vào một mẫu IN. Tệp không xếp được ghi dòng riêng ở IN00 với Mã IN "—". Tệp theo mẫu cũ chỉ là nguồn nhập (T8).

**Bước 1b. Chuẩn hóa thành IN, tách hai tệp.**

| Tệp | IN | Cách xử lý |
|---|---|---|
| KQ1.1 | IN01 | HĐĐT mua + bán trong một bảng, cột `Chieu`; cờ `Tinh_vao_tong` |
| | IN03 | Tờ khai GTGT từ XML, mỗi tờ khai một dòng; đánh dấu bản hiệu lực |
| | IN04 | BCTC và TNDN, bảng dọc; mã số theo chế độ kế toán từng năm (sheet 06); TT99 ghi "sơ bộ" đến khi đối chiếu Phụ lục IV (K7) |
| | IN05 | Một sheet cho cả NLĐ (05-1, 05-2) và quyết toán 05/QTT, cột `Loai_dong` |
| | IN14 | Lịch sử tuân thủ từ TMS, ánh xạ theo cột thật (K10) |
| KQ1.2 | IN02 | Bảng kê mua vào, bán ra: bỏ dòng tổng, MST chuẩn, khóa `Ký hiệu\|Số\|MST` |
| | IN08–IN13 | Sổ kế toán; kho, giá thành; chi tiền mặt; sao kê; giao dịch liên kết; sổ 511, 632 |

Mỗi tệp có `IN00_Manifest` của riêng nó, một dòng một tệp nguồn; **mã băm chỉ ghi ở IN00**. Dòng `File_nguon = VI_DU` là ví dụ, bị bỏ khi đếm. Không sửa số trong tệp gốc.

**Bước 1c. Kiểm kê và đánh giá (KQ1.3).**
1. `KQ1_Kiem_ke`: mỗi loại IN một dòng (13 dòng điền sẵn: IN01, 02, 03, 04, 05, 07, 08–14). Ghi kỳ phủ, kỳ thiếu, số dòng thực, bản bổ sung, mức, trạng thái chuẩn hóa, việc chuẩn hóa đã làm, số dòng loại bỏ.
2. `KQ1_Kha_nang_phan_tich`: 80 quy tắc, mỗi quy tắc một dòng (cột A–G điền sẵn). Điền khả năng chạy: `Chạy ngay` / `Chạy một phần` / `Bị chặn`, nguồn còn thiếu, gợi ý.
3. `KQ1_De_nghi_bo_sung`: suy ra từ sheet trên, xếp theo số quy tắc mở khóa. Nguồn CQT thiếu: cán bộ trích từ hệ thống. Nguồn công ty thiếu: NNT bổ sung.
4. Mức hồ sơ (cổng sang GĐ2):

| Mức | Có gì | Chạy được |
|---|---|---|
| A | Chỉ dữ liệu CQT (IN01, 03, 04, 05, 14) | KQ2.1 phần chỉ cần CQT; KQ2.2 để trống |
| B | A cộng bảng kê (IN02) | KQ2.1 đủ GTGT và NCC (sau khi lập IN07) |
| C | B cộng sổ kế toán, sổ 511/632 (IN08, IN13) | KQ2.1 đủ; KQ2.2 một phần |
| D | C cộng kho, giá thành (IN09) | KQ2.2 thêm nhóm KH |
| E | D cộng chi tiền mặt, sao kê (IN10, IN11) | KQ2.2 đầy đủ |

Mức không chặn việc dựng KQ2 và KQ3; mức chỉ cho biết quy tắc nào có số.

### 4.2 Giai đoạn 2: phân tích hồ sơ

**Mục tiêu.** Từ KQ1.1, KQ1.2, tạo bảng số liệu làm chứng cứ cho KQ3. Không kết luận.

**Bước.**
1. Đọc KQ1.3. Sao mẫu KQ2.1, KQ2.2 về hồ sơ.
2. Lập `IN07_NCC_CanhBao` trong KQ2.1 từ danh sách đối tác của IN02 và NGUỒN CHỈ ĐỊNH (mục 6A của bộ prompt cũ: sổ F0, F1, CA, công văn, danh bạ ĐKKD). Chưa có IN02 thì để trống.
3. Kiểm tổng khớp chi tiết (Q12). Ánh xạ năm BCTC với năm kê khai, năm MƠ HỒ chờ cán bộ xác nhận.
4. Điền các bảng: chỉ số liệu, chênh lệch, tỷ lệ và cờ máy theo tham số sheet 05. Ô phụ thuộc dữ liệu công ty chưa có thì **để trống** (T10).
5. Bảng quét tập (NCC, NLĐ): đưa 100% đối tượng vào bảng.
6. Nhãn "sơ bộ": TN03, TN04, TN19, TN22 (K13), mọi dòng dùng mã TT99 chưa xác minh (K7); HD07 là TREO.

| Tệp | Sheet | Nội dung | Quy tắc |
|---|---|---|---|
| KQ2.1 | `KQ2_Doi_chieu_GTGT` | PHẦN I tờ khai GTGT theo kỳ; PHẦN II bán ra (HĐĐT, bảng kê, tờ khai, TNDN, BCTC); PHẦN III mua vào, thêm TK 133 | GT01–GT06, GT10, GT13–GT15, TN01, TN02 |
| | `KQ2_NCC_mua_vao` | 100% nhà cung cấp: giá trị, tỷ trọng, nhóm F; cờ HD02, HD03 (P13), HD04 (P14), HD05 (12 tháng) | HD01–HD05 |
| | `KQ2_TNCN_LD2` | Tổng hợp QTT; đa nguồn; lịch sử từng người; tờ khai quý so QTT (TC22); lương so BCTC (TC23, P07) | LD01–LD04 |
| | `IN07_NCC_CanhBao` | Danh sách cảnh báo | đầu vào của HD |
| KQ2.2 | `KQ2_511_632` | Doanh thu 511 so giá vốn 632 theo khóa ghép (K4: theo skill doi-chieu-511-632) | DT01, DT02 |
| | `KQ2_Kho_Gia_thanh` | Nhập xuất tồn 152, 155; tồn kho so BCTC (K6; TC3; P05) | KH01, KH04, KH06 |
| | `KQ2_Gia_thanh_va_Dinh_muc` | Giá thành; giá vốn so [02] BCTC; biên lợi nhuận gộp (TC28; P18) | KH05, KH07, TN04 |
| | `KQ2_Cong_no_Tien_mat` | Công nợ 131/331 (TC20); danh mục tài khoản ngân hàng; chi tiền mặt, cụm hóa đơn vượt ngưỡng (P10, P11, P12; K5 dùng ≥) | GT16, TN17, TT01, TT02, TT03, TT05, TT06 |

Đã loại khỏi KQ2 các phần trùng: `KQ2_TNCN_LD2` Phần I so Phần IV (TC22) gộp một chỗ; `KQ2_Kho_Gia_thanh` bảng trên cùng trùng Phần 1–3 và tỷ trọng trùng sheet giá thành; bảy cột thuộc tính cảnh báo của `KQ2_NCC_mua_vao` trùng IN07 (nay chỉ còn Nhóm F, các thuộc tính khác ở lại IN07).

### 4.3 Quy tắc ghi thẳng vào KQ3 (không có bảng ở KQ2)

44 quy tắc, cột B ở KQ3 ghi `Trực tiếp`, chứng cứ ở cột Y và AB:
- DQ: DQ01, DQ02, DQ03, DQ04, DQ05, DQ06, DQ07, DQ08
- GT: GT07, GT08, GT09, GT11, GT12
- TN: TN03, TN05, TN06, TN07, TN08, TN09, TN10, TN11, TN12, TN13, TN14, TN15, TN16, TN18, TN19, TN20, TN21, TN22
- HD: HD06, HD07, HD08, HD09
- TT: TT04
- KH: KH02, KH03
- DT: DT03
- LD: LD05
- LK: LK01, LK02
- LS: LS01, LS02

### 4.4 Giai đoạn 3: kết luận, tổng hợp sai phạm, gợi ý kiểm tra trọng điểm

**Đầu vào.** KQ1.3, KQ2.1, KQ2.2, tham số sheet 05.

**Bước.**
1. Chạy 80 quy tắc: 36 trên bảng KQ2 (KQ2.1 21, KQ2.2 15), 44 trực tiếp. Mỗi quy tắc ít nhất một dòng trong KQ3, kể cả `Không phát hiện`, `Chưa kết luận`, `Không áp dụng`. Quy tắc thiếu nguồn: trạng thái KHÔNG_ĐỦ_DỮ_LIỆU, nêu IN còn thiếu ở cột AA.
2. Dòng `Có phát hiện`: Nhãn (Chắc / Sơ bộ), Mức rủi ro (NCC F0 và F1 đều Cao, K2), hệ số trọng yếu, điểm, sắc thuế, số tiền liên quan (V).
3. Chống trùng: Mã_PH, dòng chủ (T3), cột W (T2).
4. Cột Z: tham chiếu chứng cứ, ví dụ `KQ2.1!KQ2_NCC_mua_vao!A12:R12`; trống nếu `Trực tiếp`.
5. Cột AB, AC, AH cho dòng Cao và Trung bình; AI, AJ cho dòng NCC hoặc hóa đơn trọng điểm. Cột AD, AK, AL, AM để cán bộ.
6. Sắp xếp: Có phát hiện → Chưa kết luận → Không phát hiện → Không áp dụng; rồi điểm giảm dần.
7. Kiểm lại độc lập 10 phát hiện điểm cao nhất từ dữ liệu gốc.
8. Bản 28 tiêu chí chỉ xuất khi cán bộ yêu cầu, sinh từ KQ3 bằng cách chép giá trị (K3).

### 4.5 Công ty gửi thêm hồ sơ (lệnh cập nhật)

1. Chạy lại GĐ1: so IN00 để biết tệp mới, tệp đổi (mã băm khác), bản bổ sung. Xuất KQ1.1, KQ1.2, KQ1.3 phiên bản mới.
2. Xác định quy tắc bị ảnh hưởng bởi IN đổi (cột "Dữ liệu cần").
3. Điền lại **đúng các ô để trống** ở KQ2.1, KQ2.2 của quy tắc được mở khóa. Bảng khác giữ nguyên.
4. KQ3: cập nhật **cùng Mã_PH**; cột AE ghi `Nâng lên Chắc` / `Đổi kết luận` / `Đổi mức` / `Giải tỏa` / `Giữ nguyên` / `Mới`; cột AG ghi lý do và hồ sơ công ty làm thay đổi.
5. Bản cũ giữ làm lịch sử; tổng chỉ tính trên KQ3 mới nhất (T1).

### 4.6 Lệnh từng bước (module)

Các skill A0–B3 giữ tên và là module nội bộ. Một lệnh bước chỉ ghi vào đúng sheet của tệp đích và tạo đúng một phiên bản mới của tệp đó, **không tạo tệp chuyên đề riêng**. Quy tắc chưa chạy ghi CHƯA_CHẠY.

| Bước | Ghi vào |
|---|---|
| 0 (checklist) | `KQ1_Kiem_ke` |
| 0b, 1, 1b (tổng hợp đầu vào, bảng kê) | IN01, IN02, IN05 |
| 2 (GTGT–TNDN) | `KQ2_Doi_chieu_GTGT` |
| 3 (NCC mua vào) | `IN07`, `KQ2_NCC_mua_vao` |
| 4 (TNCN, đa nguồn) | `KQ2_TNCN_LD2` |
| 5 (511–632) | `KQ2_511_632` |
| 6, 7 (kho, giá thành) | `KQ2_Kho_Gia_thanh`, `KQ2_Gia_thanh_va_Dinh_muc` |
| 8 (tiền mặt, ngân hàng) | `KQ2_Cong_no_Tien_mat` |
| 9 (Master 28 tiêu chí) | bản trình bày sinh từ KQ3 |
| 10 (đề cương) | KQ3 cột AB–AD, AH |

## 5. Mẫu chốt của từng tệp

### 5.1 KQ1.1, KQ1.2

- Chỉ các sheet ở mục 3. Không chứa `KQ1_*`, không chứa sheet tham chiếu 00–09.
- `IN00_Manifest` mỗi tệp chỉ liệt kê IN của tệp đó (danh sách chọn bị giới hạn).
- `IN05_NLD_TNCN`: dòng `NLD` điền cột người lao động (Ho_ten, CCCD, MST_CN, Nguon_BK, Loai_thu_nhap, Tong_TNCT, Thue_da_KT…); dòng `QTT` điền cột quyết toán (Mau_TK, So_LD, Co_HDLD, Ca_nhan_cu_tru, TN_tinh_thue, Thue_con_phai_nop, CP_luong_BCTC…). Cột dùng chung: `Nam`, `Tong_TNCT`, `Thue_da_KT`, `File_nguon`.
- Dữ liệu ví dụ trong IN là chuỗi giả ("CÔNG TY A", MST 0000000001), đánh dấu `VI_DU`.

### 5.2 KQ1.3

| Sheet | Cột |
|---|---|
| `KQ1_Kiem_ke` (17) | STT, Mã IN, Loại hồ sơ, Tệp nguồn, Nguồn gốc, Định dạng, Kỳ phủ, Kỳ còn thiếu, Số dòng thực, Số dòng thiếu MST/ký hiệu, Có bản bổ sung, Mức, Lý do mức, Trạng thái chuẩn hóa, Việc chuẩn hóa đã làm, Số dòng loại bỏ, Ghi chú |
| `KQ1_Kha_nang_phan_tich` (11) | Mã quy tắc, Nhóm, Tên quy tắc, Dữ liệu cần (Mã IN), Cần dữ liệu công ty, Tệp phân tích, Sheet chứng cứ, Khả năng chạy, Tệp nguồn còn thiếu, IN còn thiếu, Gợi ý xử lý |
| `KQ1_De_nghi_bo_sung` (8) | Ưu tiên, Tên tài liệu/hồ sơ cần bổ sung, Mã IN, Tệp nguồn, Hướng bổ sung, Số quy tắc mở khóa, Danh sách quy tắc mở khóa, Căn cứ yêu cầu |

Đầu mỗi sheet có khối thông tin (MST, tên NNT, kỳ, mức hồ sơ A–E, tiến độ, phiên bản). BM01 không thuộc KQ1.3.

### 5.3 KQ2.1, KQ2.2

- Dòng 1 tiêu đề; dòng 2 MST, kỳ, giai đoạn 2, phiên bản, mức dữ liệu công ty; dòng 3 ghi chú: ô công ty để trống khi chưa có, không điền 0.
- Cột có tiêu đề xanh là cột nhập; cột có tiêu đề cam là công thức; ô vàng là tham số (P05, P07, P10, P11, P12, P18…).
- Không có nhận xét, mức rủi ro, kiến nghị, tổng số tiền toàn hồ sơ.
- **Sheet `KQ2_Tong_hop_co`** (đầu mỗi tệp KQ2) tổng hợp cờ máy của các sheet còn lại: mỗi dòng là một chỉ tiêu có cờ (ví dụ cờ khớp doanh thu bán ra, chênh lệch nhập xuất tồn, NCC thuộc nhóm F0/F1, vượt ngưỡng P05 hoặc P18), kèm quy tắc liên quan, số dòng đã đánh giá, **số dòng bất thường** (lệch, vượt ngưỡng, có cảnh báo), số dòng chưa đủ nguồn và cờ tổng hợp (`Có bất thường` / `Không` / `Chưa đủ nguồn`). KQ2.1 có 16 chỉ tiêu, KQ2.2 có 11. Đây là bảng đếm bằng công thức, không phải kết luận: không có mức rủi ro, không cộng vào KQ3 (T4). Dòng chưa có dữ liệu để trống, không tính là bất thường.
- Tham số P23 (ngưỡng TN17, 10%) đã điền ở `KQ2_Cong_no_Tien_mat` và sheet 05.

### 5.3a Quy ước hình thức (mọi mẫu biểu)

- **Phông chữ:** Times New Roman cho mọi ô, kể cả ô trống (phông mặc định của tệp), để dữ liệu gõ thêm cùng phông. Không đổi phông khi điền.
- **Tô màu ô khác biệt** bằng định dạng có điều kiện (tự đổi màu khi giá trị thay đổi): đỏ = lệch, vượt ngưỡng, có cảnh báo (F0, F1), lãi gộp ≤ 0, không có 632 hoặc 511, chênh lệch khác 0; cam = chưa đủ nguồn hoặc cần lưu ý (chỉ có ở một nguồn, sơ bộ, mỏng); xanh = khớp, không vượt, đủ.
- Ở KQ1.1, KQ1.2: cột `Tinh_vao_tong`, `La_ban_hieu_luc` = Không (cam); `Trang_thai_DC` (KHOP xanh, KHOP_LECH_TIEN cam, CHI_CO_BK và CHI_CO_HDDT đỏ); `Nhom_F` = F0, F1 (đỏ).
- Ở KQ1.3: Mức Thiếu, Chưa nhận (đỏ), Mỏng, PDF chưa trích (cam), Đủ (xanh); Khả năng chạy Bị chặn (đỏ), Chạy một phần (cam), Chạy ngay (xanh).
- Ở KQ3: Kết quả Có phát hiện (đỏ), Chưa kết luận (cam), Không phát hiện (xanh); Mức rủi ro Cao (đỏ), Trung bình (cam); Nhãn Sơ bộ (cam), Chắc (xanh); Trạng thái chạy KHÔNG_ĐỦ_DỮ_LIỆU, TREO (cam).
- Màu là cờ máy giúp nhìn nhanh, không thay cho kết luận ở KQ3.

### 5.4 KQ3: sổ kết quả duy nhất

Một sheet `KQ3` (39 cột A–AM; AG–AM thêm để gộp các sheet của mẫu KQ3 cũ, xem 5.5); dòng 1 tiêu đề; dòng 2–7 khối thông tin; dòng 8 dải nhóm cột; dòng 9 tiêu đề; dữ liệu từ dòng 10. Giá trị tĩnh. Một dòng là kết quả của một quy tắc trên một đối tượng.

| Cột | Tên | Cách điền |
|---|---|---|
| A | Mã_PH | `<Mã quy tắc>\|<Khóa đối tượng>\|<Kỳ>`, không đổi giữa các phiên bản |
| B | Tệp phân tích | `KQ2.1` / `KQ2.2` / `Trực tiếp` |
| C | Nhóm | DQ GT TN HD TT KH DT LD LK LS |
| D, E | Mã quy tắc, Tên | Sheet 04 |
| F | Mã TC | TC1–TC28 hoặc — |
| G, H | Đối tượng, Kỳ | |
| I–L | Chỉ tiêu A, Giá trị A, Chỉ tiêu B, Giá trị B | Quét tập không phát hiện: "Đã rà n" |
| M, N | Giá trị đo, Ngưỡng | Ngưỡng kèm mã tham số |
| O | Trạng thái chạy | ĐÃ_CHẠY / KHÔNG_ĐỦ_DỮ_LIỆU / TREO / CHƯA_TRIỂN_KHAI / KHÔNG_ÁP_DỤNG / CHƯA_CHẠY |
| P | Kết quả | Có phát hiện / Không phát hiện / Chưa kết luận / Không áp dụng |
| Q | Nhãn | Chắc / Sơ bộ |
| R | Mức rủi ro | Cao / Trung bình / Thấp / Thông tin |
| S, T | Hệ số trọng yếu, Điểm | S: 3 nếu ≥ P04, 2 nếu ≥ P03, 1 còn lại; T = hệ số mức (Cao 3, TB 2, Thấp 1, TT 0) × S; skill của bạn có công thức khác thì theo skill (K8) |
| U | Sắc thuế | GTGT / TNDN / TNCN / Khác |
| V | Số tiền liên quan | Trước chống trùng; không phải số truy thu |
| W | Số tiền tính vào tổng | 0 ≤ W ≤ V |
| X | Mã_PH chủ | Bắt buộc khi W < V |
| Y | Nguồn | Mã IN, tệp, sheet, dòng; CQT hoặc DN |
| Z | Tham chiếu chứng cứ | `tệp!sheet!ô`; trống nếu Trực tiếp |
| AA | Còn thiếu để kết luận | Bắt buộc khi Chưa kết luận |
| AB, AC | Nội dung kiểm tra, Chứng từ cần xem | Dòng Có phát hiện mức Cao, Trung bình |
| AD | Nhánh xử lý (cán bộ) | Để trống |
| AE | Thay đổi so với bản trước | Mới / Giữ nguyên / Nâng lên Chắc / Đổi kết luận / Đổi mức / Giải tỏa |
| AF | Ghi chú | |
| AG | Lý do thay đổi / hồ sơ làm thay đổi | Khi AE khác Mới và Giữ nguyên |
| AH | Câu hỏi chất vấn NNT | Dòng Cao và Trung bình |
| AI | Chuỗi bằng chứng 5 mắt xích | NCC, hóa đơn trọng điểm. Mỗi mắt xích: Có / Không / Không đủ dữ liệu (K11). M1 hóa đơn, M2 hợp đồng đơn hàng, M3 thanh toán qua NH, M4 nhập kho, M5 xuất kho tiêu hao |
| AJ | Độ vững chuỗi bằng chứng | Vững / Trung bình / Yếu / Chưa đủ dữ liệu |
| AK | Kết quả giải trình của NNT | Cán bộ điền, không tính tổng |
| AL | Số tiền đề xuất truy thu | Cán bộ điền; **không** cộng vào cột W |
| AM | Trạng thái xử lý | Chưa xử lý / Đang xử lý / Đã xử lý (cán bộ) |

Ràng buộc: đủ 80 quy tắc; Có phát hiện thì Q–U, Y bắt buộc; 0 ≤ W ≤ V; W < V thì X bắt buộc; mỗi Mã_PH một dòng.

### 5.5 Các sheet của mẫu KQ3 cũ đã về đâu

Bạn gửi lại mẫu KQ3 cũ (bản 1.1–1.2) gồm 10 sheet. Một sheet KQ3 mới gộp hết nội dung theo bảng sau:

| Sheet cũ | Về KQ3 |
|---|---|
| `KQ3_Phat_hien_tong` (80 quy tắc) | Từng dòng của KQ3, cột A–T; "Kiến nghị kiểm tra" là cột AB |
| `KQ2_Phat_hien` (xếp theo điểm) | Không còn là sheet; là bộ lọc Kết quả = Có phát hiện, sắp theo cột T |
| `KQ2_Chua_ket_luan` | Bộ lọc Kết quả = Chưa kết luận; điều còn thiếu ở cột AA |
| `KQ3_Con_thieu` | Cột AA, cộng `KQ1_De_nghi_bo_sung` (ở KQ1.3) |
| `KQ3_Thay_doi_so_voi_GD2` | Cột AE (loại thay đổi) và AG (lý do, hồ sơ làm thay đổi) |
| `KQ3_De_cuong` | Cột AB (nội dung kiểm tra), AC (chứng từ), AD (nhánh xử lý), AH (câu hỏi chất vấn); thứ tự ưu tiên là thứ tự sắp theo điểm |
| `KQ3_Master_28TC` | Không là sheet; cột F (Mã TC). Bản 28 tiêu chí sinh từ KQ3 khi cán bộ cần (K3) |
| `KQ3_Chuoi_bang_chung` | Cột AI, AJ; thuế ảnh hưởng ở cột U, V |
| `BM06_Trong_yeu_kien_nghi` | Cột E, H, V, AB, AC, AD, AK (giải trình), AL (đề xuất truy thu), AM (trạng thái). Công thức liên kết ngoài `[1]` bỏ |
| `Sheet1` rỗng | Bỏ |

Tổng đề xuất truy thu (cộng cột AL) ghi ở khối thông tin đầu sheet, tách khỏi tổng cột W.

## 6. Quy tắc nằm ở đâu

| Tệp | Sheet chứng cứ | Số QT | Quy tắc | Trong đó cần dữ liệu công ty |
|---|---|---|---|---|
| KQ2.1 | `KQ2_Doi_chieu_GTGT` | 12 | GT01, GT02, GT03, GT04, GT05, GT06, GT10, GT13, GT14, GT15, TN01, TN02 | 6 |
| KQ2.1 | `KQ2_NCC_mua_vao` | 5 | HD01, HD02, HD03, HD04, HD05 | 5 |
| KQ2.1 | `KQ2_TNCN_LD2` | 4 | LD01, LD02, LD03, LD04 | 1 |
| KQ2.2 | `KQ2_511_632` | 2 | DT01, DT02 | 2 |
| KQ2.2 | `KQ2_Kho_Gia_thanh` | 3 | KH01, KH04, KH06 | 3 |
| KQ2.2 | `KQ2_Gia_thanh_va_Dinh_muc` | 3 | TN04, KH05, KH07 | 3 |
| KQ2.2 | `KQ2_Cong_no_Tien_mat` | 7 | GT16, TN17, TT01, TT02, TT03, TT05, TT06 | 7 |
| Trực tiếp | — | 44 | DQ01, DQ02, DQ03, DQ04, DQ05, DQ06, DQ07, DQ08, GT07, GT08, GT09, GT11, GT12, TN03, TN05, TN06, TN07, TN08, TN09, TN10, TN11, TN12, TN13, TN14, TN15, TN16, TN18, TN19, TN20, TN21, TN22, HD06, HD07, HD08, HD09, TT04, KH02, KH03, DT03, LD05, LK01, LK02, LS01, LS02 | 23 |
| **Tổng** | | **80** | | **50** |


## 7. Luật không tính gộp

| Mã | Luật | Nội dung |
|---|---|---|
| T1 | Một tổng, một file | Tổng chỉ tính trong KQ3 mới nhất bằng cách cộng cột W theo sắc thuế. Không cộng số từ KQ2.1, KQ2.2 hay phiên bản KQ3 cũ. |
| T2 | Một hóa đơn, một sắc thuế, một lần tính | Đơn vị chống trùng là hóa đơn × sắc thuế (quy tắc không gắn hóa đơn thì đối tượng × kỳ × sắc thuế). (a) Cùng cấp: nhiều quy tắc trên cùng đối tượng thì dòng chủ theo T3 giữ W = V, các dòng khác W = 0 và ghi Mã_PH chủ. (b) Khác cấp: dòng chi tiết (hóa đơn, NCC) giữ W = V; dòng tổng hợp (kỳ, năm) chỉ giữ phần chưa được dòng chi tiết giải thích, W = M − tổng W của các dòng chi tiết thuộc kỳ đó, tối thiểu 0. Giá trị đầy đủ luôn còn ở V và cột I–M. Một hóa đơn ảnh hưởng GTGT và TNDN là hai sắc thuế khác nhau, mỗi sắc thuế tính một lần. |
| T3 | Chọn dòng chủ | Ưu tiên: (1) nhãn Chắc trước Sơ bộ; (2) cấp thấp hơn (hóa đơn, rồi NCC, rồi kỳ/năm); (3) điểm cao hơn; (4) nhóm theo thứ tự HD, GT, TT, TN, KH, DT, LD, LK, LS, DQ; (5) mã quy tắc nhỏ hơn. |
| T4 | KQ2 là chứng cứ, không cộng | Số trong KQ2.1 và KQ2.2 không cộng vào tổng, không đếm làm số phát hiện, không mang mức rủi ro hay nhận xét. |
| T5 | Một Mã_PH, một dòng | Mỗi Mã_PH chỉ một dòng trong một tệp KQ3. Chạy lại giữ nguyên Mã_PH, đổi cột AE; không thêm dòng cho cùng một phát hiện. |
| T6 | NCC F0 và F1 | Một NCC chỉ một dòng cho mỗi quy tắc, dù liên quan nhiều NCC F0. NCC vừa F0 vừa F1 ghi F0. Hóa đơn của NCC đó tính một lần theo T2. |
| T7 | BCTC chỉ lấy một cột | Chỉ lấy số cuối năm (B01) và số năm nay (B02) của đúng năm, đúng một nguồn chuẩn (CQT). Không cộng 'số đầu năm' hay 'năm trước'. Không xét bản BCTC, tờ khai do công ty gửi: IN03 và IN04 chỉ chứa bản cơ quan thuế (cán bộ chốt 19/10/2026). |
| T8 | Mẫu cũ là nguồn nhập, không là nguồn tính | Tệp theo mẫu cũ chỉ nạp vào IN một lần theo sheet 01b. Hai bản tổng hợp cùng chủ đề: nạp một, ghi bản kia là bản trùng; không cộng. |
| T9 | Một dữ kiện, một tệp | Mỗi dữ kiện nằm ở một tệp theo phân vai: tệp gốc → IN00 (mã băm, từng tệp) → KQ1.3 (mức đủ, từng loại IN) → KQ2 (số liệu) → KQ3 (kết luận). Nhắc lại giá trị ở tệp khác phải là tham chiếu hoặc tổng hợp suy ra, không nhập lại. |
| T10 | Ô công ty để trống, không điền 0 | Khi chưa có dữ liệu công ty, ô phụ thuộc IN02, IN08–IN13 để trống ở KQ2.1, KQ2.2; quy tắc tương ứng ở KQ3 ghi KHÔNG_ĐỦ_DỮ_LIỆU. Cập nhật điền lại đúng ô đó, không tạo bảng mới. |

## 8. Đối chiếu sáu tệp bạn gửi với bản mẫu mới

| Tệp, sheet hiện có | Xử lý |
|---|---|
| KQ1.1: ba sheet `KQ1_*` | Xóa (trùng nguyên văn KQ1.3) |
| KQ1.1: IN06 | Gộp vào IN05 (`Loai_dong`) |
| KQ1.2: sheet 00–09 và IN14 | Xóa (tham chiếu ở workbook chuẩn; IN14 ở KQ1.1) |
| KQ1.2: IN08 | Thêm |
| KQ1.2: 287 ô `#REF!` | Hết khi bỏ sheet tham chiếu |
| KQ1.3: BM01 (17 ô `#VALUE!`, liên kết ngoài), `Sheet1` rỗng | Xóa |
| KQ1.3: ba sheet `KQ1_*` | Dựng lại, thêm cột chuẩn hóa, khả năng chạy 80 quy tắc điền sẵn |
| KQ2.1: phần đánh giá rủi ro, nhận xét, cột mức rủi ro | Xóa; chuyển vào KQ3 |
| KQ2.1: `#DIV/0!` ở `KQ2_TNCN_LD2!H57`, hai `#N/A` | Hết khi dựng lại sheet |
| KQ2.1: số 0 giả ở vùng nhập GTGT | Thay bằng ô trống (T10) |
| KQ2.2: tiêu đề "KQ3_…", "Giai đoạn 3" | Đổi thành KQ2_, giai đoạn 2 |
| KQ2.2, KQ2.1: tên, MST công ty thật, tên cá nhân, nhận xét có số liệu thật | Xóa hết |
| KQ3: `Sheet1` rỗng | Dựng sheet `KQ3` mục 5.4 |

Các tệp gốc bạn gửi **không bị sửa**; sáu tệp mới là bản dựng lại.

## 9. Điểm còn mở

1. **P03, P04** và các ngưỡng đề xuất còn "tạm" (K8).
2. **Số liệu ngành** chưa có (K13).
3. **Cột thật của sổ F0/F1/CA và tệp TMS** chưa thấy (K2, K10).
4. **Thân các SKILL.md** chưa đọc: mọi chỗ "skill đang làm gì" chỉ dựa vào bản đồ skill.
