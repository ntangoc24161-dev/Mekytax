---
type: wiki
title: Playbook 3 giai đoạn phân tích hồ sơ thuế (v1.2 - 80 quy tắc & Quyết định K1-K13)
updated: 2026-10-15
---

# Playbook 3 giai đoạn phân tích hồ sơ thuế (Chuẩn v1.2)

Căn cứ: Bộ biểu mẫu v1.2 (`Bo_bieu_mau_chuan_phan_tich_ho_so_thue_v2.xlsx`, 80 quy tắc, 15 mẫu nhập IN00–IN14, 6 mẫu BM), Bản đặc tả ứng dụng, Bộ prompt P0–P4 và **Prompt bổ sung P5 (bản 2) ngày 15/10/2026**.

---

## 1. Kiến trúc năm lớp prompt (P0 – P5)

| Lớp | Dán khi nào | Làm gì |
|---|---|---|
| **P0** | Một lần, và mỗi khi đổi bộ biểu mẫu | AI đọc, hiểu bộ biểu mẫu, cập nhật năng lực skill và playbook. |
| **Q (Khối quy tắc chung)** | Lưu vào skill, hoặc dán đầu mọi lệnh | 11 nguyên tắc bất biến + 13 Quyết định K1–K13 của cán bộ kiểm tra. |
| **P1, P2, P3** | Mỗi hồ sơ, theo trạng thái dữ liệu (mức A–E) | Lệnh phân tích cho từng giai đoạn (Kiểm kê → CQT → Sổ DN). |
| **P4** | Khi NNT gửi thêm tài liệu bổ sung | Chạy lại phần mới mở khóa, xuất phiên bản mới (v+1), không ghi đè. |
| **P5** | Ghi nhận quyết định cán bộ & chế độ TT99 | Áp dụng chồng lên P0–P4: K1–K13, nạp bản 1.2, đối chiếu Phụ lục IV TT99. |

> **Nguyên tắc cốt lõi:** Giai đoạn là **trạng thái của dữ liệu**, không phải thứ tự thời gian. P1 chạy ở mọi lần tiếp nhận hồ sơ. P2 chạy khi chưa có sổ sách của công ty. P3 chạy khi công ty đã nộp sổ sách (kể cả nộp thiếu).

---

## 2. Khối Q: Quy tắc chung & Quyết định cán bộ K1–K13

```text
KHỐI QUY TẮC CHUNG (Q) - PHIÊN BẢN 1.2
1. Giai đoạn do dữ liệu quyết định. Chạy mọi quy tắc mà nguồn đã có; quy tắc thiếu nguồn ghi KHÔNG_ĐỦ_DỮ_LIỆU hoặc TREO kèm lý do. Không điền 0 khi thiếu nguồn.
2. Mỗi quy tắc có đúng một trạng thái: ĐÃ_CHẠY, KHÔNG_ĐỦ_DỮ_LIỆU, TREO, CHƯA_TRIỂN_KHAI, KHÔNG_ÁP_DỤNG (theo loại hình doanh nghiệp). Không để quy tắc không trạng thái.
3. Mỗi nguồn gắn nguồn gốc: CQT (dữ liệu cơ quan thuế) hoặc DN (công ty gửi). Khi có cả hai bản, bản CQT là chuẩn so sánh; chênh lệch là phát hiện nhóm DQ.
4. Mọi phát hiện có: mã quy tắc, giá trị đo, ngưỡng và mã tham số, mức rủi ro, hệ số trọng yếu, điểm, nhãn "chắc"/"sơ bộ", ô nguồn dạng chữ (file, sheet, dòng).
5. Phép tính do mã hoặc công thức tất định thực hiện. Mô hình chỉ đọc file lạ, ánh xạ vào IN, và viết diễn giải trên số đã tính.
6. Không sửa, không xóa file nguồn. Mọi file nhận vào ghi tên, ngày nhận, mã băm SHA-256. Chỉ cán bộ quyết định lưu trữ hay xóa.
7. Không kết luận vượt nguồn. Sao kê PDF chưa trích dòng thì ghi "PDF – chưa kết luận", không kết luận thiếu UNC. Năm BCTC chỉ gợi ý kèm độ tin cậy, cán bộ xác nhận.
8. Giá trị chưa xác minh giữ nhãn "chưa xác minh" trong kết quả.
9. Chạy lại là tăng dần: không ghi đè kết quả cũ, tăng số phiên bản (v1, v2…), và lập sheet thay đổi so với lần trước.
10. Đầu ra chỉ có sheet kết quả theo hợp đồng đầu ra (KQ1, KQ2, KQ3 mang giá trị tĩnh). Cấm xuất sheet hướng dẫn, tham chiếu, tham số, ánh xạ, danh mục quy tắc và sheet dữ liệu đầu vào.
11. Trước khi giao file, tự kiểm: mọi quy tắc có trạng thái; không có 0 giả; tổng khớp chi tiết; mọi phát hiện có ô nguồn; tên sheet đúng hợp đồng; 10 phát hiện điểm cao nhất đã kiểm lại từ dữ liệu gốc.

QUYẾT ĐỊNH CỦA CÁN BỘ (K1–K13 - NGÀY 15/10/2026):
K1. Nguồn gốc dữ liệu theo thư mục cấp 1: "1. DU LIEU THO NGANH THUE" = CQT; "2. HO SO NNT CUNG CAP" = DN. Thư mục con: 1.1 = tờ khai XML; 1.3 = Nghĩa vụ nộp thuế & Thông báo TMS; 2.5 = bảng kê; 3.3 = tra cứu đa nguồn; 4 = kết quả.
K2. Nhóm F: F0 = cảnh báo rủi ro hoặc danh sách Công an (CA); F1 = mua hàng từ F0. Mức rủi ro: NCC thuộc F0 xếp CAO; NCC thuộc F1 (tầng hai) cũng xếp CAO. NNT đang kiểm tra thuộc F1 là phát hiện Cao của HD01. IN07 có cột Nhom_F và Nguon_F0; BM04 có cột S (Nhóm F) và T (Mức rủi ro).
K3. Master 8 sheet: Nếu không tìm thấy mẫu thì bỏ qua, CHỈ DỰNG SHEET 28 TIÊU CHÍ (sheet 08). Không tự dựng 7 sheet còn lại, không báo như lỗi.
K4. Đối chiếu 511–632: Do skill doi-chieu-511-632 thực hiện. Khóa ghép và cột xuất lấy từ skill, thay khóa giả định.
K5. Ngưỡng thanh toán không dùng tiền mặt: Chốt dùng dấu ≥ ở cả hai giai đoạn: ≥ 20.000.000 đ đến 30/06/2025 và ≥ 5.000.000 đ từ 01/07/2025. Sửa mọi chỗ ghi ">20 triệu" thành ">= 20 triệu".
K6. Hàng tồn kho theo mẫu BCTC: B01-DN (TT200, TT99) dùng [141] = 152+155+156. B01-DNN (TT133) dùng ct140. ct141 = 0 trên mẫu DNN không có nghĩa HTK = 0.
K7. Chế độ kế toán Thông tư 99/2025/TT-BTC: Hiệu lực 01/01/2026, thay thế TT200/2014. Thông tư 133/2016 KHÔNG BỊ THAY THẾ, vẫn tiếp tục áp dụng cho DNNVV. Đối chiếu Phụ lục IV: DT_CHUA_THUC_HIEN (Doanh thu chờ phân bổ ngắn hạn) mang mã số 319; VON_CSH (Vốn góp chủ sở hữu) mang mã số 411. Chỉ tiêu vắng mặt trong BCTC TT99 là "vắng mặt - không có số liệu", không phải thiếu nguồn (Điều 17 khoản 4).
K8. Mốc trọng yếu: Thứ tự lấy giá trị: (1) Cán bộ chỉ định trong lệnh, (2) Giá trị trong skill AI, (3) Mặc định sheet 05 gắn nhãn "tạm".
K9. Thuế suất GTGT giảm: Giữ bảng đề xuất sheet 05 kèm số NQ 174/2024, NQ 204/2025.
K10. Lịch sử tuân thủ TMS: Đọc file Excel từ "1. DU LIEU THO NGANH THUE/1.3. Nghia vu nop thue & Thong bao" nạp vào IN14_TMS_Nghia_vu. LS01 và LS02 đọc IN14. Chưa có file thì ghi KHÔNG_ĐỦ_DỮ_LIỆU.
K11. Chứng từ công ty (sao kê, UNC, hợp đồng): Thiếu không phải là phát hiện; mắt xích thiếu ghi "Không đủ dữ liệu", quy tắc ghi KHÔNG_ĐỦ_DỮ_LIỆU.
K12. Phạm vi kỳ kiểm tra: Do cán bộ chỉ định từng hồ sơ tại sheet 00 mục A. Không có mặc định.
K13. Số liệu tham chiếu ngành: Chưa có dữ liệu: HD07 ghi TREO; TN03, TN04, TN19, TN22 chạy sơ bộ – chưa so ngành. Khung bảng ở sheet 05 để trống, không điền ước lượng.
```

---

## 3. Phân chia 80 quy tắc theo giai đoạn (Chuẩn v1.2)

| Giai đoạn | Số QT | Danh sách quy tắc | Ghi chú nguồn dữ liệu |
|---|---|---|---|
| **Giai đoạn 2** (Chỉ dữ liệu CQT) | **57** | `DQ02`–`DQ08`<br>`GT01`–`GT05`, `GT07`–`GT15`<br>`TN01`–`TN03`, `TN05`–`TN16`, `TN19`–`TN22`<br>`HD01`–`HD09`<br>`TT01`<br>`LD01`, `LD03`, `LD04`<br>`LK01`, `LK02`<br>`LS01`, `LS02` | Cần HĐĐT (IN01), Bảng kê (IN02), Tờ khai GTGT (IN03), QTT TNDN & BCTC (IN04), NLĐ & QTT TNCN (IN05, IN06), Cảnh báo rủi ro (IN07), GDLK (IN12), Lịch sử TMS (IN14).<br>- `HD01`, `HD03`, `HD08`: F0 và F1 đều xếp mức **Cao** (K2).<br>- `TT01`: Ngưỡng $\ge 20$ triệu và $\ge 5$ triệu (K5).<br>- `LS01`, `LS02`: Đọc IN14 từ thư mục 1.3 (K10).<br>- `HD07`: Ghi **TREO** do thiếu tham chiếu ngành (K13).<br>- BCTC: Quy tắc so BCTC mang nhãn **"sơ bộ"** khi chưa có sổ DN. |
| **Giai đoạn 3** (Cần dữ liệu DN) | **23** | `DQ01`<br>`GT06`, `GT16`<br>`TN04`, `TN17`, `TN18`<br>`TT02`–`TT06`<br>`KH01`–`KH07`<br>`DT01`–`DT03`<br>`LD02`, `LD05` | Cần Sổ kế toán tổng hợp (IN08), Sổ kho và Giá thành (IN09), Sổ chi tiết tiền mặt (IN10), Sao kê ngân hàng và UNC (IN11), Sổ chi tiết 511/632 (IN13).<br>- `TT02`, `GT16`: Ngưỡng $\ge 20$ triệu và $\ge 5$ triệu (K5).<br>- `DT01`–`DT03`: Ghép theo skill `doi-chieu-511-632` (K4).<br>- `KH04`: B01-DN dùng [141], B01-DNN dùng ct140 (K6). |

---

## 4. Bảng "Thiếu dữ liệu thì làm gì" (Cập nhật IN14 và K1–K13)

| Nguồn thiếu | Số QT bị chặn | Quy tắc chịu ảnh hưởng | Hành động chuẩn hóa | Đề nghị NNT bổ sung |
|---|---|---|---|---|
| **Sổ kho & Giá thành (IN09)** | 10 | `KH01`–`KH07`, `TN04`, `TN05`, `TN22` | Ghi nhận `KHÔNG_ĐỦ_DỮ_LIỆU`. Tại GĐ2 chỉ đánh giá sơ bộ TC3 từ BCTC. Không kết luận thiếu kho hay xuất khống. | Sổ chi tiết Nhập–Xuất–Tồn kho TK 152, 155, 156; Bảng tính giá thành 154; Bảng định mức kỹ thuật. |
| **Sao kê NH & UNC (IN11)** | 5 | `GT16`, `TT02`, `TT03`, `TT05`, `TT06` | Nếu IN00 có sao kê PDF scan chưa trích dòng: ghi `PDF – chưa kết luận`, ghi nhận `KHÔNG_ĐỦ_DỮ_LIỆU`, **cấm kết luận thiếu chứng từ thanh toán** (K11). | File sao kê ngân hàng định dạng Excel; Giấy báo Nợ / Ủy nhiệm chi gốc các món $\ge 20$tr ($\ge 5$tr). |
| **Sổ chi tiết 511/632 (IN13)** | 4 | `DT01`, `DT02`, `DT03`, `KH07` | Ghi nhận `KHÔNG_ĐỦ_DỮ_LIỆU`. Sheet `KQ3_511_632` không tạo. Vẫn đối chiếu tổng doanh thu ở GĐ2. | Sổ chi tiết Doanh thu 511 (PS Có) và Giá vốn 632 (PS Nợ) theo từng dòng bút toán (K4). |
| **Sổ kế toán tổng hợp (IN08)** | 4 | `GT06`, `TN17`, `TN18`, `TT04` | Ghi nhận `KHÔNG_ĐỦ_DỮ_LIỆU`. Không kết luận chênh lệch thuế mua vào với Nợ 133 theo tháng. | Sổ Nhật ký chung, Sổ Cái TK 133, 111, 112, 131, 331, 3388, 642. |
| **Lịch sử tuân thủ TMS (IN14)** | 2 | `LS01`, `LS02` | Chưa có file tại thư mục 1.3: ghi nhận `KHÔNG_ĐỦ_DỮ_LIỆU` (K10). Có file thì ánh xạ theo cột thật. | Xuất file Excel nghĩa vụ nộp thuế, nợ thuế, kết luận TTKT từ TMS vào thư mục 1.3. |
| **Sổ chi tiết tiền mặt (IN10)** | 2 | `TT02`, `TT04` | Ghi nhận `KHÔNG_ĐỦ_DỮ_LIỆU`. Quét tạm thời trên Bảng kê mua vào theo cụm cùng ngày (quy tắc `TT01`). | Sổ chi tiết tiền mặt TK 1111 và các phiếu chi $\ge 20$tr ($\ge 5$tr). |
| **Tham chiếu ngành (K13)** | 5 | `HD07`, `TN03`, `TN04`, `TN19`, `TN22` | `HD07` ghi **TREO**; các quy tắc còn lại chạy theo ngưỡng cố định, kết luận ghi **sơ bộ – chưa so ngành**. | Bổ sung số liệu trung vị ngành sau khi có thống kê TCT. |

---

## 5. Các câu lệnh thực thi (P1 – P5)

### Lệnh P1: Kiểm kê hồ sơ & Đánh giá mức độ
```text
Giai đoạn 1 – Kiểm kê hồ sơ <FOLDER>, MST <MST>, kỳ <KY>. Áp dụng khối Q và K1-K13.
1. Quét toàn bộ thư mục, kể cả file ngoài cấu trúc 4 bước. Xếp mỗi file vào một nguồn IN00–IN14 (1. DU LIEU THO NGANH THUE = CQT, 2. HO SO NNT CUNG CAP = DN). File lạ liệt kê riêng.
2. Mỗi nguồn ghi: nguồn gốc (CQT/DN), định dạng (XML/Excel/PDF/ảnh), kỳ phủ, số dòng, mã băm SHA-256. Đánh giá: Đủ / Mỏng / Thiếu / PDF chưa trích / Chưa nhận.
3. Rà soát 80 quy tắc: chạy được ngay, chạy một phần, bị chặn vì nguồn nào.
4. Kết luận mức hồ sơ A–E và giai đoạn tiếp theo (2 hoặc 3).
5. Lập danh mục đề nghị bổ sung cho công ty xếp theo số quy tắc mở khóa.
Đầu ra: KQ1_<MST>_<KY>_v1.xlsx gồm KQ1_Kiem_ke, KQ1_Kha_nang_phan_tich, KQ1_De_nghi_bo_sung.
```

### Lệnh P2: Phân tích bằng dữ liệu cơ quan thuế
```text
Giai đoạn 2 – Phân tích bằng dữ liệu CQT cho <FOLDER>, MST <MST>, kỳ <KY>. Áp dụng khối Q và K1-K13. Công ty chưa gửi sổ sách.
1. Chuẩn hóa một lần vào IN01–IN07, IN12 và IN14 (thư mục 1.3): HĐĐT (gắn Tinh_vao_tong, ký hiệu|số|MST chuẩn), bảng kê (bỏ Tổng cộng, chuẩn MST BT6), tờ khai GTGT, BCTC-TNDN, TNCN, cảnh báo rủi ro, TMS.
2. Chạy 57 quy tắc GĐ2. Quy tắc so BCTC ghi "sơ bộ". HD07 ghi TREO. HD01, HD03, HD08 xếp F0 và F1 đều mức Cao (K2). TT01 dùng ngưỡng >= 20tr và >= 5tr (K5). LS01, LS02 đọc IN14 (K10).
3. 100% NCC mua vào kèm cờ F0/F1 (mức Cao) và ĐKKD (cột S, T).
4. Ánh xạ năm BCTC sheet 07 kèm độ tin cậy; cờ MƠ HỒ chờ cán bộ duyệt.
5. 23 quy tắc GĐ3 ghi KHÔNG_ĐỦ_DỮ_LIỆU và nêu nguồn cần xin.
Đầu ra: KQ2_<MST>_<KY>_v1.xlsx gồm 7 sheet theo hợp đồng đầu ra.
```

### Lệnh P3: Hoàn chỉnh hồ sơ khi có sổ sách doanh nghiệp
```text
Giai đoạn 3 – Hoàn chỉnh hồ sơ <FOLDER>, MST <MST>, kỳ <KY>. Áp dụng khối Q và K1-K13. Đã có KQ2, công ty vừa gửi sổ sách.
1. Chạy lại P1 cho phần mới nhận, cập nhật mức hồ sơ C/D/E.
2. Chuẩn hóa nguồn công ty vào IN08–IN11, IN13. Lệch giữa bản DN và CQT ghi phát hiện DQ07.
3. Chạy 23 quy tắc GĐ3. Ghép 511-632 theo skill doi-chieu-511-632 (K4). B01-DN dùng [141], B01-DNN dùng ct140 (K6). Lọc theo loại hình: <LOẠI HÌNH>.
4. Nâng quy tắc "sơ bộ" GĐ2 thành "chắc" hoặc giải tỏa rủi ro; ghi rõ khác biệt so với KQ2.
5. Dựng chuỗi bằng chứng 5 mắt xích cho NCC/HĐ trọng điểm; mắt xích thiếu ghi "Không đủ dữ liệu" (K11).
6. Lập Master 28 tiêu chí 7 cột (nếu không có mẫu 8 sheet thì chỉ dựng 28 TC - K3). Đề cương kiểm tra xếp theo điểm có cột Nhánh xử lý (để cán bộ chọn).
7. Kiểm tra độc lập 10 phát hiện điểm cao nhất từ dữ liệu gốc.
Đầu ra: KQ3_<MST>_<KY>_v1.xlsx gồm 9 sheet theo hợp đồng đầu ra.
```

### Lệnh P4: Nhận hồ sơ bổ sung sau khi đã phân tích
```text
Công ty gửi thêm hồ sơ vào <FOLDER> (MST <MST>, kỳ <KY>). Áp dụng khối Q và K1-K13.
1. So manifest mới với manifest cũ: file mới, file đổi mã băm SHA-256.
2. Chỉ chạy lại quy tắc có nguồn mới/đổi. Các quy tắc khác giữ nguyên kết quả cũ.
3. Xuất phiên bản kế tiếp (v<n+1>), không ghi đè; sheet KQ3_Thay_doi_so_voi_GD2 ghi biến động.
4. Cập nhật danh mục còn thiếu (KQ3_Con_thieu).
```

### Lệnh P5: Ghi nhận quyết định cán bộ và áp dụng TT99
```text
Áp dụng Prompt bổ sung P5 (bản 2) cho <FOLDER>, MST <MST>, kỳ <KY>.
1. Ghi nhận 13 quyết định K1–K13 và nạp workbook 1.2 (32 sheet, IN14 TMS, cột TT99, F0/F1 mức Cao, ngưỡng >=).
2. Đối chiếu mã TT99: B01-DN, B02-DN (DT_CHUA_THUC_HIEN mã 319, VON_CSH mã 411). Thông tư 133 không bị thay thế.
3. Chạy lại đúng các quy tắc bị ảnh hưởng:
   - K5: TT01, TT02, GT16 (các HĐ đúng bằng 20.000.000 đ hoặc cụm đúng ngưỡng).
   - K2: HD01, HD03, HD08 (xếp F1 mức Cao).
   - K6, K7: B01-DN [141], B01-DNN ct140; quy tắc đọc BCTC kỳ từ 2026.
   - K10: LS01, LS02 đọc file 1.3 nếu có.
   - K13: Đổi nhãn HD07 thành TREO; TN03, TN04, TN19, TN22 thành "sơ bộ - chưa so ngành".
4. Xuất phiên bản kế tiếp (v<n+1>) kèm bảng thay đổi.
```
