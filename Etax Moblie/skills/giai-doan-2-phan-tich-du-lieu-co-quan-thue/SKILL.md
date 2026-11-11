---
name: Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế
description: "Phân tích 57 quy tắc thuế khi chưa có sổ DN: chuẩn hóa IN01-IN07/12, đối chiếu GTGT-TNDN-BCTC, 100% NCC rủi ro, phân loại Chắc/Sơ bộ (KQ2)."
group: Tài chính
updated: '2026-10-15'
---

# Giai đoạn 2 – Phân tích bằng dữ liệu cơ quan thuế

Thực hiện phân tích toàn diện hồ sơ khi Người nộp thuế (NNT) chưa cung cấp sổ kế toán, hoặc chỉ mới có nguồn dữ liệu do cơ quan thuế (CQT) quản lý và tờ khai do NNT nộp (hồ sơ mức A hoặc B). Chuẩn hóa một lần các nguồn dữ liệu vào các bảng phẳng IN01–IN07 và IN12, chạy 57 quy tắc rủi ro của giai đoạn 2, bóc tách 100% nhà cung cấp mua vào, xác định ánh xạ năm BCTC kèm cờ độ tin cậy, và phân loại minh bạch kết luận "Chắc chắn" hay "Sơ bộ".

> Áp dụng nghiêm ngặt **Khối Q** (Quy tắc chung): Mô hình ngôn ngữ không tự tính toán số liệu; mọi phát hiện phải có mã quy tắc, ngưỡng tham số, điểm và trích dẫn ô nguồn dạng chữ; quy tắc thiếu nguồn ghi `KHÔNG_ĐỦ_DỮ_LIỆU`, không điền số 0 giả.

---

## Khi nào kích hoạt

- Sau khi hoàn thành Giai đoạn 1 (`giai-doan-1-kiem-ke-danh-gia`) và xác định hồ sơ ở mức **A** hoặc **B**.
- NNT chưa nộp hoặc nộp chưa đủ sổ sách kế toán nội bộ (chưa có Sổ 511/632, Sổ XNT kho, Sổ chi tiết tiền mặt, Sao kê ngân hàng).
- Các câu lệnh/yêu cầu: "phân tích giai đoạn 2", "chạy phân tích dữ liệu cơ quan thuế", "đối chiếu tờ khai HĐĐT bảng kê", "lệnh P2".

---

## Đầu vào

### Nguồn dữ liệu CQT bắt buộc:
1. **IN01_HDDT**: Dữ liệu HĐĐT chiều Mua vào và Bán ra (Portal CQT).
2. **IN02_BangKe**: Bảng kê mua vào, bán ra các kỳ của NNT (thư mục 2.5).
3. **IN03_ToKhai_GTGT**: Tờ khai 01/GTGT chính thức và các kỳ bổ sung XML (thư mục 1.1).
4. **IN04_BCTC_TNDN**: Tờ khai QTT TNDN (03/TNDN) và BCTC XML (TT200, TT133, TT99).
5. **IN05_NLD_TNCN**: Bảng kê người lao động (05-1/BK-TNCN và 05-2/BK-TNCN XML).
6. **IN06_QTT_TNCN**: Tờ khai quyết toán TNCN và 4 quý tạm tính XML.
7. **IN07_NCC_CanhBao**: Danh sách cảnh báo rủi ro (CVCBRR F0/F1, danh bạ ĐKKD, Sổ Công an - có cột `Nhom_F` và `Nguon_F0`).
8. **IN12_GDLK**: Thông tin giao dịch liên kết (kèm tờ khai TNDN).
9. **IN14_TMS_Nghia_vu**: Lịch sử tuân thủ từ hệ thống TMS (thư mục `1.3. Nghia vu nop thue & Thong bao` - Quyết định K10).

---

## Chuẩn hóa và khớp hóa đơn

1. **Chuẩn hóa MST đối tác (BT6):**
   - Loại bỏ ký tự lạ, khoảng trắng, dấu chấm; giữ số 0 đầu; chuẩn hóa về 10 số (doanh nghiệp chính) hoặc 13 số có dấu `-` (chi nhánh / đơn vị phụ thuộc).
   - Kiểm tra tính hợp lệ của MST 10 số bằng thuật toán cơ quan thuế. Lưu song song cả `MST_goc` và `MST_chuan`.
2. **Xác định cờ `Tinh_vao_tong` của HĐĐT:**
   - `Có`: Hóa đơn mới, Hóa đơn thay thế, Hóa đơn điều chỉnh tăng/giảm. (HĐ điều chỉnh giảm lưu giá trị âm).
   - `Không`: Hóa đơn đã bị xóa bỏ, Hóa đơn đã bị thay thế/hủy bỏ (lấy từ dữ liệu trạng thái HĐĐT).
3. **Khóa đối chiếu hóa đơn theo tầng:**
   - Gộp các dòng hàng của cùng một hóa đơn trước khi đối chiếu (tránh lỗi lệch số dòng).
   - **Tầng 1 (Khớp tuyệt đối):** `Ky_hieu_HD | So_HD | MST_chuan` (Ký hiệu thật, bỏ các ký hiệu Q1–Q4).
   - **Tầng 2 (Khớp mờ):** `So_HD + MST_chuan + Ngày (± 3 ngày) + Tiền (± sai số P02/P01)`.
   - Kết quả gán vào `IN02.Trang_thai_DC`: `KHOP`, `KHOP_LECH_TIEN`, `CHI_CO_BK`, hoặc `CHI_CO_HDDT`.

---

## Phạm vi 57 quy tắc chạy trong Giai đoạn 2 (Chuẩn v1.2)

| Nhóm | Số QT | Mã quy tắc cụ thể |
|---|---|---|
| **Chất lượng dữ liệu** | 7 | `DQ02` (dòng Tổng cộng lẫn BK), `DQ03` (MST sai định dạng), `DQ04` (thiếu ký hiệu HĐ), `DQ05` (HĐ trùng BK), `DQ06` (HĐĐT trùng), `DQ07` (lệch bản CQT vs DN), `DQ08` (bản bổ sung không đủ dòng). |
| **GTGT** | 14 | `GT01` (DT bán ra [34] vs HĐĐT), `GT02` (DT [34] vs BK bán ra), `GT03` (Thuế BR [35] vs HĐĐT/BK), `GT04` (Mua vào [23] vs HĐĐT), `GT05` (Thuế MV [24] vs HĐĐT/BK), `GT07` (Thuế suất BR lệch nhóm), `GT08` (Thuế suất MV lệch nhóm), `GT09` (HĐ lập cuối kỳ), `GT10` (HĐ áp sai giảm thuế 8% - K9), `GT11` (Thuế [43] vs BCTC [152]/[162]), `GT12` (Thuế [43] > 10%*(HTK+TSCĐ)), `GT13` (Thuế [43] tăng liên tục), `GT14` (Thuế [40] âm), `GT15` (Kê khai khấu trừ âm 4 quý liên tiếp). |
| **TNDN & BCTC** | 19 | `TN01` (DT GTGT vs DT BCTC [01]), `TN02` (DT QTT TNDN vs DT BCTC), `TN03` (LN gộp âm - TC5, sơ bộ chưa so ngành - K13), `TN05` (HTK+XDCB > 30% DT - TC3, K6), `TN06` (Người mua trả trước > 5% DT - TC2), `TN07` (DT chưa thực hiện / chờ phân bổ > 0 - TC4, TT99 mã 319), `TN08` (CP phải trả trích trước lớn - TC6), `TN09` (CP trả trước lớn - TC7), `TN10` (Miễn giảm ưu đãi TNDN - TC8), `TN11` (Dự phòng giảm giá HTK lớn), `TN12` (Lỗ lũy kế vượt VCSH - K6), `TN13` (DT giảm nhưng HTK tăng), `TN14` (Lãi vay / EBITDA > 30% - TC17), `TN15` (Lỗ liên tục 2 năm, VON_CSH TT99 mã 411), `TN16` (DT tăng nhưng thuế TNDN nộp giảm), `TN19` (Thuế TNDN thực tế vs TNDN danh nghĩa, sơ bộ chưa so ngành), `TN20` (Khấu hao tài sản - chưa triển khai), `TN21` (Doanh thu tài chính bất thường), `TN22` (Tỷ lệ giá vốn/DT thương mại, sơ bộ chưa so ngành - K13). |
| **Hóa đơn & NCC** | 9 | `HD01` (Mua vào từ NCC cảnh báo F0 hoặc F1 - K2 xếp mức **Cao**; NNT thuộc F1 là phát hiện Cao), `HD02` (Mua vào từ NCC ĐKKD bất thường), `HD03` (Tỷ trọng mua NCC rủi ro F0/F1 > 5% - TC26), `HD04` (Tỷ trọng Top-5 NCC > 50%), `HD05` (NCC mới thành lập quy mô lớn), `HD06` (Benford phân phối số đầu), `HD07` (Đơn giá mặt hàng lệch chuẩn - **TREO** do thiếu tham chiếu ngành - K13), `HD08` (Doanh nghiệp nằm trên sổ CVCBRR bên mua), `HD09` (Quan hệ mua bán qua lại NNT - NCC). |
| **Tiền mặt** | 1 | `TT01` (Hóa đơn mua vào đạt ngưỡng: $\ge 20$ triệu đến 30/06/2025; $\ge 5$ triệu từ 01/07/2025 - Quyết định K5). |
| **Lao động - TNCN** | 3 | `LD01` (Lệch QTT TNCN vs tạm tính 4 quý), `LD03` (Lao động nguồn 05-2 không khấu trừ 10%), `LD04` (Trùng số CCCD/CMND trên tờ khai). |
| **Giao dịch liên kết** | 2 | `LK01` (Không nộp phụ lục GDLK khi có quan hệ liên kết), `LK02` (Chi phí lãi vay vượt mức khống chế 30% khi có GDLK). |
| **Lịch sử tuân thủ** | 2 | `LS01` (Lịch sử nợ đọng thuế, phạt VPHC - đọc IN14 từ thư mục 1.3 - K10), `LS02` (Bộ chỉ số rủi ro tuân thủ TMS - đọc IN14 - K10). |

---

## Phân loại kết luận: "Chắc chắn" vs "Sơ bộ"

- **Nhãn "Chắc chắn" (Chắc):** Các quy tắc đối chiếu trực tiếp dữ liệu cơ quan thuế và kê khai thuế:
  - Nhóm Chất lượng dữ liệu: `DQ02`–`DQ08`.
  - Nhóm GTGT đối chiếu nội bộ: `GT01`–`GT05`, `GT07`–`GT10`, `GT14`, `GT15`.
  - Nhóm NCC rủi ro: `HD01`–`HD04`, `HD08`.
  - Nhóm Lao động - TNCN: `LD01`, `LD03`, `LD04`.
  - Nhóm Lịch sử & GDLK: `LK01`, `LK02`, `LS01`.
- **Nhãn "Sơ bộ":** Các quy tắc so sánh BCTC và ngưỡng thanh toán mà chưa có sổ kế toán / chứng từ thực tế:
  - `TN01`, `TN02`: Lệch DT giữa BCTC và tờ khai (cần sổ 511 để xác minh cut-off, doanh thu chưa xuất HĐ).
  - `TN03`, `TN04`: Biên lợi nhuận gộp, giá vốn (cần Sổ 632 và sổ kho).
  - `TN05` (TC3), `TN06` (TC2), `TN07`–`TN14`, `TN22`: Cần Sổ Cái tài khoản 152/155/156, 131, 331, 242, 335... để khẳng định bản chất.
  - `TT01`: Khẳng định chắc chắn về việc hóa đơn đạt ngưỡng bắt buộc không dùng tiền mặt, nhưng kết luận về việc "đã thanh toán hay chưa" là sơ bộ (phải chờ sao kê NH / UNC ở GĐ3).

---

## Xử lý 23 quy tắc thuộc Giai đoạn 3

- Toàn bộ 23 quy tắc Giai đoạn 3 (`DQ01`, `GT06`, `GT16`, `TN04`, `TN17`, `TN18`, `TT02`–`TT06`, `KH01`–`KH07`, `DT01`–`DT03`, `LD02`, `LD05`) được ghi nhận trạng thái:
  `KHÔNG_ĐỦ_DỮ_LIỆU`
- Ghi rõ nguồn còn thiếu cần yêu cầu NNT cung cấp (IN08, IN09, IN10, IN11, IN13).
- **Tuyệt đối không kết luận vượt nguồn:** Nếu chỉ có sao kê PDF chưa trích dòng, ghi `PDF – chưa kết luận`, không được tự ý kết luận DN không có chứng từ thanh toán ngân hàng.

---

## Hợp đồng đầu ra

Xuất file Excel kết quả **chỉ chứa các giá trị tĩnh**:
`KQ2_<MST>_<KY>_v<n>.xlsx` (mẫu chuẩn tại `KQ2_MAU_Phan_tich_CQT.xlsx`) gồm 7 sheet:
1. **`KQ2_Doi_chieu_GTGT`**: Bảng đối chiếu GTGT theo kỳ (Tờ khai [34]/[23] vs Bảng kê vs HĐĐT).
2. **`KQ2_Doi_chieu_TNDN_BCTC`**: Đối chiếu Doanh thu, Chi phí, TNDN vs BCTC.
3. **`KQ2_Nam_BCTC`**: Ánh xạ năm BCTC với năm kê khai kèm độ tin cậy và cờ MƠ HỒ/KHÁC NĂM.
4. **`KQ2_NCC_mua_vao`**: 100% NCC mua vào kèm cờ F0/F1/CA/ĐKKD, giá trị, thuế và tỷ trọng.
5. **`KQ2_TNCN_LD`**: Kết quả đối chiếu QTT TNCN, tạm tính quý, lao động 05-2 và trùng CCCD.
6. **`KQ2_Phat_hien`**: Danh sách phát hiện rủi ro sắp xếp theo thứ tự tổng điểm (kèm nhãn Chắc / Sơ bộ).
7. **`KQ2_Chua_ket_luan`**: Bảng tổng hợp các quy tắc bị chặn, nguồn còn thiếu và điều cấm kết luận vội.

---

## Câu lệnh thực thi mẫu

### Lệnh đầy đủ (P2):
```text
Giai đoạn 2 – Phân tích bằng dữ liệu cơ quan thuế cho <FOLDER>, MST <MST>, kỳ <KY>. Áp dụng khối Q. Công ty chưa gửi sổ sách.
1. Chuẩn hóa một lần vào IN01–IN07 và IN12: HĐĐT mua vào và bán ra (gắn Tinh_vao_tong, ký hiệu|số|MST chuẩn), bảng kê (bỏ dòng Tổng cộng, giữ MST gốc và MST chuẩn, xử lý bản bổ sung), tờ khai GTGT, QTT TNDN và BCTC nộp thuế, TNCN, cảnh báo/công văn, giao dịch liên kết. Kiểm tổng khớp chi tiết.
2. Chạy 57 quy tắc giai đoạn 2 (DQ02–DQ08; GT01–GT05, GT07–GT15; TN01–TN03, TN05–TN16, TN19–TN22; HD01–HD09; TT01; LD01, LD03, LD04; LK01, LK02; LS01, LS02). Quy tắc nào cần sổ mới kết luận chắc thì ghi "sơ bộ".
3. NCC mua vào: 100% nhà cung cấp, không rút gọn, kèm cờ rủi ro.
4. Ánh xạ năm BCTC với năm kê khai kèm độ tin cậy; năm MƠ HỒ hoặc KHÁC NĂM chờ tôi xác nhận.
5. Với 23 quy tắc giai đoạn 3 ghi KHÔNG_ĐỦ_DỮ_LIỆU và nêu nguồn cần xin.
6. Nêu rõ: từ dữ liệu hiện có có thể kết luận gì, chưa thể kết luận gì.
Đầu ra: KQ2_<MST>_<KY>_v1.xlsx gồm KQ2_Doi_chieu_GTGT, KQ2_Doi_chieu_TNDN_BCTC, KQ2_Nam_BCTC, KQ2_NCC_mua_vao, KQ2_TNCN_LD, KQ2_Phat_hien, KQ2_Chua_ket_luan.
```

### Lệnh siêu ngắn:
```text
Phân tích GĐ2 CQT <FOLDER> MST <MST> kỳ <KY> xuất KQ2
```
