---
type: wiki
title: Playbook lệnh theo bước - Phân tích hồ sơ thuế
updated: 2026-10-05
---

# Playbook lệnh theo bước (công ty mới)

Mục tiêu: làm **từng module riêng** khi chưa đủ hồ sơ, không chờ đủ 100% dữ liệu rồi mới chạy. Mỗi bước có đầu vào tối thiểu, câu lệnh copy-paste, file ra, và điều kiện bỏ qua.

Thay placeholder:
- `<FOLDER>` = đường dẫn tuyệt đối thư mục DN (`...\<TEN>_<MST>`)
- `<MST>`, `<TEN>`, `<KY>` = mã số thuế, tên ngắn, niên độ (vd `2022-2025`)

Chuẩn Master cuối: **8 sheet**, 28 tiêu chí đủ 7 cột, font Times New Roman. Không dùng bản rút gọn.

## Thứ tự tối ưu (đường nhanh)

```
0 Kiểm kê → 1 Bảng kê → 2 GTGT/TNDN/HĐĐT → 3 NCC rủi ro
→ 4 TNCN → 5 511-632 → 6 XNT → 7 Giá vốn/Giá thành
→ 8 Công nợ/Tiền mặt → 9 Master 28 TC → 10 Đề cương
```

Làm song song được khi đã có dữ liệu: **(2)+(3)+(4)**; **(5)+(6)+(8)**.

---

## Bước 0 - Kiểm kê & checklist (luôn làm trước)

**Cần:** `<FOLDER>` đã theo cấu trúc 4 bước.

**Lệnh:**
```
Kiểm kê hồ sơ tại <FOLDER>
MST <MST>, kỳ <KY>.
Liệt kê đủ/thiếu theo checklist 4 bước.
Xuất/cập nhật CHECKLIST_TONG_HOP_HO_SO.xlsx tại 4. KET QUA PHAN TICH.
Nêu rõ module nào chạy được ngay, module nào đang thiếu đầu vào.
```

**Ra:** checklist + danh sách module sẵn sàng.
**Thiếu gì thì dừng:** chưa có folder tuyệt đối / chưa biết MST-kỳ.

---

## Bước 1 - Gộp bảng kê mua vào / bán ra

**Cần tối thiểu:** file bảng kê quý/tháng trong `1.2` hoặc `2.5`.

**Lệnh:**
```
Tổng hợp bảng kê mua vào bán ra của <FOLDER>
lưu tại 2. HO SO NNT CUNG CAP\2.5. Bang ke mua vao ban ra\
thành Bang_ke_mua_vao_ban_ra_tong_hop_<KY>.xlsx
2 sheet: mua vao, ban ra.
Lọc bỏ dòng Tổng cộng, giữ đối tác tên TỔNG CÔNG TY.
Map cột theo header, khử trùng kỳ bổ sung (lấy bản BS mới nhất).
Sao lưu thêm vào 4. KET QUA PHAN TICH\cau-thanh\
```

**Ra:** file 2 sheet + tổng HĐ/doanh số/thuế theo năm.
**Chưa có bảng kê:** bỏ qua; dùng tạm `HDDT.xlsx` nhưng ghi rõ độ phủ mỏng.

---

## Bước 2 - Đối chiếu GTGT / TNDN / HĐĐT (CBRR)

**Cần:** XML GTGT + TNDN (+ Bảng kê hoặc HĐĐT).

**Lệnh:**
```
Đối chiếu tờ khai GTGT - QTT TNDN - bảng kê - HĐĐT chuẩn CBRR
cho <TEN> MST <MST> tại <FOLDER>
kỳ <KY>.
Xuất Tong_hop_doi_chieu_GTGT_TNDN_<MST>.xlsx
(2 sheet chuẩn CBRR nếu đủ dữ liệu).
Bóc tách kỳ BS, HĐ hủy/thay thế, nguyên nhân lệch.
```

**Ra:** bảng đối chiếu theo quý/năm + chênh lệch.
**Thiếu XML quý:** vẫn chạy các năm có; liệt kê quý thiếu, không bịa số.

---

## Bước 3 - Đối chiếu NCC rủi ro (F0/F1/Công an/Danh bạ)

**Cần:** bảng kê mua vào (hoặc HĐĐT mua) + sổ CVCBRR + (tuỳ chọn) Công an + danh bạ.

**Lệnh:**
```
Đối chiếu NCC mua vào với CVCBRR và danh bạ cho <FOLDER>
Phần 1: NNT <MST> với sheet F1 và Bao_cao_cong_van.xlsx, ghi số CV + nguồn trích.
Phần 2: người bán sheet mua vao với sheet F0.
Xuất Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx đủ chuẩn Vạn Phát Thành:
- Sheet Tong_hop_doi_chieu: Phần 1 + Bảng 2.1 F0 + Bảng 2.2 ĐỦ 100% NCC + Phần 3 nhận xét
- Sheet Chi_tiet_hoa_don: đủ cột chuẩn, gắn cờ F0/ĐKKD
Không rút gọn chỉ lấy NCC bất thường.
```

**Ra:** file NCC 2 sheet.
**Thiếu danh bạ:** vẫn đối F0/F1/CA; ghi thiếu trạng thái ĐKKD.

---

## Bước 4 - TNCN & (nếu có) đa nguồn

**Cần:** XML QTT TNCN; tốt hơn có tờ khai quý + bảng lương + file tra cứu đa nguồn.

**Lệnh:**
```
Tổng hợp QTT TNCN và danh sách NLĐ MST <MST> kỳ <KY>
từ XML tại <FOLDER>\1. DU LIEU THO NGANH THUE\1.1. To khai XML
xuất vào 1.4 và cau-thanh.
Nếu có file tra cứu đa nguồn ở 3.3 thì đối chiếu TC22/TC23.
```

**Ra:** bảng TNCN theo năm + NLD.
**Thiếu đa nguồn:** vẫn lập QTT; đánh dấu TC22/TC23 chờ bổ sung.

---

## Bước 5 - Đối chiếu 511-632 (DN thương mại / có sổ)

**Cần:** sổ chi tiết 511 (PS Có), 632 (PS Nợ); tốt hơn có bảng kê bán ra để chuẩn tên hàng.

**Lệnh:**
```
Đối chiếu tự động doanh thu - giá vốn theo skill doi-chieu-511-632
cho công ty <TEN> MST <MST> tại <FOLDER> kỳ <KY>.
Xuất Doi_chieu_511_632_<MST>_<KY>.xlsx/.xlsm vào cau-thanh
và nhúng sheet Doi_chieu_511_632 vào Master.
Báo lãi gộp theo năm, dòng P<=0, dòng không khớp 632.
```

**Ra:** tool 511-632 + tỷ lệ lãi gộp.
**Chưa có sổ 511/632:** treo bước này; vẫn làm Bước 2 từ tờ khai.

---

## Bước 6 - XNT & Kho (TC3)

**Cần:** sổ XNT 152/155 (SX) hoặc 156 (thương mại) + BCTC.

**Lệnh:**
```
Phân tích XNT kho tại <FOLDER> kỳ <KY>.
DN thương mại: TK 156. DN sản xuất: 152+155 (+156 nếu có).
Đối chiếu tồn CK với BCTC: DNN dùng ct140; đầy đủ dùng [141] = 152+155+156.
Đối chiếu xuất kho với GV 632.
Xuất/cập nhật sheet Tong_hop_XNT_va_Kho, không để giá trị 0 nếu đã có sổ.
```

**Ra:** bảng XNT theo năm + tỷ lệ HTK/DT.
**Bẫy:** đừng kết luận HTK=0 chỉ vì ct141=0 trên BCTC DNN.

---

## Bước 7 - Giá thành / cơ cấu chi phí (TC28)

**Cần (SX):** sổ 154/621/622/627 + bảng giá thành + định mức.
**Cần (thương mại):** GV 632 + chi phí BH/QL từ TNDN/BCTC (+ kết quả 511-632).

**Lệnh SX:**
```
Phân tích giá thành và định mức tại <FOLDER> kỳ <KY>.
Bóc 621/622/627, khớp GV với BCTC [02] và TK 632.
Cập nhật sheet Gia_thanh_va_Dinh_muc.
```

**Lệnh thương mại:**
```
DN thương mại: lập sheet Gia_thanh_va_Dinh_muc theo cơ cấu
Doanh thu - GV 632 - lãi gộp - chi phí bán hàng/QLDN (từ TNDN),
không bắt buộc 621/622/627.
```

---

## Bước 8 - Công nợ & tiền mặt / ngân hàng

**Cần:** sổ 131/331/111/112; sao kê NH (Excel tốt hơn PDF).

**Lệnh:**
```
Phân tích công nợ 131/331 và quét chi tiền mặt TK 111 tại <FOLDER> kỳ <KY>.
Ngưỡng: >20 triệu đến 06/2025; >=5 triệu từ 07/2025.
Phân loại mục đích chi; nêu rõ sao kê đang PDF thì chưa kết luận thiếu UNC.
Cập nhật sheet Cong_no_va_Ngan_hang bằng số liệu thật.
```

---

## Bước 9 - Tái lập Master 8 sheet + đủ 28 tiêu chí

**Cần:** tối thiểu đã có Bước 2 (và càng nhiều bước trước càng tốt). Có thể lập khung 28 TC sớm rồi cập nhật dần.

**Lệnh:**
```
Tái lập đầy đủ Master Excel theo chuẩn Vạn Phát Thành cho <FOLDER>
MST <MST> kỳ <KY>:
- Đủ 28/28 tiêu chí, 7 cột chuẩn, không rút gọn
- Đủ 8 sheet: Danh_gia_28_tieu_chi_rui_ro, Tong_hop_GTGT_TNDN,
  Tong_hop_NCC_CVCBRR, Tong_hop_TNCN_va_Da_nguon, Tong_hop_XNT_va_Kho,
  Gia_thanh_va_Dinh_muc, Doi_chieu_511_632, Cong_no_va_Ngan_hang
- Sheet nào thiếu dữ liệu: ghi rõ "Thiếu nguồn ..." chứ không để 0 giả
- Đồng bộ: 4. KET QUA PHAN TICH\, cau-thanh\, 7. Sao luu tong hop\, ket-qua\, sao-luu\
Map năm BCTC theo doanh thu nếu kyKKhai lệch.
```

**Ra:** `Phan_tich_ho_so_cong_ty_<TEN>_<MST>.xlsx`

---

## Bước 10 - Đề cương / trọng tâm kiểm tra thực địa

**Cần:** Master đã có ít nhất các tiêu chí Cao/Trung bình.

**Lệnh:**
```
Từ Master <TEN> MST <MST>, lập đề cương kiểm tra tại trụ sở:
liệt kê tiêu chí Cao/Trung bình, mẫu chứng từ cần xem,
NCC/HĐ trọng điểm, nội dung đã giải tỏa.
```

---

## Gói lệnh siêu ngắn (khi đã quen)

| Module | Câu siêu ngắn |
|---|---|
| Kiểm kê | `Kiểm kê checklist <FOLDER>` |
| Bảng kê | `Gộp bảng kê <FOLDER> vào 2.5` |
| GTGT-TNDN | `Đối chiếu CBRR GTGT-TNDN-HĐĐT <FOLDER>` |
| NCC | `Đối chiếu NCC F0/F1/CA chuẩn VPT <FOLDER>` |
| TNCN | `Tổng hợp TNCN <MST> <FOLDER>` |
| 511-632 | `Đối chiếu 511-632 skill mới <FOLDER>` |
| XNT | `Phân tích XNT kho <FOLDER>` |
| Giá thành | `Phân tích giá thành/định mức <FOLDER>` hoặc `Cơ cấu GV thương mại <FOLDER>` |
| Công nợ/TM | `Quét 111 và công nợ 131/331 <FOLDER>` |
| Master 28 | `Tái lập Master 8 sheet đủ 28 TC chuẩn VPT <FOLDER>` |
| Đề cương | `Lập đề cương từ Master <FOLDER>` |

---

## Nguyên tắc tốc độ

1. **Không chờ đủ mới chạy** - chạy module nào đủ đầu vào trước.
2. **Một lệnh = một file ra** - dễ nối tiếp và kiểm chứng.
3. **Thiếu thì ghi thiếu** trên sheet/báo cáo; không để 0 giả.
4. **Song song:** sau Bước 1 có thể giao đồng thời Bước 2+3+4.
5. **Chốt Master** chỉ khi 28 TC đủ 7 cột và các sheet có số thật hoặc nhãn thiếu nguồn rõ ràng.
