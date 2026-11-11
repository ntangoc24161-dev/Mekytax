---
type: wiki
title: Playbook tất cả skill phân tích hồ sơ thuế
updated: 2026-11-05
---

# Playbook tất cả skill phân tích hồ sơ thuế

Catalog lệnh copy-paste của **mọi skill** phục vụ phân tích hồ sơ kiểm tra thuế tại brain này. Mỗi skill một khối: cần gì, lệnh, file ra, thiếu thì làm gì.

Thay placeholder:

- `<FOLDER>` = đường dẫn tuyệt đối thư mục DN (`...\<TEN>_<MST>`)
- `<MST>`, `<TEN>`, `<KY>` = mã số thuế, tên ngắn ASCII, niên độ (vd `2022-2025`)

Chuẩn cuối: Master **8 sheet**, **28/28** tiêu chí đủ 7 cột, font Times New Roman. Thiếu nguồn thì ghi thiếu, không để số 0 giả. Ô text Excel không bắt đầu bằng `=`.

Playbook chạy theo bước (công ty mới): `playbook-lenh-theo-buoc.md`.

Song song được khi đã có dữ liệu: **GTGT + NCC + TNCN**; **511-632 + XNT + công nợ**.

---

## Bản đồ skill

### A. Ba skill giai đoạn chủ lực (Pipeline Skills v1.1)

| Giai đoạn | Skill | Việc chính | Đầu ra |
|---|---|---|---|
| **1** | `giai-doan-1-kiem-ke-danh-gia` | Kiểm kê manifest IN00–IN13, xác định mức hồ sơ A–E, rà soát 80 quy tắc, danh mục đề nghị bổ sung | `KQ1_<MST>_<KY>_v1.xlsx` |
| **2** | `giai-doan-2-phan-tich-du-lieu-co-quan-thue` | Chuẩn hóa CQT, chạy 57 quy tắc CQT (khi chưa có sổ DN), 100% NCC rủi ro, phân loại Chắc/Sơ bộ | `KQ2_<MST>_<KY>_v1.xlsx` |
| **3** | `giai-doan-3-hoan-chinh-ho-so` | Hoàn chỉnh 23 quy tắc sổ DN, chuỗi bằng chứng 5 mắt xích, Master 28 tiêu chí 7 cột, đề cương kiểm tra | `KQ3_<MST>_<KY>_v1.xlsx` |

Chi tiết lệnh 3 giai đoạn: `skills/phan-tich-ho-so-thue/playbook-3-giai-doan.md`.

### B. Đường ống phân tích theo module (nạp đúng slug, skill trục giữ Master)

| Bước | Skill | File playbook riêng |
|---|---|---|
| 0 | `to-chuc-ho-so-kiem-tra-thue` | `skills/to-chuc-ho-so-kiem-tra-thue/playbook-lenh.md` |
| 0b + 1 | `tong-hop-dau-vao-hddt-bang-ke-nld` | `skills/tong-hop-dau-vao-hddt-bang-ke-nld/playbook-lenh.md` |
| 1b | `lam-sach-cd1-mst-hoa-don` | `skills/lam-sach-cd1-mst-hoa-don/playbook-lenh.md` |
| 2 | `doi-chieu-gtgt-tndn-cbrr` | `skills/doi-chieu-gtgt-tndn-cbrr/playbook-lenh.md` |
| 3 | `doi-chieu-ncc-mua-vao-rui-ro` | `skills/doi-chieu-ncc-mua-vao-rui-ro/playbook-lenh.md` |
| 4 | `tong-hop-tncn-da-nguon` | `skills/tong-hop-tncn-da-nguon/playbook-lenh.md` |
| 5 | `doi-chieu-511-632` | `skills/doi-chieu-511-632/playbook-lenh.md` |
| 6 | `phan-tich-xnt-kho` | `skills/phan-tich-xnt-kho/playbook-lenh.md` |
| 7 | `phan-tich-gia-thanh-dinh-muc` | `skills/phan-tich-gia-thanh-dinh-muc/playbook-lenh.md` |
| 8 | `phan-tich-cong-no-tien-mat` | `skills/phan-tich-cong-no-tien-mat/playbook-lenh.md` |
| 9-10 | `phan-tich-ho-so-thue` | `skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md` |

### B. Skill cạnh (không nằm trên đường ống 0-10)

| Skill | Việc |
|---|---|
| `mo-ho-so-ttkt` | Checklist nghiệp vụ mở hồ sơ: BCTC, GDLK, nhánh thiếu thuế/trốn thuế |
| `tong-hop-ben-mua-tu-so-cvcbrr` | Tổng hợp bên mua trên sổ CVCBRR (khác NCC bên bán của 1 DN) |
| `quan-ly-rui-ro-tuan-thu-thue` | Khung chọn đối tượng / bộ chỉ số TCT, không mổ 1 bộ sổ |
| `xem-ho-so-thue` | Mở XML, dựng tờ khai iTax Viewer (shared-skills) |
| `thue-tncn` | Chính sách TNCN (shared-skills); số liệu DN dùng `tong-hop-tncn-da-nguon` |
| `doi-chieu-mat-hang-mua-ban` | HĐĐT mua vs bán theo mặt hàng, SL, đơn giá; bắt chuỗi trung gian. Khác `doi-chieu-ncc-mua-vao-rui-ro` / `doi-chieu-511-632` / `phan-tich-xnt-kho`. Playbook: `skills/doi-chieu-mat-hang-mua-ban/playbook-lenh.md` |
| `tao-bang-tinh-tu-ptdn` | Tạo phôi `0. BIEN BAN\Bang tinh.xlsx` từ mẫu, đổi F5/F6/F7, xóa số mẫu |
| `nhap-bang-tinh-tu-bao-cao-hdkd` | Nhập KQKDSXKD từ PTDN02 (+ BCTC NamNay); GTGT/EBITDA từ KQ2.1 `*_FINAL` |
| `boc-tach-bang-tinh-phan-tich` | Tỷ lệ chi phí/doanh thu, phân tích Mục III DI-01 từ Bang tinh |

Không dùng cho phân tích số liệu 1 DN: `xu-phat-vphc-thue-hoa-don`, `tra-loi-cong-van`, `soan-thao-van-ban`.

---

## A0 - Tổ chức hồ sơ (`to-chuc-ho-so-kiem-tra-thue`)

**Cần:** đường dẫn tuyệt đối; suy được MST + tên thì làm luôn.

**Lệnh:**
```
Tổ chức hồ sơ 4 bước tại <FOLDER>
MST <MST>, tên <TEN>, kỳ <KY> theo skill to-chuc-ho-so-kiem-tra-thue.
Tạo thư mục <TEN>_<MST>, phân loại file, xuất 1 checklist A4
tại 4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx.
Xóa nguồn cũ sau khi đối soát. Không phân tích số liệu.
Nêu module phân tích nào chạy được ngay.
```

**Ra:** folder chuẩn + 1 checklist A4.
**Thiếu:** chưa có folder tuyệt đối thì dừng.

---

## A0b - Đầu vào HĐĐT / bảng kê / NLĐ (`tong-hop-dau-vao-hddt-bang-ke-nld`)

**Cần:** `1.2` và/hoặc `2.5`; XML TNCN nếu làm sheet NLĐ.

**Lệnh:**
```
Tổng hợp dữ liệu đầu vào trước phân tích cho <FOLDER>
MST <MST> kỳ <KY> theo mẫu 02_Tong_hop_dau_vao_HDDT_BangKe_NLD_MAU.xlsx:
1) HĐĐT CQ thuế → 2 sheet HDDT_Mua_vao, HDDT_Ban_ra (lọc HĐ hủy/thay thế).
2) Bảng kê công ty → 2 sheet Bang_ke_Mua_vao, Bang_ke_Ban_ra (bỏ dòng Tổng cộng).
3) NLĐ qua các năm → 1 sheet Tong_hop_NLD; ghi nguồn 05-1 hoặc 05-2.
   Nếu nguồn 05-2 mà công ty không trích 10% → Đề xuất kiểm tra cam kết = Có.
4) Điền Tom_tat_kiem_tra (Đủ/Thiếu/Mỏng).
Lưu Tong_hop_dau_vao_HDDT_BangKe_NLD_<MST>_<KY>.xlsx vào cau-thanh.
Không trộn HĐĐT với Bảng kê chung một sheet chi tiết.
```

**Lệnh gộp bảng kê riêng:**
```
Tổng hợp bảng kê mua vào bán ra của <FOLDER>
lưu tại 2. HO SO NNT CUNG CAP\2.5. Bang ke mua vao ban ra\
thành Bang_ke_mua_vao_ban_ra_tong_hop_<KY>.xlsx
2 sheet: mua vao, ban ra.
Lọc bỏ dòng Tổng cộng, giữ đối tác tên TỔNG CÔNG TY.
Map cột theo header, khử trùng kỳ bổ sung (lấy bản BS mới nhất).
Sao lưu thêm vào 4. KET QUA PHAN TICH\cau-thanh\
```

**Ra:** file mẫu 02 (7 sheet) và/hoặc bảng kê 2 sheet.
**Thiếu 1 nguồn:** vẫn xuất sheet có dữ liệu; Tom_tat ghi Thiếu/Mỏng.

---

## A1b - Làm sạch CD1 (`lam-sach-cd1-mst-hoa-don`)

**Cần:** bảng kê hoặc HĐĐT đã gộp; tốt hơn có NKC.

**Lệnh:**
```
Chạy CD1 BT5-BT9 tại <FOLDER> kỳ <KY> và gắn vào bảng biểu (không tạo mẫu 03 tách):
1) BT6 chuẩn hóa MST trên bảng kê + HĐĐT.
2) BT5 đối chiếu HĐĐT ↔ bảng kê: cột trùng khớp / không khớp + thống kê.
3) BT7 cụm cùng ngày cùng NCC (ngưỡng 20tr đến 06/2025; 5tr từ 07/2025).
4) BT8/BT9 thuế theo tháng BKMV vs HĐĐT vs Nợ 133 NKC; khóa SoHD|MST / SoHD.
Lưu Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx vào cau-thanh; nhúng tóm tắt vào Master.
Tham chiếu skills/lam-sach-cd1-mst-hoa-don/references/excel-cd1-hv-bt5-bt9.md.
```

**Ra:** file chuyên đề `Doi_chieu_BT5_BT9_*.xlsx` (không phải mẫu 03).
**Bỏ qua khi:** đã chuẩn hóa trong pipeline trước và ghi rõ trên báo cáo.

---

## A2 - GTGT / TNDN / HĐĐT (`doi-chieu-gtgt-tndn-cbrr`)

**Cần:** XML GTGT + TNDN (+ bảng kê hoặc HĐĐT).

**Lệnh:**
```
Đối chiếu tờ khai GTGT - QTT TNDN - bảng kê - HĐĐT chuẩn CBRR
cho <TEN> MST <MST> tại <FOLDER> kỳ <KY>.
Xuất Tong_hop_doi_chieu_GTGT_TNDN_<MST>.xlsx
(2 sheet chuẩn CBRR nếu đủ dữ liệu).
Bóc tách kỳ BS, HĐ hủy/thay thế, nguyên nhân lệch.
Nếu có HĐĐT mua + BKMV + NKC: đối chiếu thêm khóa SoHD|MST và thuế BKMV vs PS Nợ 133* theo tháng (CD1 BT8/BT9).
```

**Ra:** `Tong_hop_doi_chieu_GTGT_TNDN_*.xlsx` + nhúng Master `Tong_hop_GTGT_TNDN`.
**Thiếu XML quý:** chạy năm có; liệt kê quý thiếu.

---

## A3 - NCC rủi ro (`doi-chieu-ncc-mua-vao-rui-ro`)

**Cần:** bảng kê mua và/hoặc HĐĐT mua + sổ CVCBRR.

**Lệnh:**
```
Đối chiếu NCC mua vào với CVCBRR và danh bạ cho <FOLDER>
Phần 1: NNT <MST> với sheet F1 và Bao_cao_cong_van.xlsx, ghi số CV + nguồn trích.
Phần 2: người bán sheet mua vao với sheet F0.
Xuất Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx đủ chuẩn Vạn Phát Thành:
- Sheet Tong_hop_doi_chieu: Phần 1 + Bảng 2.1 F0 + Bảng 2.2 ĐỦ 100% NCC + Phần 3 nhận xét
- Sheet Chi_tiet_hoa_don: đủ cột chuẩn, gắn cờ F0/ĐKKD
- Cột Nguồn dữ liệu = Bảng kê / HĐĐT / Chung; có Bảng 2.3 theo kỳ.
Không rút gọn chỉ lấy NCC bất thường.
```

**Ra:** file NCC 2 sheet + Master `Tong_hop_NCC_CVCBRR`.
**Thiếu danh bạ:** vẫn đối F0/F1/CA; ghi thiếu ĐKKD.

---

## A4 - TNCN (`tong-hop-tncn-da-nguon`)

**Cần:** XML QTT TNCN. Chính sách: skill chung `thue-tncn`.

**Lệnh:**
```
Tổng hợp QTT TNCN và danh sách NLĐ MST <MST> kỳ <KY>
từ XML tại <FOLDER>\1. DU LIEU THO NGANH THUE\1.1. To khai XML
xuất vào 1.4 và cau-thanh.
Nếu có file tra cứu đa nguồn ở 3.3 thì đối chiếu TC22/TC23, trùng CCCD.
Xuất Tong_hop_TNCN_va_Da_nguon_<MST>_<KY>.xlsx.
```

**Ra:** bảng TNCN + NLĐ + Master `Tong_hop_TNCN_va_Da_nguon`.
**Thiếu đa nguồn:** vẫn lập QTT; ghi thiếu `3.3`.

---

## A5 - 511 vs 632 (`doi-chieu-511-632`)

**Cần:** sổ chi tiết 511 (PS Có), 632 (PS Nợ). Tool: `skills/doi-chieu-511-632/assets/Tool_DC_511_632.xlsm`.

**Lệnh:**
```
Đối chiếu tự động doanh thu - giá vốn theo skill doi-chieu-511-632
cho công ty <TEN> MST <MST> tại <FOLDER> kỳ <KY>.
Xuất Doi_chieu_511_632_<MST>_<KY>.xlsx/.xlsm vào cau-thanh
và nhúng sheet Doi_chieu_511_632 vào Master.
Báo lãi gộp theo năm, dòng P<=0, dòng không khớp 632.
```

**Ra:** `Doi_chieu_511_632_*.xlsx` + tỷ lệ lãi gộp.
**Chưa có sổ:** treo bước này; vẫn làm A2 từ tờ khai.

---

## A6 - XNT kho (`phan-tich-xnt-kho`)

**Cần:** sổ XNT 152/155 (SX) hoặc 156 (thương mại) + BCTC.

**Lệnh:**
```
Phân tích XNT kho tại <FOLDER> kỳ <KY>.
DN thương mại: TK 156. DN sản xuất: 152+155 (+156 nếu có).
Đối chiếu tồn CK với BCTC: DNN dùng ct140; đầy đủ dùng [141] = 152+155+156.
Đối chiếu xuất kho với GV 632.
Xuất/cập nhật sheet Tong_hop_XNT_va_Kho, không để giá trị 0 nếu đã có sổ.
```

**Ra:** bảng XNT theo năm + Master `Tong_hop_XNT_va_Kho`.
**Bẫy:** đừng kết luận HTK=0 chỉ vì ct141=0 trên BCTC DNN.

---

## A7 - Giá thành / GV (`phan-tich-gia-thanh-dinh-muc`)

**Cần SX:** 154/621/622/627 + bảng giá thành + định mức.
**Cần thương mại:** GV 632 + CP BH/QL từ TNDN/BCTC.

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

**Ra:** Master `Gia_thanh_va_Dinh_muc`.

---

## A8 - Công nợ / tiền mặt (`phan-tich-cong-no-tien-mat`)

**Cần:** sổ 131/331/111/112; sao kê NH (Excel tốt hơn PDF).

**Lệnh:**
```
Phân tích công nợ 131/331 và quét chi tiền mặt TK 111 tại <FOLDER> kỳ <KY>.
Ngưỡng: >= 20 triệu đến 06/2025; >= 5 triệu từ 07/2025 (K5 đã chốt >=).
Quét cụm HĐ mua cùng ngày cùng NCC bằng SUMIFS (CD1 BT7A/BT7B); flag nhóm vượt ngưỡng.
Phân loại mục đích chi; nêu rõ sao kê đang PDF thì chưa kết luận thiếu UNC.
Cập nhật sheet Cong_no_va_Ngan_hang bằng số liệu thật.
```

**Ra:** `Phan_tich_Tien_mat_va_Ngan_hang_*.xlsx` + Master `Cong_no_va_Ngan_hang`.

---

## A9 - Master 8 sheet + 28 tiêu chí (`phan-tich-ho-so-thue`)

**Cần:** tối thiểu đã có A2. Có thể lập khung 28 TC sớm rồi cập nhật dần.

**Lệnh:**
```
Tái lập đầy đủ Master Excel theo chuẩn Vạn Phát Thành cho <FOLDER>
MST <MST> kỳ <KY>:
- Đủ 28/28 tiêu chí, 7 cột chuẩn, không rút gọn
- Đủ 8 sheet: Danh_gia_28_tieu_chi_rui_ro, Tong_hop_GTGT_TNDN,
  Tong_hop_NCC_CVCBRR, Tong_hop_TNCN_va_Da_nguon, Tong_hop_XNT_va_Kho,
  Gia_thanh_va_Dinh_muc, Doi_chieu_511_632, Cong_no_va_Ngan_hang
- Sheet nào thiếu dữ liệu: ghi rõ "Thiếu nguồn ..." chứ không để 0 giả
- Mỗi sheet có khối đánh giá / nhận xét / kết luận / trọng tâm kiểm tra
- Đồng bộ: 4. KET QUA PHAN TICH\, cau-thanh\, 7. Sao luu tong hop\, ket-qua\, sao-luu\
Map năm BCTC theo doanh thu nếu kyKKhai lệch.
```

**Ra:** `Phan_tich_ho_so_cong_ty_<TEN>_<MST>.xlsx`

---

## A10 - Đề cương (`phan-tich-ho-so-thue`)

**Cần:** Master đã có tiêu chí Cao/Trung bình.

**Lệnh:**
```
Từ Master <TEN> MST <MST>, lập đề cương kiểm tra tại trụ sở:
liệt kê tiêu chí Cao/Trung bình, mẫu chứng từ cần xem,
NCC/HĐ trọng điểm, nội dung đã giải tỏa.
```

---

## B1 - Mở hồ sơ TTKT (`mo-ho-so-ttkt`)

Checklist nghiệp vụ (BCTC, tờ khai, HĐĐT, GDLK, nhánh xử lý). Không sắp folder, không chạy engine XML.

**Lệnh:**
```
Mở hồ sơ TTKT cho <TEN> MST <MST> kỳ <KY> tại <FOLDER>
theo skill mo-ho-so-ttkt, checklist A→F.
Đối chiếu BCTC - tờ khai - HĐĐT, GDLK nếu có.
Phân nhánh thủ tục / thiếu thuế / trốn thuế.
Liệt kê 5-10 vấn đề Cao/Trung/Thấp và hồ sơ còn thiếu.
Không sắp thư mục, không xuất Master 8 sheet.
```

---

## B2 - Bên mua trên sổ CVCBRR (`tong-hop-ben-mua-tu-so-cvcbrr`)

Khác với đối chiếu NCC bên bán của 1 DN (`doi-chieu-ncc-mua-vao-rui-ro`).

**Lệnh:**
```
Tổng hợp đơn vị bên mua từ sổ CVCBRR
file SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx sheet THEO_DOI.
Xuất Chi_tiet (chỉ dòng đã có bên mua, sắp ngày CV)
và Tong_hop theo MST mua.
Font Times New Roman. Không bịa doanh thu nếu sổ trống cột tiền.
```

---

## B3 - QLRR chọn đối tượng (`quan-ly-rui-ro-tuan-thu-thue`)

Khung chọn NNT / bộ chỉ số TCT. Không mổ 1 bộ sổ DN.

**Lệnh:**
```
Áp khung QLRR tuân thủ thuế (Luật 108/2025, TT 31/2021)
cho nghiệp vụ <TTKT / hoàn thuế / hóa đơn / hộ KD>.
Nêu bộ chỉ số TCT đúng nghiệp vụ, mức tuân thủ/rủi ro,
biện pháp hỗ trợ / giám sát / kiểm tra.
Không chạy đối chiếu XML/HĐĐT của một MST cụ thể.
```

---

## Gói lệnh siêu ngắn

| Skill | Câu siêu ngắn |
|---|---|
| `to-chuc-ho-so-kiem-tra-thue` | `Tổ chức / kiểm kê checklist <FOLDER>` |
| `tong-hop-dau-vao-hddt-bang-ke-nld` | `Tổng hợp đầu vào HĐĐT-Bảng kê-NLĐ <FOLDER>` |
| `tong-hop-dau-vao-hddt-bang-ke-nld` | `Gộp bảng kê <FOLDER> vào 2.5` |
| `lam-sach-cd1-mst-hoa-don` | `Chuẩn hóa MST + quét HĐ trùng <FOLDER>` |
| `doi-chieu-gtgt-tndn-cbrr` | `Đối chiếu CBRR GTGT-TNDN-HĐĐT <FOLDER>` |
| `lam-sach-cd1-mst-hoa-don` | `Đối chiếu BKMV vs HĐĐT vs Nợ 133 <FOLDER>` |
| `doi-chieu-ncc-mua-vao-rui-ro` | `Đối chiếu NCC F0/F1/CA chuẩn VPT <FOLDER>` |
| `tong-hop-tncn-da-nguon` | `Tổng hợp TNCN <MST> <FOLDER>` |
| `doi-chieu-511-632` | `Đối chiếu 511-632 <FOLDER>` |
| `phan-tich-xnt-kho` | `Phân tích XNT kho <FOLDER>` |
| `phan-tich-gia-thanh-dinh-muc` | `Phân tích giá thành/định mức <FOLDER>` hoặc `Cơ cấu GV thương mại <FOLDER>` |
| `phan-tich-cong-no-tien-mat` | `Quét 111, cụm cùng ngày NCC và công nợ 131/331 <FOLDER>` |
| `phan-tich-ho-so-thue` | `Tái lập Master 8 sheet đủ 28 TC chuẩn VPT <FOLDER>` |
| `phan-tich-ho-so-thue` | `Lập đề cương từ Master <FOLDER>` |
| `mo-ho-so-ttkt` | `Checklist mở hồ sơ TTKT A-F <FOLDER>` |
| `tong-hop-ben-mua-tu-so-cvcbrr` | `Tổng hợp bên mua sổ CVCBRR` |
| `quan-ly-rui-ro-tuan-thu-thue` | `Khung QLRR chọn đối tượng nghiệp vụ <X>` |

---

## Nguyên tắc

1. **Không chờ đủ mới chạy** - module nào đủ đầu vào thì chạy trước.
2. **Một lệnh = một file ra.**
3. **Thiếu thì ghi thiếu**; không để 0 giả.
4. Skill trục `phan-tich-ho-so-thue` giữ Master 8 sheet, 28 TC, đề cương. Không lặp thân skill vệ tinh.
