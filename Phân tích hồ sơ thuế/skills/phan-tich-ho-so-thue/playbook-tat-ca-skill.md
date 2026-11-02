---
type: wiki
title: Playbook tất cả skill phân tích hồ sơ thuế
updated: 2026-10-23
---

# Playbook tất cả skill phân tích hồ sơ thuế

Catalog lệnh copy-paste. **Chuẩn v1.4 vòng 2:** đầu ra giao sáu tệp KQ1.1-KQ3; GĐ3 thêm 2 file phụ (dấu hiệu + giải trình NNT 10 sheet). Skill A0-B3 là module nội bộ, ghi đúng sheet đích, không giao tệp chuyên đề kiểu cũ.

**Phòng 6 thư ký:** Thủ hồ sơ GĐ1, Dữ liệu đầu vào, KQ2.1 CQT, KQ2.2 sổ công ty, Kết luận GĐ3, Trưởng phòng. Workflow 7 bước `phan-tich-ho-so-thue-3-giai-doan`.

Thay placeholder:

- `<FOLDER>` = đường dẫn tuyệt đối thư mục DN (`...\<TEN>_<MST>`)
- `<MST>`, `<TEN>`, `<KY>` = mã số thuế, tên ngắn ASCII, niên độ (vd `2022-2025`)

Thiếu nguồn thì ô trống, không để số 0 giả. Ô text Excel không bắt đầu bằng `=`. Font Times New Roman.

Playbook chạy theo bước: `playbook-lenh-theo-buoc.md`. Lệnh 1-4: `playbook-3-giai-doan.md`.

Song song được khi đã có dữ liệu: **GTGT + NCC + TNCN**; **511-632 + XNT + công nợ**.

---

## Bản đồ skill

### A. Ba skill giai đoạn chủ lực (Pipeline Skills v1.4)

| Giai đoạn | Skill | Việc chính | Đầu ra |
|---|---|---|---|
| **1** | `giai-doan-1-kiem-ke-danh-gia` | Kiểm kê IN00-IN14, mức A-E, 80 quy tắc, đề nghị bổ sung | KQ1.1, KQ1.2, KQ1.3 |
| **2** | `giai-doan-2-phan-tich-du-lieu-co-quan-thue` | Số liệu, chênh lệch, tỷ lệ, cờ máy; lập IN07 | KQ2.1, KQ2.2 |
| **3** | `giai-doan-3-hoan-chinh-ho-so` | 80 quy tắc + 28 tiêu chí, chống trùng T1-T10, dấu hiệu + giải trình NNT | KQ3 2 sheet + 2 file phụ |

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
Ghi KQ1.1!IN01,IN05 và KQ1.2!IN02. Không giao file mẫu 02.
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

**Ra:** phiên bản mới của KQ1.1 và/hoặc KQ1.2.
**Thiếu 1 nguồn:** ô trống, không điền 0.

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
Gắn cột vào IN01/IN02/KQ2. Không giao Doi_chieu_BT5_BT9.
Tham chiếu skills/lam-sach-cd1-mst-hoa-don/references/excel-cd1-hv-bt5-bt9.md.
```

**Ra:** phiên bản mới của KQ1.1/KQ1.2 (cột CD1 trên IN).
**Bỏ qua khi:** đã chuẩn hóa trong pipeline trước và ghi rõ trên báo cáo.

---

## A2 - GTGT / TNDN / HĐĐT (`doi-chieu-gtgt-tndn-cbrr`)

**Cần:** XML GTGT + TNDN (+ bảng kê hoặc HĐĐT).

**Lệnh:**
```
Đối chiếu tờ khai GTGT - QTT TNDN - bảng kê - HĐĐT chuẩn CBRR
cho <TEN> MST <MST> tại <FOLDER> kỳ <KY>.
Ghi KQ2.1!KQ2_Doi_chieu_GTGT (số, chênh lệch, cờ máy; nhận xét nguyên nhân trọng yếu khi lệch lớn - 3.4).
Bóc tách kỳ BS, HĐ hủy/thay thế. Không xuất Tong_hop_doi_chieu_GTGT_TNDN.
Nếu có HĐĐT mua + BKMV + NKC: đối chiếu thêm khóa SoHD|MST và thuế BKMV vs PS Nợ 133* theo tháng (CD1 BT8/BT9).
```

**Ra:** phiên bản mới của KQ2.1.
**Thiếu XML quý:** chạy năm có; liệt kê quý thiếu.

---

## A3 - NCC rủi ro (`doi-chieu-ncc-mua-vao-rui-ro`)

**Cần:** bảng kê mua và/hoặc HĐĐT mua + sổ CVCBRR.

**Lệnh:**
```
Đối chiếu NCC mua vào với CVCBRR và danh bạ cho <FOLDER>
Phần 1: NNT <MST> với sheet F1 và Bao_cao_cong_van.xlsx, ghi số CV + nguồn trích.
Phần 2: người bán sheet mua vao với sheet F0.
Ghi KQ2.1!IN07_NCC_CanhBao và KQ2_NCC_mua_vao. 100% NCC. Giữ HD04 Top-5 gộp theo người bán. Nhận xét module Phần 1/3 (3.4); mức Cao chỉ ở KQ3.
Không xuất Tong_hop_NCC_mua_vao_CV_trang_thai. Không rút gọn.
```

**Ra:** phiên bản mới của KQ2.1.
**Thiếu danh bạ:** vẫn đối F0/F1/CA; ghi thiếu ĐKKD.

---

## A4 - TNCN (`tong-hop-tncn-da-nguon`)

**Cần:** XML QTT TNCN. Chính sách: skill chung `thue-tncn`.

**Lệnh:**
```
Tổng hợp QTT TNCN và danh sách NLĐ MST <MST> kỳ <KY>
từ XML tại <FOLDER>\1. DU LIEU THO NGANH THUE\1.1. To khai XML
Ghi KQ2.1!KQ2_TNCN_LD2. Kèm nhận xét module ngắn (3.4).
Nếu có file tra cứu đa nguồn ở 3.3 thì đối chiếu TC22/TC23, trùng CCCD.
Không xuất Tong_hop_TNCN_va_Da_nguon.
```

**Ra:** phiên bản mới của KQ2.1.
**Thiếu đa nguồn:** vẫn lập QTT; ghi thiếu `3.3`.

---

## A5 - 511 vs 632 (`doi-chieu-511-632`)

**Cần:** sổ chi tiết 511 (PS Có), 632 (PS Nợ). Tool: `skills/doi-chieu-511-632/assets/Tool_DC_511_632.xlsm`.

**Lệnh:**
```
Đối chiếu tự động doanh thu - giá vốn theo skill doi-chieu-511-632
cho công ty <TEN> MST <MST> tại <FOLDER> kỳ <KY>.
Ghi KQ2.2!KQ2_511_632 (tổng hợp năm + Đánh giá tuân thủ) và KQ2_511_632_Chi_tiet (bán dưới giá). Không xuất Doi_chieu_511_632_*.
Báo lãi gộp theo năm, dòng P<=0, dòng không khớp 632 (cờ máy).
```

**Ra:** phiên bản mới của KQ2.2.
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
Ghi KQ2.2!KQ2_Kho_Gia_thanh, không để giá trị 0 nếu đã có sổ.
```

**Ra:** phiên bản mới của KQ2.2.
**Bẫy:** đừng kết luận HTK=0 chỉ vì ct141=0 trên BCTC DNN.

---

## A7 - Giá thành / GV (`phan-tich-gia-thanh-dinh-muc`)

**Cần SX:** 154/621/622/627 + bảng giá thành + định mức.
**Cần thương mại:** GV 632 + CP BH/QL từ TNDN/BCTC.

**Lệnh SX:**
```
Phân tích giá thành và định mức tại <FOLDER> kỳ <KY>.
Bóc 621/622/627, khớp GV với BCTC [02] và TK 632.
Ghi KQ2.2!KQ2_Gia_thanh_va_Dinh_muc.
```

**Lệnh thương mại:**
```
DN thương mại: lập sheet Gia_thanh_va_Dinh_muc theo cơ cấu
Doanh thu - GV 632 - lãi gộp - chi phí bán hàng/QLDN (từ TNDN),
không bắt buộc 621/622/627. Ghi KQ2.2!KQ2_Gia_thanh_va_Dinh_muc.
```

**Ra:** phiên bản mới của KQ2.2.

---

## A8 - Công nợ / tiền mặt (`phan-tich-cong-no-tien-mat`)

**Cần:** sổ 131/331/111/112; sao kê NH (Excel tốt hơn PDF).

**Lệnh:**
```
Phân tích công nợ 131/331 và quét chi tiền mặt TK 111 tại <FOLDER> kỳ <KY>.
Ngưỡng: >= 20 triệu đến 06/2025; >= 5 triệu từ 07/2025 (K5 đã chốt >=).
Quét cụm HĐ mua cùng ngày cùng NCC bằng SUMIFS (CD1 BT7A/BT7B); flag nhóm vượt ngưỡng.
Phân loại mục đích chi; nêu rõ sao kê đang PDF thì chưa kết luận thiếu UNC.
Ghi KQ2.2!KQ2_Cong_no_Tien_mat bằng số liệu thật. Không xuất file chuyên đề.
```

**Ra:** phiên bản mới của KQ2.2.

---

## A9-A11 - GĐ3 đầy đủ (`giai-doan-3-hoan-chinh-ho-so`)

Sheet `Danh_gia_28_tieu_chi_rui_ro` nằm trong KQ3 là chuẩn vòng 2. Không xuất Master 8 sheet tách ngoài.

**Lệnh:**
```
Hoàn chỉnh GĐ3 hồ sơ <FOLDER> MST <MST> kỳ <KY> xuất KQ3 và chi tiết giải trình.
1) Sheet KQ3 80 quy tắc + sheet Danh_gia_28_tieu_chi_rui_ro.
2) Đề cương cột AB, AC, AH; AD để trống cho cán bộ. Không tách KQ3_De_cuong kiểu cũ.
3) Xuất KQ3_Danh_gia_Dau_hieu_Bat_thuong và KQ3_NNT_Giai_trinh_Chi_tiet (10 sheet có số liệu thật).
Áp Q14-Q20.
```

**Ra:** KQ3 2 sheet + 2 file phụ GĐ3.

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
Không sắp thư mục, không xuất Master 8 sheet, không xuất workbook trọn gói.
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
| `giai-doan-1-kiem-ke-danh-gia` | `Kiểm kê GĐ1 <FOLDER> MST <MST> kỳ <KY> xuất KQ1.1 KQ1.2 KQ1.3` |
| `giai-doan-2-phan-tich-du-lieu-co-quan-thue` | `Phân tích GĐ2 <FOLDER> xuất KQ2.1 KQ2.2` |
| `giai-doan-3-hoan-chinh-ho-so` | `Hoàn chỉnh GĐ3 <FOLDER> xuất KQ3 và chi tiết giải trình` |
| `mo-ho-so-ttkt` | `Checklist mở hồ sơ TTKT A-F <FOLDER>` |
| `tong-hop-ben-mua-tu-so-cvcbrr` | `Tổng hợp bên mua sổ CVCBRR` |
| `quan-ly-rui-ro-tuan-thu-thue` | `Khung QLRR chọn đối tượng nghiệp vụ <X>` |

---

## Nguyên tắc

1. **Không chờ đủ mới chạy** - module nào đủ đầu vào thì chạy trước.
2. **Một lệnh = một phiên bản mới của tệp đích** trong sáu tệp 1.4.
3. **Thiếu thì ô trống**; không để 0 giả.
4. Skill trục `phan-tich-ho-so-thue` giữ Hướng dẫn thường trực P7. Không giao mẫu cũ, PL, workbook trọn gói, Master 8 sheet tách ngoài. Sheet 28 tiêu chí nằm trong KQ3 là chuẩn.
