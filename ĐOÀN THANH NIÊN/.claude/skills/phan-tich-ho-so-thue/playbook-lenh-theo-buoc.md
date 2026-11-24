---
type: wiki
title: Playbook lệnh theo bước - Phân tích hồ sơ thuế
updated: 2026-10-15
---

# Playbook lệnh theo bước (công ty mới)

Catalog mọi skill (kèm skill cạnh QLRR / mở hồ sơ / bên mua CVCBRR): `playbook-tat-ca-skill.md`.

Mục tiêu: làm **từng module riêng** khi chưa đủ hồ sơ, không chờ đủ 100% dữ liệu rồi mới chạy. Mỗi bước có đầu vào tối thiểu, câu lệnh copy-paste, file ra, và điều kiện bỏ qua.

Thay placeholder:
- `<FOLDER>` = đường dẫn tuyệt đối thư mục DN (`...\<TEN>_<MST>`)
- `<MST>`, `<TEN>`, `<KY>` = mã số thuế, tên ngắn, niên độ (vd `2022-2025`)

Chuẩn Master cuối: **8 sheet**, 28 tiêu chí đủ 7 cột, font Times New Roman. Không dùng bản rút gọn.
Kỹ thuật Excel CD1 (MST, HĐ trùng, cụm cùng ngày, BKMV-NKC 133): skill `lam-sach-cd1-mst-hoa-don` và `skills/phan-tich-ho-so-thue/references/excel-cd1-hv-bt5-bt9.md`.

## Thứ tự tối ưu (đường nhanh) và Ánh xạ 3 giai đoạn

Quy trình module hóa được ánh xạ vào **3 giai đoạn** (chi tiết xem `playbook-3-giai-doan.md`):

- **Giai đoạn 1 (Kiểm kê & Đánh giá mức A–E):** Bước 0. Skill chủ: `giai-doan-1-kiem-ke-danh-gia`. Xuất `KQ1_<MST>_<KY>_v1.xlsx`.
- **Giai đoạn 2 (Dữ liệu CQT, 57 quy tắc khi chưa có sổ DN):** Bước 0b, 1, 1b, 2, 3, 4. Skill chủ: `giai-doan-2-phan-tich-du-lieu-co-quan-thue`. Xuất `KQ2_<MST>_<KY>_v1.xlsx`.
- **Giai đoạn 3 (Hoàn chỉnh hồ sơ & Sổ sách DN, 23 quy tắc):** Bước 5, 6, 7, 8, 9, 10. Skill chủ: `giai-doan-3-hoan-chinh-ho-so`. Xuất `KQ3_<MST>_<KY>_v1.xlsx`.

```
0 Kiểm kê (GĐ1) → 0b Tổng hợp đầu vào (HĐĐT + Bảng kê + NLĐ)
→ 1 Bảng kê (nếu chưa gộp) → 1b Làm sạch MST/HĐ trùng → 2 GTGT/TNDN/HĐĐT → 3 NCC rủi ro
→ 4 TNCN (Hết GĐ2) → 5 511-632 → 6 XNT → 7 Giá vốn/Giá thành
→ 8 Công nợ/Tiền mặt (cụm cùng ngày) → 9 Master 28 TC → 10 Đề cương (Hết GĐ3)
```

Làm song song được khi đã có dữ liệu: **(2)+(3)+(4)**; **(5)+(6)+(8)**.

---

## Bước 0 - Kiểm kê & checklist (luôn làm trước)

**Skill:** `to-chuc-ho-so-kiem-tra-thue` (dựng folder 4 bước, phân loại file, 1 checklist A4, xóa nguồn cũ). Playbook này không lặp lại cây thư mục.

**Cần:** đường dẫn tuyệt đối; nếu folder chưa chuẩn thì skill kia tạo `<TEN>_<MST>` trước.

**Lệnh:**
```
Tổ chức / kiểm kê hồ sơ tại <FOLDER>
MST <MST>, kỳ <KY> theo skill to-chuc-ho-so-kiem-tra-thue.
Xuất 1 checklist A4 tại 4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx.
Nêu module phân tích nào chạy được ngay.
```

**Ra:** folder chuẩn + checklist + danh sách module sẵn sàng.
**Thiếu gì thì dừng:** chưa có folder tuyệt đối / chưa biết MST-kỳ.

---

## Bước 0b - Tổng hợp dữ liệu đầu vào (HĐĐT + Bảng kê + NLĐ)

**Skill:** `tong-hop-dau-vao-hddt-bang-ke-nld` (gộp 3 nguồn độc lập, mẫu 02; không trộn HĐĐT với bảng kê trên 1 sheet chi tiết).

**Cần tối thiểu:** folder `1.2` (HĐĐT) và/hoặc `2.5` (bảng kê); XML/phụ lục TNCN nếu có NLĐ.

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

**Ra:** 1 file 7 sheet (Huong_dan, Tom_tat, 4 sheet HĐ/BK, Tong_hop_NLD).
**Thiếu một nguồn:** vẫn xuất các sheet có dữ liệu; đánh Thiếu/Mỏng trên Tom_tat, không để số 0 giả.

---

## Bước 1 - Gộp bảng kê mua vào / bán ra

**Skill:** `tong-hop-dau-vao-hddt-bang-ke-nld` (phần bảng kê 2 sheet; lọc Tổng cộng).

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

## Bước 1b - Làm sạch MST & đối chiếu HĐ (CD1 BT5-BT9, không tách mẫu 03)

**Skill:** `lam-sach-cd1-mst-hoa-don`. Công thức đầy đủ: `skills/lam-sach-cd1-mst-hoa-don/references/excel-cd1-hv-bt5-bt9.md`.

**Cần:** bảng kê hoặc HĐĐT đã gộp; tốt hơn có NKC/sổ cái để BT8/BT9.

**Chốt 2026-10-11:** không xuất riêng file `03_Lam_sach_CD1_BT5_BT9` làm kết quả DN. Gắn kỹ thuật vào bảng biểu nghiệp vụ:
- **BT6** → cột `MST_chuan` trên bảng kê mua vào/bán ra (và mọi hồ sơ có MST công ty gửi).
- **BT5** → đối chiếu HĐĐT CQ thuế ↔ bảng kê NNT: cột `KHOP` / `CHI_CO_BANG_KE` / `CHI_CO_HDDT` + thống kê không khớp.
- **BT7** → cụm cùng ngày cùng MST vượt ngưỡng, gắn sheet `Cong_no_va_Ngan_hang` (sao kê NH nếu có).
- **BT8/BT9** → thuế GTGT mua vào theo tháng/năm: bảng kê ↔ HĐĐT (loại HĐ hủy/thay thế) ↔ PS Nợ TK `133*`; gắn `Tong_hop_GTGT_TNDN` Phần III.

**Lệnh:**
```
Chạy CD1 BT5-BT9 tại <FOLDER> kỳ <KY> và gắn vào bảng biểu (không tạo mẫu 03 tách):
1) BT6 chuẩn hóa MST trên bảng kê + HĐĐT.
2) BT5 đối chiếu HĐĐT ↔ bảng kê: cột trùng khớp / không khớp + thống kê.
3) BT7 cụm cùng ngày cùng NCC (ngưỡng 20tr đến 06/2025; 5tr từ 07/2025).
4) BT8/BT9 thuế theo tháng BKMV vs HĐĐT vs Nợ 133 NKC; khóa SoHD|MST / SoHD.
Lưu Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx vào cau-thanh; nhúng tóm tắt vào Master.
Tham chiếu references/excel-cd1-hv-bt5-bt9.md.
```

**Ra:** file chuyên đề trong `cau-thanh/` + cột/phụ lục trên Master; không cần file 03 riêng.
**Bỏ qua khi:** đã chuẩn hóa/đối chiếu trong pipeline trước đó và ghi rõ trong báo cáo.

---

## Bước 2 - Đối chiếu GTGT / TNDN / HĐĐT (CBRR)

**Skill:** `doi-chieu-gtgt-tndn-cbrr` (2 sheet chuẩn CBRR). CD1 BT8/BT9 mua vào vs 133: `lam-sach-cd1-mst-hoa-don`.

**Cần:** XML GTGT + TNDN (+ Bảng kê hoặc HĐĐT).

**Lệnh:**
```
Đối chiếu tờ khai GTGT - QTT TNDN - bảng kê - HĐĐT chuẩn CBRR
cho <TEN> MST <MST> tại <FOLDER>
kỳ <KY>.
Xuất Tong_hop_doi_chieu_GTGT_TNDN_<MST>.xlsx
(2 sheet chuẩn CBRR nếu đủ dữ liệu).
Bóc tách kỳ BS, HĐ hủy/thay thế, nguyên nhân lệch.
Nếu có HĐĐT mua + BKMV + NKC: đối chiếu thêm khóa SoHD|MST và thuế BKMV vs PS Nợ 133* theo tháng (CD1 BT8/BT9).
```

**Ra:** bảng đối chiếu theo quý/năm + chênh lệch.
**Thiếu XML quý:** vẫn chạy các năm có; liệt kê quý thiếu, không bịa số.

---

## Bước 3 - Đối chiếu NCC rủi ro (F0/F1/Công an/Danh bạ)

**Skill:** `doi-chieu-ncc-mua-vao-rui-ro`.

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

**Skill:** `tong-hop-tncn-da-nguon`. Chính sách TNCN: skill chung `thue-tncn`.

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

**Skill:** `doi-chieu-511-632`.

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

**Skill:** `phan-tich-xnt-kho`.

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

**Skill:** `phan-tich-gia-thanh-dinh-muc`.

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

**Skill:** `phan-tich-cong-no-tien-mat`. Cụm cùng ngày CD1 BT7: `lam-sach-cd1-mst-hoa-don`.

**Cần:** sổ 131/331/111/112; sao kê NH (Excel tốt hơn PDF).

**Lệnh:**
```
Phân tích công nợ 131/331 và quét chi tiền mặt TK 111 tại <FOLDER> kỳ <KY>.
Ngưỡng: >= 20 triệu đến 06/2025; >= 5 triệu từ 07/2025 (K5 đã chốt >=).
Quét cụm HĐ mua cùng ngày cùng NCC bằng SUMIFS (CD1 BT7A/BT7B); flag nhóm vượt ngưỡng.
Phân loại mục đích chi; nêu rõ sao kê đang PDF thì chưa kết luận thiếu UNC.
Cập nhật sheet Cong_no_va_Ngan_hang bằng số liệu thật.
```

---

## Bước 9 - Tái lập Master 8 sheet + đủ 28 tiêu chí

**Quyết định K3:** Nếu không tìm thấy mẫu thì bỏ qua, chỉ dựng sheet 28 tiêu chí (`08_Master_28TC`), không tự dựng 7 sheet còn lại và không báo như lỗi.

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

**Bắt buộc trên Master (chuẩn Vạn Phát Thành):** mỗi sheet chuyên đề phải có khối **đánh giá / nhận xét / kết luận / trọng tâm kiểm tra / hồ sơ cần giải trình** - không chỉ bảng số. Thiếu dữ liệu thì viết hiện trạng `thiếu/chưa đủ đầu vào`, không bỏ trống cả khối kết luận. Mẫu `06_*_MAU.xlsx` đã có sẵn các khung này.

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

| Module | Skill | Câu siêu ngắn |
|---|---|---|
| Kiểm kê | `to-chuc-ho-so-kiem-tra-thue` | `Tổ chức / kiểm kê checklist <FOLDER>` |
| Đầu vào HĐĐT+BK+NLĐ | `tong-hop-dau-vao-hddt-bang-ke-nld` | `Tổng hợp đầu vào HĐĐT-Bảng kê-NLĐ <FOLDER>` |
| Bảng kê | `tong-hop-dau-vao-hddt-bang-ke-nld` | `Gộp bảng kê <FOLDER> vào 2.5` |
| Làm sạch khóa | `lam-sach-cd1-mst-hoa-don` | `Chuẩn hóa MST + quét HĐ trùng <FOLDER>` |
| GTGT-TNDN | `doi-chieu-gtgt-tndn-cbrr` | `Đối chiếu CBRR GTGT-TNDN-HĐĐT <FOLDER>` |
| BKMV-NKC 133 | `lam-sach-cd1-mst-hoa-don` | `Đối chiếu BKMV vs HĐĐT vs Nợ 133 <FOLDER>` |
| NCC | `doi-chieu-ncc-mua-vao-rui-ro` | `Đối chiếu NCC F0/F1/CA chuẩn VPT <FOLDER>` |
| TNCN | `tong-hop-tncn-da-nguon` | `Tổng hợp TNCN <MST> <FOLDER>` |
| 511-632 | `doi-chieu-511-632` | `Đối chiếu 511-632 skill mới <FOLDER>` |
| XNT | `phan-tich-xnt-kho` | `Phân tích XNT kho <FOLDER>` |
| Giá thành | `phan-tich-gia-thanh-dinh-muc` | `Phân tích giá thành/định mức <FOLDER>` hoặc `Cơ cấu GV thương mại <FOLDER>` |
| Công nợ/TM | `phan-tich-cong-no-tien-mat` | `Quét 111, cụm cùng ngày NCC và công nợ 131/331 <FOLDER>` |
| Master 28 | `phan-tich-ho-so-thue` | `Tái lập Master 8 sheet đủ 28 TC chuẩn VPT <FOLDER>` |
| Đề cương | `phan-tich-ho-so-thue` | `Lập đề cương từ Master <FOLDER>` |

---

## Nguyên tắc tốc độ

1. **Không chờ đủ mới chạy** - chạy module nào đủ đầu vào trước.
2. **Một lệnh = một file ra** - dễ nối tiếp và kiểm chứng.
3. **Thiếu thì ghi thiếu** trên sheet/báo cáo; không để 0 giả.
4. **Song song:** sau Bước 1 có thể giao đồng thời Bước 2+3+4.
5. **Chốt Master** chỉ khi 28 TC đủ 7 cột và các sheet có số thật hoặc nhãn thiếu nguồn rõ ràng.
