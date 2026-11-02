---
type: wiki
title: Playbook lệnh theo bước - Phân tích hồ sơ thuế v1.4
updated: 2026-10-23
---

# Playbook lệnh theo bước (v1.4 vòng 2)

Catalog đủ skill: `playbook-tat-ca-skill.md`. Lệnh giai đoạn: `playbook-3-giai-doan.md`.

Một lệnh bước = một phiên bản mới của **tệp đích** trong bộ 1.4 vòng 2. Không tạo tệp chuyên đề (`Tong_hop_*`, `Doi_chieu_*`, Master 8 sheet tách ngoài) để giao. Quy tắc chưa chạy ghi `CHƯA_CHẠY`.

**Agent mặc định:** bước 0 → `thu-ho-so-kiem-ke`; 0b/1/1b → `chuyen-vien-du-lieu-dau-vao`; 2-4 + EBITDA → `phan-tich-vien-cqt`; 5-8 → `phan-tich-vien-so-cty`; 9-11 → `phan-tich-vien-so-sach-dn`.

Placeholder: `<FOLDER>`, `<MST>`, `<TEN>`, `<KY>`.

## Ánh xạ bước -> sheet đích

| Bước | Skill | Tệp đích | Sheet |
|---|---|---|---|
| 0 | `to-chuc-ho-so-kiem-tra-thue` + `giai-doan-1-kiem-ke-danh-gia` | KQ1.3 | `KQ1_Kiem_ke` |
| 0b, 1, 1b | `tong-hop-dau-vao-hddt-bang-ke-nld`, `lam-sach-cd1-mst-hoa-don` | KQ1.1 / KQ1.2 | IN01, IN02, IN05 |
| 2 | `doi-chieu-gtgt-tndn-cbrr` | KQ2.1 | `KQ2_Doi_chieu_GTGT` (+ nhận xét 3.4) |
| 2b | `giai-doan-2-phan-tich-du-lieu-co-quan-thue` | KQ2.1 | `KQ2_EBITDA_Phan_tich` |
| 3 | `doi-chieu-ncc-mua-vao-rui-ro` | KQ2.1 | `IN07_NCC_CanhBao`, `KQ2_NCC_mua_vao` (HD04 bắt buộc) |
| 4 | `tong-hop-tncn-da-nguon` | KQ2.1 | `KQ2_TNCN_LD2` |
| 5 | `doi-chieu-511-632` | KQ2.2 | `KQ2_511_632`, `KQ2_511_632_Chi_tiet` |
| 6, 7 | `phan-tich-xnt-kho`, `phan-tich-gia-thanh-dinh-muc` | KQ2.2 | `KQ2_Kho_Gia_thanh`, `KQ2_Gia_thanh_va_Dinh_muc` |
| 8 | `phan-tich-cong-no-tien-mat` | KQ2.2 | `KQ2_Cong_no_Tien_mat` |
| 9-10 | `giai-doan-3-hoan-chinh-ho-so` | KQ3 | sheet `KQ3` + `Danh_gia_28_tieu_chi_rui_ro`; cột AB-AD, AH |
| 11 | `giai-doan-3-hoan-chinh-ho-so` | phụ lục GĐ3 | `KQ3_Danh_gia_Dau_hieu_Bat_thuong`, `KQ3_NNT_Giai_trinh_Chi_tiet` (10 sheet) |

Khi chạy theo phòng 6 vai: gộp bước 2-4+2b vào một lệnh KQ2.1; gộp 5-8 vào một lệnh KQ2.2; gộp 9-11 vào một lệnh GĐ3.

## Bước 0 - Kiểm kê

**Lệnh:**
```
Tổ chức / kiểm kê hồ sơ tại <FOLDER> MST <MST> kỳ <KY>.
Sao mẫu KQ1.1 KQ1.2 KQ1.3 rồi điền. Checklist A4 nếu cần tổ chức folder thì ghi thêm vào KQ1_Kiem_ke, không giao CHECKLIST_TONG_HOP_HO_SO như sản phẩm phân tích.
Nêu module nào chạy được ngay.
```

**Ra:** KQ1.1, KQ1.2, KQ1.3 v1. **Thiếu folder/MST/kỳ:** dừng hoặc hỏi một câu.

## Bước 0b / 1 / 1b - Đầu vào HĐĐT, bảng kê, NLĐ, CD1

**Lệnh:**
```
Tổng hợp đầu vào và CD1 BT5-BT9 cho <FOLDER> MST <MST> kỳ <KY>.
Ghi HĐĐT vào KQ1.1!IN01; bảng kê vào KQ1.2!IN02 (đủ mua+bán, bỏ Tổng cộng, MST_chuan, Trang_thai_DC); NLĐ/QTT vào KQ1.1!IN05 (Loai_dong NLD|QTT).
Không trộn nguồn. Không xuất file mẫu 02 hay Doi_chieu_BT5_BT9 để giao.
Tạo đúng một phiên bản mới của KQ1.1 và/hoặc KQ1.2.
```

**Thiếu một nguồn:** sheet có dữ liệu vẫn điền; ô kia trống, không điền 0.

## Bước 2 - GTGT / TNDN

**Lệnh:**
```
Đối chiếu GTGT-TNDN-HĐĐT cho <FOLDER> MST <MST> kỳ <KY>.
Ghi vào KQ2.1!KQ2_Doi_chieu_GTGT. Số, chênh lệch, cờ Khớp|Lệch|Chưa đủ nguồn; khối nhận xét nguyên nhân trọng yếu khi lệch lớn (3.4).
Kỳ trống bỏ dòng. Ô bảng kê/sổ trống nếu chưa có IN02/IN08.
Một phiên bản mới của KQ2.1. Không xuất Tong_hop_doi_chieu_GTGT_TNDN.
```

## Bước 3 - NCC

**Lệnh:**
```
Đối chiếu 100% NCC F0/F1/CA cho <FOLDER> MST <MST> kỳ <KY>.
Ghi IN07 và KQ2_NCC_mua_vao trong KQ2.1. Giữ HD04 Top-5 gộp theo MST người bán. Nhận xét module Phần 1/3 (3.4).
F0 và F1 đều cờ nhóm F (mức Cao chỉ ở KQ3). Không rút gọn. Không xuất Tong_hop_NCC.
Một phiên bản mới của KQ2.1. Quy tắc khác CHƯA_CHẠY.
```

## Bước 4 - TNCN

**Lệnh:**
```
Tổng hợp QTT TNCN và NLĐ MST <MST> kỳ <KY> tại <FOLDER>.
Ghi KQ2.1!KQ2_TNCN_LD2. Có 3.3 thì đối TC22/TC23. Thiếu 3.3: vẫn lập QTT, ô đa nguồn trống.
Kèm nhận xét module ngắn (3.4). Không xuất Tong_hop_TNCN. Một phiên bản mới của KQ2.1.
```

## Bước 5 - 511-632

**Lệnh:**
```
Đối chiếu 511-632 theo skill doi-chieu-511-632 cho <FOLDER> MST <MST> kỳ <KY>.
Ghi KQ2.2!KQ2_511_632 (tổng hợp năm + Đánh giá tuân thủ & Rủi ro) và KQ2_511_632_Chi_tiet (bán dưới giá).
Chưa có sổ: để trống, không suy từ tờ khai. Không xuất Doi_chieu_511_632_*. Một phiên bản mới của KQ2.2.
```

## Bước 6 - XNT kho

**Lệnh:**
```
Phân tích XNT kho tại <FOLDER> kỳ <KY>.
Ghi KQ2.2!KQ2_Kho_Gia_thanh. B01-DN [141]; B01-DNN ct140 (K6). Nhận xét module 3.4.
Thiếu sổ: ô trống, không kết luận xuất khống. Một phiên bản mới của KQ2.2.
```

## Bước 7 - Giá thành

**Lệnh:**
```
Phân tích giá thành / cơ cấu GV tại <FOLDER> kỳ <KY>.
Ghi KQ2.2!KQ2_Gia_thanh_va_Dinh_muc. Nhận xét module 3.4. Một phiên bản mới của KQ2.2.
```

## Bước 8 - Công nợ / tiền mặt

**Lệnh:**
```
Quét 111, cụm cùng ngày, 131/331, sao kê 112 tại <FOLDER> kỳ <KY>.
Ngưỡng >= 20tr đến 30/06/2025; >= 5tr từ 01/07/2025. P23 = 10% (TN17).
Ghi KQ2.2!KQ2_Cong_no_Tien_mat. PDF chưa OCR: PDF - chưa kết luận. Nhận xét module 3.4.
Không xuất Phan_tich_Tien_mat_va_Ngan_hang. Một phiên bản mới của KQ2.2.
```

## Bước 9-10 - KQ3 (80 quy tắc + 28 tiêu chí)

Sheet `Danh_gia_28_tieu_chi_rui_ro` nằm trong KQ3 là chuẩn vòng 2 (không còn "chỉ khi cán bộ yêu cầu"). Điền đề cương cột AB-AC-AH; AD để trống cho cán bộ. Không tách sheet KQ3_De_cuong kiểu cũ.

## Bước 11 - Phụ lục GĐ3

Xuất `KQ3_Danh_gia_Dau_hieu_Bat_thuong` và `KQ3_NNT_Giai_trinh_Chi_tiet` (10 sheet có số liệu thật). Áp Q14-Q20.

## Gói siêu ngắn

| Module | Câu |
|---|---|
| GĐ1 | `Kiểm kê GĐ1 <FOLDER> MST <MST> kỳ <KY> xuất KQ1.1 KQ1.2 KQ1.3` |
| KQ2.1 | `Phân tích KQ2.1 <FOLDER> MST <MST> kỳ <KY>` |
| KQ2.2 | `Phân tích KQ2.2 <FOLDER> MST <MST> kỳ <KY>` |
| GĐ2 đủ | `Phân tích GĐ2 <FOLDER> MST <MST> kỳ <KY> xuất KQ2.1 KQ2.2` |
| GĐ3 | `Hoàn chỉnh GĐ3 <FOLDER> MST <MST> kỳ <KY> xuất KQ3 và chi tiết giải trình` |

## Nguyên tắc

1. Không chờ đủ mới chạy. Ô thiếu để trống.
2. Một lệnh = một phiên bản mới của tệp đích (hoặc một lệnh GĐ3 = KQ3 + 2 file phụ).
3. Hỏi PL / workbook trọn gói / mẫu 01-07 / Master 8 sheet tách ngoài: từ chối, đề nghị bộ 1.4 vòng 2.
