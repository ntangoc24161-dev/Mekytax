---
name: Phân tích hồ sơ kiểm tra thuế
description: Quy trình 3 giai đoạn phân tích hồ sơ thuế theo bộ biểu mẫu v1.1 (80 quy tắc), kết nối dữ liệu CQT và sổ sách DN, chốt Master 28 tiêu chí (KQ1-3).
group: Tài chính
updated: '2026-10-15'
learned_updated: '2026-10-15'
---

# Phân tích hồ sơ kiểm tra thuế (Chuẩn v1.2 – 80 quy tắc rủi ro & Quyết định K1–K13)

Quy trình tự động hóa phân tích toàn diện hồ sơ kiểm tra thuế theo **kiến trúc 3 giai đoạn độc lập** căn cứ theo Bộ biểu mẫu chuẩn v1.2 (`Bo_bieu_mau_chuan_phan_tich_ho_so_thue_v2.xlsx`, 80 quy tắc rủi ro, 15 mẫu nhập IN00–IN14, 6 mẫu BM) và Quyết định cán bộ ngày 15/10/2026 (Prompt P5 bản 2):
1. **Giai đoạn 1 – Kiểm kê & Đánh giá hồ sơ:** Quét manifest IN00, xác định mức hồ sơ A–E, rà soát 80 quy tắc, lập danh mục đề nghị NNT bổ sung (`KQ1`).
2. **Giai đoạn 2 – Phân tích dữ liệu Cơ quan thuế (CQT):** Chuẩn hóa một lần IN01–IN07/12/14, chạy 57 quy tắc CQT khi chưa có sổ DN, rà soát 100% NCC rủi ro (F0/F1 đều mức Cao - K2), phân loại kết luận Chắc / Sơ bộ (`KQ2`).
3. **Giai đoạn 3 – Hoàn chỉnh hồ sơ & Sổ sách doanh nghiệp:** Chạy 23 quy tắc sổ sách, nâng cấp quy tắc sơ bộ, dựng chuỗi bằng chứng 5 mắt xích, tổng hợp Master 28 tiêu chí (7 cột chuẩn, K3) và lập Đề cương kiểm tra phân nhánh xử lý (`KQ3`).

---

## Ba kỹ năng giai đoạn chủ lực (Pipeline Skills)

- **Giai đoạn 1:** `giai-doan-1-kiem-ke-danh-gia` (Kiểm kê IN00–IN14, mức A–E, xuất KQ1).
- **Giai đoạn 2:** `giai-doan-2-phan-tich-du-lieu-co-quan-thue` (Phân tích 57 quy tắc CQT, xuất KQ2).
- **Giai đoạn 3:** `giai-doan-3-hoan-chinh-ho-so` (Hoàn chỉnh 23 quy tắc sổ sách, chuỗi bằng chứng, Master 28 TC, xuất KQ3).

---

## Phân hệ nghiệp vụ chuyên sâu (Sub-modules)

Các skill chuyên đề phục vụ chi tiết từng phần việc trong giai đoạn 2 và 3:
- **Tổ chức thư mục 4 bước + checklist A4:** `to-chuc-ho-so-kiem-tra-thue` (A0).
- **Gộp HĐĐT / Bảng kê / NLĐ (mẫu 02):** `tong-hop-dau-vao-hddt-bang-ke-nld` (A0b + A1).
- **Làm sạch MST, HĐ trùng, BKMV-NKC 133 (CD1 BT5-BT9):** `lam-sach-cd1-mst-hoa-don` (A1b).
- **Đối chiếu GTGT / TNDN / HĐĐT chuẩn CBRR:** `doi-chieu-gtgt-tndn-cbrr` (A2).
- **100% NCC mua vào vs CVCBRR / Công an / ĐKKD:** `doi-chieu-ncc-mua-vao-rui-ro` (A3, F0/F1 xếp Cao).
- **QTT TNCN + đa nguồn (TC22/TC23):** `tong-hop-tncn-da-nguon` (A4).
- **Đối chiếu Doanh thu 511 vs Giá vốn 632:** `doi-chieu-511-632` (A5 - Quyết định K4).
- **Phân tích XNT kho 152/155/156 (TC3):** `phan-tich-xnt-kho` (A6 - Quyết định K6).
- **Giá thành / Định mức / Cơ cấu GV (TC28):** `phan-tich-gia-thanh-dinh-muc` (A7 - TN22).
- **Công nợ 131/331, tiền mặt 111, sao kê 112 (TC20):** `phan-tich-cong-no-tien-mat` (A8 - ngưỡng >= 20tr theo K5).
- **Checklist mở hồ sơ TTKT, GDLK:** `mo-ho-so-ttkt` (B1).
- **Bên mua trên sổ CVCBRR:** `tong-hop-ben-mua-tu-so-cvcbrr` (B2).
- **Khung quản lý rủi ro tuân thủ:** `quan-ly-rui-ro-tuan-thu-thue` (B3).

---

## Khối Q: 11 nguyên tắc vận hành & Quyết định cán bộ K1–K13

1. **Giai đoạn do dữ liệu quyết định:** Chạy mọi quy tắc mà nguồn đã có; thiếu nguồn ghi `KHÔNG_ĐỦ_DỮ_LIỆU` hoặc `TREO` kèm lý do. Không điền số 0 giả.
2. **Mỗi quy tắc đúng một trạng thái:** `ĐÃ_CHẠY`, `KHÔNG_ĐỦ_DỮ_LIỆU`, `TREO`, `CHƯA_TRIỂN_KHAI`, `KHÔNG_ÁP_DỤNG` (theo loại hình DN). Không để quy tắc không trạng thái.
3. **Mỗi nguồn gắn nguồn gốc:** CQT hoặc DN (K1: suy từ thư mục cấp 1). Khi có cả hai bản, bản CQT là chuẩn so sánh; chênh lệch là phát hiện nhóm DQ.
4. **Mọi phát hiện có đầy đủ thông tin:** Mã quy tắc, giá trị đo, ngưỡng áp dụng (kèm mã tham số P), mức rủi ro, hệ số trọng yếu, điểm, nhãn "Chắc"/"Sơ bộ", ô nguồn dạng chữ (file, sheet, dòng).
5. **Tính toán tất định:** Phép tính do mã hoặc công thức tất định thực hiện. AI chỉ đọc file, ánh xạ vào IN và viết diễn giải trên số đã tính.
6. **Không sửa, không xóa file nguồn gốc:** Mọi file nhận vào ghi tên, ngày nhận, mã băm SHA-256. Cán bộ quyết định lưu trữ hay xóa.
7. **Không kết luận vượt nguồn:** Sao kê PDF chưa trích dòng ghi `PDF – chưa kết luận`, không kết luận thiếu UNC (K11). Năm BCTC chỉ gợi ý kèm độ tin cậy.
8. **Quyết định cán bộ K1–K13 (15/10/2026):**
   - K1: Thư mục 1 = CQT; 2 = DN.
   - K2: F0 và F1 đều xếp mức **Cao**; CA là F0; NNT thuộc F1 là phát hiện Cao của HD01.
   - K3: Master 8 sheet nếu không có mẫu thì **chỉ dựng sheet 28 tiêu chí (sheet 08)**.
   - K4: Đối chiếu 511–632 theo skill `doi-chieu-511-632`.
   - K5: Ngưỡng thanh toán không dùng tiền mặt chốt dùng $\ge$ ở cả hai giai đoạn: $\ge 20$ triệu (đến 30/06/2025) và $\ge 5$ triệu (từ 01/07/2025).
   - K6: Hàng tồn kho B01-DN dùng [141], B01-DNN dùng ct140.
   - K7: Chế độ kế toán TT99 hiệu lực 01/01/2026; TT133 không bị thay thế; DT_CHUA_THUC_HIEN mã 319, VON_CSH mã 411.
   - K8: Mốc trọng yếu ưu tiên: Lệnh cán bộ → Skill AI → Mặc định sheet 05.
   - K9: Thuế suất GTGT giảm giữ bảng đề xuất, bổ sung số NQ 174, NQ 204.
   - K10: Lịch sử tuân thủ TMS đọc từ thư mục 1.3 vào IN14_TMS_Nghia_vu; LS01, LS02 đọc IN14.
   - K11: Thiếu chứng từ DN (sao kê, UNC, hợp đồng) không phải là phát hiện; mắt xích thiếu ghi "Không đủ dữ liệu".
   - K12: Phạm vi kỳ kiểm tra do cán bộ chỉ định từng hồ sơ.
   - K13: Số liệu tham chiếu ngành: HD07 ghi TREO; TN03, TN04, TN19, TN22 chạy sơ bộ – chưa so ngành.
9. **Chạy lại tăng dần:** Không ghi đè kết quả cũ, tăng phiên bản (v1, v2…), lập sheet theo dõi thay đổi (`KQ3_Thay_doi_so_voi_GD2`).
10. **Hợp đồng đầu ra chỉ gồm sheet kết quả tĩnh:** File Excel kết quả (KQ1, KQ2, KQ3) không mang công thức liên sheet; không xuất các sheet dữ liệu thô IN hoặc sheet tham chiếu 00–06.
11. **Tự kiểm tra trước khi bàn giao:** Mọi quy tắc có trạng thái; không có số 0 giả; tổng khớp chi tiết; 10 phát hiện điểm cao nhất đã kiểm lại từ dữ liệu gốc.

---

## Hợp đồng đầu ra theo giai đoạn

| File kết quả | Sheet bên trong | Nội dung |
|---|---|---|
| **`KQ1_<MST>_<KY>_v<n>.xlsx`** | `KQ1_Kiem_ke`<br>`KQ1_Kha_nang_phan_tich`<br>`KQ1_De_nghi_bo_sung` | Kiểm kê toàn bộ nguồn IN00–IN14, trạng thái mức hồ sơ A–E, rà soát 80 quy tắc, danh mục hồ sơ cần đề nghị DN bổ sung xếp theo số quy tắc được mở khóa. |
| **`KQ2_<MST>_<KY>_v<n>.xlsx`** | `KQ2_Doi_chieu_GTGT`<br>`KQ2_Doi_chieu_TNDN_BCTC`<br>`KQ2_Nam_BCTC`<br>`KQ2_NCC_mua_vao`<br>`KQ2_TNCN_LD`<br>`KQ2_Phat_hien`<br>`KQ2_Chua_ket_luan` | Kết quả đối chiếu GTGT, TNDN vs BCTC; ánh xạ năm BCTC; 100% NCC kèm cờ rủi ro; kết quả TNCN & LD; danh sách phát hiện GĐ2 xếp theo điểm (Chắc / Sơ bộ); quy tắc bị chặn và điều cấm kết luận vội. |
| **`KQ3_<MST>_<KY>_v<n>.xlsx`** | `KQ3_Phat_hien_tong`<br>`KQ3_511_632`<br>`KQ3_Kho_Gia_thanh`<br>`KQ3_Cong_no_Tien_mat`<br>`KQ3_Chuoi_bang_chung`<br>`KQ3_Master_28TC`<br>`KQ3_De_cuong`<br>`KQ3_Thay_doi_so_voi_GD2`<br>`KQ3_Con_thieu` | Tổng hợp cả 80 quy tắc rủi ro; đối chiếu 511–632; phân tích kho và giá thành; công nợ và tiền mặt; chuỗi bằng chứng 5 mắt xích; Master 28 tiêu chí 7 cột; Đề cương kiểm tra phân nhánh xử lý; biến động so với GĐ2; nội dung còn thiếu. |

---

## Định dạng số tiền (Bắt buộc)

- **Hiển thị** (chat, Markdown, báo cáo): chuẩn Việt Nam - dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (openpyxl `number_format`): dùng mã invariant `#,##0` (nguyên) hoặc `#,##0.00` (có phần lẻ).
- **Tỷ lệ %**: hiển thị `12,34%` (dấu `,` phần lẻ).
- Cấm dạng Mỹ `1,234,567.89` trong câu trả lời và báo cáo tiếng Việt.

---

## Playbook và Lệnh thực thi

- **Playbook 3 giai đoạn chuẩn:** `skills/phan-tich-ho-so-thue/playbook-3-giai-doan.md` (chứa P0, Q, P1, P2, P3, P4 và bảng xử lý thiếu dữ liệu).
- **Playbook theo bước module:** `skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md`.
- **Catalog toàn bộ lệnh:** `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md`.

---

## Lịch sử học tập

- [2026-10-15] Nâng cấp toàn diện lên bộ biểu mẫu v1.1 (80 quy tắc rủi ro, 14 mẫu nhập IN00–IN13, 6 mẫu phân tích BM01–BM06).
- [2026-10-15] Tách kiến trúc thành 3 skill giai đoạn: `giai-doan-1-kiem-ke-danh-gia`, `giai-doan-2-phan-tich-du-lieu-co-quan-thue`, `giai-doan-3-hoan-chinh-ho-so`.
- [2026-10-15] Chuẩn hóa hợp đồng đầu ra KQ1, KQ2, KQ3 mang giá trị tĩnh và Khối Q 11 quy tắc.
- [2026-10-15] Gắn bảng ánh xạ chỉ tiêu BCTC TT200/TT133 qua sheet 06 (mã `HTK_GOP`) và ánh xạ năm BCTC sheet 07.
