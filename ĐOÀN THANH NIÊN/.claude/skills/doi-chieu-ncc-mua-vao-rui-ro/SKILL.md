---
name: Đối chiếu NCC mua vào với CVCBRR và danh bạ
description: Đối chiếu NCC từ Bảng kê + HĐĐT với CVCBRR F0/F1, Công an, danh bạ; ghi nguồn Bảng kê/HĐĐT/Chung; Bảng 2.3 theo kỳ.
group: Tài chính
origin: javis-learned
status: active
created: 2026-09-30
updated: '2026-10-09'
learned_updated: '2026-10-07'
---
## Khi nào dùng

- Đối chiếu nhà cung cấp / bên bán trên **Bảng kê mua vào** và **HĐĐT mua vào (CQ thuế)** với Sổ CVCBRR (F0, F1), Sổ Công an (`Bao_cao_cong_van.xlsx`) và Danh bạ NNT.
- Xuất Master sheet `Tong_hop_NCC_CVCBRR` hoặc file riêng `Tong_hop_NCC_mua_vao_CV_trang_thai_*.xlsx` (có Bảng 2.1 / 2.2 / 2.3 + Phần 3).
- So khớp với bản chuẩn Vạn Phát Thành khi nghi file bị rút gọn.

## Chuẩn bị

1. **Bảng kê mua vào**: `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx` (sheet `mua vao` đã gộp, đã lọc dòng Tổng cộng).
2. **HĐĐT mua vào**: file HĐĐT portal / hệ thống CQ thuế (sheet mua vào đã lọc HĐ hủy/thay thế khi cần).
3. **Sổ CVCBRR**: `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` (F0 bên bán, F1 bên mua).
4. **Sổ Công an**: `Bao_cao_cong_van.xlsx`.
5. **Danh bạ NNT**: trạng thái ĐKKD.

## Quy tắc nguồn dữ liệu (bắt buộc)

- Chạy rủi ro NCC trên **cả hai nguồn**: Bảng kê và HĐĐT.
- **Một bảng duy nhất** cho danh sách NCC (không tách 2 bảng riêng theo nguồn).
- Cột **`Nguồn dữ liệu`** ghi đúng một trong ba giá trị:
  - `Bảng kê` - chỉ có trên bảng kê công ty
  - `HĐĐT` - chỉ có trên HĐĐT CQ thuế
  - `Chung` - có ở cả hai nguồn
- Có thêm cột đếm `Số dòng BK` / `Số dòng HĐĐT` (và/hoặc checkbox Có trong Bảng kê? / Có trong HĐĐT?) để kiểm chứng nhanh.

## Cách chạy / Quy trình xuất

Font **Times New Roman**. Mẫu: `thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/` (`06_Phan_tich_ho_so_cong_ty_MAU.xlsx` sheet `Tong_hop_NCC_CVCBRR` hoặc `05_Tong_hop_NCC_mua_vao_CV_trang_thai_MAU.xlsx`).

### 1. Sheet tổng hợp (`Tong_hop_NCC_CVCBRR` / `Tong_hop_doi_chieu` / `BM04_NCC_mua_vao`)

- **Quy tắc xếp mức K2 + F0 cho NNT (chốt 11/11/2026):**
  - **F0 = NCC (nhà cung cấp):** công ty bị cảnh báo rủi ro hoặc trong danh sách mua bán hóa đơn của Công an (CA). Đối chiếu tất cả NCC với sheet F0.
  - **F1 = NCC (nhà cung cấp):** công ty mua hàng có sử dụng hóa đơn của F0 (tầng hai). Đối chiếu cả NCC để phát hiện F1.
  - **F0 = NNT (người nộp thuế):** kiểm F0 TRƯỚC để xác định NNT có phải công ty bị cảnh báo không → ghi KQ1.3 + KQ3 mức rủi zo Cao từ đầu.
  - **F1 = NNT (người nộp thuế):** kiểm F1 để phát hiện NNT có mua HĐ từ F0 không → ghi Phần 1 IN07/KQ2_NCC_mua_vao, phát hiện **HD01 mức Cao**.
  - **Mức rủi zo NCC:** NCC thuộc F0 xếp **Cao**; NCC thuộc F1 (tầng hai) cũng xếp mức **Cao**. 
  - IN07 ghi nhận cột `Nhom_F` (F0/F1) và `Nguon_F0` (CANH_BAO/CA); BM04 ghi cột S (`Nhóm F`) và T (`Mức rủi zo NCC`).
- **Phần 1 (bắt buộc):** 
  - Kiểm **F0 của NNT trước**: NNT có nằm trong sheet F0 không? Nếu có, ghi rõ "NNT [tên] (MST [mst]) bị cảnh báo rủi zo thuế - Công văn [số]" (rủi zo NNT tự thân) → ghi KQ1.3 + KQ3 mức Cao từ đầu.
  - Kiểm **F1 của NNT tiếp**: NNT có nằm trong sheet F1 không? Nếu có, ghi rõ "NNT [tên] (MST [mst]) đã mua HĐ từ F0 [tên F0]" → phát hiện **HD01 mức Cao** (mua HĐ rủi zo).
  - Đối chiếu NNT với Sổ Công an: kiểm NNT có trong danh sách CA không?
  - Kết luận: an toàn / có cảnh báo / mục tiêu kiểm tra lớn.
- **Phần 2 (1 khối duy nhất, gắn nguồn):**
  - **Bảng 2.1 (F0/F1):** NCC dính công văn cảnh báo; có cột `Nguồn dữ liệu`, `Số dòng BK`, `Số dòng HĐĐT`, kỳ mua, CV/F0, ĐKKD...
  - **Bảng 2.2:** Liệt kê **100% NCC** (union Bảng kê ∪ HĐĐT), có dòng **Tổng cộng toàn bộ**; đánh dấu nhóm trạng thái ĐKKD bất thường. **Cấm** chỉ giữ nhóm bất thường.
  - **Bảng 2.3 (theo kỳ kê khai):** Từ NCC rủi ro đã chọn ở 2.1/2.2 (F0 / F1 / ĐKKD bất thường / chọn lọc), liệt kê chi tiết từng dòng HĐ theo **kỳ kê khai** tương ứng: MST/tên NCC, nguồn, kỳ, số HĐ, ngày HĐ, mặt hàng, giá trị chưa thuế, thuế GTGT, loại rủi ro, số CV, có trên Bảng kê?, có trên HĐĐT?, ghi chú. Một NCC có thể xuất hiện nhiều kỳ.
- **Phần 3 (bắt buộc):** (1) Đánh giá rủi ro NNT, (2) Đánh giá rủi ro NCC theo nhóm, (3) Đề xuất trọng tâm kiểm tra chứng từ (ưu tiên NCC `Chung` lệch số liệu giữa 2 nguồn, NCC chỉ có một phía, cụm HĐ cùng ngày vượt ngưỡng tiền mặt).

### 2. Sheet / lớp chi tiết hóa đơn

- File riêng có sheet `Bang_2_3_Ky_ke_khai_CT` làm lớp chi tiết theo kỳ cho NCC rủi ro.
- Nếu vẫn xuất `Chi_tiet_hoa_don` toàn kỳ: sắp **`Ngay_HD` tăng dần**; đủ nhóm chứng từ kế toán, nhóm F0 và cột `Trạng thái hoạt động ĐKKD`; nên có cột nguồn / cờ có trên Bảng kê và HĐĐT.

## Định dạng số tiền (bắt buộc)

- **Hiển thị** (chat, Markdown, báo cáo, ô text): chuẩn Việt Nam - dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (openpyxl `number_format`): dùng mã invariant `#,##0` (nguyên) hoặc `#,##0.00` (có phần lẻ). Excel locale VN sẽ hiện thành dấu `.` / `,`. Không mô tả nhầm là "dấu phẩy ngăn cách hàng nghìn".
- **Tỷ lệ %**: hiển thị `12,34%` (dấu `,` phần lẻ).
- Cấm dạng Mỹ `1,234,567.89` trong câu trả lời / báo cáo tiếng Việt.

## Bẫy

- **Bẫy rút gọn:** chỉ xuất NCC bất thường / cắt Phần 3 / thiếu Bảng 2.3 / thiếu cột nguồn -> mất NCC chủ lực hoặc mất dấu vết lệch giữa Bảng kê và HĐĐT.
- Chỉ chạy trên Bảng kê mà bỏ HĐĐT (hoặc ngược lại) khi cả hai nguồn đều có.
- Tách thành 2 bảng riêng theo nguồn thay vì 1 bảng + cột `Nguồn dữ liệu`.
- Nhầm F1 (bên mua) với F0 (bên bán).
- Khớp sót MST có/không gạch nhánh (chuẩn hóa 10 và 13 số; bổ sung 0 đầu theo CD1 BT6 trước khi map F0/danh bạ).
- Sort chi tiết theo MST thay vì `Ngay_HD` / kỳ kê khai khi mục tiêu là truy ngược kỳ.
- Ô text bắt đầu bằng `=` gây lỗi Removed Records Formula.
- Bỏ qua cụm nhiều HĐ cùng ngày cùng NCC vượt ngưỡng thanh toán không dùng tiền mặt (20tr đến 06/2025, 5tr từ 07/2025) khi đề xuất trọng tâm kiểm tra UNC.

## Kiểm chứng

- Tổng doanh số mọi NCC trên Bảng 2.2 khớp union nguồn đã dùng (ghi rõ nếu một phía thiếu phủ).
- Có đủ Phần 1 + 2.1 + 2.2 (100% NCC) + **2.3 (NCC rủi ro theo kỳ)** + Phần 3.
- Mọi dòng 2.1/2.2/2.3 có `Nguồn dữ liệu` ∈ {Bảng kê, HĐĐT, Chung}.
- MST đã qua chuẩn hóa 10/13 số trước khi gắn nhãn F0/ĐKKD.
- Font Times New Roman.

## Liên kết quy trình hồ sơ

Playbook lệnh: `playbook-lenh.md`.

Skill này là phân hệ chuyên sâu của `phan-tich-ho-so-thue` tại **Bước 4.2 / playbook Bước 3**. Đầu vào: bảng kê + HĐĐT đã gộp (`tong-hop-dau-vao-hddt-bang-ke-nld`), MST đã chuẩn (`lam-sach-cd1-mst-hoa-don`). Cụm HĐ cùng ngày vượt ngưỡng tiền mặt khi đề xuất UNC: `phan-tich-cong-no-tien-mat`. Kết quả nhúng Master sheet `Tong_hop_NCC_CVCBRR`.

## Lịch sử (tự học)
- [2026-10-15] Gắn Liên kết quy trình vào skill trục; nhắc đầu vào đầu-vào / CD1 / công nợ.
- [2026-10-09] Chốt định dạng số tiền chuẩn VN đồng loạt: hiển thị dấu `.` hàng nghìn và `,` phần lẻ; Excel `#,##0` / `#,##0.00`.
- [2026-10-05] User chỉ ra file Tong_hop_NCC của T&T thiếu thông tin so với Vạn Phát Thành; skill cũ mô tả Bảng 2.2 như chỉ liệt kê NCC bất thường, dễ dẫn tới bản rút gọn. Cập nhật bắt buộc 100% NCC, Phần 3 và đủ cột Chi_tiet.
- [2026-10-07] Bổ sung chuẩn hóa MST CD1 BT6 và nhắc quét cụm HĐ cùng ngày vượt ngưỡng tiền mặt khi đề xuất trọng tâm kiểm tra chứng từ.
- [2026-10-07] Bổ sung nguồn HĐĐT ngang hàng Bảng kê; cột `Nguồn dữ liệu` = Bảng kê / HĐĐT / Chung (1 bảng duy nhất); thêm Bảng 2.3 đối chiếu theo kỳ kê khai cho NCC rủi ro (chi tiết HĐ, mặt hàng, giá trị, thuế).
