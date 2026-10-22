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
> **Hợp đồng đầu ra v1.4:** module nội bộ. Kết quả ghi vào sheet đích của sáu tệp KQ (xem skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md). Không giao tệp chuyên đề Tong_hop_*/Doi_chieu_* / Master 8 sheet như sản phẩm.

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

## Cách chạy / Quy trình xuất (v1.4 vòng 2, 19/10/2026)

Font **Times New Roman**, chữ đứng (không in nghiêng), tiêu đề cột in đậm. Ghi vào hai sheet của bộ KQ1.4: `IN07_NCC_CanhBao` (KQ2.1, tổng hợp theo NCC) và `KQ2_NCC_mua_vao` (KQ2.1, chi tiết theo hóa đơn). Không còn xuất `Tong_hop_NCC_CVCBRR`/`BM04`/file riêng làm sản phẩm giao.

### 1. `IN07_NCC_CanhBao` - tổng hợp theo NCC (một dòng/NCC)

- **Quy tắc xếp mức K2 (15/10/2026):**
  - F0 = công ty bị cảnh báo rủi ro hoặc trong danh sách mua bán hóa đơn của Công an (CA).
  - F1 = công ty mua hàng có sử dụng hóa đơn của F0 (tầng hai).
  - **Mức rủi ro:** NCC thuộc F0 xếp **Cao**; NCC thuộc F1 (tầng hai) cũng xếp mức **Cao**. NNT đang kiểm tra thuộc F1 nghĩa là đã mua hàng từ F0 → ghi nhận phát hiện **Cao** của quy tắc `HD01`.
  - Cột `Nhom_F` (F0/F1) và `Nguon_F0` (CANH_BAO/CA).
- Liệt kê **100% NCC** (union Bảng kê ∪ HĐĐT), có dòng **Tổng cộng toàn bộ** ở CUỐI bảng (cách 1 dòng trống, không chèn giữa dữ liệu - Q19/lỗi dòng 147 TT_BIEN_HOA); đánh dấu nhóm trạng thái ĐKKD bất thường. **Cấm** chỉ giữ nhóm bất thường.
- Cột `Nguồn dữ liệu` ∈ {Bảng kê, HĐĐT, Chung}; thêm `Số dòng BK` / `Số dòng HĐĐT`.

### 2. `KQ2_NCC_mua_vao` - MỘT bảng chi tiết hóa đơn duy nhất (bổ sung/thay thế 19/10/2026 vòng 2)

Thay cho Bảng 2.1/2.2/2.3 cũ (gây lẫn 2 bảng không phân cách - lỗi dòng 147/148 TT_BIEN_HOA): một bảng phẳng, **mỗi dòng = một hóa đơn mua vào**, lấy **100% hóa đơn hợp lệ** (không chỉ NCC rủi ro) - đây là bảng khai triển chi tiết của `IN07_NCC_CanhBao`.

Cột bắt buộc: MST người bán (chuẩn hóa), Tên người bán, Số HĐ, Ngày HĐ (`dd/mm/yyyy`, không giờ), Mặt hàng, Giá trị chưa thuế, Thuế GTGT, Nguồn dữ liệu (Bảng kê/HĐĐT/Chung), Có công văn cảnh báo? (Có/Không + số CV nếu có), Nhóm rủi ro (F0/F1/Không, lấy từ IN07 của chính người bán đó), Trạng thái ĐKKD (Đang hoạt động/Bỏ địa chỉ kinh doanh/Ngừng kinh doanh/Khác).

Sắp **`Ngày HĐ` tăng dần** theo từng người bán, hoặc theo người bán rồi theo ngày (nêu rõ cách sắp ở đầu sheet). Nếu cần dòng Tổng cộng: cuối bảng, cách 1 dòng trống.

**HD04 - tỷ trọng Top-5 NCC (P14, mặc định 50%) bắt buộc giữ.** Không dùng công thức `LARGE` trên cột giá trị của bảng chi tiết hóa đơn (sai vì Top-5 là theo người bán, không theo dòng HĐ). Cách tính đúng: gộp tổng mua theo MST người bán (pivot/code), lấy 5 NCC lớn nhất / tổng mua kỳ kiểm tra, ghi số vào ô tỷ trọng Top-5 và cờ `Vượt ngưỡng` | `Không vượt` | `Chưa đủ nguồn` (ô H5 để `KQ2_Tong_hop_co` đọc). Được thêm một bảng Top-5 ngắn (MST, tên, giá trị, tỷ trọng) cách bảng hóa đơn bằng 1 dòng trống.

### 3. Phần đánh giá (ngoại lệ 3.4 skill trục, bắt buộc có ở `KQ2_NCC_mua_vao`)

Khối narrative ở đầu hoặc cuối sheet `KQ2_NCC_mua_vao` (cách bảng chi tiết bằng dòng trống rõ ràng):
- **Phần 1:** Đối chiếu NNT đang xét với F1 và Sổ Công an; ghi kết luận an toàn/có cảnh báo.
- **Phần 3:** (1) Đánh giá rủi ro NNT, (2) Đánh giá rủi ro NCC theo nhóm, (3) Đề xuất trọng tâm kiểm tra chứng từ (ưu tiên NCC lệch số liệu giữa 2 nguồn, NCC chỉ có một phía, cụm HĐ cùng ngày vượt ngưỡng tiền mặt). Đây là nhận xét module (3.4), KHÔNG phải kết luận hồ sơ - mức rủi ro/điểm/đề cương cuối cùng vẫn chỉ ở KQ3.

## Định dạng số tiền (bắt buộc)

- **Hiển thị** (chat, Markdown, báo cáo, ô text): chuẩn Việt Nam - dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (openpyxl `number_format`): dùng mã invariant `#,##0` (nguyên) hoặc `#,##0.00` (có phần lẻ). Excel locale VN sẽ hiện thành dấu `.` / `,`. Không mô tả nhầm là "dấu phẩy ngăn cách hàng nghìn".
- **Tỷ lệ %**: hiển thị `12,34%` (dấu `,` phần lẻ).
- Cấm dạng Mỹ `1,234,567.89` trong câu trả lời / báo cáo tiếng Việt.

## Bẫy

- **Bẫy rút gọn:** `KQ2_NCC_mua_vao` chỉ xuất NCC bất thường thay vì 100% hóa đơn, hoặc cắt Phần 1/3 -> mất dấu vết lệch giữa Bảng kê và HĐĐT.
- Chỉ chạy trên Bảng kê mà bỏ HĐĐT (hoặc ngược lại) khi cả hai nguồn đều có.
- Tách `KQ2_NCC_mua_vao` thành nhiều bảng rời theo nguồn hoặc theo nhóm rủi ro thay vì 1 bảng phẳng + cột `Nguồn dữ liệu`/`Nhóm rủi ro`.
- Chèn dòng Tổng cộng hoặc header phụ giữa vùng dữ liệu không có dòng trống ngăn cách (lỗi dòng 147/148 TT_BIEN_HOA 19/10) - dòng Tổng cộng (nếu có) luôn ở cuối, cách 1 dòng trống.
- Nhầm F1 (bên mua) với F0 (bên bán).
- Khớp sót MST có/không gạch nhánh (chuẩn hóa 10 và 13 số; bổ sung 0 đầu theo CD1 BT6 trước khi map F0/danh bạ).
- Sort chi tiết theo MST thay vì `Ngày HĐ` khi mục tiêu là truy ngược thời gian.
- Ô text bắt đầu bằng `=` gây lỗi Removed Records Formula.
- Bỏ qua cụm nhiều HĐ cùng ngày cùng NCC vượt ngưỡng thanh toán không dùng tiền mặt (20tr đến 06/2025, 5tr từ 07/2025) khi đề xuất trọng tâm kiểm tra UNC.

## Kiểm chứng

- Tổng doanh số mọi NCC trên `IN07_NCC_CanhBao` khớp union nguồn đã dùng (ghi rõ nếu một phía thiếu phủ).
- `KQ2_NCC_mua_vao` có đủ 100% hóa đơn mua vào (không chỉ nhóm rủi ro) + khối Phần 1 + Phần 3.
- Mọi dòng `IN07_NCC_CanhBao`/`KQ2_NCC_mua_vao` có `Nguồn dữ liệu` ∈ {Bảng kê, HĐĐT, Chung}.
- MST đã qua chuẩn hóa 10/13 số trước khi gắn nhãn F0/ĐKKD.
- Không có dòng Tổng cộng hoặc header phụ lẫn giữa vùng dữ liệu (kiểm bằng mắt dòng ngay trước/sau mỗi "Tổng cộng").
- Font Times New Roman, chữ đứng, tiêu đề đậm, ngày `dd/mm/yyyy` không giờ.

## Liên kết quy trình hồ sơ

Playbook lệnh: `playbook-lenh.md`.

Skill này là phân hệ chuyên sâu của `phan-tich-ho-so-thue` tại **Bước 3 GĐ2** (xem `giai-doan-2-phan-tich-du-lieu-co-quan-thue`). Đầu vào: bảng kê + HĐĐT đã gộp (`tong-hop-dau-vao-hddt-bang-ke-nld`), MST đã chuẩn (`lam-sach-cd1-mst-hoa-don`). Cụm HĐ cùng ngày vượt ngưỡng tiền mặt khi đề xuất UNC: `phan-tich-cong-no-tien-mat`. Kết quả ghi vào `IN07_NCC_CanhBao` + `KQ2_NCC_mua_vao` của `KQ2.1_Phan_tich_CQT`.

## Lịch sử (tự học)
- [2026-10-19 vòng 2] Bỏ cấu trúc Bảng 2.1/2.2/2.3 lồng nhau (gây lẫn dữ liệu dòng 147/148 TT_BIEN_HOA); `KQ2_NCC_mua_vao` đổi thành 1 bảng chi tiết hóa đơn duy nhất (100%, không chỉ rủi ro) là bảng khai triển của `IN07_NCC_CanhBao`; Phần 1/3 chuyển thành khối narrative trong `KQ2_NCC_mua_vao` (ngoại lệ 3.4). Định dạng tiêu đề: đậm, chữ đứng.
- [2026-10-19 vòng 2b] Cán bộ chốt giữ HD04 (Top-5 NCC / P14): tính bằng gộp theo MST người bán, không `LARGE` trên dòng hóa đơn.
- [2026-10-15] Gắn Liên kết quy trình vào skill trục; nhắc đầu vào đầu-vào / CD1 / công nợ.
- [2026-10-09] Chốt định dạng số tiền chuẩn VN đồng loạt: hiển thị dấu `.` hàng nghìn và `,` phần lẻ; Excel `#,##0` / `#,##0.00`.
- [2026-10-05] User chỉ ra file Tong_hop_NCC của T&T thiếu thông tin so với Vạn Phát Thành; skill cũ mô tả Bảng 2.2 như chỉ liệt kê NCC bất thường, dễ dẫn tới bản rút gọn. Cập nhật bắt buộc 100% NCC, Phần 3 và đủ cột Chi_tiet.
- [2026-10-07] Bổ sung chuẩn hóa MST CD1 BT6 và nhắc quét cụm HĐ cùng ngày vượt ngưỡng tiền mặt khi đề xuất trọng tâm kiểm tra chứng từ.
- [2026-10-07] Bổ sung nguồn HĐĐT ngang hàng Bảng kê; cột `Nguồn dữ liệu` = Bảng kê / HĐĐT / Chung (1 bảng duy nhất); thêm Bảng 2.3 đối chiếu theo kỳ kê khai cho NCC rủi ro (chi tiết HĐ, mặt hàng, giá trị, thuế).
