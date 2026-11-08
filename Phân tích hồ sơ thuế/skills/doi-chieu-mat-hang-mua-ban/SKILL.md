---
name: Đối chiếu mặt hàng mua vào bán ra
description: "Đối chiếu mặt hàng mua-bán HĐĐT theo tiền, số lượng, đơn giá; lọc HĐ hủy, cắt kỳ chồng, bắt chuỗi trung gian."
group: Tài chính
status: active
created: 2026-11-03
updated: 2026-11-03
---

# Đối chiếu mặt hàng mua vào bán ra

So khớp **HĐĐT mua vào** với **HĐĐT bán ra** theo tên hàng, tiền, số lượng và đơn giá. Mục tiêu: hàng mua có bán tương xứng không, có chuỗi trung gian (mua rồi bán cùng ngày, cộng biên mỏng) không. Khác `doi-chieu-ncc-mua-vao-rui-ro` (NCC vs F0/F1), `doi-chieu-511-632` (sổ 511 vs 632), `phan-tich-xnt-kho` (sổ kho vs BCTC).

Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "đối chiếu mặt hàng", "mua vào bán ra có tương xứng", "hàng mua có bán không", "chuỗi trung gian hóa đơn", "biên lãi theo kg".
- Có file HĐĐT (portal / tổng hợp) đủ sheet mua vào và bán ra của một NNT.
- Cần chứng cứ trước khi kết luận thiếu đầu vào hoặc hợp thức hóa hóa đơn.

Không dùng cho đối chiếu NCC với sổ CVCBRR, không thay sổ 511/632, không thay XNT kho.

## Chuẩn bị

1. **MST NNT**, kỳ, thư mục hồ sơ. Thiếu file thì nêu rõ, không bịa.
2. File HĐĐT có sheet **Mua vao** và **Ban ra** (tên sheet có thể lệch dấu; dò theo nội dung). Header thường ở dòng 3, dữ liệu từ dòng 4.
3. Cột tối thiểu: ngày HĐ, số HĐ, ký hiệu, MST đối tác, tên đối tác, tên hàng, thành tiền, thuế suất. Cột số lượng / đơn giá / ĐVT nếu có thì dùng.
4. Thuế GTGT: nếu không có cột thuế thì **thành tiền × thuế suất**. Không giả 0.
5. MST trống 100% (`CHUA_RO_MST`): suy NNT qua **một họ ký hiệu HĐ bán ra duy nhất**, hỏi chủ xác nhận trước khi gắn nhãn hồ sơ.

## Quy trình

### 1. Nạp và lọc

1. Bỏ dòng trống, dòng Tổng cộng, dòng ghi chú / diễn giải (so `tinh_chat` sau khi **gỡ dấu**).
2. Bỏ HĐ hủy, HĐ thay thế, HĐ bị thay. Lọc **trước** khi so SL theo tháng.
3. Chuẩn hóa MST 10/13 số (bổ sung 0 đầu). `fold` tên hàng: gỡ dấu, thường, gom khoảng trắng.

### 2. Cửa sổ kỳ chồng (bắt buộc khi có SL)

Kỳ chồng = giao của khoảng ngày mua và khoảng ngày bán sau lọc. Chỉ so SL / tiền hàng hóa trên kỳ này. Phần mua trước cửa sổ và bán ngoài cửa sổ ghi riêng, không lấy làm "thiếu đầu vào".

Thiếu cột SL: chỉ so tiền, ghi **phủ chưa đủ**, không kết luận thiếu hàng.

### 3. Nhóm hàng và khóa tên

Nhóm gợi ý (đổi theo ngành NNT, không cứng CTL):

- A hàng hóa chính (vd thép, nông sản)
- B hàng hóa phụ (vd vỏ xe, vật tư)
- C dịch vụ / thuê / xăng / CNTT
- H khác
- Z trống tên

`item_key`: gỡ dấu; bỏ hậu tố kiểu `x C`; thống nhất `6.0 mm` / `6.0mm`; thống nhất biến thể chính tả (vd Trand/Strand). Không kết luận "chỉ bán, không mua" khi tên lệch mà **SL tháng của nhóm chính ≈ 1,000**.

### 4. Số liệu phải ra

1. Tổng dòng, số HĐ, số NCC / khách, thành tiền, thuế - mua vs bán trên kỳ chồng.
2. Theo nhóm: dòng, tiền, SL (nếu có), tỷ lệ SL bán/mua.
3. Theo tháng nhóm chính: SL mua, SL bán, tỷ lệ. Tỷ lệ 1,000 từng tháng = khớp vật lý.
4. Theo `item_key`: tiền mua, tiền bán, SL; cột "chỉ bán" phải đối chiếu lại khóa tên trước khi kết luận thiếu.
5. Cặp mua-bán **cùng SL** (sai số tương đối ≤ 0,5%) trong 0-2 ngày: đếm cùng ngày, đo **đồng/ĐVT** (vd đ/kg) trước, % biên lãi sau. Cụm cộng cố định (+10 / +15 / +25 đ/kg) là dấu trung gian mạnh hơn biên 0,05-0,11%.
6. Đối tác **hai chiều** (vừa NCC vừa khách): đủ 100% đối tác + dòng **Tổng cộng**. Cấm chỉ giữ bất thường.
7. Chuỗi A → NNT → B cùng ngày trùng tiền / trùng SL.

### 5. Xuất file

Font **Times New Roman**. Excel số thật `#,##0` / `#,##0.00`. Ngày `dd/mm/yyyy`. Chat/Word: dấu `.` nghìn, dấu `,` lẻ.

Sáu sheet, sắp chi tiết theo `Ngay_HD` (không theo MST khi truy ngược kỳ):

0. `Tom tat ket luan` - tóm tắt trước: số NCC/khách, tỷ lệ SL nhóm chính, số cặp cùng ngày, số đối tác hai chiều, kết luận khớp vật lý hay trung gian.
1. `Tong hop nhom hang`
2. `Doi chieu mat hang`
3. `Cap mua-ban trung khop`
4. `Dien bien theo thang`
5. `Doi tac` - 100% + Tổng cộng

Word cùng thư mục khi được hỏi. Không xóa file nguồn. File đang mở thì lưu bản `_cap_nhat`.

Trong pipeline hồ sơ thuế v1.4: skill này là **skill cạnh**. Không tự dựng sheet KQ2 mới. Chứng cứ đưa vào nhận xét `KQ2_Kho_Gia_thanh` hoặc dấu hiệu KQ3 khi cán bộ hỏi; không giao `DOI_CHIEU_*` như sản phẩm bộ 1.4.

## Định dạng số tiền

- Hiển thị: `1.234.567` hoặc `1.234.567,89`.
- Excel: `#,##0` / `#,##0.00`.
- Tỷ lệ: `12,34%` hoặc `1,000` (SL).
- Cấm dạng Mỹ `1,234,567.89`.

## Bẫy

- Kết luận thiếu đầu vào vì file chưa có cột SL (CTL: mua 631 tỷ vs bán 1.249 tỷ là thiếu SL, không phải thiếu hàng).
- So SL khi chưa lọc HĐ hủy: tháng lệch giả.
- So toàn bộ dải ngày mua (kể cả trước khi có bán) rồi bảo thiếu hàng.
- Khớp theo tên thô: `x C`, `6.0 mm` vs `6.0mm` thành "chỉ bán".
- Chỉ nhìn % biên lãi, bỏ mốc cộng đồng/ĐVT trên cặp cùng kg.
- Rút gọn sheet đối tác, bỏ dòng Tổng cộng.
- Nhầm skill này với F0/F1 (`doi-chieu-ncc-mua-vao-rui-ro`) hoặc 511/632.
- Xóa file nguồn. Ô text bắt đầu bằng `=`.

## Kiểm chứng

- Tổng tiền / SL sheet tổng = tổng dòng đã lọc trên kỳ chồng.
- Nhóm chính: từng tháng có tỷ lệ SL; nếu = 1,000 thì không được kết luận thiếu hàng theo tên.
- Sheet đối tác đủ 100% + Tổng cộng.
- Cặp trùng khớp ghi được đ/ĐVT, không chỉ %.
- Font Times New Roman. Nguồn thiếu thì ghi thiếu, không giả 0.

## Liên kết quy trình hồ sơ

Skill cạnh của `phan-tich-ho-so-thue`. Đầu vào HĐĐT: `gop-hoa-don-mua-ban` / `tong-hop-dau-vao-hddt-bang-ke-nld`. Sau khi có sổ kho: `phan-tich-xnt-kho`. Sau khi có sổ 511/632: `doi-chieu-511-632`. Gắn nhãn F0/F1: `doi-chieu-ncc-mua-vao-rui-ro`.
