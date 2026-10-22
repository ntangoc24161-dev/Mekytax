---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
---
Khi đối chiếu NCC mua vào với CVCBRR/danh bạ:

1. Chạy trên **hai nguồn**: Bảng kê mua vào (công ty) và HĐĐT mua vào (CQ thuế).
2. **Một bảng duy nhất** cho danh sách NCC; cột `Nguồn dữ liệu` ghi `Bảng kê` / `HĐĐT` / `Chung` (có ở cả hai). Không tách 2 bảng riêng theo nguồn.
3. Bổ sung **Bảng 2.3 đối chiếu theo kỳ kê khai** cho NCC rủi ro đã chọn (F0 / ĐKKD bất thường / chọn lọc): kỳ kê khai, số HĐ, ngày HĐ, mặt hàng, giá trị chưa thuế, thuế GTGT, loại rủi ro, căn cứ CV, cờ có trên Bảng kê/HĐĐT.

**Vì sao:** Cần thấy NCC chỉ xuất hiện một phía hoặc lệch giữa Bảng kê và HĐĐT, đồng thời truy ngược nhanh NCC rủi ro về đúng kỳ kê khai và dòng HĐ để kiểm tra chứng từ.

**Áp dụng:** Sheet `Tong_hop_NCC_CVCBRR` trong Master MAU; file `Tong_hop_NCC_mua_vao_CV_trang_thai_MAU.xlsx` (sheet Bang_2_1 / 2_2 / 2_3); skill `doi-chieu-ncc-mua-vao-rui-ro`.
