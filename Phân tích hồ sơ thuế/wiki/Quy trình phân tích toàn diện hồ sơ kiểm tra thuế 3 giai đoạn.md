---
type: wiki
status: active
tags: [wiki, TTKT, phan-tich-ho-so, quy-trinh]
created: '2026-10-01'
updated: '2026-10-01'
source: '[[phan-tich-ho-so-thue]]'
---
# Quy trình phân tích toàn diện hồ sơ kiểm tra thuế 3 giai đoạn

Quy trình chuẩn hóa phục vụ công tác thanh tra - kiểm tra thuế tại cơ quan thuế và tại trụ sở Người nộp thuế (NNT), tích hợp giữa dữ liệu điện tử ngành Thuế và hồ sơ nội bộ doanh nghiệp giải trình/cung cấp.

## 1. Giai đoạn 1: Phân tích trước kiểm tra tại bàn (Dữ liệu ngành Thuế)

1. **Chuẩn hóa & Gộp bảng kê HĐĐT:** Gộp các file bảng kê rời/HDDT portal thành 1 file 2 sheet (`mua vao`, `ban ra`). Lọc bỏ dòng nhãn `Tổng cộng` trên cả hai sheet, bảo toàn các hóa đơn của khách hàng có tên "TỔNG CÔNG TY...". Map theo header để tránh lệch cột giữa các sheet.
2. **Đối chiếu 3 chiều GTGT - TNDN - HĐĐT (10 năm, 40 quý):** Đối chiếu Doanh thu bán ra HĐĐT với Chỉ tiêu [34] Tờ khai GTGT (bóc tách hóa đơn hủy, tờ khai bổ sung) và Chỉ tiêu [01] Doanh thu thuần trên BCTC.
3. **Rà soát Nhà cung cấp (NCC) mua vào với CVCBRR & Danh bạ NNT:** Phân loại 3 nhóm đối chiếu: (1) NNT với F1/Công an; (2) Toàn bộ NCC mua vào với F0 (cảnh báo bán rủi ro cao); (3) Rà soát NCC có trạng thái ĐKKD bất thường (đóng MST, tạm ngừng, giải thể, bỏ địa chỉ KD).
4. **Bóc tách Quyết toán TNCN & Rà soát đa nguồn (2020-2025):** Bóc tách danh sách lao động, đối chiếu Tờ khai Quý vs Quyết toán năm (TC22) và Lương BCTC (TC23). Tra cứu 100% lao động trên hệ thống Quản lý thuế ngành Thuế để xác định thu nhập đa nguồn và phát hiện trùng số CCCD/CMND trên XML.
5. **Chấm điểm 28 tiêu chí rủi ro (QĐ 1109/QĐ-TCT & Tool 4.4):** Chấm điểm toàn diện, khoanh vùng các tiêu chí đỏ: Tiêu chí 3 (Tồn kho/DT > 30%), Tiêu chí 28 (Giá vốn > 90% DT, biên lãi gộp mỏng), Tiêu chí 26 (NCC rủi ro F0), Tiêu chí 20 (Nợ phải trả NCC lớn).

## 2. Giai đoạn 2: Phân tích chuyên sâu dữ liệu NNT cung cấp/giải trình

1. **Phân tích Sổ Nhập - Xuất - Tồn Kho (TK 152, TK 155, TK 156):**
   - Trích xuất số liệu Tồn ĐK, Nhập TK, Xuất TK, Tồn CK (Số lượng & Thành tiền) từng năm.
   - **Đối chiếu số dư cuối kỳ:** So sánh Tổng Tồn CK (TK 152 + TK 155) với Chỉ tiêu [141] trên Bảng CĐKT BCTC.
   - Rà soát cơ cấu tồn kho (tỷ trọng NVL vs Thành phẩm), kiểm tra tính hợp lý của tỷ lệ HTK/DT để giải tỏa Tiêu chí 3.
2. **Phân tích Tập hợp Chi phí, Giá thành & Bảng Định mức:**
   - Bóc tách chi phí cấu thành giá thành: Chi phí NVL trực tiếp (TK 621), Nhân công trực tiếp (TK 622), Sản xuất chung (TK 627).
   - Đối chiếu Chi phí Nhân công TK 622 với QTT TNCN và Bảng lương nhân viên.
   - **Đối chiếu Giá vốn:** So sánh Giá vốn xuất bán TK 632 trên Sổ XNT TP / Giá thành với Chỉ tiêu [02] Giá vốn hàng bán trên BCTC.
   - Phân tích Bảng định mức tiêu hao nguyên vật liệu kỹ thuật so với thực tế xuất dùng TK 621 để giải tỏa rủi ro Tiêu chí 28.
3. **Phân tích Sao kê Ngân hàng & Thanh toán không dùng tiền mặt:**
   - Kiểm tra danh mục tài khoản ngân hàng của NNT (đã đăng ký với CQT theo Mẫu 08-MST chưa).
   - Rà soát điều kiện thanh toán không dùng tiền mặt đối với hóa đơn mua vào >= 20 triệu đồng (Thông tư 219, TT 78, TT 96).
   - Đối chiếu dòng tiền sao kê ngân hàng với các NCC thuộc diện cảnh báo rủi ro F0 / đóng MST / tạm ngừng hoạt động.
4. **Phân tích Sổ Chi tiết Công nợ (TK 131, TK 331) & Chi phí phân bổ:**
   - Rà soát công nợ phải thu (TK 131) và người mua trả tiền trước (TK 131 Có) để phát hiện doanh thu chưa kê khai.
   - Rà soát nợ phải trả NCC (TK 331) để giải tỏa rủi ro Tiêu chí 20.
   - Kiểm tra Bảng khấu hao TSCĐ (TK 214) và Bảng phân bổ Chi phí trả trước (TK 242).

## 3. Giai đoạn 3: Tích hợp Master Excel & Lập Báo cáo / Đề cương kiểm tra

1. **Tích hợp File Master Excel `Phan_tich_ho_so_cong_ty_<Tên>.xlsx` gồm 7 Sheet chuẩn hóa:**
   - Sheet 1: `Danh_gia_28_tieu_chi_rui_ro`
   - Sheet 2: `Tong_hop_GTGT_TNDN`
   - Sheet 3: `Tong_hop_NCC_CVCBRR`
   - Sheet 4: `Tong_hop_TNCN_va_Da_nguon`
   - Sheet 5: `Tong_hop_XNT_va_Kho`
   - Sheet 6: `Gia_thanh_va_Dinh_muc`
   - Sheet 7: `Cong_no_va_Ngan_hang`
2. **Lập Báo cáo kết quả & Đề cương kiểm tra tại trụ sở NNT:**
   - Báo cáo rõ các điểm đã giải tỏa (khớp đúng).
   - Báo cáo rõ các trọng tâm rủi ro cần kiểm tra thực địa tại trụ sở NNT.
   - Báo cáo danh mục các nội dung thiếu dữ liệu / chứng từ chưa thực hiện được để yêu cầu NNT cung cấp khi công bố Quyết định kiểm tra.
