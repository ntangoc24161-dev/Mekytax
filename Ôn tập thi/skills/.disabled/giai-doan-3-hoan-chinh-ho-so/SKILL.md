---
name: Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN
description: "Hoàn chỉnh hồ sơ khi có sổ sách DN: 23 quy tắc GĐ3, chuỗi bằng chứng 5 mắt xích, Master 28 tiêu chí 7 cột, đề cương kiểm tra (KQ3)."
group: Tài chính
updated: '2026-10-15'
---

# Giai đoạn 3 – Hoàn chỉnh hồ sơ khi có sổ sách doanh nghiệp

Thực hiện phân tích toàn diện chuyên sâu khi Người nộp thuế (NNT) đã cung cấp sổ kế toán, sổ kho, giá thành, sổ chi tiết tiền mặt, sao kê ngân hàng và sổ chi tiết TK 511/632 (hồ sơ đạt mức C, D hoặc E). Chạy 23 quy tắc rủi ro chuyên sâu của Giai đoạn 3, rà soát nâng các quy tắc Giai đoạn 2 từ "Sơ bộ" lên "Chắc chắn" (hoặc giải tỏa rủi ro), dựng chuỗi bằng chứng 5 mắt xích cho các giao dịch trọng điểm, tổng hợp Master 28 tiêu chí chuẩn 7 cột và lập Đề cương kiểm tra tại trụ sở NNT phân nhánh xử lý theo thẩm quyền.

> Áp dụng nghiêm ngặt **Khối Q** (Quy tắc chung): Khi có cả bản CQT và bản DN thì bản CQT là chuẩn so sánh; chênh lệch là phát hiện nhóm DQ; ứng dụng không tự ý gán nhãn "Trốn thuế" (cán bộ tự chọn tại cột Nhánh xử lý); kiểm lại độc lập 10 phát hiện điểm cao nhất từ dữ liệu gốc trước khi xuất báo cáo.

---

## Khi nào kích hoạt

- Đã có kết quả Giai đoạn 2 (`KQ2_<MST>_<KY>_v<n>.xlsx`), NNT vừa bàn giao bổ sung sổ sách kế toán.
- Hồ sơ kiểm tra đã đạt mức **C** (có sổ kế toán tổng hợp, 511/632), mức **D** (có thêm kho, giá thành) hoặc mức **E** (có thêm chi tiền mặt, sao kê ngân hàng).
- Các câu lệnh/yêu cầu: "phân tích giai đoạn 3", "chạy hoàn chỉnh hồ sơ", "đối chiếu sổ kế toán với tờ khai", "dựng chuỗi bằng chứng", "tái lập Master 28 tiêu chí", "lập đề cương kiểm tra", "lệnh P3".

---

## Đầu vào

Bên cạnh dữ liệu CQT đã chuẩn hóa ở GĐ2 (IN01–IN07, IN12), tiếp nhận và chuẩn hóa các nguồn DN:
1. **IN08_So_KeToan**: Sổ Nhật ký chung, Sổ Cái TK 111, 112, 131, 133, 152, 154, 155, 156, 331, 3388, 511, 632, 641, 642.
2. **IN09_Kho_GiaThanh**: Sổ Nhập–Xuất–Tồn kho (TK 152/155/156), Bảng tính giá thành đơn vị, Bảng định mức tiêu hao kỹ thuật.
3. **IN10_Chi_tien_mat**: Sổ chi tiết tiền mặt TK 1111 (chi tiết từng phiếu chi).
4. **IN11_Sao_ke_NH**: Dữ liệu sao kê ngân hàng chi tiết (Excel/dòng trích xuất) và danh mục Ủy nhiệm chi.
5. **IN13_So_511_632**: Sổ chi tiết phát sinh Có TK 511 và phát sinh Nợ TK 632 theo từng dòng bút toán (đầu vào của tool đối chiếu 511–632).

---

## Quy trình thực hiện

1. **Kiểm kê lại phần mới nộp:** Tự động gọi lại P1 cho các file mới nhận, cập nhật bảng manifest IN00 và nâng mức hồ sơ từ A/B lên C/D/E.
2. **Chuẩn hóa và đối chiếu CQT vs DN:**
   - So sánh BCTC và tờ khai do DN gửi với bản đã nộp trên cổng CQT. Nếu có sai lệch số liệu, lập phát hiện nhóm `DQ07` (không tự ý bỏ bản nào).
   - Chuẩn hóa các sổ kế toán vào IN08–IN11 và IN13.
3. **Chạy 23 quy tắc chuyên sâu Giai đoạn 3:**
   - Lọc theo loại hình doanh nghiệp (`00_Huong_dan`): Nếu là DN Thương mại thì bỏ qua `KH03`, `KH05`, `LD05` (ghi nhận `KHÔNG_ÁP_DỤNG`); nếu là DN Sản xuất thì áp dụng đầy đủ.
   - Nguồn nào còn thiếu thì quy tắc tương ứng ghi `KHÔNG_ĐỦ_DỮ_LIỆU`, không dừng toàn bộ giai đoạn.
4. **Nâng cấp kết luận quy tắc GĐ2:**
   - Các quy tắc ở GĐ2 mang nhãn "Sơ bộ" (như `TN01`, `TN03`, `TN05`–`TN14`, `TT01`): đối chiếu với sổ kế toán thực tế để:
     - Nâng lên kết luận **"Chắc chắn"** nếu phát hiện sai phạm thực sự.
     - Hoặc **"Giải tỏa"** rủi ro nếu DN giải trình và chứng minh được bằng hạch toán hợp lý (ví dụ: giải tỏa TC3 bằng Sổ XNT kho khớp đúng [141] theo B01-DN hoặc ct140 theo B01-DNN - Quyết định K6).
   - Ghi nhận chi tiết vào sheet `KQ3_Thay_doi_so_voi_GD2`.
5. **Dựng Chuỗi bằng chứng 5 mắt xích (cho các NCC/HĐ trọng điểm - Quyết định K11):**
   - Mắt xích 1: Hóa đơn điện tử hợp pháp (trạng thái CQT, không bị hủy/thay thế).
   - Mắt xích 2: Hợp đồng kinh tế / Đơn đặt hàng / Báo giá / Biên bản nghiệm thu giao nhận.
   - Mắt xích 3: Chứng từ thanh toán ngân hàng (Ủy nhiệm chi, Giấy báo Nợ khớp số tiền và tài khoản đăng ký).
   - Mắt xích 4: Chứng từ nhập kho thực tế (Phiếu nhập kho, thẻ kho, sổ chi tiết NVL/hàng hóa).
   - Mắt xích 5: Sử dụng / Xuất bán thực tế (Phiếu xuất kho sản xuất, bảng phân bổ, hợp đồng bán ra).
   - **Quy tắc K11:** Thiếu chứng từ công ty cung cấp (sao kê, UNC, hợp đồng) không phải là phát hiện vi phạm. Mắt xích thiếu nguồn ghi nhận là **"Không đủ dữ liệu"** (không ghi "Không").
6. **Tổng hợp Master 28 tiêu chí rủi ro (7 cột chuẩn - Quyết định K3):**
   - Lấy kết quả từ BM05 tổng hợp thành sheet `KQ3_Master_28TC` theo đúng 7 cột chuẩn: `ID`, `Tên tiêu chí`, `Phân loại`, `Ngưỡng rủi ro`, `Kết quả phân tích trên hồ sơ`, `Mức độ rủi ro`, `Kiến nghị kiểm tra`.
   - **Quyết định K3:** Nếu không tìm thấy file mẫu Master 8 sheet thì bỏ qua: **chỉ dựng sheet 28 tiêu chí (sheet 08)**, không tự dựng 7 sheet còn lại và không báo như lỗi.
7. **Đối chiếu Doanh thu 511 vs Giá vốn 632 (Quyết định K4):**
   - Thực hiện theo logic và cấu trúc cột của skill `doi-chieu-511-632`. Khóa ghép và cột xuất lấy từ skill, thay thế các khóa giả định.
8. **Lập Đề cương kiểm tra tại trụ sở NNT:**
   - Sắp xếp các nội dung rủi ro theo tổng điểm giảm dần.
   - Cột **Nhánh xử lý**: để trống cho cán bộ lựa chọn (`1. Sai sót thủ tục`, `2. Khai sai dẫn đến thiếu thuế`, `3. Có dấu hiệu trốn thuế`). Ứng dụng không tự ý gán nhãn vi phạm.
   - Nêu rõ danh mục mẫu chứng từ NNT cần xuất trình và bộ câu hỏi chất vấn.
9. **Tự kiểm tra độc lập (Self-check):**
   - Soát lại 10 phát hiện có điểm cao nhất trực tiếp từ dữ liệu gốc (IN01, IN02, IN08, IN09, IN11, IN13). Đảm bảo không có số 0 giả, tổng khớp chi tiết 100%.

---

## Chi tiết 23 quy tắc Giai đoạn 3

1. **`DQ01`** (Thiếu nguồn dữ liệu đầu vào theo quy định).
2. **`GT06`** (Đối chiếu BKMV vs PS Nợ TK 133 trên NKC theo tháng/năm - BT8/BT9).
3. **`GT16`** (Hóa đơn mua vào đạt ngưỡng thanh toán tiền mặt nhưng không có chứng từ thanh toán ngân hàng).
4. **`TN04`** (Tỷ lệ Giá vốn / Doanh thu thuần cao bất thường - TC28).
5. **`TN17`** (Số dư Nợ phải trả người bán TK 331 lớn bất thường so với quy mô hoạt động).
6. **`TN18`** (Chi phí quản lý doanh nghiệp TK 642 tăng đột biến không tương xứng doanh thu).
7. **`TT02`** (Phiếu chi tiền mặt thanh toán tiền mua hàng vượt ngưỡng P10/P11).
8. **`TT03`** (Chi tiền mặt nhiều lần trong cùng 1 ngày cho cùng 1 NCC có tổng giá trị vượt ngưỡng - BT7).
9. **`TT04`** (Phát sinh vay mượn tiền mặt nội bộ / cá nhân lớn TK 3388, TK 341).
10. **`TT05`** (Công nợ phải thu khách hàng TK 131 tồn đọng lâu năm không thu hồi).
11. **`TT06`** (Nghiệp vụ rút tiền gửi ngân hàng nhập quỹ tiền mặt với giá trị lớn bất thường).
12. **`KH01`** (Sổ kho: Tồn kho âm tại một số thời điểm trong kỳ).
13. **`KH02`** (Hàng tồn kho luân chuyển chậm, ứ đọng kéo dài không trích dự phòng).
14. **`KH03`** (Định mức tiêu hao NVL sản xuất thực tế vượt quá định mức kỹ thuật đã ban hành).
15. **`KH04`** (Đối chiếu Tổng tồn cuối kỳ TK 152+155+156 với Chỉ tiêu [141]/[140] trên BCTC).
16. **`KH05`** (Cơ cấu chi phí 621, 622, 627 bất thường trong giá thành sản phẩm).
17. **`KH06`** (Nghiệp vụ xuất kho bán hàng hoặc điều chuyển không có hóa đơn chứng từ kèm theo).
18. **`KH07`** (Đối chiếu giá trị xuất kho thành phẩm/hàng hóa với Giá vốn hàng bán TK 632).
19. **`DT01`** (Đối chiếu 511 vs 632: Bán dưới giá vốn, lợi nhuận gộp âm từng dòng bút toán).
20. **`DT02`** (Đối chiếu 511 vs 632: Dòng ghi nhận Doanh thu 511 nhưng không có Giá vốn 632).
21. **`DT03`** (Đối chiếu 511 vs 632: Dòng ghi nhận Giá vốn 632 nhưng không có Doanh thu 511).
22. **`LD02`** (Lệch giữa Chi phí lương trên BCTC [TK 641, 642, 622] với TNCT trên QTT TNCN - TC23).
23. **`LD05`** (Chi phí nhân công trực tiếp vượt mức tương quan với doanh thu sản xuất).

---

## Hợp đồng đầu ra

Xuất file Excel kết quả tích lũy **chỉ chứa các giá trị tĩnh**:
`KQ3_<MST>_<KY>_v<n>.xlsx` (mẫu chuẩn tại `KQ3_MAU_Hoan_chinh_ho_so.xlsx`) gồm 9 sheet:
1. **`KQ3_Phat_hien_tong`**: Tổng hợp cả 80 quy tắc rủi ro, mỗi quy tắc có 1 trạng thái chuẩn.
2. **`KQ3_511_632`**: Kết quả đối chiếu chi tiết 511 vs 632 (lãi gộp, bán dưới giá, dòng không khớp).
3. **`KQ3_Kho_Gia_thanh`**: Bảng phân tích nhập xuất tồn kho và giá thành định mức.
4. **`KQ3_Cong_no_Tien_mat`**: Kết quả rà soát công nợ 131/331, quét chi tiền mặt 111 và sao kê ngân hàng 112.
5. **`KQ3_Chuoi_bang_chung`**: Bảng đánh giá 5 mắt xích cho các NCC và hóa đơn trọng điểm.
6. **`KQ3_Master_28TC`**: Bảng Master 28 tiêu chí rủi ro đủ 7 cột chuẩn.
7. **`KQ3_De_cuong`**: Đề cương kiểm tra tại trụ sở NNT xếp theo điểm, có cột nhánh xử lý.
8. **`KQ3_Thay_doi_so_voi_GD2`**: Bảng theo dõi các quy tắc đổi trạng thái, nâng thành chắc, hoặc giải tỏa so với GĐ2.
9. **`KQ3_Con_thieu`**: Bảng tổng hợp các nội dung DN vẫn chưa cung cấp và kiến nghị xử lý.

---

## Câu lệnh thực thi mẫu

### Lệnh đầy đủ (P3):
```text
Giai đoạn 3 – Hoàn chỉnh hồ sơ <FOLDER>, MST <MST>, kỳ <KY>. Áp dụng khối Q. Đã có kết quả giai đoạn 2 (KQ2 v<n>), công ty vừa gửi sổ sách.
1. Chạy P1 lại cho phần mới nhận, cập nhật mức hồ sơ.
2. Chuẩn hóa nguồn công ty vào IN08–IN11, IN13 (và IN04, IN03 nếu công ty gửi bản riêng). So từng nguồn công ty với bản cơ quan thuế; chênh lệch ghi thành phát hiện DQ.
3. Chạy 23 quy tắc giai đoạn 3 cho nguồn đã có. Nguồn còn thiếu thì quy tắc ghi KHÔNG_ĐỦ_DỮ_LIỆU, không dừng cả giai đoạn. Chọn nhánh thương mại hay sản xuất theo loại hình: <LOẠI HÌNH>.
4. Chạy lại các quy tắc giai đoạn 2 đang ở nhãn "sơ bộ" mà nay có sổ; nâng thành "chắc" hoặc đổi kết luận, ghi rõ khác gì so với KQ2.
5. Dựng chuỗi bằng chứng 5 mắt xích (hóa đơn, hợp đồng/đơn hàng, thanh toán, nhập kho, sử dụng) cho NCC và HĐ trọng điểm; chấm điểm và ước tính tác động thuế.
6. Lập Master 28 tiêu chí, đề cương xếp theo điểm, danh mục nội dung còn thiếu.
7. Kiểm lại độc lập 10 phát hiện điểm cao nhất từ dữ liệu gốc.
Đầu ra: KQ3_<MST>_<KY>_v1.xlsx theo hợp đồng đầu ra mục 4, có KQ3_Thay_doi_so_voi_GD2 và KQ3_Con_thieu.
```

### Lệnh siêu ngắn:
```text
Hoàn chỉnh GĐ3 hồ sơ <FOLDER> MST <MST> kỳ <KY> loại hình <LOẠI HÌNH> xuất KQ3
```
