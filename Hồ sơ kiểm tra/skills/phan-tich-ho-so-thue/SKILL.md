---
name: Phân tích hồ sơ kiểm tra thuế
description: Phân tích hồ sơ kiểm tra thuế từ XML tờ khai/BCTC, Excel hóa đơn, sổ kế toán XNT, giá thành, công nợ, sao kê ngân hàng và 28 tiêu chí rủi ro.
group: Tài chính
updated: '2026-11-05'
learned_updated: '2026-11-05'
---
# Phân tích hồ sơ kiểm tra thuế

Quy trình tự động hóa phân tích toàn diện hồ sơ kiểm tra thuế theo chu trình 4 bước chuẩn hóa: kết hợp dữ liệu điện tử ngành Thuế (XML tờ khai GTGT, TNDN, TNCN, BCTC, HĐĐT) với hồ sơ, sổ kế toán và sao kê ngân hàng do NNT cung cấp; rà soát khoảng trống dữ liệu để bổ sung nguồn tra cứu quản lý thuế; tiến hành phân tích chuyên sâu 3 chiều và đánh giá 28 tiêu chí rủi ro. Kết quả tích hợp thành 01 file Master Excel (7 sheet chuẩn hóa) + Báo cáo rủi ro Markdown + Đề cương kiểm tra tại trụ sở NNT.

## Khi nào dùng

- "phân tích hồ sơ thuế", "đối chiếu hóa đơn tờ khai", "chấm rủi ro BCTC", "28 tiêu chí", "phân tích sổ nhập xuất tồn", "đối chiếu giá thành định mức", "rà soát sao kê ngân hàng", "quét chi tiền mặt TK 111".
- Đoàn kiểm tra chuẩn bị kiểm tra tại bàn hoặc lập Đề cương kiểm tra tại trụ sở NNT.
- NNT đã nộp XML tờ khai và/hoặc cung cấp thư mục hồ sơ kế toán (`CTY GUI` gồm Sổ Nhật ký chung, Sổ Chi tiết TK, Sổ XNT, Bảng giá thành, Định mức, Sổ công nợ 131/331, Bảng lương, Sao kê ngân hàng).

## Chuỗi Bảng tính (skill cạnh)

Khi cần phôi `Bang tinh.xlsx` cho DI-01 / biên bản (không thay sáu tệp KQ của quy trình 1.4 ở shared-skills):

1. `tao-bang-tinh-tu-ptdn` - tạo `0. BIEN BAN\`, copy mẫu, đổi `F5/F6/F7`, xóa số mẫu.
2. `nhap-bang-tinh-tu-bao-cao-hdkd` - KQKDSXKD từ PTDN02 (+ BCTC `NamNay`); GTGT/EBITDA từ `KQ2.1_*_FINAL`.
3. `boc-tach-bang-tinh-phan-tich` - tỷ lệ chi phí, Mục III DI-01.

Engine quét hồ sơ bỏ qua thư mục `BIEN BAN` và file PTDN01/PTDN02; chuỗi trên chạy bằng skill cạnh, không nhét PTDN vào IN GĐ1.

**Chuẩn quy trình hiện hành:** bản shared-skills `phan-tich-ho-so-thue` (v1.4, sáu tệp KQ1.1-KQ3). Bản skill trong brain này giữ lịch sử 4 bước / Master; khi lệch ưu tiên shared + playbook 3 giai đoạn.

## Quy trình phân tích 4 bước chuẩn hóa

### Bước 1: Thu thập và chuẩn hóa dữ liệu thô ngành Thuế (Trước kiểm tra tại bàn)
1. **Dữ liệu Tờ khai thuế điện tử (XML):**
   - Tờ khai GTGT mẫu 842 (các kỳ tháng/quý và tờ khai bổ sung KHBS).
   - Tờ khai Quyết toán TNDN mẫu 892 kèm Phụ lục kết quả hoạt động SXKD (Phụ lục 03-1A/TNDN).
   - Tờ khai Quyết toán TNCN mẫu 953 kèm Phụ lục 05-1/BK-TNCN và 23+ kỳ tờ khai tạm tính Quý (mẫu 05/KK-TNCN).
   - Báo cáo tài chính mẫu 402 (Bảng CĐKT, Báo cáo KQKD, Báo cáo LCTT, Thuyết minh BCTC).
2. **Dữ liệu Hóa đơn điện tử (HĐĐT) & Bảng kê:**
   - Dữ liệu HĐĐT trích xuất từ Cổng thông tin điện tử ngành Thuế / Portal HĐĐT (`HDDT.xlsx`).
   - Bảng kê mua vào, bán ra các năm/quý.
   - Gộp bảng kê thành 1 file Excel 2 sheet (`mua vao`, `ban ra`), tự động lọc bỏ các dòng nhãn `Tổng cộng`, bảo toàn dòng đối tác "TỔNG CÔNG TY...".
3. **Dữ liệu Quản lý thuế & Dòng tiền nghĩa vụ:**
   - Dữ liệu theo dõi nộp thuế, tờ khai thông báo nợ, thông tin đăng ký thuế ban đầu.

### Bước 2: Thu thập và tiếp nhận hồ sơ do Người nộp thuế (NNT) cung cấp
1. **Hệ thống Sổ kế toán nội bộ:**
   - Sổ Nhật ký chung, Sổ Cái các tài khoản tổng hợp.
   - Sổ Chi tiết Tiền mặt (TK 1111), Sổ Chi tiết Tiền gửi Ngân hàng (TK 1121).
   - Sổ Chi tiết Phải thu khách hàng (TK 131) và Phải trả người bán (TK 331).
   - Sổ Chi tiết Vay mượn nội bộ / cá nhân (TK 336, 3388, 341).
2. **Sổ Kho và Tập hợp Chi phí - Giá thành:**
   - Sổ Tổng hợp & Chi tiết Nhập - Xuất - Tồn Kho Nguyên vật liệu (TK 152), Kho Thành phẩm (TK 155), Kho Hàng hóa (TK 156).
   - Sổ Tập hợp Chi phí sản xuất (TK 154 / 621, 622, 627) và Bảng tính giá thành đơn vị sản phẩm.
   - Bảng định mức tiêu hao nguyên vật liệu kỹ thuật đã ban hành qua các năm.
3. **Hồ sơ Lao động, Tiền lương và Khấu hao:**
   - Bảng thanh toán tiền lương và Bảng chấm công 12 tháng/năm.
   - Bảng trích Khấu hao TSCĐ (TK 214) theo Thông tư 45/2013/TT-BTC và Bảng phân bổ Chi phí trả trước (TK 242).
4. **Hồ sơ Dòng tiền & Chứng từ thanh toán:**
   - File sao kê chi tiết tài khoản ngân hàng (PDF và file dữ liệu Excel) của tất cả ngân hàng giao dịch.
   - Chứng từ gốc thanh toán không dùng tiền mặt (Ủy nhiệm chi, Giấy báo Nợ) kèm theo các hóa đơn mua vào lớn.

### Bước 3: Xử lý dữ liệu tổng hợp, rà soát khoảng trống & Bổ sung dữ liệu thiếu
1. **Rà soát khoảng trống dữ liệu ban đầu:**
   - Kiểm tra tính đầy đủ của các kỳ kê khai (đủ 4 quý/năm, đủ 10 năm hoặc phạm vi kiểm tra).
   - Rà soát định dạng sao kê ngân hàng (yêu cầu trích xuất bổ sung file Excel nếu NNT chỉ cung cấp PDF scan).
2. **Bổ sung nguồn dữ liệu đối chiếu ngoại ngành và Quản lý thuế:**
   - Tra cứu và tích hợp Sổ theo dõi Công văn cảnh báo rủi ro (CVCBRR F0 bên bán rủi ro cao, F1 bên mua).
   - Tích hợp Sổ Công an theo dõi các vụ án hóa đơn và Danh bạ trạng thái ĐKKD NNT (đóng MST, tạm ngừng, không hoạt động tại ĐCKD, giải thể).
   - Tra cứu dữ liệu nguồn thu nhập cá nhân toàn quốc từ hệ thống Quản lý thuế ngành Thuế: xác định tình trạng lao động có 1 nguồn hay nhiều nguồn thu nhập, số thuế đã khấu trừ 10% tại nơi khác, kiểm tra trùng số CCCD/CMND trên tờ khai XML.
   - Rà soát danh mục tài khoản ngân hàng thực tế so với đăng ký thuế (Mẫu 08-MST / Thông tư 105).
3. **Yêu cầu NNT giải trình và cung cấp bổ sung chứng từ còn thiếu:**
   - Phát hành Phiếu yêu cầu cung cấp thông tin, tài liệu bổ sung (Ủy nhiệm chi cho các NCC rủi ro F0, Hợp đồng mượn tiền cá nhân TK 3388, Hồ sơ nhân sự xác minh trùng CCCD).

### Bước 4: Tiến hành phân tích chuyên sâu hồ sơ trên cơ sở dữ liệu tích hợp
1. **Đối chiếu 3 chiều GTGT - TNDN - HĐĐT:**
   - Đối chiếu Doanh thu HĐĐT bán ra vs Chỉ tiêu [34] Tờ khai GTGT (theo từng quý và cả năm).
   - Bóc tách hóa đơn hủy, hóa đơn điều chỉnh/thay thế; giải trình nguyên nhân chênh lệch.
   - Đối chiếu Doanh thu BCTC Chỉ tiêu [01] vs Doanh thu kê khai thuế GTGT.
2. **Rà soát Nhà cung cấp (NCC) mua vào với dữ liệu cảnh báo rủi ro:**
   - Đối chiếu toàn bộ NCC với Sổ CVCBRR (F0, F1), Sổ Công an và Danh bạ NNT.
   - Bóc tách nhóm NCC rủi ro F0/Công an và nhóm có trạng thái ĐKKD bất thường để khoanh vùng kiểm tra chứng từ.
3. **Bóc tách Quyết toán TNCN & Rà soát đa nguồn:**
   - Đối chiếu khớp đúng giữa Tờ khai tạm tính quý (05/KK-TNCN) vs Quyết toán năm (TC22) và Chi phí lương BCTC (TC23).
   - Đánh giá nghĩa vụ tự quyết toán và thuế TNCN nộp thừa/thiếu của người lao động.
4. **Chấm điểm toàn diện 28 tiêu chí rủi ro (QĐ 1109/QĐ-TCT & Tool 4.4):**
   - Đánh giá rủi ro BCTC và tờ khai, xác định các tiêu chí đỏ: TC3, TC28, TC26, TC20, TC5.
5. **Phân tích Sổ Nhập - Xuất - Tồn Kho (TK 152, 155, 156) giải tỏa Tiêu chí 3:**
   - Đối chiếu Tổng Tồn CK (TK 152 + 155 + 156) với Chỉ tiêu [141] trên BCTC (khớp đúng 100%).
   - Bóc tách cơ cấu NVL vs Thành phẩm, rà soát tỷ lệ HTK/Doanh thu để giải tỏa hoặc xác nhận rủi ro TC3.
6. **Phân tích Tập hợp Chi phí, Giá thành & Bảng Định mức giải tỏa Tiêu chí 28:**
   - Bóc tách cơ cấu chi phí 621 (NVL), 622 (Nhân công), 627 (SXC) trong giá thành sản phẩm.
   - Đối chiếu Giá vốn xuất bán TK 632 với Chỉ tiêu [02] Giá vốn hàng bán trên BCTC (khớp đúng 100%).
   - Đối chiếu định mức tiêu hao kỹ thuật với số lượng nguyên vật liệu xuất dùng thực tế.
7. **Phân tích Sổ Chi tiết Tiền mặt (TK 111) & Sao kê Ngân hàng (TK 112):**
   - Quét toàn bộ phiếu chi tiền mặt vượt ngưỡng quy định: trên 20 triệu đồng (giai đoạn đến 06/2025) và từ 5 triệu đồng trở lên (từ 07/2025 đến nay theo Luật Thuế GTGT 48/2024).
   - Phân loại mục đích chi: Nộp tiền vào TK ngân hàng (112), Trả nợ vay/mượn cá nhân (336/3388), Chi trả lương (334), Mua hàng hóa dịch vụ.
   - Quét các cụm nhiều hóa đơn mua vào trong 1 ngày của cùng 1 NCC có tổng giá trị vượt ngưỡng thanh toán không dùng tiền mặt.
8. **Phân tích Sổ Chi tiết Công nợ (TK 131, TK 331) giải tỏa Tiêu chí 20:**
   - Rà soát công nợ phải thu (131) và người mua trả tiền trước (131 Có) để kiểm tra ghi nhận doanh thu.
   - Rà soát nợ phải trả NCC (331) theo dõi việc thanh toán qua ngân hàng đối với các nhà cung cấp lớn.
9. **Tích hợp Master Excel 7 Sheet & Lập Báo cáo / Đề cương kiểm tra:**
   - Tự động xuất file Master Excel `Phan_tich_ho_so_cong_ty_<Tên>.xlsx` gồm 7 Sheet chuẩn hóa:
     - Sheet 1: `Danh_gia_28_tieu_chi_rui_ro`
     - Sheet 2: `Tong_hop_GTGT_TNDN`
     - Sheet 3: `Tong_hop_NCC_CVCBRR`
     - Sheet 4: `Tong_hop_TNCN_va_Da_nguon`
     - Sheet 5: `Tong_hop_XNT_va_Kho`
     - Sheet 6: `Gia_thanh_va_Dinh_muc`
     - Sheet 7: `Cong_no_va_Ngan_hang`
   - Lập Báo cáo tóm tắt rủi ro, phân định rõ nội dung đã giải tỏa và các trọng tâm rủi ro cần kiểm tra thực địa.

## Bẫy nghiệp vụ cần tránh

- **Bẫy HĐ mua vào mỏng giả:** File `HDDT.xlsx` portal thường thiếu dữ liệu mua vào so với folder `Bang ke mua vao ban ra` rời. Phải đối chiếu độ phủ trước khi đánh giá lệch mua vào.
- **Bẫy dòng Tổng cộng:** Sau khi gộp bảng kê mua/bán, bắt buộc lọc bỏ các dòng nhãn `Tổng cộng` (thường mỗi quý 1 dòng); cẩn thận giữ lại các dòng khách hàng/đối tác có tên "TỔNG CÔNG TY...".
- **Bẫy lệch cột giữa các sheet:** Sheet `ban ra` và `mua vao` có thể khác cấu trúc cột; phải map theo tên header (Doanh số, Thuế, Thuế suất) thay vì số thứ tự cột cố định.
- **Bẫy thiếu kỳ tờ khai:** Khi doanh thu bảng kê cả năm lệch tờ khai, phải bóc tách theo từng quý để kiểm tra xem có thiếu file XML của quý nào không (ví dụ Q4).
- **Bẫy tồn kho BCTC vs Sổ XNT:** Phải cộng cả TK 152 (NVL), TK 155 (Thành phẩm) và TK 156 (Hàng hóa) mới bằng Chỉ tiêu [141] trên BCTC; nếu chỉ lấy TK 152 sẽ kết luận sai là lệch sổ sách.
- **Bẫy chi phí nhân công 622:** Chi phí 622 trên bảng giá thành thường là lương công nhân trực tiếp sản xuất, trong khi QTT TNCN bao gồm cả lương quản lý (TK 642) và bán hàng (TK 641); cần đối chiếu tổng thể với bảng thanh toán tiền lương.
- **Bẫy sao kê ngân hàng dạng PDF:** NNT thường gửi PDF scan hoặc PDF trích xuất; nếu chưa có file Excel thì không kết luận vội là không có chứng từ thanh toán, mà ghi rõ hiện trạng và kiến nghị yêu cầu file Excel / kiểm tra chọn mẫu UNC gốc.
- **Bẫy chi tiền mặt nhiều lần trong ngày:** Theo TT 219/2013 và Luật 48/2024, mua nhiều hóa đơn trong cùng một ngày của cùng một người bán có tổng giá trị vượt ngưỡng (20 triệu hoặc 5 triệu từ 07/2025) vẫn bắt buộc phải có chứng từ thanh toán qua ngân hàng.

## Kiểm chứng kết quả

- File Master Excel `Phan_tich_ho_so_cong_ty_<Tên>.xlsx` có đủ 7 sheet chuẩn hóa, font Times New Roman, định dạng số có dấu phẩy ngăn cách hàng nghìn, các tiêu chí rủi ro cao tô màu đỏ nổi bật.
- Toàn bộ số liệu đối chiếu 3 chiều (GTGT, TNDN, TNCN, BCTC, XNT, Giá thành, Công nợ, Tiền mặt) khớp đúng 100% với số liệu gốc.
- Đồng bộ lưu trữ tại thư mục `ket-qua/`, `7. Sao luu tong hop/` trên máy và thư mục `sao-luu/` trong Second Brain.
- Báo cáo rõ ràng, ngắn gọn, phân định rành mạch giữa nội dung đã làm được và nội dung thiếu dữ liệu cần thực hiện tiếp. Không dùng ký tự em dash.

## Lịch sử (tự học)
- [2026-09-30] Xử lý hồ sơ Vạn Phát Thành: gộp 20 file bảng kê rời thành 1 file 2 sheet; lọc bỏ dòng Tổng cộng nhưng giữ "TỔNG CÔNG TY"; map header tránh lệch cột bán ra vs mua vào; đối chiếu khớp 100% doanh thu 10 năm sau bóc tách 3 HĐ hủy năm 2025.
- [2026-10-01] Bóc tách 11 năm QTT TNCN và 25 bản ghi tra cứu đa nguồn: 100% lao động chỉ có 1 nguồn thu nhập, thuế TNCN = 0 đ; phát hiện trùng CCCD giữa 2 nhân sự năm 2025.
- [2026-10-01] Mở rộng quy trình phân tích hồ sơ NNT gửi: xử lý 10 năm Sổ XNT NVL (TK 152), Sổ XNT TP (TK 155), Bảng tính giá thành sản phẩm, Bảng định mức tiêu hao kỹ thuật, Sổ công nợ 131/331 và 48 file PDF sao kê ngân hàng (ACB, Indovina, Techcombank, HDBank). Tích hợp Master Excel thành 7 Sheet hoàn chỉnh.
- [2026-10-02] Bổ sung phân tích Sổ chi tiết Tiền mặt TK 1111 (quét 636 phiếu chi 10 năm, bóc tách 186 món vượt ngưỡng 20 triệu / 5 triệu từ 07/2025) và chuẩn hóa toàn bộ quy trình phân tích hồ sơ kiểm tra thuế thành 4 bước khép kín.
