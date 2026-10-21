---
type: wiki
status: active
tags: [wiki, TTKT, phan-tich-ho-so, quy-trinh, checklist]
created: '2026-10-02'
updated: '2026-10-02'
source: '[[phan-tich-ho-so-thue]]'
---
# Quy trình phân tích toàn diện hồ sơ kiểm tra thuế 4 bước và Bộ checklist hồ sơ

Quy trình chuẩn hóa phục vụ công tác thanh tra - kiểm tra thuế tại cơ quan thuế và tại trụ sở Người nộp thuế (NNT), tích hợp khép kín 4 bước từ dữ liệu thô ban đầu, tiếp nhận hồ sơ NNT cung cấp, rà soát khoảng trống và bổ sung dữ liệu đến phân tích chuyên sâu đa chiều [[phan-tich-ho-so-thue]].

---

## 1. Khung quy trình phân tích 4 bước chuẩn hóa

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
   - Dữ liệu theo dõi nộp thuế, thông báo nợ thuế, thông tin đăng ký thuế ban đầu.

---

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

---

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

---

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

---

## 2. Bảng ma trận đối chiếu loại hồ sơ, định dạng và mục tiêu kiểm tra

| Nhóm hồ sơ | Loại tài liệu | Nguồn cung cấp | Định dạng chuẩn | Mục tiêu kiểm tra đối chiếu trọng tâm |
|:---|:---|:---|:---:|:---|
| **Dữ liệu Tờ khai thuế** | Tờ khai GTGT (842), QTT TNDN (892), QTT TNCN (953), BCTC (402), KHBS | Cơ quan thuế chuẩn bị | XML điện tử | Đối chiếu 3 chiều doanh thu, thuế phát sinh, chi phí lương và cấu trúc tài chính 10 năm |
| **Dữ liệu Hóa đơn** | HĐĐT bán ra/mua vào, Bảng kê chi tiết | CQT / Portal HĐĐT | Excel / XML | Bóc tách HĐ hủy, điều chỉnh, chi tiết từng NCC và xác định cụm HĐ trong 1 ngày |
| **Sổ Kho & XNT** | Sổ Tổng hợp & Chi tiết XNT (TK 152, 155, 156) | NNT cung cấp | Excel / Data | Đối chiếu khớp Chỉ tiêu [141] BCTC, rà soát tỷ lệ HTK/DT (giải tỏa TC3), hàng chậm luân chuyển |
| **Giá thành & Định mức** | Sổ TK 154 (621, 622, 627), Bảng tính Z SP, Định mức kỹ thuật | NNT cung cấp | Excel / Scan có dấu | Đối chiếu khớp Chỉ tiêu [02] Giá vốn BCTC, kiểm tra tỷ lệ lãi gộp (TC28), hao hụt nguyên liệu |
| **Công nợ & Vay mượn** | Sổ Chi tiết TK 131, TK 331, TK 336, TK 3388, TK 341 | NNT cung cấp | Excel / Data | Giải tỏa nợ NCC lớn (TC20), phát hiện doanh thu chưa khai (131 Có), kiểm tra mượn tiền không lãi |
| **Tiền mặt & Ngân hàng** | Sổ Chi tiết TK 111, Sổ TK 112, Sao kê Ngân hàng, UNC | NNT cung cấp | Excel + PDF có dấu | Quét chi tiền mặt vượt ngưỡng (20 triệu / 5 triệu), đối chiếu chứng từ không dùng tiền mặt TT 219 |
| **Lao động & Tiền lương** | Bảng thanh toán lương 12 tháng, Bảng chấm công, HĐLĐ | NNT cung cấp | Excel / PDF ký nhận | Đối chiếu khớp Chi phí 622, QTT TNCN (TC22, TC23), xác minh người lao động có thật |
| **Tra cứu Cảnh báo & Đa nguồn** | Sổ CVCBRR F0/F1, Sổ Công an, Danh bạ NNT, Tra cứu TNCN | Hệ thống QLT bổ sung | Excel / Tra cứu QLT | Nhận diện NCC rủi ro F0, NNT bỏ trốn/đóng MST, xác định người làm nhiều nơi và trùng CCCD |

---

## 3. Mẫu Phiếu yêu cầu cung cấp thông tin, tài liệu chuẩn gửi NNT

```markdown
CỤC THUẾ TỈNH / TP ...                     CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
CHI CỤC THUẾ ...                                Độc lập - Tự do - Hạnh phúc
Số: ... /PYC-CCT                                 ----------------------------
                                                 ..., ngày ... tháng ... năm 2026

                       PHIẾU YÊU CẦU CUNG CẤP THÔNG TIN, TÀI LIỆU
                 (Phục vụ công tác kiểm tra thuế tại trụ sở Người nộp thuế)

Kính gửi: CÔNG TY ... (Mã số thuế: ...)

Căn cứ Quyết định kiểm tra thuế số .../QĐ-CCT ngày .../.../2026 của Chi cục trưởng Chi cục Thuế ...;
Để phục vụ công tác kiểm tra thuế theo đúng kế hoạch và quy định của Luật Quản lý thuế số 38/2019/QH14, Đoàn kiểm tra yêu cầu Công ty chuẩn bị và cung cấp đầy đủ các hồ sơ, tài liệu kế toán liên quan đến niên độ kiểm tra từ năm 20... đến năm 20... như sau:

1. Hệ thống Sổ sách Kế toán tổng hợp và chi tiết (dạng file dữ liệu Excel và bản in có ký tên, đóng dấu):
   - Sổ Nhật ký chung, Sổ Cái các tài khoản phát sinh;
   - Sổ Chi tiết Quỹ tiền mặt (TK 1111) và Sổ Tiền gửi ngân hàng (TK 1121);
   - Sổ Chi tiết Phải thu khách hàng (TK 131) và Sổ Phải trả người bán (TK 331);
   - Sổ Chi tiết Vay, mượn tiền cá nhân/tổ chức (TK 336, 3388, 341) kèm Hợp đồng vay/mượn;
   - Sổ Trích khấu hao TSCĐ (TK 214) và Bảng phân bổ chi phí trả trước (TK 242).

2. Hồ sơ Kho, Tập hợp Chi phí sản xuất và Tính giá thành:
   - Sổ Tổng hợp và Sổ Chi tiết Nhập - Xuất - Tồn Kho Nguyên vật liệu (TK 152), Kho Thành phẩm (TK 155), Kho Hàng hóa (TK 156);
   - Sổ Tập hợp Chi phí sản xuất (TK 154 / TK 621, 622, 627) và Bảng tính giá thành đơn vị sản phẩm chi tiết theo từng đơn hàng/lô hàng;
   - Bảng định mức tiêu hao nguyên vật liệu kỹ thuật chính do Giám đốc doanh nghiệp phê duyệt ban hành;
   - Biên bản kiểm kê quỹ tiền mặt và Biên bản kiểm kê hàng tồn kho thực tế tại thời điểm khóa sổ 31/12 các năm tài chính.

3. Hồ sơ Dòng tiền và Chứng từ thanh toán Ngân hàng:
   - File dữ liệu trích xuất Excel (Internet Banking) và bản sao kê ngân hàng có xác nhận của ngân hàng đối với toàn bộ các tài khoản thanh toán mở tại các Ngân hàng TMCP trong suốt thời kỳ kiểm tra;
   - Chứng từ gốc thanh toán không dùng tiền mặt (Ủy nhiệm chi, Giấy báo Nợ) kèm theo các hóa đơn mua hàng hóa, dịch vụ có giá trị lớn hoặc mua của các Nhà cung cấp thuộc diện cảnh báo rủi ro theo danh sách Đoàn kiểm tra thông báo.

4. Hồ sơ Lao động và Tiền lương:
   - Bảng chấm công, Bảng thanh toán tiền lương 12 tháng/năm có chữ ký nhận của người lao động hoặc chứng từ chuyển khoản lương qua ngân hàng;
   - Hợp đồng lao động, bản sao CCCD của toàn bộ nhân sự có tên trên Tờ khai Quyết toán thuế TNCN;
   - Hồ sơ giải trình về các trường hợp lao động có thông tin trùng số CCCD/CMND trên hệ thống.

Thời gian và địa điểm cung cấp:
Hồ sơ, tài liệu yêu cầu cung cấp gửi về Đoàn kiểm tra chậm nhất vào ngày .../.../2026 tại trụ sở Công ty hoặc qua thư điện tử chính thức của Đoàn kiểm tra: ...@gdt.gov.vn.

Đoàn kiểm tra thông báo để Công ty biết và phối hợp thực hiện./.

                                                          TRƯỞNG ĐOÀN KIỂM TRA
                                                         (Ký, ghi rõ họ và tên)
```


## Mâu thuẫn
- Quan điểm mới ([[conversations/2026-10-05]]): Master Excel phân tích hồ sơ công ty chuẩn hiện hành gồm **8 sheet** (thêm `Doi_chieu_511_632`), không còn dừng ở 7 sheet. (mục tiêu) [[conversations/2026-10-05]]

Sheet `Danh_gia_28_tieu_chi_rui_ro` phải đủ **28/28 tiêu chí** với **7 cột** chuẩn Vạn Phát Thành; cấm bản rút gọn tool v1 chỉ còn vài dòng. (mục tiêu) [[conversations/2026-10-05]]

Đối chiếu tồn kho sổ XNT với BCTC DNN lấy **ct140** (không mặc định [141]/ct141). (mục tiêu) [[conversations/2026-10-05]]

Doanh nghiệp thương mại: sheet 
- Cần xác minh (append _open-questions).
