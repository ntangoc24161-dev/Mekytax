# Javis Index (tầng vận hành)

> Tự sinh từ file - ĐỪNG sửa tay. Chỉ mục mọi năng lực của Javis trong brain này để bất kỳ AI/engine đọc 1 chỗ là hiểu Javis làm được gì. Song song `wiki/index.md` (tri thức).

**Tổng quan:** 26 agents · 57 skills · 9 workflows (6 bật) · 1 loops (1 bật) · 14 plugins (13 chạy)

## Agents
### AI
- **Trợ lý tri thức** (`tri-thuc-second-brain`) - Người giữ wiki Second Brain sống và trả lời có trích dẫn. · skills: ingest-source, query-wiki, lint-wiki
### Chung
- **phong-van-thu** (`phong-van-thu`) - 
### Tài chính
- **Chuyên viên công nợ và tiền mặt** (`chuyen-vien-cong-no-tien-mat`) - Quét chi tiền mặt 111, cụm HĐ cùng ngày, công nợ 131/331, sao kê 112. Ngưỡng K5 dùng >=. · skills: phan-tich-cong-no-tien-mat
- **Chuyên viên doanh thu giá vốn** (`chuyen-vien-doanh-thu-gia-von`) - Đối chiếu sổ 511 với 632 theo khóa HĐ/ngày/mặt hàng, lãi gộp, bán dưới giá (K4). · skills: doi-chieu-511-632
- **Chuyên viên dữ liệu đầu vào** (`chuyen-vien-du-lieu-dau-vao`) - Gộp HĐĐT, bảng kê, NLĐ theo mẫu 02; làm sạch CD1 BT5-BT9; không trộn nguồn. · skills: tong-hop-dau-vao-hddt-bang-ke-nld, lam-sach-cd1-mst-hoa-don
- **Chuyên viên GTGT-TNDN** (`chuyen-vien-gtgt-tndn`) - Đối chiếu bảng kê, tờ khai GTGT, QTT TNDN, HĐĐT theo 2 sheet chuẩn CBRR, bóc nguyên nhân trọng yếu. · skills: doi-chieu-gtgt-tndn-cbrr
- **Chuyên viên kho và giá thành** (`chuyen-vien-kho-gia-thanh`) - Đối chiếu XNT 152/155/156 với BCTC, xuất kho với 632, bóc giá thành/định mức, giải tỏa TC3 và TC28. · skills: phan-tich-xnt-kho, phan-tich-gia-thanh-dinh-muc
- **Chuyên viên NCC và hóa đơn** (`chuyen-vien-ncc-hoa-don`) - Đối chiếu 100% NCC với CVCBRR F0/F1, Công an, danh bạ; F0 và F1 đều mức Cao (K2). · skills: doi-chieu-ncc-mua-vao-rui-ro, tong-hop-ben-mua-tu-so-cvcbrr
- **Chuyên viên TNCN** (`chuyen-vien-tncn`) - Tổng hợp QTT TNCN, danh sách NLĐ, đối chiếu đa nguồn TC22/TC23, trùng CCCD. · skills: tong-hop-tncn-da-nguon
- **Cập nhật công văn cảnh báo rủi ro** (`cập-nhật-công-văn-cảnh-báo-rủi-ro`) - Bóc tách công văn cảnh báo đã số hóa và nạp dòng mới vào sheet F0 bên bán, F1 bên mua · skills: theo-doi-cv-canh-bao-rui-ro
- **Đối chiếu rủi ro NCC** (`doi-chieu-rui-ro-ncc`) - Soi nhà cung cấp mua vào với F0/F1, danh sách Công an và danh bạ NNT, lập Bảng 2.3 theo kỳ · skills: doi-chieu-ncc-mua-vao-rui-ro, theo-doi-cv-canh-bao-rui-ro
- **Khai thác dữ liệu hồ sơ** (`khai-thac-du-lieu-ho-so`) - trợ lý hỗ trợ - xuất bảng kê tiêu chí cung cấp và đối chiếu số liệu thuế khi công văn hoặc soạn thảo cần bằng chứng · skills: xuat-danh-sach-ho-so-ca, doi-chieu-ncc-mua-vao-rui-ro, phan-tich-cong-no-tien-mat, phan-tich-ho-so-thue, giai-doan-1-kiem-ke-danh-gia, giai-doan-2-phan-tich-du-lieu-co-quan-thue, giai-doan-3-hoan-chinh-ho-so
- **Trợ lý khởi tạo hồ sơ** (`khoi-tao-ho-so`) - Chuyên viên văn bản khởi tạo hồ sơ kiểm tra và cập nhật dữ liệu công ty. · skills: lap-ho-so-kiem-tra, nhap-bang-tinh-tu-bao-cao-hdkd, boc-tach-bang-tinh-phan-tich, xem-ho-so-thue
- **Lập hồ sơ kiểm tra** (`lap-ho-so-kiem-tra`) - Cán bộ Thuế cơ sở 3 lập hồ sơ kiểm tra trên app nội bộ và xuất bộ Word. · skills: lap-ho-so-kiem-tra
- **Phân tích viên dữ liệu CQT** (`phan-tich-vien-cqt`) - Chạy 57 quy tắc khi chưa có sổ DN, nhãn Chắc/Sơ bộ, xuất KQ2. Chỉ sau mức A hoặc B, hoặc để giữ phần Chắc từ CQT. · skills: giai-doan-2-phan-tich-du-lieu-co-quan-thue
- **Phân tích viên sổ sách DN** (`phan-tich-vien-so-sach-dn`) - Chạy 23 quy tắc sổ DN, chuỗi 5 mắt xích, nâng sơ bộ thành chắc hoặc giải tỏa, xuất KQ3. Chỉ khi mức C/D/E. · skills: giai-doan-3-hoan-chinh-ho-so
- **Pháp chế & chính sách thuế** (`phap-che-chinh-sach-thue`) - trợ lý hỗ trợ - trả lời câu hỏi luật/chính sách có trích dẫn điều khoản khi Điều phối hoặc Soạn thảo vướng căn cứ · skills: thue-gtgt, thue-tncn, thue-tndn, hoa-don-dien-tu, luat-quan-ly-thue, xu-ly-rui-ro-thue-hoa-don, xu-phat-vphc-thue-hoa-don, quy-trinh-hoan-thue, quy-trinh-kiem-tra-thue, huong-dan-kiem-tra-chuyen-de, kiem-tra-thue-thuong-mai-dien-tu, luat-thuong-mai-dien-tu, quy-che-phoi-hop-kiem-tra-thue, luat-phong-chong-tham-nhung
- **Tiếp nhận & Điều phối công văn** (`quản-lý-hồ-sơ-công-an`) - đầu vào phòng cung cấp hồ sơ Công an - tiếp nhận, bóc tách, tra cứu NNT, phân Tổ KT, cập nhật sổ và đóng thư mục · skills: theo-doi-cung-cap-ca, tra-cuu-nnt-tu-dong, tra-cuu-ioffice-dong-nai, pdf-sang-md, dong-bo-thu-muc-ca
- **Số hóa OCR công văn** (`so-hoa-ocr-cong-van`) - Chuyển PDF và ảnh scan công văn cảnh báo thành Markdown sạch, giữ heading và bảng · skills: pdf-sang-md
- **Trợ lý soạn thảo biên bản và hồ sơ về** (`soan-bien-ban-ho-so-ve`) - Chuyên viên văn bản soạn biên bản, hồ sơ về và quyết định xử phạt trình ký. · skills: soan-thao-van-ban, dien-bien-ban-tu-bang-tinh, xu-phat-vphc-thue-hoa-don
- **Trợ lý soạn thảo hồ sơ đi** (`soan-ho-so-di`) - Chuyên viên văn bản soạn hồ sơ đi và trình ký theo đúng thể thức. · skills: ho-so-kiem-tra-ke-hoach, soan-thao-van-ban
- **Soạn thảo trả lời Công an** (`soan-thao-tra-loi-cong-an`) - đầu ra giấy tờ của phòng - dự thảo công văn phúc đáp và tờ trình đúng thể thức NĐ 30/2020 từ gói bàn giao của Điều phối · skills: tra-loi-cong-van, soan-thao-van-ban
- **Soạn thảo văn bản** (`soan-thao-van-ban`) - Chuẩn hóa thể thức NĐ 30 và soạn văn bản hành chính TCS3; phúc đáp Công an để soan-thao-tra-loi-cong-an · skills: soan-thao-van-ban
- **Thủ hồ sơ và Kiểm kê GĐ1** (`thu-ho-so-kiem-ke`) - Tổ chức cây thư mục 4 bước, checklist A4, manifest IN00-IN14, mức A-E, xuất KQ1. Cấm phân tích rủi ro. · skills: to-chuc-ho-so-kiem-tra-thue, giai-doan-1-kiem-ke-danh-gia
- **Tra cứu và tải công văn iOffice** (`tra-cuu-tai-cv-ioffice`) - Đăng nhập VNPT iOffice Đồng Nai, lọc công văn cảnh báo, tải tệp đính kèm về thư mục lưu trữ · skills: tra-cuu-ioffice-dong-nai
- **Trưởng phòng Phân tích hồ sơ thuế** (`truong-phong-phan-tich-ho-so`) - Điều phối P0-P5, chọn giai đoạn theo mức A-E, chốt Master 28 tiêu chí và đề cương, áp K1-K13. · skills: phan-tich-ho-so-thue, query-wiki

## Skills
### AI
- **Ingest Source** (`ingest-source`) - Tiêu hoá một source thô vào Second Brain, chưng cất thành tri thức wiki tích luỹ.
- **Javis Builder** (`javis-builder`) - Tạo hoặc sửa năng lực của Javis: agent, skill, workflow, loop, plugin. Kèm mẫu file chuẩn và luật chống trùng.
- **Lint Wiki** (`lint-wiki`) - Rà soát sức khoẻ wiki của Second Brain, trả về danh sách vấn đề. Không tự sửa hàng loạt.
- **Notes** (`notes`) - Lưu tin nhắn hiện tại nguyên văn vào sources/ (kèm ảnh), tự chưng cất lên wiki nếu note đáng.
- **Query Wiki** (`query-wiki`) - Khai thác tri thức trong Second Brain: tổng hợp, so sánh, giả thuyết. Trả lời có trích dẫn.
### Marketing
- **HTML sang Webcake** (`html-to-webcake`) - Chuyển trang HTML thành file .pke mở được trong trình dựng Webcake, giữ đúng màu, cỡ chữ, ảnh và bố cục của bản gốc.
### Năng suất
- **Luật Thương mại điện tử 2025** (`luat-thuong-mai-dien-tu`) - Luật Thương mại điện tử 2025: phân loại nền tảng số, hợp đồng điện tử, livestream, tiếp thị liên kết, logistics, thanh toán, CSDL TMĐT quốc gia.
- **PaddleOCR** (`paddleocr`) - OCR ảnh và PDF bằng PaddleOCR qua MCP (tool ocr). Bảng/layout: đổi model PP-StructureV3.
- **PDF sang Markdown** (`pdf-sang-md`) - Chuyển PDF/ảnh scan thành file Markdown (.md) cho Second Brain, giữ heading/bảng, lọc header/footer.
- **Phân tích câu hỏi sinh đôi và Cheat Sheet ôn thi** (`phan-tich-cau-hoi-sinh-doi-cheat-sheet`) - Phân tích câu hỏi sinh đôi, tìm keyword cốt lõi, bóc tách bẫy đáp án, lập Cheat Sheet mapping và tóm tắt siêu tốc < 300 từ dạng Word/Excel ôn thi.
- **Soạn thảo và chuẩn hóa văn bản hành chính** (`soan-thao-van-ban`) - Dự thảo, rà soát thể thức 29 loại VB hành chính theo NĐ 30/2020, viết hoa, chữ ký số và ký hiệu trình ký TCS3.
- **Tiến trình supervisor tự động** (`tien-trinh-supervisor-tu-dong`) - Supervisor tự động điều phối, giám sát chu trình giải đề khảo thí cuốn chiếu, tự phục hồi khi gián đoạn mạng và kiểm kê đến khi đủ mục tiêu.
- **Tra cứu và tải văn bản iOffice Đồng Nai** (`tra-cuu-ioffice-dong-nai`) - Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
- **Trạm tài liệu** (`tram-tai-lieu`) - Chuyển PDF/ảnh scan ra Word, Excel, TXT, PDF tìm kiếm được; xoay/ghép/tách trang; so chữ hai file; theo dõi thư mục.
### Tài chính
- **Bóc tách Bảng tính phân tích tài chính NNT** (`boc-tach-bang-tinh-phan-tich`) - Bóc tách Bảng tính NNT (sheet KQKDSXKD): lấy KQKD theo năm, tính tỷ lệ chi phí/doanh thu, điền phân tích Mục III và đề xuất nội dung kiểm tra DI-01.
- **Điền biên bản từ Bảng tính** (`dien-bien-ban-tu-bang-tinh`) - Điền BBXN/BBKT từ Bảng tính (DTCHEP), đồng bộ QĐ riêng, chèn ngày công bố/báo cáo và dọn bak/backup.
- **Đối chiếu doanh thu 511 và giá vốn 632** (`doi-chieu-511-632`) - Đối chiếu sổ chi tiết TK 511 với TK 632 theo Tool_DC_511_632: khớp HĐ/ngày CT/mặt hàng, lãi gộp và bán dưới giá.
- **Đối chiếu CUP GDLK** (`doi-chieu-cup-gdlk`) - Đối chiếu giao dịch liên kết và độc lập theo phương pháp CUP, kiểm tra giá mua bán và lệch xuất nhập tồn kho.
- **Đối chiếu tờ khai GTGT và TNDN chuẩn CBRR** (`doi-chieu-gtgt-tndn-cbrr`) - Đối chiếu Bảng kê - Tờ khai GTGT - Quyết toán TNDN - HĐĐT chuẩn 2 sheet CBRR, bóc tách nguyên nhân trọng yếu và xuất Excel.
- **Đối chiếu NCC mua vào với CVCBRR và danh bạ** (`doi-chieu-ncc-mua-vao-rui-ro`) - Đối chiếu NCC từ Bảng kê + HĐĐT với CVCBRR F0/F1, Công an, danh bạ; ghi nguồn Bảng kê/HĐĐT/Chung; Bảng 2.3 theo kỳ.
- **Rà soát và đồng bộ thư mục Công văn hoàn tất** (`dong-bo-thu-muc-ca`) - Rà soát kiểm tra file công văn trả lời đã ký số trong toàn bộ thư mục công văn năm 2026 và tự động cập nhật hậu tố -R vào tên thư mục.
- **Giai đoạn 1 - Kiểm kê & Đánh giá hồ sơ thuế** (`giai-doan-1-kiem-ke-danh-gia`) - Kiểm kê manifest IN00-IN13, xác định mức hồ sơ A-E, rà soát khả năng chạy 80 quy tắc, xếp danh mục đề nghị DN bổ sung (KQ1).
- **Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế** (`giai-doan-2-phan-tich-du-lieu-co-quan-thue`) - Phân tích 57 quy tắc thuế khi chưa có sổ DN: chuẩn hóa IN01-IN07/12, đối chiếu GTGT-TNDN-BCTC, 100% NCC rủi ro, phân loại Chắc/Sơ bộ (KQ2).
- **Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN** (`giai-doan-3-hoan-chinh-ho-so`) - Hoàn chỉnh hồ sơ khi có sổ sách DN: 23 quy tắc GĐ3, chuỗi bằng chứng 5 mắt xích, Master 28 tiêu chí 7 cột, đề cương kiểm tra (KQ3).
- **Hồ sơ kiểm tra theo kế hoạch** (`ho-so-kiem-tra-ke-hoach`) - Lập và xuất bộ hồ sơ kiểm tra thuế theo kế hoạch (DI-01 đến DI-07, tổng hợp trình ký) căn cứ QĐ 1412, QĐ 1161, Luật 108/2025.
- **Chính sách và kiểm tra hóa đơn điện tử** (`hoa-don-dien-tu`) - Chính sách và kiểm tra hóa đơn điện tử: NĐ 254/2026, TT 91/2026, Luật 108: HĐĐT máy tính tiền, xử lý sai sót, xăng dầu, rủi ro HĐĐT.
- **Nghiệp vụ kiểm tra thuế chuyên đề và hướng dẫn ngành** (`huong-dan-kiem-tra-chuyen-de`) - Kiểm tra thuế chuyên đề: 2 hệ thống sổ sách, trung gian thanh toán, chuyển nhượng vốn gián tiếp, số định danh cá nhân, QLRR HĐĐT 325, chi phí GTGT.
- **Chính sách và kiểm tra thuế thương mại điện tử** (`kiem-tra-thue-thuong-mai-dien-tu`) - Nghiệp vụ kiểm tra thuế TMĐT: đối chiếu doanh thu sàn, dòng tiền trung gian thanh toán, logistics COD, rủi ro livestream, affiliate, thuế nhà thầu.
- **Làm sạch MST và hóa đơn CD1** (`lam-sach-cd1-mst-hoa-don`) - Chuẩn hóa MST, quét HĐ trùng, cụm cùng ngày 20tr/5tr, đối chiếu BKMV-HĐĐT-NKC 133 theo CD1 BT5-BT9.
- **Lập hồ sơ kiểm tra thuế** (`lap-ho-so-kiem-tra`) - Vận hành app hồ sơ kiểm tra thuế: nhập vụ, import Excel, xuất Word đi, biên bản, phát sinh và về.
- **Luật Phòng chống tham nhũng** (`luat-phong-chong-tham-nhung`) - Cẩm nang pháp lý và thực thi Luật Phòng, chống tham nhũng 36/2018/QH14: quy tắc ứng xử, kiểm soát TSTN, xung đột lợi ích, xử lý vi phạm.
- **Luật Quản lý thuế và Thông tư 89** (`luat-quan-ly-thue`) - Luật Quản lý thuế 108/2025, NĐ 252/2026, TT 89/2026: TPR rủi ro, khai bổ sung 5 năm, hoàn thuế tự động, mẫu biểu KTT, số định danh cá nhân.
- **Mở hồ sơ TTKT** (`mo-ho-so-ttkt`) - Checklist mở hồ sơ TTKT: chọn đối tượng, đối chiếu BCTC-tờ khai-HĐĐT, GDLK, phân nhánh thiếu thuế/trốn thuế.
- **Nhập số liệu Bảng tính từ báo cáo HĐKD ngành thuế** (`nhap-bang-tinh-tu-bao-cao-hdkd`) - Điền sheet KQKDSXKD của Bảng tính từ báo cáo phân tích HĐKD ngành thuế (PTDN02) theo mã chỉ tiêu; năm hoặc ô thiếu thì bù bằng BCTC sau cùng.
- **Phân tích công nợ và tiền mặt** (`phan-tich-cong-no-tien-mat`) - Quét chi tiền mặt TK 111, cụm HĐ cùng ngày, công nợ 131/331 và sao kê ngân hàng 112.
- **Phân tích giá thành và định mức** (`phan-tich-gia-thanh-dinh-muc`) - Bóc giá thành 621/622/627 hoặc cơ cấu GV thương mại, khớp 632 với BCTC [02], giải tỏa TC28.
- **Phân tích hồ sơ kiểm tra thuế** (`phan-tich-ho-so-thue`) - Quy trình 3 giai đoạn phân tích hồ sơ thuế theo bộ biểu mẫu v1.1 (80 quy tắc), kết nối dữ liệu CQT và sổ sách DN, chốt Master 28 tiêu chí (KQ1-3).
- **Phân tích XNT kho** (`phan-tich-xnt-kho`) - Đối chiếu XNT kho 152/155/156 với BCTC HTK, xuất kho với giá vốn 632, giải tỏa TC3.
- **Quản lý rủi ro tuân thủ thuế** (`quan-ly-rui-ro-tuan-thu-thue`) - QLRR tuân thủ thuế: Luật 108/2025, TT 31/2021, phân loại rủi ro, bộ chỉ số TCT, 5 bước xử lý rủi ro.
- **Quy chế và phối hợp kiểm tra thuế** (`quy-che-phoi-hop-kiem-tra-thue`) - Quy chế kiểm tra kiến thức công chức QĐ 58, phối hợp liên ngành tố giác tội phạm QCPH 01 và cấp HĐĐT từng lần sau giải thể.
- **Quy trình và quản lý rủi ro hoàn thuế GTGT** (`quy-trinh-hoan-thue`) - Quy trình hoàn thuế QĐ 679/QĐ-TCT và QLRR hoàn thuế GTGT: phân loại kiểm trước/hoàn trước, bộ CSTC nhóm I-II-III, chuỗi cung ứng và kiểm tra sau hoàn.
- **Quy trình kiểm tra thuế QĐ 1161** (`quy-trinh-kiem-tra-thue`) - Quy trình kiểm tra thuế theo QĐ 1161/QĐ-CT và Luật 108/2025: thời hạn kiểm tra, phân nhóm đối tượng, digital audit và xử lý kết quả.
- **Theo dõi cung cấp hồ sơ Công an** (`theo-doi-cung-cap-ca`) - Bóc tách công văn điều tra, tách file gộp, tra cứu địa chỉ/masothue, phân tổ KT, cán bộ KT1 và cập nhật Bao_cao_cong_van.xlsx sheet NHAP_HO_SO.
- **Theo dõi công văn cảnh báo rủi ro** (`theo-doi-cv-canh-bao-rui-ro`) - Bóc tách công văn cảnh báo thuế/HĐĐT, cập nhật độc lập sheet F0 bên bán và F1 bên mua sổ SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.
- **Chính sách và kiểm tra thuế GTGT** (`thue-gtgt`) - Chính sách và kiểm tra thuế GTGT: Luật 48, Luật 149, NĐ 181, NĐ 359, NĐ 144, TT 69: thuế suất, ngưỡng 500tr, không dùng tiền mặt từ 5tr, khấu trừ…
- **Chính sách và kiểm tra thuế TNCN** (`thue-tncn`) - Chính sách và kiểm tra thuế TNCN: Luật 109/2025, NĐ 253/2026, TT 87/2026: biểu 5 bậc, giảm trừ 15.5tr/6.2tr, khấu trừ 5tr 10%, tài sản số, vàng.
- **Chính sách và kiểm tra thuế TNDN** (`thue-tndn`) - Chính sách và kiểm tra thuế TNDN: Luật 67/2025, NĐ 320/2025, TT 20/2026, TT 45: thuế suất 15-17%, không dùng tiền mặt từ 5tr, chi phí được trừ, ưu…
- **Tổ chức hồ sơ kiểm tra thuế** (`to-chuc-ho-so-kiem-tra-thue`) - Tổ chức thư mục hồ sơ kiểm tra thuế 4 bước, phân loại file, checklist đủ/thiếu A4, gom 1 nơi.
- **Tổng hợp bên mua từ sổ CVCBRR** (`tong-hop-ben-mua-tu-so-cvcbrr`) - Tổng hợp đơn vị bên mua từ sổ CVCBRR: xuất Chi_tiet theo ngày CV và Tong_hop theo MST mua.
- **Tổng hợp đầu vào HĐĐT, bảng kê và NLĐ** (`tong-hop-dau-vao-hddt-bang-ke-nld`) - Gộp HĐĐT portal, bảng kê NNT (lọc Tổng cộng) và NLĐ QTT TNCN thành file mẫu 02, không trộn nguồn.
- **Tổng hợp TNCN và tra cứu đa nguồn** (`tong-hop-tncn-da-nguon`) - Tổng hợp QTT TNCN, tờ khai quý, danh sách NLĐ và tra cứu đa nguồn; đối chiếu TC22/TC23.
- **Tra cứu NNT và phân Tổ Kiểm tra tự động** (`tra-cuu-nnt-tu-dong`) - Tra cứu MST, CQ thuế quản lý; chỉ phân KT1-KT3/CBQL nếu thuộc TCS3 (Vĩnh Cửu=KT2); CQ khác ghi tên vào ô Tổ KT.
- **Trả lời công văn Công an** (`tra-loi-cong-van`) - Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho Công an theo quy chuẩn thể thức và dữ liệu DB NNT.xls
- **Xem hồ sơ thuế điện tử** (`xem-ho-so-thue`) - Mở XML hồ sơ thuế, dựng tờ khai bằng mẫu iTax Viewer, kiểm chữ ký số, xuất HTML.
- **Quản lý rủi ro thuế và hóa đơn điện tử** (`xu-ly-rui-ro-thue-hoa-don`) - Quản lý rủi ro thuế và HĐĐT: QĐ 1264/2026 hệ số K, CV 1125/2026 siết cảnh báo, CV 1873/2022 25 dấu hiệu vi phạm, CV 2383/2021 giám sát HĐ.
- **Xử phạt vi phạm hành chính về thuế và hóa đơn** (`xu-phat-vphc-thue-hoa-don`) - Xử phạt VPHC thuế và hóa đơn: NĐ 125, NĐ 310, NĐ 291, NĐ 68: sai thời điểm, không lập HĐ, trốn thuế, trao đổi thông tin QT, xử lý DN 03-06.
- **Xuất danh sách tổng hợp chi tiết hồ sơ Công an** (`xuat-danh-sach-ho-so-ca`) - Xuất Excel tổng hợp chi tiết hồ sơ công văn Công an theo từng cột tiêu chí tờ khai, BCTC, hóa đơn, vi phạm, nợ thuế tại thư mục CV MOI.

## Workflows
### Tài chính
- **Dây chuyền hồ sơ kiểm tra** (`day-chuyen-ho-so-kiem-tra`) - False · 3 bước [khoi-tao-ho-so -> soan-ho-so-di -> soan-bien-ban-ho-so-ve] · Nối 3 trợ lý văn bản theo vòng đời vụ kiểm tra: khởi tạo dữ liệu công ty, soạn hồ sơ đi trình ký, rồi soạn biên bản và hồ sơ về.
- **Dây chuyền xử lý công văn cảnh báo** (`day-chuyen-xu-ly-cv-canh-bao`) - active · 5 bước [tra-cuu-tai-cv-ioffice -> so-hoa-ocr-cong-van -> cập-nhật-công-văn-cảnh-báo-rủi-ro -> doi-chieu-rui-ro-ncc -> soan-thao-van-ban] · Tải iOffice, OCR, nạp sổ F0/F1, đối chiếu NCC rồi soạn văn bản cảnh báo.
- **Lập hồ sơ kiểm tra** (`lap-ho-so-kiem-tra`) - False · 5 bước [lap-ho-so-kiem-tra -> lap-ho-so-kiem-tra -> lap-ho-so-kiem-tra -> lap-ho-so-kiem-tra -> lap-ho-so-kiem-tra] · Nhập hoặc import một vụ kiểm tra thuế rồi xuất bộ Word hồ sơ đi, biên bản, phát sinh và hồ sơ về.
- **Phân tích hồ sơ thuế 3 giai đoạn** (`phan-tich-ho-so-thue-3-giai-doan`) - active · 12 bước [thu-ho-so-kiem-ke -> truong-phong-phan-tich-ho-so -> chuyen-vien-du-lieu-dau-vao -> chuyen-vien-gtgt-tndn -> chuyen-vien-ncc-hoa-don -> chuyen-vien-tncn -> phan-tich-vien-cqt -> chuyen-vien-doanh-thu-gia-von -> chuyen-vien-kho-gia-thanh -> chuyen-vien-cong-no-tien-mat -> phan-tich-vien-so-sach-dn -> truong-phong-phan-tich-ho-so] · Chuỗi KQ1-KQ2-KQ3 theo mức A-E, bộ mẫu v1.2, 11 trợ lý phòng phân tích.
- **Quản lý công văn công an** (`quản-lý-công-văn-công-an`) - active · 2 bước [quản-lý-hồ-sơ-công-an -> soan-thao-tra-loi-cong-an] · Tiếp nhận-phân tổ-ghi sổ rồi dự thảo công văn và tờ trình cho DN thuộc KT1-KT2-KT3.
### Văn thư
- **Quy trình điều khiển tài liệu** (`quy-trinh-dieu-khien-tai-lieu`) - False · 6 bước [yeu-cau-tai-lieu -> phong-van-thu -> yeu-cau-tai-lieu -> phong-van-thu -> yeu-cau-tai-lieu -> phong-van-thu] · TẮT. Trùng ban hành văn bản đi. Dùng quy-trinh-soan-thao-va-ban-hanh-van-ban-di.
- **Quy trình lưu trữ và số hóa hồ sơ** (`quy-trinh-luu-tru-va-so-hoa-ho-so`) - active · 3 bước [phong-van-thu -> so-hoa-ocr-cong-van -> tri-thuc-second-brain] · Thu hồi, kiểm kê, số hóa OCR và lưu trữ hồ sơ tài liệu vào Second Brain phục vụ tra cứu
- **Quy trình soạn thảo và ban hành văn bản đi** (`quy-trinh-soan-thao-va-ban-hanh-van-ban-di`) - active · 3 bước [phong-van-thu -> soan-thao-van-ban -> phong-van-thu] · Kiểm thể thức NĐ 30, chuẩn hóa Word TCS3, cấp số và phát hành văn bản đi
- **Quy trình tiếp nhận và xử lý công văn đến** (`quy-trinh-tiep-nhan-va-xu-ly-cong-van-den`) - active · 2 bước [so-hoa-ocr-cong-van -> phong-van-thu] · OCR công văn đến, vào sổ Văn thư, phân phòng thụ lý và đặt hạn theo độ khẩn

## Loops
- **Văn thư điều phối và đôn đốc** (`van-thu-dieu-phoi-don-doc`) - bật · custom/full · mỗi 60 phút

## Plugins (tool/hook native cho mọi engine)
- **Thời gian & ngày** (`datetime-vn`) - bundled/chạy · tools: javis_now, javis_date_add · Xem ngày giờ hiện tại theo múi giờ đã cấu hình và tính ngày tương đối (mai, mốt, N ngày nữa, tuần trước). Thuần stdlib, chỉ đọc, không cần mạng.
- **Điều khiển app trên máy** (`desktop-apps`) - bundled/chạy · tools: javis_app_list, javis_app_open, javis_app_close · Mở, đóng và liệt kê ứng dụng trên CHÍNH MÁY đang chạy Javis (Windows, macOS, Linux có màn hình) theo lời người dùng, ví dụ "mở Chrome", "tắt Excel", "đang mở app nào". Tự chặn và nói rõ khi Javis chạy trong Docker trên máy chủ, vì khi đó máy chạy Javis không phải máy của người dùng.
- **Theo dõi Facebook (Apify)** (`fb-monitor-apify`) - bundled/chạy · tools: fb_monitor · Theo dõi Trang/Nhóm CÔNG KHAI Facebook tìm bài nhiều share qua dịch vụ Apify. Chỉ đọc, không đụng tài khoản cá nhân, chạy tốt trên VPS. Dùng token của kết nối "facebook-monitor".
- **Tạo ảnh (ChatGPT)** (`image-chatgpt`) - bundled/chạy · tools: javis_generate_image · Tạo ảnh từ mô tả bằng GÓI ChatGPT đang đăng nhập (OAuth) - không cần OpenAI API key. Dùng Codex Responses API + tool image_generation (gpt-image-2). Ảnh lưu vào attachments/ của vault để nhúng thẳng vào chat. Cần đã kết nối ChatGPT ở trang Model.
- **Đấu thêm MCP** (`javis-connect`) - bundled/chạy · tools: javis_add_mcp · Đấu một MCP server vào kho Kết nối của Javis ngay từ chat, để nó HIỆN ra trang Kết nối như mọi tài khoản khác. Trước tool này Javis chỉ còn đường `claude mcp add` - thứ rơi vào config riêng của Claude Code, không bộ não nào khác thấy và người dùng cũng không thấy trên trang Kết nối.
- **Đặt việc định kỳ & nhắc hẹn** (`javis-schedule`) - bundled/chạy · tools: javis_schedule · Tạo/liệt kê/huỷ việc chạy định kỳ và nhắc hẹn ngay từ chat. Tự chọn kho - việc lặp và bền thì ghi Javis/loops/<slug>.md để sửa được trong Obsidian; nhắc một lần hoặc lịch cron thì vào kho nhắc hẹn (đã có sẵn cron 5 trường). Thay cho việc gõ YAML tay hoặc curl.
- **Giao việc Kanban** (`javis-task`) - bundled/chạy · tools: javis_task · Giao một việc nền vào hàng đợi Kanban và xem việc đang chạy tới đâu, ngay từ chat. Trước tool này chỉ engine chạy được lệnh máy (Claude Code, Codex) mới giao việc được, vì đường duy nhất là curl POST /kanban/task - năm engine API đứng ngoài.
- **Điều khiển dashboard** (`javis-ui`) - bundled/chạy · tools: javis_ui · Mở trang, mở file trong brain, mở một việc Kanban hoặc cuộn khung chat trên dashboard đang mở, theo lời người dùng nói hoặc gõ. Dashboard xác nhận đã làm xong ngay trong lượt; không có tab nào mở thì tool nói rõ.
- **Lịch sử quy trình** (`javis-workflow`) - bundled/chạy · tools: javis_workflow · Xem các lần chạy quy trình gần đây và kết quả từng lần, ngay từ chat. Trước tool này, chạy quy trình xong là kết quả mất, hỏi "quy trình chạy gần nhất ra sao" Javis không có gì để tra.
- **Meta Ads (Graph API)** (`meta-ads-graph`) - bundled/chạy · tools: meta_ads_accounts, meta_ads_insights, meta_ads_campaigns, meta_ads_get · Đọc số liệu quảng cáo Facebook/Instagram (tài khoản ads, chiến dịch, hiệu suất) qua Graph API, dùng token của kết nối "Meta Ads (tự tạo app)". CHỈ ĐỌC, không tiêu tiền.
- **Facebook Trang (Graph API)** (`meta-pages-graph`) - bundled/chạy · tools: fb_pages_list, fb_page_posts, fb_page_comments, fb_page_post, fb_page_photo, fb_page_album, fb_page_video, fb_page_edit, fb_page_delete, fb_page_reply · Quản lý Trang/Fanpage Facebook qua Graph API - liệt kê Trang, đọc bài và bình luận (chỉ đọc), đăng bài/ảnh/video/album, sửa chữ, xoá bài và trả lời bình luận (toàn quyền). Dùng token của kết nối "Facebook Trang (tự tạo app)".
- **Nhật ký dùng tool** (`tool-audit`) - bundled/tắt · tools: javis_tool_stats · hooks: post_tool_call · Đếm số lần MỖI tool được engine gọi (qua hook post_tool_call) và cho xem thống kê tool hay dùng. Đây là ví dụ minh hoạ cơ chế HOOK của plugin. Mặc định TẮT - bật qua POST /plugins/toggle (slug=tool-audit) để thử.
- **Đọc video YouTube** (`youtube-read`) - bundled/chạy · tools: javis_youtube_read · Đọc lời thoại (phụ đề) của video YouTube từ link để tóm tắt. Tự đổi qua 6 kiểu trình phát rồi tới yt-dlp khi YouTube chặn máy chủ. Không cần đăng nhập, không cần API key.
- **Gửi ảnh & file qua Zalo** (`zalo-image`) - bundled/chạy · tools: zalo_send_image · Gửi ẢNH hoặc FILE qua Zalo kèm lời nhắn, bằng chính tài khoản đã quét QR ở trang Kết nối. Bù đúng chỗ MCP chuẩn của zalo-agent-cli còn thiếu - tool zalo_send_message của nó chỉ nhận chữ.

## Cờ sức khoẻ
- Workflow trỏ agent KHÔNG tồn tại: yeu-cau-tai-lieu
- Agent chưa workflow nào dùng: khai-thac-du-lieu-ho-so, phap-che-chinh-sach-thue
