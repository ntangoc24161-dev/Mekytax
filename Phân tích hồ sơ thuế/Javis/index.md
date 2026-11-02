# Javis Index (tầng vận hành)

> Tự sinh từ file - ĐỪNG sửa tay. Chỉ mục mọi năng lực của Javis trong brain này để bất kỳ AI/engine đọc 1 chỗ là hiểu Javis làm được gì. Song song `wiki/index.md` (tri thức).

**Tổng quan:** 11 agents · 33 skills · 1 workflows (1 bật) · 0 loops (0 bật) · 15 plugins (14 chạy)

## Agents
- **Chuyên viên công nợ và tiền mặt** (`chuyen-vien-cong-no-tien-mat`) - Quét 111/131/331/112, ghi KQ2.2!KQ2_Cong_no_Tien_mat. Ngưỡng K5 và P23. · skills: phan-tich-cong-no-tien-mat
- **Chuyên viên doanh thu giá vốn** (`chuyen-vien-doanh-thu-gia-von`) - Đối chiếu 511-632, ghi KQ2.2!KQ2_511_632 (K4). · skills: doi-chieu-511-632
- **Chuyên viên dữ liệu đầu vào** (`chuyen-vien-du-lieu-dau-vao`) - Gộp HĐĐT, bảng kê, NLĐ, CD1 vào IN01 IN02 IN05 của KQ1.1/KQ1.2. · skills: tong-hop-dau-vao-hddt-bang-ke-nld, lam-sach-cd1-mst-hoa-don
- **Chuyên viên GTGT-TNDN** (`chuyen-vien-gtgt-tndn`) - Đối chiếu GTGT TNDN HĐĐT, ghi KQ2.1!KQ2_Doi_chieu_GTGT, chỉ số liệu và cờ máy. · skills: doi-chieu-gtgt-tndn-cbrr
- **Chuyên viên kho và giá thành** (`chuyen-vien-kho-gia-thanh`) - XNT và giá thành, ghi KQ2.2!KQ2_Kho_Gia_thanh và KQ2_Gia_thanh_va_Dinh_muc. · skills: phan-tich-xnt-kho, phan-tich-gia-thanh-dinh-muc
- **Chuyên viên NCC và hóa đơn** (`chuyen-vien-ncc-hoa-don`) - 100% NCC vs F0/F1/CA, ghi IN07 và KQ2_NCC_mua_vao trong KQ2.1. · skills: doi-chieu-ncc-mua-vao-rui-ro, tong-hop-ben-mua-tu-so-cvcbrr
- **Chuyên viên TNCN** (`chuyen-vien-tncn`) - QTT TNCN, NLĐ, đa nguồn; ghi KQ2.1!KQ2_TNCN_LD2. · skills: tong-hop-tncn-da-nguon
- **Phân tích viên dữ liệu CQT** (`phan-tich-vien-cqt`) - Lệnh 2, xuất KQ2.1 và KQ2.2 chỉ số liệu và cờ máy. Không kết luận. · skills: giai-doan-2-phan-tich-du-lieu-co-quan-thue
- **Phân tích viên sổ sách DN** (`phan-tich-vien-so-sach-dn`) - Lệnh 3, xuất KQ3 một sheet 39 cột, 80 quy tắc, chống trùng T1-T10. · skills: giai-doan-3-hoan-chinh-ho-so
- **Thủ hồ sơ và Kiểm kê GĐ1** (`thu-ho-so-kiem-ke`) - Tổ chức cây 4 bước, Lệnh 1, xuất KQ1.1 KQ1.2 KQ1.3. Cấm phân tích rủi ro. · skills: to-chuc-ho-so-kiem-tra-thue, giai-doan-1-kiem-ke-danh-gia
- **Trưởng phòng Phân tích hồ sơ thuế** (`truong-phong-phan-tich-ho-so`) - Điều phối Lệnh 1-4 theo quy trình 1.4, chốt sáu tệp KQ, từ chối mẫu cũ. · skills: phan-tich-ho-so-thue, query-wiki

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
- **PaddleOCR** (`paddleocr`) - OCR ảnh và PDF bằng PaddleOCR qua MCP (tool ocr). Bảng/layout: đổi model PP-StructureV3.
- **PDF sang Markdown** (`pdf-sang-md`) - Chuyển PDF/ảnh scan thành file Markdown (.md) cho Second Brain, giữ heading/bảng, lọc header/footer.
- **Phân tích câu hỏi sinh đôi và Cheat Sheet ôn thi** (`phan-tich-cau-hoi-sinh-doi-cheat-sheet`) - Phân tích câu hỏi sinh đôi, tìm keyword cốt lõi, bóc tách bẫy đáp án, lập Cheat Sheet mapping và tóm tắt siêu tốc < 300 từ dạng Word/Excel ôn thi.
- **Soạn thảo và chuẩn hóa văn bản hành chính** (`soan-thao-van-ban`) - Dự thảo, rà soát thể thức 29 loại VB hành chính theo NĐ 30/2020, viết hoa, chữ ký số và ký hiệu trình ký TCS3.
- **Tiến trình supervisor tự động** (`tien-trinh-supervisor-tu-dong`) - Supervisor tự động điều phối, giám sát chu trình giải đề khảo thí cuốn chiếu, tự phục hồi khi gián đoạn mạng và kiểm kê đến khi đủ mục tiêu.
- **Tra cứu và tải văn bản iOffice Đồng Nai** (`tra-cuu-ioffice-dong-nai`) - Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
- **Trạm tài liệu** (`tram-tai-lieu`) - Chuyển PDF/ảnh scan ra Word, Excel, TXT, PDF tìm kiếm được; xoay/ghép/tách trang; so chữ hai file; theo dõi thư mục.
### Tài chính
- **Đối chiếu doanh thu 511 và giá vốn 632** (`doi-chieu-511-632`) - Đối chiếu sổ chi tiết TK 511 với TK 632 theo Tool_DC_511_632: khớp HĐ/ngày CT/mặt hàng, lãi gộp và bán dưới giá.
- **Đối chiếu tờ khai GTGT và TNDN chuẩn CBRR** (`doi-chieu-gtgt-tndn-cbrr`) - Đối chiếu Bảng kê - Tờ khai GTGT - Quyết toán TNDN - HĐĐT chuẩn 2 sheet CBRR, bóc tách nguyên nhân trọng yếu và xuất Excel.
- **Đối chiếu NCC mua vào với CVCBRR và danh bạ** (`doi-chieu-ncc-mua-vao-rui-ro`) - Đối chiếu NCC từ Bảng kê + HĐĐT với CVCBRR F0/F1, Công an, danh bạ; ghi nguồn Bảng kê/HĐĐT/Chung; Bảng 2.3 theo kỳ.
- **Giai đoạn 1 - Kiểm kê & Đánh giá hồ sơ thuế** (`giai-doan-1-kiem-ke-danh-gia`) - Kiểm kê IN00-IN14 thành KQ1.1 (CQT), KQ1.2 (công ty), KQ1.3 (mức A-E, 80 quy tắc, đề nghị bổ sung).
- **Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế** (`giai-doan-2-phan-tich-du-lieu-co-quan-thue`) - Phân tích số liệu KQ2.1 (CQT+IN07) và KQ2.2 (sổ DN): số, chênh lệch, tỷ lệ, cờ máy; không kết luận.
- **Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN** (`giai-doan-3-hoan-chinh-ho-so`) - KQ3 một sheet 39 cột, 80 quy tắc, chống trùng T1-T10; kết luận, mức, điểm, đề cương, chuỗi bằng chứng.
- **Làm sạch MST và hóa đơn CD1** (`lam-sach-cd1-mst-hoa-don`) - Chuẩn hóa MST, quét HĐ trùng, cụm cùng ngày 20tr/5tr, đối chiếu BKMV-HĐĐT-NKC 133 theo CD1 BT5-BT9.
- **Mở hồ sơ TTKT** (`mo-ho-so-ttkt`) - Checklist mở hồ sơ TTKT: chọn đối tượng, đối chiếu BCTC-tờ khai-HĐĐT, GDLK, phân nhánh thiếu thuế/trốn thuế.
- **Phân tích công nợ và tiền mặt** (`phan-tich-cong-no-tien-mat`) - Quét chi tiền mặt TK 111, cụm HĐ cùng ngày, công nợ 131/331 và sao kê ngân hàng 112.
- **Phân tích giá thành và định mức** (`phan-tich-gia-thanh-dinh-muc`) - Bóc giá thành 621/622/627 hoặc cơ cấu GV thương mại, khớp 632 với BCTC [02], giải tỏa TC28.
- **Phân tích hồ sơ kiểm tra thuế** (`phan-tich-ho-so-thue`) - Quy trình 1.4: 6 tệp KQ1.1-KQ3, 80 quy tắc; sao mẫu rồi điền; KQ2 chỉ số liệu, kết luận chỉ ở KQ3.
- **Phân tích XNT kho** (`phan-tich-xnt-kho`) - Đối chiếu XNT kho 152/155/156 với BCTC HTK, xuất kho với giá vốn 632, giải tỏa TC3.
- **Quản lý rủi ro tuân thủ thuế** (`quan-ly-rui-ro-tuan-thu-thue`) - QLRR tuân thủ thuế: Luật 108/2025, TT 31/2021, phân loại rủi ro, bộ chỉ số TCT, 5 bước xử lý rủi ro.
- **Tổ chức hồ sơ kiểm tra thuế** (`to-chuc-ho-so-kiem-tra-thue`) - Tổ chức thư mục hồ sơ kiểm tra thuế 4 bước, phân loại file, checklist đủ/thiếu A4, gom 1 nơi.
- **Tổng hợp bên mua từ sổ CVCBRR** (`tong-hop-ben-mua-tu-so-cvcbrr`) - Tổng hợp đơn vị bên mua từ sổ CVCBRR: xuất Chi_tiet theo ngày CV và Tong_hop theo MST mua.
- **Tổng hợp đầu vào HĐĐT, bảng kê và NLĐ** (`tong-hop-dau-vao-hddt-bang-ke-nld`) - Gộp HĐĐT portal, bảng kê NNT (lọc Tổng cộng) và NLĐ QTT TNCN thành file mẫu 02, không trộn nguồn.
- **Tổng hợp TNCN và tra cứu đa nguồn** (`tong-hop-tncn-da-nguon`) - Tổng hợp QTT TNCN, tờ khai quý, danh sách NLĐ và tra cứu đa nguồn; đối chiếu TC22/TC23.
- **Trả lời công văn Công an** (`tra-loi-cong-van`) - Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho cơ quan Công an theo quy chuẩn thể thức và dữ liệu quản lý thuế.
- **Xử phạt vi phạm hành chính về thuế và hóa đơn** (`xu-phat-vphc-thue-hoa-don`) - Xử phạt VPHC thuế và hóa đơn: NĐ 125, NĐ 310, NĐ 291, NĐ 68: sai thời điểm, không lập HĐ, trốn thuế, trao đổi thông tin QT, xử lý DN 03-06.
- **Xem hồ sơ thuế điện tử** (`xem-ho-so-thue`) - Mở XML hồ sơ thuế, dựng tờ khai bằng mẫu iTax Viewer, kiểm chữ ký số, xuất HTML.

## Workflows
- **Phân tích hồ sơ thuế 3 giai đoạn** (`phan-tich-ho-so-thue-3-giai-doan`) - active · 12 bước [thu-ho-so-kiem-ke -> truong-phong-phan-tich-ho-so -> chuyen-vien-du-lieu-dau-vao -> chuyen-vien-gtgt-tndn -> chuyen-vien-ncc-hoa-don -> chuyen-vien-tncn -> phan-tich-vien-cqt -> chuyen-vien-doanh-thu-gia-von -> chuyen-vien-kho-gia-thanh -> chuyen-vien-cong-no-tien-mat -> phan-tich-vien-so-sach-dn -> truong-phong-phan-tich-ho-so] · Chuỗi Lệnh 1-3 quy trình 1.4: KQ1.1/1.2/1.3 rồi KQ2.1/2.2 rồi KQ3; module ghi sheet đích.

## Loops
_(chưa có)_

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
- **Trạm tài liệu** (`tram-tai-lieu`) - bundled/chạy · tools: javis_tram_tai_lieu · OCR và chuyển đổi PDF/ảnh scan sang Word, Excel, Markdown, TXT hoặc PDF tìm kiếm được; xoay, tách, ghép, xóa trang; so sánh nội dung 2 file. Chạy local bằng Python venv .venv-paddleocr.
- **Đọc video YouTube** (`youtube-read`) - bundled/chạy · tools: javis_youtube_read · Đọc lời thoại (phụ đề) của video YouTube từ link để tóm tắt. Tự đổi qua 6 kiểu trình phát rồi tới yt-dlp khi YouTube chặn máy chủ. Không cần đăng nhập, không cần API key.
- **Gửi ảnh & file qua Zalo** (`zalo-image`) - bundled/chạy · tools: zalo_send_image · Gửi ẢNH hoặc FILE qua Zalo kèm lời nhắn, bằng chính tài khoản đã quét QR ở trang Kết nối. Bù đúng chỗ MCP chuẩn của zalo-agent-cli còn thiếu - tool zalo_send_message của nó chỉ nhận chữ.
