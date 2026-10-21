# Javis Index (tầng vận hành)

> Tự sinh từ file - ĐỪNG sửa tay. Chỉ mục mọi năng lực của Javis trong brain này để bất kỳ AI/engine đọc 1 chỗ là hiểu Javis làm được gì. Song song `wiki/index.md` (tri thức).

**Tổng quan:** 1 agents · 30 skills · 1 workflows (1 bật) · 0 loops (0 bật) · 14 plugins (13 chạy)

## Agents
- **Quản lý hồ sơ công an** (`quản-lý-hồ-sơ-công-an`) - công chức thuế được giao nhiệm vụ cung cấp hồ sơ theo yêu cầu của cơ quan công an, thuế thành phố và cơ quan khác · skills: tra-cuu-ioffice-dong-nai, theo-doi-cung-cap-ca, tra-cuu-nnt-tu-dong, tra-loi-cong-van, xuat-danh-sach-ho-so-ca, dong-bo-thu-muc-ca, pdf-sang-md

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
- **PDF sang Markdown** (`pdf-sang-md`) - Chuyển PDF/ảnh scan thành file Markdown (.md) cho Second Brain, giữ heading/bảng, lọc header/footer.
- **Phân tích câu hỏi sinh đôi và Cheat Sheet ôn thi** (`phan-tich-cau-hoi-sinh-doi-cheat-sheet`) - Phân tích câu hỏi sinh đôi, tìm keyword cốt lõi, bóc tách bẫy đáp án, lập Cheat Sheet mapping và tóm tắt siêu tốc < 300 từ dạng Word/Excel ôn thi.
- **Tiến trình supervisor tự động** (`tien-trinh-supervisor-tu-dong`) - Supervisor tự động điều phối, giám sát chu trình giải đề khảo thí cuốn chiếu, tự phục hồi khi gián đoạn mạng và kiểm kê đến khi đủ mục tiêu.
- **Tra cứu và tải văn bản iOffice Đồng Nai** (`tra-cuu-ioffice-dong-nai`) - Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
### Tài chính
- **Rà soát và đồng bộ thư mục Công văn hoàn tất** (`dong-bo-thu-muc-ca`) - Rà soát kiểm tra file công văn trả lời đã ký số trong toàn bộ thư mục công văn năm 2026 và tự động cập nhật hậu tố -R vào tên thư mục.
- **Chính sách và kiểm tra hóa đơn điện tử** (`hoa-don-dien-tu`) - Chính sách và kiểm tra hóa đơn điện tử: NĐ 254/2026, TT 91/2026, Luật 108: HĐĐT máy tính tiền, xử lý sai sót, xăng dầu, rủi ro HĐĐT.
- **Nghiệp vụ kiểm tra thuế chuyên đề và hướng dẫn ngành** (`huong-dan-kiem-tra-chuyen-de`) - Kiểm tra thuế chuyên đề: 2 hệ thống sổ sách, trung gian thanh toán, chuyển nhượng vốn gián tiếp, số định danh cá nhân, QLRR HĐĐT 325, chi phí GTGT.
- **Chính sách và kiểm tra thuế thương mại điện tử** (`kiem-tra-thue-thuong-mai-dien-tu`) - Nghiệp vụ kiểm tra thuế TMĐT: đối chiếu doanh thu sàn, dòng tiền trung gian thanh toán, logistics COD, rủi ro livestream, affiliate, thuế nhà thầu.
- **Luật Phòng chống tham nhũng** (`luat-phong-chong-tham-nhung`) - Cẩm nang pháp lý và thực thi Luật Phòng, chống tham nhũng 36/2018/QH14: quy tắc ứng xử, kiểm soát TSTN, xung đột lợi ích, xử lý vi phạm.
- **Luật Quản lý thuế và Thông tư 89** (`luat-quan-ly-thue`) - Luật Quản lý thuế 108/2025, NĐ 252/2026, TT 89/2026: TPR rủi ro, khai bổ sung 5 năm, hoàn thuế tự động, mẫu biểu KTT, số định danh cá nhân.
- **Phân tích hồ sơ kiểm tra thuế** (`phan-tich-ho-so-thue`) - Phân tích hồ sơ kiểm tra thuế từ XML tờ khai/BCTC và Excel hóa đơn, đối chiếu HĐ-tờ khai, 28 tiêu chí rủi ro.
- **Quy chế và phối hợp kiểm tra thuế** (`quy-che-phoi-hop-kiem-tra-thue`) - Quy chế kiểm tra kiến thức công chức QĐ 58, phối hợp liên ngành tố giác tội phạm QCPH 01 và cấp HĐĐT từng lần sau giải thể.
- **Quy trình và quản lý rủi ro hoàn thuế GTGT** (`quy-trinh-hoan-thue`) - Quy trình hoàn thuế QĐ 679/QĐ-TCT và QLRR hoàn thuế GTGT: phân loại kiểm trước/hoàn trước, bộ CSTC nhóm I-II-III, chuỗi cung ứng và kiểm tra sau hoàn.
- **Quy trình kiểm tra thuế QĐ 1161** (`quy-trinh-kiem-tra-thue`) - Quy trình kiểm tra thuế theo QĐ 1161/QĐ-CT và Luật 108/2025: thời hạn kiểm tra, phân nhóm đối tượng, digital audit và xử lý kết quả.
- **Theo dõi cung cấp hồ sơ Công an** (`theo-doi-cung-cap-ca`) - Bóc tách công văn điều tra, tách file gộp, tra cứu CSDL NNT/cổng thông tin, phân tổ KT, cán bộ KT1 và cập nhật Bao_cao_cong_van.xlsx.
- **Chính sách và kiểm tra thuế GTGT** (`thue-gtgt`) - Chính sách và kiểm tra thuế GTGT: Luật 48, Luật 149, NĐ 181, NĐ 359, NĐ 144, TT 69: thuế suất, ngưỡng 500tr, không dùng tiền mặt từ 5tr, khấu trừ…
- **Chính sách và kiểm tra thuế TNCN** (`thue-tncn`) - Chính sách và kiểm tra thuế TNCN: Luật 109/2025, NĐ 253/2026, TT 87/2026: biểu 5 bậc, giảm trừ 15.5tr/6.2tr, khấu trừ 5tr 10%, tài sản số, vàng.
- **Chính sách và kiểm tra thuế TNDN** (`thue-tndn`) - Chính sách và kiểm tra thuế TNDN: Luật 67/2025, NĐ 320/2025, TT 20/2026, TT 45: thuế suất 15-17%, không dùng tiền mặt từ 5tr, chi phí được trừ, ưu…
- **Tra cứu NNT và phân Tổ Kiểm tra tự động** (`tra-cuu-nnt-tu-dong`) - Tra cứu MST qua DB NNT.xls và cổng thông tin, chuẩn hóa địa chỉ, phân Tổ KT1-KT2-KT3 và xác định cán bộ quản lý KT1 tự động.
- **Trả lời công văn Công an** (`tra-loi-cong-van`) - Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho Công an theo quy chuẩn thể thức và dữ liệu DB NNT.xls
- **Quản lý rủi ro thuế và hóa đơn điện tử** (`xu-ly-rui-ro-thue-hoa-don`) - Quản lý rủi ro thuế và HĐĐT: QĐ 1264/2026 hệ số K, CV 1125/2026 siết cảnh báo, CV 1873/2022 25 dấu hiệu vi phạm, CV 2383/2021 giám sát HĐ.
- **Xử phạt vi phạm hành chính về thuế và hóa đơn** (`xu-phat-vphc-thue-hoa-don`) - Xử phạt VPHC thuế và hóa đơn: NĐ 125, NĐ 310, NĐ 291, NĐ 68: sai thời điểm, không lập HĐ, trốn thuế, trao đổi thông tin QT, xử lý DN 03-06.
- **Xuất danh sách tổng hợp chi tiết hồ sơ Công an** (`xuat-danh-sach-ho-so-ca`) - Xuất Excel tổng hợp chi tiết hồ sơ công văn Công an theo từng cột tiêu chí tờ khai, BCTC, hóa đơn, vi phạm, nợ thuế tại thư mục CV MOI.

## Workflows
- **Quản lý công văn công an** (`quản-lý-công-văn-công-an`) - active · 1 bước [quản-lý-hồ-sơ-công-an] · Tiếp nhận, bóc tách công văn công an, tra cứu NNT, cập nhật sổ theo dõi, lập dự thảo trả lời Tổ KT1, xuất Excel chi tiết và đồng bộ thư mục.

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
- **Đọc video YouTube** (`youtube-read`) - bundled/chạy · tools: javis_youtube_read · Đọc lời thoại (phụ đề) của video YouTube từ link để tóm tắt. Tự đổi qua 6 kiểu trình phát rồi tới yt-dlp khi YouTube chặn máy chủ. Không cần đăng nhập, không cần API key.
- **Gửi ảnh & file qua Zalo** (`zalo-image`) - bundled/chạy · tools: zalo_send_image · Gửi ẢNH hoặc FILE qua Zalo kèm lời nhắn, bằng chính tài khoản đã quét QR ở trang Kết nối. Bù đúng chỗ MCP chuẩn của zalo-agent-cli còn thiếu - tool zalo_send_message của nó chỉ nhận chữ.
