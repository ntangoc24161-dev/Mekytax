# Javis Index (tầng vận hành)

> Tự sinh từ file - ĐỪNG sửa tay. Chỉ mục mọi năng lực của Javis trong brain này để bất kỳ AI/engine đọc 1 chỗ là hiểu Javis làm được gì. Song song `wiki/index.md` (tri thức).

**Tổng quan:** 9 agents · 40 skills · 4 workflows (2 bật) · 2 loops (2 bật) · 28 plugins (27 chạy)

## Agents
- **Trưởng phòng cung cấp hồ sơ công an** (`javis-core.conflict-remote-20261111-133623`) - Trợ lý hệ thống nhận báo cáo và hỗ trợ công việc trong brain này.
- **Javis Core** (`javis-core`) - Trợ lý hệ thống nhận báo cáo và hỗ trợ công việc trong brain này.
- **khai-thac-du-lieu-ho-so** (`khai-thac-du-lieu-ho-so`) - 
- **Kiểm soát & Điều phối quy trình** (`kiem-soat-dieu-phoi-quy-trinh`) - quản lý chất lượng, kiểm soát chéo dữ liệu bóc tách, nghiệm thu từng bước và phân phối nhiệm vụ cho các trợ lý · model sonnet · skills: theo-doi-cung-cap-ca, tra-cuu-nnt-tu-dong, tra-loi-cong-van, in-phan-cong-kt1, soan-thao-van-ban, dong-bo-thu-muc-ca, xuat-danh-sach-ho-so-ca · tags: Tài chính
- **phap-che-chinh-sach-thue** (`phap-che-chinh-sach-thue`) - 
- **quản-lý-hồ-sơ-công-an** (`quản-lý-hồ-sơ-công-an`) - 
- **soan-thao-tra-loi-cong-an** (`soan-thao-tra-loi-cong-an`) - 
- **Theo dõi CV trả lời đã ký** (`theo-doi-cv-tra-loi-da-ky`) - công chức thuế chuyên trách theo dõi & bóc tách công văn trả lời Công an đã ký số, cập nhật trạng thái và tào tác danh sách tổng hợp (Bước 5 - Giám sát chất lượng) · model sonnet · skills: dong-bo-thu-muc-ca, xuat-danh-sach-ho-so-ca, soan-thao-van-ban · tags: Tài chính, giám-sát
- **Theo dõi & Phân công công văn** (`theo-doi-phan-cong`) - theo dõi tiến độ, kiểm soát công văn mới, rà soát cán bộ xử lý và in ấn bàn giao hồ sơ Tổ KT1 · model sonnet · skills: in-phan-cong-kt1, theo-doi-cung-cap-ca, tra-cuu-nnt-tu-dong · tags: Tài chính

## Skills
### AI
- **Ingest Source** (`ingest-source`) - Tiêu hoá một source thô vào Second Brain, chưng cất thành tri thức wiki tích luỹ.
- **Javis Builder** (`javis-builder`) - Tạo hoặc sửa năng lực của Javis: agent, skill, workflow, loop, plugin. Kèm mẫu file chuẩn và luật chống trùng.
- **Lint Wiki** (`lint-wiki`) - Rà soát sức khoẻ wiki của Second Brain, trả về danh sách vấn đề. Không tự sửa hàng loạt.
- **Notes** (`notes`) - Lưu tin nhắn hiện tại nguyên văn vào sources/ (kèm ảnh), tự chưng cất lên wiki nếu note đáng.
- **Query Wiki** (`query-wiki`) - Khai thác tri thức trong Second Brain: tổng hợp, so sánh, giả thuyết. Trả lời có trích dẫn.
### Marketing
- **HTML sang Webcake** (`html-to-webcake`) · [TẮT] - Chuyển trang HTML thành file .pke mở được trong trình dựng Webcake, giữ đúng màu, cỡ chữ, ảnh và bố cục của bản gốc.
### Năng suất
- **in-phan-cong-kt1** (`in-phan-cong-kt1`) - In đồng loạt công văn Công an giao Tổ KT1 trạng thái Mới, mỗi cán bộ một PDF chẵn trang, có dòng phân công đầu trang, đổi sang Đang xử lý.
- **Luật Thương mại điện tử 2025** (`luat-thuong-mai-dien-tu`) - Luật Thương mại điện tử 2025: phân loại nền tảng số, hợp đồng điện tử, livestream, tiếp thị liên kết, logistics, thanh toán, CSDL TMĐT quốc gia.
- **PaddleOCR** (`paddleocr`) - OCR ảnh và PDF bằng PaddleOCR qua MCP (tool ocr). Bảng/layout: đổi model PP-StructureV3.
- **PDF sang Markdown** (`pdf-sang-md`) - Chuyển PDF/ảnh scan thành file Markdown (.md) cho Second Brain, giữ heading/bảng, lọc header/footer.
- **Soạn thảo và chuẩn hóa văn bản hành chính** (`soan-thao-van-ban`) - Dự thảo, rà soát thể thức 29 loại VB hành chính theo NĐ 30/2020, viết hoa, chữ ký số và ký hiệu trình ký TCS3.
- **Tra cứu và tải văn bản iOffice Đồng Nai** (`tra-cuu-ioffice-dong-nai`) - Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
- **Trạm tài liệu** (`tram-tai-lieu`) - Chuyển PDF/ảnh scan ra Word, Excel, TXT, PDF tìm kiếm được; xoay/ghép/tách trang; so chữ hai file; theo dõi thư mục.
- **Phân tích câu hỏi sinh đôi và Cheat Sheet ôn thi** (`phan-tich-cau-hoi-sinh-doi-cheat-sheet`) · [TẮT] - Phân tích câu hỏi sinh đôi, tìm keyword cốt lõi, bóc tách bẫy đáp án, lập Cheat Sheet mapping và tóm tắt siêu tốc < 300 từ dạng Word/Excel ôn thi.
- **Tiến trình supervisor tự động** (`tien-trinh-supervisor-tu-dong`) · [TẮT] - Supervisor tự động điều phối, giám sát chu trình giải đề khảo thí cuốn chiếu, tự phục hồi khi gián đoạn mạng và kiểm kê đến khi đủ mục tiêu.
### Security
- **Security Audit** (`security-audit`) - Rà soát bảo mật mã nguồn, API, dịch vụ: tìm lỗ hổng có bằng chứng, kiểm chứng độc lập từng lỗi, báo cáo kèm cách sửa. Từ Cloudflare.
### Tài chính
- **Đối chiếu NCC mua vào với CVCBRR và danh bạ** (`doi-chieu-ncc-mua-vao-rui-ro`) - Đối chiếu NCC từ Bảng kê + HĐĐT với CVCBRR F0/F1, Công an, danh bạ; ghi nguồn Bảng kê/HĐĐT/Chung; Bảng 2.3 theo kỳ.
- **Rà soát và đồng bộ thư mục Công văn hoàn tất** (`dong-bo-thu-muc-ca`) - Rà soát kiểm tra file công văn trả lời đã ký số trong toàn bộ thư mục công văn năm 2026 và tự động cập nhật hậu tố -R vào tên thư mục.
- **Giai đoạn 1 - Kiểm kê & Đánh giá hồ sơ thuế** (`giai-doan-1-kiem-ke-danh-gia`) - Kiểm kê manifest IN00-IN13, xác định mức hồ sơ A-E, rà soát khả năng chạy 80 quy tắc, xếp danh mục đề nghị DN bổ sung (KQ1).
- **Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế** (`giai-doan-2-phan-tich-du-lieu-co-quan-thue`) - Phân tích 57 quy tắc thuế khi chưa có sổ DN: chuẩn hóa IN01-IN07/12, đối chiếu GTGT-TNDN-BCTC, 100% NCC rủi ro, phân loại Chắc/Sơ bộ (KQ2).
- **Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN** (`giai-doan-3-hoan-chinh-ho-so`) - Hoàn chỉnh hồ sơ khi có sổ sách DN: 23 quy tắc GĐ3, chuỗi bằng chứng 5 mắt xích, Master 28 tiêu chí 7 cột, đề cương kiểm tra (KQ3).
- **Chính sách và kiểm tra hóa đơn điện tử** (`hoa-don-dien-tu`) - Chính sách và kiểm tra hóa đơn điện tử: NĐ 254/2026, TT 91/2026, Luật 108: HĐĐT máy tính tiền, xử lý sai sót, xăng dầu, rủi ro HĐĐT.
- **Nghiệp vụ kiểm tra thuế chuyên đề và hướng dẫn ngành** (`huong-dan-kiem-tra-chuyen-de`) - Kiểm tra thuế chuyên đề: 2 hệ thống sổ sách, trung gian thanh toán, chuyển nhượng vốn gián tiếp, số định danh cá nhân, QLRR HĐĐT 325, chi phí GTGT.
- **Chính sách và kiểm tra thuế thương mại điện tử** (`kiem-tra-thue-thuong-mai-dien-tu`) - Nghiệp vụ kiểm tra thuế TMĐT: đối chiếu doanh thu sàn, dòng tiền trung gian thanh toán, logistics COD, rủi ro livestream, affiliate, thuế nhà thầu.
- **Luật Phòng chống tham nhũng** (`luat-phong-chong-tham-nhung`) - Cẩm nang pháp lý và thực thi Luật Phòng, chống tham nhũng 36/2018/QH14: quy tắc ứng xử, kiểm soát TSTN, xung đột lợi ích, xử lý vi phạm.
- **Luật Quản lý thuế và Thông tư 89** (`luat-quan-ly-thue`) - Luật Quản lý thuế 108/2025, NĐ 252/2026, TT 89/2026: TPR rủi ro, khai bổ sung 5 năm, hoàn thuế tự động, mẫu biểu KTT, số định danh cá nhân.
- **Phân tích công nợ và tiền mặt** (`phan-tich-cong-no-tien-mat`) - Quét chi tiền mặt TK 111, cụm HĐ cùng ngày, công nợ 131/331 và sao kê ngân hàng 112.
- **Phân tích hồ sơ kiểm tra thuế** (`phan-tich-ho-so-thue`) - Quy trình 3 giai đoạn phân tích hồ sơ thuế theo bộ biểu mẫu v1.1 (80 quy tắc), kết nối dữ liệu CQT và sổ sách DN, chốt Master 28 tiêu chí (KQ1-3).
- **Quy chế và phối hợp kiểm tra thuế** (`quy-che-phoi-hop-kiem-tra-thue`) - Quy chế kiểm tra kiến thức công chức QĐ 58, phối hợp liên ngành tố giác tội phạm QCPH 01 và cấp HĐĐT từng lần sau giải thể.
- **Quy trình và quản lý rủi ro hoàn thuế GTGT** (`quy-trinh-hoan-thue`) - Quy trình hoàn thuế QĐ 679/QĐ-TCT và QLRR hoàn thuế GTGT: phân loại kiểm trước/hoàn trước, bộ CSTC nhóm I-II-III, chuỗi cung ứng và kiểm tra sau hoàn.
- **Quy trình kiểm tra thuế QĐ 1161** (`quy-trinh-kiem-tra-thue`) - Quy trình kiểm tra thuế theo QĐ 1161/QĐ-CT và Luật 108/2025: thời hạn kiểm tra, phân nhóm đối tượng, digital audit và xử lý kết quả.
- **Theo dõi cung cấp hồ sơ Công an** (`theo-doi-cung-cap-ca`) - Bóc tách CV điều tra, tra cứu CQT quản lý, phân tổ KT (Vĩnh Cửu vào KT2; CQT khác ghi tên vào Tổ KT), gán CBQL KT1 và cập nhật Bao_cao_cong_van.xlsx.
- **Chính sách và kiểm tra thuế GTGT** (`thue-gtgt`) - Chính sách và kiểm tra thuế GTGT: Luật 48, Luật 149, NĐ 181, NĐ 359, NĐ 144, TT 69: thuế suất, ngưỡng 500tr, không dùng tiền mặt từ 5tr, khấu trừ…
- **Chính sách và kiểm tra thuế TNCN** (`thue-tncn`) - Chính sách và kiểm tra thuế TNCN: Luật 109/2025, NĐ 253/2026, TT 87/2026: biểu 5 bậc, giảm trừ 15.5tr/6.2tr, khấu trừ 5tr 10%, tài sản số, vàng.
- **Chính sách và kiểm tra thuế TNDN** (`thue-tndn`) - Chính sách và kiểm tra thuế TNDN: Luật 67/2025, NĐ 320/2025, TT 20/2026, TT 45: thuế suất 15-17%, không dùng tiền mặt từ 5tr, chi phí được trừ, ưu…
- **Tra cứu NNT và phân Tổ Kiểm tra tự động** (`tra-cuu-nnt-tu-dong`) - Tra cứu MST qua DB NNT và cổng thông tin, đọc cơ quan quản lý, phân Tổ KT1-KT3; MST ngoài TCS3 thì không gán CBQL.
- **Trả lời công văn Công an** (`tra-loi-cong-van`) - Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho Công an theo quy chuẩn thể thức và dữ liệu DB NNT.xls
- **Quản lý rủi ro thuế và hóa đơn điện tử** (`xu-ly-rui-ro-thue-hoa-don`) - Quản lý rủi ro thuế và HĐĐT: QĐ 1264/2026 hệ số K, CV 1125/2026 siết cảnh báo, CV 1873/2022 25 dấu hiệu vi phạm, CV 2383/2021 giám sát HĐ.
- **Xử phạt vi phạm hành chính về thuế và hóa đơn** (`xu-phat-vphc-thue-hoa-don`) - Xử phạt VPHC thuế và hóa đơn: NĐ 125, NĐ 310, NĐ 291, NĐ 68: sai thời điểm, không lập HĐ, trốn thuế, trao đổi thông tin QT, xử lý DN 03-06.
- **Xuất danh sách tổng hợp chi tiết hồ sơ Công an** (`xuat-danh-sach-ho-so-ca`) - Xuất Excel tổng hợp chi tiết hồ sơ công văn Công an theo từng cột tiêu chí tờ khai, BCTC, hóa đơn, vi phạm, nợ thuế tại thư mục CV MOI.

## Workflows
### Tài chính
- **Quản lý công văn công an** (`quản-lý-công-văn-công-an`) - active · 4 bước [khai-thac-du-lieu-ho-so -> theo-doi-phan-cong -> soan-thao-tra-loi-cong-an -> quản-lý-hồ-sơ-công-an] · Tiếp nhận, bóc tách công văn, kiểm soát chéo dữ liệu, phân công in ấn KT1, lập dự thảo phúc đáp và nghiệm thu đồng bộ thư mục.
- **Theo dõi & cập nhật CV trả lời đã ký** (`theo-doi-cv-tra-loi-da-ky`) - active · 1 bước [theo-doi-cv-tra-loi-da-ky] · Quét thư mục công văn năm 2026, phát hiện CV trả lời đã ký số, bóc tách dữ liệu, xác nhận danh sách DN, cập nhật trạng thái sổ theo dõi.
### Vận hành
- **Bóc tách & Phân công công văn Công an** (`boc-tach-phan-cong-cong-van`) - off · 2 bước [khai-thac-du-lieu-ho-so -> Theo dõi & Phân công công văn] · Bóc tách danh sách doanh nghiệp từ CV Công an, tra cứu DB NNT, phân công cho Tổ KT tương ứng.
- **Soạn thảo & Quản lý hồ sơ công an** (`soan-thao-quan-ly-ho-so`) - off · 2 bước [soan-thao-tra-loi-cong-an -> quản-lý-hồ-sơ-công-an] · Soạn dự thảo công văn trả lời và tờ trình lãnh đạo theo yêu cầu, lưu file và cập nhật sổ theo dõi.

## Loops
- **Hồ sơ cung cấp công an** (`ho-so-cung-cap-cong-an.conflict-remote-20261111-133623`) - bật · custom/full · mỗi 1440 phút
- **Hồ sơ cung cấp công an** (`ho-so-cung-cap-cong-an`) - bật · custom/full · mỗi 720 phút

## Plugins (tool/hook native cho mọi engine)
- **Date & time** (`datetime-vn`) - bundled/chạy · tools: javis_now, javis_date_add · Shows the current date and time in the timezone picked in Settings and works out relative dates (tomorrow, in N days, last week). Pure stdlib, read-only, no network needed.
- **Control apps on this computer** (`desktop-apps`) - bundled/chạy · tools: javis_app_list, javis_app_open, javis_app_close · Opens, closes and lists apps on THE MACHINE running Javis (Windows, macOS, Linux with a screen) when you ask, e.g. "open Chrome", "close Excel", "which apps are open". Blocks itself and says so when Javis runs in Docker on a server, because then the machine running Javis is not yours.
- **Facebook monitor (Apify)** (`fb-monitor-apify`) - bundled/chạy · tools: fb_monitor · Watches PUBLIC Facebook Pages and Groups for highly shared posts through the Apify service. Read-only, never touches a personal account, runs fine on a VPS. Uses the token of the "facebook-monitor" connection.
- **Create and view images (ChatGPT)** (`image-chatgpt`) - bundled/chạy · tools: javis_generate_image, javis_describe_image · Creates images from a description on the ChatGPT PLAN you are signed in to (OAuth), no OpenAI API key needed. Uses the Codex Responses API with the image_generation tool (gpt-image-2). Images are saved to the vault's attachments/ so they embed straight into chat. Also a tool to VIEW images: ChatGPT looks at an image in the brain and describes it in words, so even engines that cannot see images (OpenRouter, Gemini API...) know what is in it. Needs ChatGPT connected on the Models page.
- **Connect an MCP server** (`javis-connect`) - bundled/chạy · tools: javis_add_mcp · Adds an MCP server to Javis's connection store right from chat, so it SHOWS on the Connections page like any other account. Before this tool the only route was `claude mcp add`, which lands in Claude Code's own config: no other brain sees it and neither do you on the Connections page.
- **Javis Core's name and style** (`javis-core-name`) - bundled/chạy · tools: javis_core_name · Lets the person give this brain's Javis Core its own name and way of talking, right in chat. The name shows on the Partners page. Only the owner's own sessions see this tool.
- **GitHub in chat** (`javis-github`) - bundled/chạy · tools: javis_github · Shows the chat's GitHub repos (branch, PR, CI, merge), queues the next job, which Javis starts on its own once the current PR is merged, and cross-reviews a PR when the user asks.
- **Goals (Resonance)** (`javis-goal`) - bundled/chạy · tools: javis_goal, javis_submit_deliverable · Lets the brain set up a lasting goal when the user hands over work that needs following up after the chat turn, or update an open goal when the user adds to it. The host checks the proposal against SMART and stores it with revisions. Only shown in brains with Resonance switched on.
- **Background jobs** (`javis-job`) - bundled/chạy · tools: javis_job · Tracks background commands an engine leaves running (render, build, download). Say in advance "when it's done, do this next", see running jobs, stop one. When it finishes Javis reports back to the right chat and carries on by itself.
- **Cookies** (`javis-login`) - bundled/chạy · tools: javis_login · Opens a web page that needs a sign-in, using a sign-in the owner stored on the Cookies page. Returns the page text only, never the cookies.
- **Memory from chat** (`javis-memory`) - bundled/chạy · tools: javis_memory · Lets Javis list, edit and forget the brain's memories, switch new memories on or off, edit the shared profile and revoke an assistant's lessons when the person asks in chat. Replaces the memory and lesson lists that used to sit in the settings column. Only the owner's own sessions see this tool.
- **My tasks and habits** (`javis-mine`) - bundled/chạy · tools: javis_mytask · Adds, edits and deletes tasks, appointments and habits in the Mine tab of the Tasks page (turn a repeat into an appointment, change times and repeats, skip one time), lists them, and ticks them done when the user says so, right from chat.
- **Notification commander** (`javis-notify`) - bundled/chạy · tools: javis_notify · See and change which notices ring. Each notice group has a mode (everything, important and up, urgent only, off); filter or always ring one kind of notice; undo a change Javis Core made.
- **Plugins from chat** (`javis-plugin`) - bundled/chạy · tools: javis_plugin · Lets Javis create, read and edit your own plugins in place when you ask in chat, instead of handing you a zip to download and upload again. A new plugin, or one whose code just changed, always stays OFF until you read it and turn it on in the Plugins page. Only the owner's own sessions see this tool.
- **Group bot reply judge** (`javis-reply-policy`) - bundled/chạy · tools: javis_reply_policy, javis_reply_policy_tune · Lets Javis read the reply judge's numbers (when a group bot spoke or stayed silent, why, where it was marked wrong) and tune it when the owner asks, with undo. Only the owner's own sessions see these two tools; customer-facing bots never do.
- **Recurring jobs & reminders** (`javis-schedule`) - bundled/chạy · tools: javis_schedule · Create, list and cancel recurring jobs and reminders right from chat. Picks the right store itself: lasting repeat work goes to Javis/loops/<slug>.md so you can edit it in Obsidian; one-off reminders or cron schedules go to the reminder store (5-field cron built in). Replaces hand-writing YAML or curl.
- **Kanban tasks** (`javis-task`) - bundled/chạy · tools: javis_task · Queues a background task on the Kanban board and shows how far running tasks have got, right from chat. Before this tool only engines that run shell commands (Claude Code, Codex) could queue work, because the only route was curl POST /kanban/task; the five API engines were left out.
- **Team chat management** (`javis-team`) - bundled/chạy · tools: javis_team · Creates, edits, archives and deletes team chats, adds or removes assistants, picks the leader, and creates new assistants with their own model then adds them, right from chat on every engine. Before this tool it took clicks on the Teams page or driving a browser.
- **Dashboard control** (`javis-ui`) - bundled/chạy · tools: javis_ui · Opens a page, a file in the brain or a Kanban task, or scrolls the chat on the open dashboard, as you say or type. The dashboard confirms it is done within the same turn; if no tab is open, the tool says so.
- **Web page watch** (`javis-watch`) - bundled/chạy · tools: javis_watch · Watches a web page on a schedule without spending tokens: alerts when the page changes, when some text appears, or when a price drops below a threshold (understands 1.250.000đ, 250k, $20). Only calls AI when summaries are on.
- **Workflow history** (`javis-workflow`) - bundled/chạy · tools: javis_workflow · Shows recent workflow runs and the result of each, right from chat. Before this tool a finished workflow's result was lost, and asking "how did the last workflow run go" left Javis nothing to look up.
- **Messenger health check** (`messenger-health`) - bundled/chạy · tools: messenger_check, messenger_repair · Checks each Facebook Page's Messenger bot connection (Page token, permissions, webhook, Page subscription, running bot) and repairs the webhook when the bot goes silent.
- **Meta Ads (Graph API)** (`meta-ads-graph`) - bundled/chạy · tools: meta_ads_accounts, meta_ads_insights, meta_ads_campaigns, meta_ads_get · Reads Facebook/Instagram ad numbers (ad accounts, campaigns, performance) through the Graph API, using the token of the "Meta Ads (your own app)" connection. READ-ONLY, spends no money.
- **Facebook Pages (Graph API)** (`meta-pages-graph`) - bundled/chạy · tools: fb_pages_list, fb_page_posts, fb_page_comments, fb_page_post, fb_page_photo, fb_page_album, fb_page_video, fb_page_edit, fb_page_delete, fb_page_reply · Manages Facebook Pages through the Graph API: list Pages, read posts and comments (read-only), publish posts, photos, videos and albums, edit text, delete posts and reply to comments (full access). Uses the Pages connected on the Facebook Messenger card (Connections page), or the store's Facebook Pages connection when present.
- **Tool usage log** (`tool-audit`) - bundled/tắt · tools: javis_tool_stats · hooks: post_tool_call · Counts how many times EACH tool is called by an engine (through the post_tool_call hook) and shows the most used tools. It is a demo of the plugin HOOK mechanism. OFF by default: turn it on with POST /plugins/toggle (slug=tool-audit) to try it.
- **Trạm tài liệu** (`tram-tai-lieu`) - bundled/chạy · tools: javis_tram_tai_lieu · OCR và chuyển đổi PDF/ảnh scan sang Word, Excel, Markdown, TXT hoặc PDF tìm kiếm được; xoay, tách, ghép, xóa trang; so sánh nội dung 2 file. Chạy local bằng Python venv .venv-paddleocr.
- **Voice notes** (`voice-note`) - bundled/chạy · tools: javis_voice_note · Javis records a short spoken message (up to 1000 characters) as an mp3 in attachments/voice/ and embeds it in the answer. The chat shows a player, Telegram receives the audio file. Uses the free Edge voices, picked by the language of the text.
- **Read YouTube videos** (`youtube-read`) - bundled/chạy · tools: javis_youtube_read · Reads the transcript (subtitles) of a YouTube video from its link so it can be summarised. Cycles through 6 player types, then yt-dlp, when YouTube blocks the server. No sign-in, no API key.

## Cờ sức khoẻ
- Workflow trỏ agent KHÔNG tồn tại: Theo dõi & Phân công công văn
- Agent chưa workflow nào dùng: javis-core, javis-core.conflict-remote-20261111-133623, kiem-soat-dieu-phoi-quy-trinh, phap-che-chinh-sach-thue
- Skill đang tắt: html-to-webcake, phan-tich-cau-hoi-sinh-doi-cheat-sheet, tien-trinh-supervisor-tu-dong
