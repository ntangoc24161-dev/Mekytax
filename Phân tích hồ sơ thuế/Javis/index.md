# Javis Index (tầng vận hành)

> Tự sinh từ file - ĐỪNG sửa tay. Chỉ mục mọi năng lực của Javis trong brain này để bất kỳ AI/engine đọc 1 chỗ là hiểu Javis làm được gì. Song song `wiki/index.md` (tri thức).

**Tổng quan:** 7 agents · 39 skills · 1 workflows (1 bật) · 0 loops (0 bật) · 28 plugins (27 chạy)

## Agents
- **chuyen-vien-du-lieu-dau-vao** (`chuyen-vien-du-lieu-dau-vao`) - 
- **Javis Core** (`javis-core`) - Trợ lý hệ thống nhận báo cáo và hỗ trợ công việc trong brain này.
- **phan-tich-vien-cqt** (`phan-tich-vien-cqt`) - 
- **phan-tich-vien-so-cty** (`phan-tich-vien-so-cty`) - 
- **phan-tich-vien-so-sach-dn** (`phan-tich-vien-so-sach-dn`) - 
- **thu-ho-so-kiem-ke** (`thu-ho-so-kiem-ke`) - 
- **truong-phong-phan-tich-ho-so** (`truong-phong-phan-tich-ho-so`) - 

## Skills
### AI
- **Ingest Source** (`ingest-source`) - Tiêu hoá một source thô vào Second Brain, chưng cất thành tri thức wiki tích luỹ.
- **Javis Builder** (`javis-builder`) - Tạo hoặc sửa năng lực của Javis: agent, skill, workflow, loop, plugin. Kèm mẫu file chuẩn và luật chống trùng.
- **Lint Wiki** (`lint-wiki`) - Rà soát sức khoẻ wiki của Second Brain, trả về danh sách vấn đề. Không tự sửa hàng loạt.
- **Notes** (`notes`) - Lưu tin nhắn hiện tại nguyên văn vào sources/ (kèm ảnh), tự chưng cất lên wiki nếu note đáng.
- **Query Wiki** (`query-wiki`) - Khai thác tri thức trong Second Brain: tổng hợp, so sánh, giả thuyết. Trả lời có trích dẫn.
### Chung
- **doi-chieu-ty-suat-sinh-loi** (`doi-chieu-ty-suat-sinh-loi`) - # Đối chiếu tỷ suất sinh lợi ngành (Kiểm tra TNDN bất thường)
- **giam-thue-tndn** (`giam-thue-tndn`) - # Giảm thuế TNDN, TNCN & GTGT (2020-2026)
- **khau-hao-tscđ** (`khau-hao-tscđ`) - # Khấu hao tài sản cố định (Quy trình & Phương pháp)
### Marketing
- **HTML sang Webcake** (`html-to-webcake`) · [TẮT] - Chuyển trang HTML thành file .pke mở được trong trình dựng Webcake, giữ đúng màu, cỡ chữ, ảnh và bố cục của bản gốc.
### Năng suất
- **PaddleOCR** (`paddleocr`) - OCR ảnh và PDF bằng PaddleOCR qua MCP (tool ocr). Bảng/layout: đổi model PP-StructureV3.
- **PDF sang Markdown** (`pdf-sang-md`) - Chuyển PDF/ảnh scan thành file Markdown (.md) cho Second Brain, giữ heading/bảng, lọc header/footer.
- **Trạm tài liệu** (`tram-tai-lieu`) - Chuyển PDF/ảnh scan ra Word, Excel, TXT, PDF tìm kiếm được; xoay/ghép/tách trang; so chữ hai file; theo dõi thư mục.
- **Phân tích câu hỏi sinh đôi và Cheat Sheet ôn thi** (`phan-tich-cau-hoi-sinh-doi-cheat-sheet`) · [TẮT] - Phân tích câu hỏi sinh đôi, tìm keyword cốt lõi, bóc tách bẫy đáp án, lập Cheat Sheet mapping và tóm tắt siêu tốc < 300 từ dạng Word/Excel ôn thi.
- **Soạn thảo và chuẩn hóa văn bản hành chính** (`soan-thao-van-ban`) · [TẮT] - Dự thảo, rà soát thể thức 29 loại VB hành chính theo NĐ 30/2020, viết hoa, chữ ký số và ký hiệu trình ký TCS3.
- **Tiến trình supervisor tự động** (`tien-trinh-supervisor-tu-dong`) · [TẮT] - Supervisor tự động điều phối, giám sát chu trình giải đề khảo thí cuốn chiếu, tự phục hồi khi gián đoạn mạng và kiểm kê đến khi đủ mục tiêu.
- **Tra cứu và tải văn bản iOffice Đồng Nai** (`tra-cuu-ioffice-dong-nai`) · [TẮT] - Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
### Security
- **Security Audit** (`security-audit`) - Rà soát bảo mật mã nguồn, API, dịch vụ: tìm lỗ hổng có bằng chứng, kiểm chứng độc lập từng lỗi, báo cáo kèm cách sửa. Từ Cloudflare.
### Tài chính
- **Đối chiếu doanh thu 511 và giá vốn 632** (`doi-chieu-511-632`) - Đối chiếu sổ chi tiết TK 511 với TK 632 theo Tool_DC_511_632: khớp HĐ/ngày CT/mặt hàng, lãi gộp và bán dưới giá.
- **Đối chiếu tờ khai GTGT và TNDN chuẩn CBRR** (`doi-chieu-gtgt-tndn-cbrr`) - Đối chiếu Bảng kê - Tờ khai GTGT - Quyết toán TNDN - HĐĐT chuẩn 2 sheet CBRR, bóc tách nguyên nhân trọng yếu và xuất Excel.
- **Đối chiếu mặt hàng mua vào bán ra** (`doi-chieu-mat-hang-mua-ban`) - Đối chiếu mặt hàng mua-bán HĐĐT theo tiền, số lượng, đơn giá; lọc HĐ hủy, cắt kỳ chồng, bắt chuỗi trung gian.
- **Đối chiếu NCC mua vào với CVCBRR và danh bạ** (`doi-chieu-ncc-mua-vao-rui-ro`) - Đối chiếu NCC từ Bảng kê + HĐĐT với CVCBRR F0/F1, Công an, danh bạ; ghi nguồn Bảng kê/HĐĐT/Chung; Bảng 2.3 theo kỳ.
- **Giai đoạn 1 - Kiểm kê & Đánh giá hồ sơ thuế** (`giai-doan-1-kiem-ke-danh-gia`) - Kiểm kê manifest IN00-IN13, xác định mức hồ sơ A-E, rà soát khả năng chạy 80 quy tắc, xếp danh mục đề nghị DN bổ sung (KQ1).
- **Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế** (`giai-doan-2-phan-tich-du-lieu-co-quan-thue`) - Phân tích 57 quy tắc thuế khi chưa có sổ DN: chuẩn hóa IN01-IN07/12, đối chiếu GTGT-TNDN-BCTC, 100% NCC rủi ro, phân loại Chắc/Sơ bộ (KQ2).
- **Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN** (`giai-doan-3-hoan-chinh-ho-so`) - Hoàn chỉnh hồ sơ khi có sổ sách DN: 23 quy tắc GĐ3, chuỗi bằng chứng 5 mắt xích, Master 28 tiêu chí 7 cột, đề cương kiểm tra (KQ3).
- **Gộp hóa đơn mua vào bán ra** (`gop-hoa-don-mua-ban`) - Gộp hóa đơn mua vào/bán ra thành Excel 2 sheet, mỗi công ty 1 file; mini app HTML hoặc exe không cần admin.
- **Làm sạch MST và hóa đơn CD1** (`lam-sach-cd1-mst-hoa-don`) - Chuẩn hóa MST, quét HĐ trùng, cụm cùng ngày 20tr/5tr, đối chiếu BKMV-HĐĐT-NKC 133 theo CD1 BT5-BT9.
- **Lập hồ sơ giải thể tại bàn** (`lap-ho-so-giai-the-tai-ban`) - Lập hồ sơ giải thể tại bàn từ folder NNT: nạp Bang tinh từ phôi chuẩn, bóc XML/HĐĐT, điền Word theo TB ngừng MST.
- **Mở hồ sơ TTKT** (`mo-ho-so-ttkt`) - Checklist mở hồ sơ TTKT: chọn đối tượng, đối chiếu BCTC-tờ khai-HĐĐT, GDLK, phân nhánh thiếu thuế/trốn thuế.
- **Phân tích công nợ và tiền mặt** (`phan-tich-cong-no-tien-mat`) - Quét chi tiền mặt TK 111, cụm HĐ cùng ngày, công nợ 131/331 và sao kê ngân hàng 112.
- **Phân tích giá thành và định mức** (`phan-tich-gia-thanh-dinh-muc`) - Bóc giá thành 621/622/627 hoặc cơ cấu GV thương mại, khớp 632 với BCTC [02], giải tỏa TC28.
- **Phân tích hồ sơ kiểm tra thuế** (`phan-tich-ho-so-thue`) - Quy trình 1.4: 6 tệp KQ1.1-KQ3, 80 quy tắc; sao mẫu rồi điền; KQ2 chỉ số liệu, kết luận chỉ ở KQ3.
- **Phân tích XNT kho** (`phan-tich-xnt-kho`) - Đối chiếu XNT kho 152/155/156 với BCTC HTK, xuất kho với giá vốn 632, giải tỏa TC3.
- **Tổ chức hồ sơ kiểm tra thuế** (`to-chuc-ho-so-kiem-tra-thue`) - Tổ chức thư mục hồ sơ kiểm tra thuế 4 bước, phân loại file, checklist đủ/thiếu A4, gom 1 nơi.
- **Tổng hợp bên mua từ sổ CVCBRR** (`tong-hop-ben-mua-tu-so-cvcbrr`) - Tổng hợp đơn vị bên mua từ sổ CVCBRR: xuất Chi_tiet theo ngày CV và Tong_hop theo MST mua.
- **Tổng hợp đầu vào HĐĐT, bảng kê và NLĐ** (`tong-hop-dau-vao-hddt-bang-ke-nld`) - Gộp HĐĐT portal, bảng kê NNT (lọc Tổng cộng) và NLĐ QTT TNCN thành file mẫu 02, không trộn nguồn.
- **Tổng hợp TNCN và tra cứu đa nguồn** (`tong-hop-tncn-da-nguon`) - Tổng hợp QTT TNCN, tờ khai quý, danh sách NLĐ và tra cứu đa nguồn; đối chiếu TC22/TC23.
- **Quản lý rủi ro tuân thủ thuế** (`quan-ly-rui-ro-tuan-thu-thue`) · [TẮT] - QLRR tuân thủ thuế: Luật 108/2025, TT 31/2021, phân loại rủi ro, bộ chỉ số TCT, 5 bước xử lý rủi ro.
- **Trả lời công văn Công an** (`tra-loi-cong-van`) · [TẮT] - Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho cơ quan Công an theo quy chuẩn thể thức và dữ liệu quản lý thuế.
- **Xử phạt vi phạm hành chính về thuế và hóa đơn** (`xu-phat-vphc-thue-hoa-don`) · [TẮT] - Xử phạt VPHC thuế và hóa đơn: NĐ 125, NĐ 310, NĐ 291, NĐ 68: sai thời điểm, không lập HĐ, trốn thuế, trao đổi thông tin QT, xử lý DN 03-06.

## Workflows
- **Phân tích hồ sơ thuế 3 giai đoạn** (`phan-tich-ho-so-thue-3-giai-doan`) - active · 7 bước [thu-ho-so-kiem-ke -> truong-phong-phan-tich-ho-so -> chuyen-vien-du-lieu-dau-vao -> phan-tich-vien-cqt -> phan-tich-vien-so-cty -> phan-tich-vien-so-sach-dn -> truong-phong-phan-tich-ho-so] · Chuỗi Lệnh 1-3 quy trình 1.4 vòng 2: KQ1 rồi KQ2.1/KQ2.2 rồi KQ3 + phụ lục GĐ3; phòng 6 vai.

## Loops
_(chưa có)_

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
- Agent chưa workflow nào dùng: javis-core
- Skill đang tắt: html-to-webcake, phan-tich-cau-hoi-sinh-doi-cheat-sheet, quan-ly-rui-ro-tuan-thu-thue, soan-thao-van-ban, tien-trinh-supervisor-tu-dong, tra-cuu-ioffice-dong-nai, tra-loi-cong-van, xu-phat-vphc-thue-hoa-don
