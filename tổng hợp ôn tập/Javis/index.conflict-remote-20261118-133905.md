# Javis Index (tầng vận hành)

> Tự sinh từ file - ĐỪNG sửa tay. Chỉ mục mọi năng lực của Javis trong brain này để bất kỳ AI/engine đọc 1 chỗ là hiểu Javis làm được gì. Song song `wiki/index.md` (tri thức).

**Tổng quan:** 0 agents · 26 skills · 0 workflows (0 bật) · 0 loops (0 bật) · 27 plugins (26 chạy)

## Agents
_(chưa có)_

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
- **Chuẩn hóa và xuất ngân hàng trắc nghiệm** (`chuan-hoa-ngan-hang-trac-nghiem`) - Chuẩn hóa ngân hàng trắc nghiệm Excel sang Word: gom nhóm chuyên đề, sắp câu sinh đôi, tô vàng keyword, in đậm đáp án chuẩn, format in ấn.
- **Truy cập và tổng hợp đề thi** (`lay-va-tong-hop-de-thi`) - Truy cập hệ thống khảo thí thuế, tự động lấy đề, dò đáp án chuẩn và tổng hợp ngân hàng câu hỏi vào bảng tính Excel.
- **PaddleOCR** (`paddleocr`) - OCR ảnh và PDF bằng PaddleOCR qua MCP (tool ocr). Bảng/layout: đổi model PP-StructureV3.
- **PDF sang Markdown** (`pdf-sang-md`) - Chuyển PDF/ảnh scan thành file Markdown (.md) cho Second Brain, giữ heading/bảng, lọc header/footer.
- **Phân tích câu hỏi sinh đôi và Cheat Sheet ôn thi** (`phan-tich-cau-hoi-sinh-doi-cheat-sheet`) - Phân tích câu hỏi sinh đôi, tìm keyword cốt lõi, bóc tách bẫy đáp án, lập Cheat Sheet mapping và tóm tắt siêu tốc < 300 từ dạng Word/Excel ôn thi.
- **Soạn thảo và chuẩn hóa văn bản hành chính** (`soan-thao-van-ban`) - Dự thảo, rà soát thể thức 29 loại VB hành chính theo NĐ 30/2020, viết hoa, chữ ký số và ký hiệu trình ký TCS3.
- **Tiến trình supervisor tự động** (`tien-trinh-supervisor-tu-dong`) - Supervisor tự động điều phối, giám sát chu trình giải đề khảo thí cuốn chiếu, tự phục hồi khi gián đoạn mạng và kiểm kê đến khi đủ mục tiêu.
- **Tra cứu và tải văn bản iOffice Đồng Nai** (`tra-cuu-ioffice-dong-nai`) - Tự động đăng nhập VNPT iOffice Đồng Nai, tra cứu văn bản đến/đi theo trích yếu, ngày tháng và tải tệp đính kèm về máy tính.
- **Trạm tài liệu** (`tram-tai-lieu`) - Chuyển PDF/ảnh scan ra Word, Excel, TXT, PDF tìm kiếm được; xoay/ghép/tách trang; so chữ hai file; theo dõi thư mục.
### Security
- **Security Audit** (`security-audit`) - Rà soát bảo mật mã nguồn, API, dịch vụ: tìm lỗ hổng có bằng chứng, kiểm chứng độc lập từng lỗi, báo cáo kèm cách sửa. Từ Cloudflare.
### Tài chính
- **Đối chiếu NCC mua vào với CVCBRR và danh bạ** (`doi-chieu-ncc-mua-vao-rui-ro`) - Đối chiếu NCC từ Bảng kê + HĐĐT với CVCBRR F0/F1, Công an, danh bạ; ghi nguồn Bảng kê/HĐĐT/Chung; Bảng 2.3 theo kỳ.
- **Giai đoạn 1 - Kiểm kê & Đánh giá hồ sơ thuế** (`giai-doan-1-kiem-ke-danh-gia`) - Kiểm kê manifest IN00-IN13, xác định mức hồ sơ A-E, rà soát khả năng chạy 80 quy tắc, xếp danh mục đề nghị DN bổ sung (KQ1).
- **Giai đoạn 2 - Phân tích dữ liệu cơ quan thuế** (`giai-doan-2-phan-tich-du-lieu-co-quan-thue`) - Phân tích 57 quy tắc thuế khi chưa có sổ DN: chuẩn hóa IN01-IN07/12, đối chiếu GTGT-TNDN-BCTC, 100% NCC rủi ro, phân loại Chắc/Sơ bộ (KQ2).
- **Giai đoạn 3 - Hoàn chỉnh hồ sơ & Sổ sách DN** (`giai-doan-3-hoan-chinh-ho-so`) - Hoàn chỉnh hồ sơ khi có sổ sách DN: 23 quy tắc GĐ3, chuỗi bằng chứng 5 mắt xích, Master 28 tiêu chí 7 cột, đề cương kiểm tra (KQ3).
- **Ôn tập Luật Phòng chống tham nhũng** (`on-tap-luat-phong-chong-tham-nhung`) - Ôn tập thi trắc nghiệm Luật Phòng chống tham nhũng 2018: 96 điều, 10 chương, bẫy đề thi, thời hạn, định mức kê khai TSTN, xử lý trách nhiệm.
- **Ôn tập nghiệp vụ thuế tháng 09/2026** (`on-tap-nghiep-vu-thue-thang-09-2026`) - Ôn tập chuyên môn nghiệp vụ thuế đợt 09/2026: 14 chuyên đề, ngân hàng 1.204 câu, 27 bộ đề, gom câu sinh đôi, highlight keyword và xuất Word/Excel.
- **Phân tích công nợ và tiền mặt** (`phan-tich-cong-no-tien-mat`) - Quét chi tiền mặt TK 111, cụm HĐ cùng ngày, công nợ 131/331 và sao kê ngân hàng 112.
- **Phân tích hồ sơ kiểm tra thuế** (`phan-tich-ho-so-thue`) - Quy trình 3 giai đoạn phân tích hồ sơ thuế theo bộ biểu mẫu v1.1 (80 quy tắc), kết nối dữ liệu CQT và sổ sách DN, chốt Master 28 tiêu chí (KQ1-3).
- **Trả lời công văn Công an** (`tra-loi-cong-van`) - Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho cơ quan Công an theo quy chuẩn thể thức và dữ liệu quản lý thuế.
- **Xử phạt vi phạm hành chính về thuế và hóa đơn** (`xu-phat-vphc-thue-hoa-don`) - Xử phạt VPHC thuế và hóa đơn: NĐ 125, NĐ 310, NĐ 291, NĐ 68: sai thời điểm, không lập HĐ, trốn thuế, trao đổi thông tin QT, xử lý DN 03-06.

## Workflows
_(chưa có)_

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
