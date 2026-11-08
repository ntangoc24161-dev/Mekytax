---
type: wiki
status: active
tags: [wiki, van-hanh, dieu-phoi]
created: 2026-10-23
updated: 2026-11-06
source: Ban-do-lien-ket-phong-brain + skill ban-giao-lien-phong + loop van-thu + cắt token 2026-10-23 + case CTL
---
# Phương pháp liên kết phòng: cách làm và đánh giá

Căn cứ [[Ban-do-lien-ket-phong-brain]], skill `ban-giao-lien-phong`, loop `van-thu-dieu-phoi-don-doc`, quyết định hub Văn thư, và kết quả cắt token / vụ CTL.

## Phương pháp đang dùng (hub + bàn giao)

Javis không gọi agent chéo brain. Phương pháp bù bằng 4 lớp:

1. **Hub Văn thư** - một cửa nhận việc, ghi sổ, phân phòng, đôn đốc, kiểm thể thức VB đi.
2. **Bản đồ** - việc → brain đích → agent/workflow chuẩn.
3. **Checklist bàn giao** - gói đủ đường dẫn, MST, đầu ra, lệnh "đổi brain → chạy Y".
4. **Chạy chuyên môn đúng brain** - workflow/agent tại brain đích; báo kết quả về sổ.

## Cách làm chi tiết (từng bước)

### A. Việc mới vào hệ

1. Người dùng giao việc (chat VĂN THƯ, Telegram, hoặc bỏ file vào luồng văn thư).
2. Agent `phong-van-thu` hoặc loop đọc sổ `06 - Văn thư/so-van-ban-den.md`.
3. Dòng `Mới` được khớp bảng phân việc: ghi Phòng thụ lý, Hạn chót, `Đã phân`.
4. Ghi `Kết quả báo lại` dạng: `Brain: … · Chạy: … · Đầu ra: …`.

### B. Bàn giao sang phòng chuyên

1. Mở skill `ban-giao-lien-phong` (hoặc agent Văn thư tự điền checklist).
2. Đủ tối thiểu: việc, brain đích, agent/workflow, đường dẫn hồ sơ, đầu ra, hạn.
3. Người dùng (hoặc Kanban đúng brain) **đổi brain** trên dashboard.
4. Vào **Cộng sự** → chọn workflow/agent đúng map → chạy.
5. File kết quả để đúng thư mục vụ; ghi lại sổ Văn thư khi xong.

### C. Trong một brain (không cần bàn giao)

Dùng workflow nội bộ: nhiều agent nối `{{prev}}`. Ví dụ `phan-tich-ho-so-thue-3-giai-doan`, `day-chuyen-xu-ly-cv-canh-bao`, `quản-lý-công-văn-công-an`.

### D. Báo về và đôn đốc

- Loop 180 phút: quá hạn chưa có `Kết quả báo lại` → đánh `Quá hạn` và báo.
- Văn bản đi: kiểm thể thức NĐ 30 trước khi cấp số (đúng vai hub Văn thư).

## Kết quả đã đạt được (hiện trạng đo / đã dựng)

- Có map + skill + loop/agent Văn thư đã chỉnh theo brain đích.
- Cắt token: orch off hầu hết brain; skill lệch tắt; loop giãn 180 phút; Default/Phân tích/Hồ sơ kiểm tra/Công an/Cảnh báo/Ôn tập đã thu gọn.
- Case CTL chứng minh tách đúng phòng có lợi: bảng kê/tờ khai ở **Phân tích hồ sơ thuế**; đối chiếu mặt hàng + dự thảo CV ở **Cập nhật công văn cảnh báo**.

## Ưu điểm

- **Khớp giới hạn Javis**: không phụ thuộc tính năng gọi chéo brain (không có).
- **Đúng quyết định nghiệp vụ**: một đầu mối kiểm soát sổ và VB phát hành.
- **Giảm trộn việc**: phân tích ≠ soạn hồ sơ ≠ cảnh báo ≠ công an; catalog skill mỗi brain nhẹ hơn sau cắt.
- **Memory đúng chỗ**: facts CTL mặt hàng không lẫn với facts phân tích tờ khai.
- **Tiết kiệm token có kiểm soát**: orch off + đúng brain tránh Default/phòng ma chạy nặng.
- **Truy vết được**: sổ + checklist + wiki map; trả lời sau này có nguồn.

## Nhược điểm / rủi ro

- **Không tự động end-to-end**: sau phân sổ vẫn cần người đổi brain / bấm workflow (hoặc bật orch đúng brain).
- **Ma sát thao tác**: nhiều brain = nhiều lần chuyển context; dễ quên báo `Kết quả báo lại`.
- **Mirror agent trong VĂN THƯ**: chạy nhanh tại hub được, nhưng dễ làm việc “nhầm vault” (mất facts brain chuyên).
- **Hai nguồn sự thật**: cùng vụ có thể có file trên đĩa + facts ở 2 brain; cần kỷ luật đường dẫn trong checklist.
- **Orch off toàn hệ**: an toàn token nhưng bàn giao Kanban tự chạy gần như tắt; phụ thuộc tay nhiều hơn.
- **Map phải bảo trì**: thêm loại việc / đổi slug workflow mà không cập nhật wiki+loop thì phân sai.

## Khi nào phương pháp này đủ / khi nào cần đổi

- **Đủ** khi số vụ vừa phải, anh chấp nhận 1 bước đổi brain, ưu tiên đúng memory và tiết kiệm token.
- **Cần siết thêm** nếu hay quên bàn giao: bắt buộc ô `Kết quả báo lại` có dòng Brain/Chạy trước khi coi là Đã phân (loop đã yêu cầu).
- **Cân nhắc gộp** nếu 2 brain ít dùng và trùng 80% agent: gộp để giảm ma sát (đổi lấy catalog nặng hơn).
- **Không chờ** “agent A tự gọi agent B ở brain khác”: đó không phải hướng Javis hiện có.

## Điểm backlog đã chốt quy tắc (2026-11-06)

Chi tiết đầy đủ: [[Backlog-van-hanh-lien-phong]].

1. **Handoff có xác nhận** - skill + `phong-van-thu` hỏi rồi mới hướng dẫn đổi brain / `/brain`.
2. **Mirror** - nguồn chân lý = brain chuyên; danh sách gỡ khỏi hub chờ chủ duyệt.
3. **Ghim model** - haiku việc nhẹ, sonnet chốt; đã ghim 4 agent hub.
4. **Orch** - mặc định off; bật tạm từng phòng khi có chuỗi vụ rõ rồi tắt lại.
5. **Archive** - đề xuất IFOFFIC, KIỂM TRA THUẾ, Brain tính cup, tổng hợp ôn tập; chưa xoá.
6. **Kênh** - `/brain` gắn phiên Telegram/Zalo; khuyến nghị cửa vào VĂN THƯ, tránh Default.
