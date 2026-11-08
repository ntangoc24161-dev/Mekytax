---
type: wiki
status: active
tags: [wiki, van-hanh, dieu-phoi]
created: 2026-11-06
updated: 2026-11-06
source: backlog chủ 2026-11-06 + [[Phuong-phap-lien-ket-phong-uu-nhuoc]] + [[Ban-do-lien-ket-phong-brain]] + docs 11-telegram / 26-kenh-zalo-bot + javis-builder model field
---
# Backlog vận hành liên phòng (6 điểm)

Trang này chốt **hiện trạng đo được**, **quy tắc**, và việc đã/chưa làm. Căn cứ [[Phuong-phap-lien-ket-phong-uu-nhuoc]], [[Ban-do-lien-ket-phong-brain]], skill `ban-giao-lien-phong`.

## 1. Handoff có xác nhận (gợi ý brain + trợ lý, hỏi rồi mới đổi)

**Hiện trạng:** Javis không tự đổi brain giúp người dùng. Skill và agent Văn thư chỉ ghi checklist + hướng dẫn đổi tay.

**Quy tắc UX (đã nhúng skill `ban-giao-lien-phong` + agent `phong-van-thu`):**

1. Khi việc thuộc brain khác: nêu rõ Brain đích + Agent/Workflow trong 1 khối bàn giao.
2. **Hỏi xác nhận** trước khi coi là đã bàn giao (nút JAVIS_ASK trên dashboard; Telegram/Zalo thì danh sách số).
3. Chỉ sau khi chủ chọn "Đồng ý đổi" / trả lời đúng: nhắc thao tác `đổi brain X → Cộng sự → agent/workflow Y` (web) hoặc `/brain <tên>` rồi gọi agent (Telegram/Zalo).
4. Không nói "đã chuyển" nếu chủ chưa xác nhận và chưa đổi brain thật.

**Mục tiêu / kế hoạch (chủ muốn, 2026-11-06):** sau xác nhận bàn giao, Javis **tự đổi brain + mở agent/workflow đích** (không chỉ hướng dẫn tay). **Hiện trạng:** chưa có trên nền tảng; mới ghi `_open-questions`. Workaround: anh `/brain` hoặc chọn brain trên dashboard / đứng sẵn đúng brain chuyên (lối 1).

## 2. Mirror agent: nguồn chân lý một chiều = brain chuyên

**Hiện trạng sau duyệt chủ (2026-11-06):**

| Nơi | Số agent active | Ghi chú |
|---|---|---|
| VĂN THƯ | **4** | Chỉ hub; 22 mirror + 5 workflow chuyên đã chuyển `agents/_archive-mirror-2026-11-06/` và `workflows/_archive-mirror-2026-11-06/` |
| Phân tích hồ sơ thuế | 6 | Nguồn chân lý phân tích |
| Hồ sơ kiểm tra | 5 | Nguồn chân lý soạn HS |
| Cung cấp thông tin công an | 4 active (+1 conflict-remote) | Nguồn chân lý Công an |
| Cập nhật công văn cảnh báo | 5 active (+1 conflict-remote) | Nguồn chân lý cảnh báo |
| Ôn tập thi / IFOFFIC / … | 0 | Chỉ skill |

**Giữ tại VĂN THƯ (hub):** `phong-van-thu`, `so-hoa-ocr-cong-van`, `soan-thao-van-ban`, `tri-thuc-second-brain`.

**Quy tắc đồng bộ:**

- Sửa nghiệp vụ agent → sửa ở **brain chuyên**.
- File `*.conflict-remote-*` ở Công an / Cảnh báo: chọn 1 bản rồi xoá bản conflict (còn mở).
- Khôi phục mirror tạm: Move file từ `_archive-mirror-2026-11-06/` lên lại `agents/` hoặc `workflows/`.

## 3. Ghim model theo agent

Frontmatter agent hỗ trợ `model:` (`""` mặc định | `haiku` | `sonnet` | `opus` | ...). [[javis-builder]]

**Chính sách đề xuất:**

| Loại việc | Model | Ví dụ agent |
|---|---|---|
| Trao đổi nhẹ, phân sổ, OCR, checklist, hỏi lại | `haiku` | `phong-van-thu`, `so-hoa-ocr-cong-van`, `khoi-tao-ho-so`, `tra-cuu-tai-cv-ioffice` |
| Chốt luận, dự thảo VB quan trọng, Master 28, phúc đáp Công an | `sonnet` | `truong-phong-phan-tich-ho-so`, `soan-thao-tra-loi-cong-an`, `soan-thao-van-ban`, `cập-nhật-công-văn-cảnh-báo-rủi-ro` |
| Việc cực khó / dài (hiếm) | `opus` hoặc để trống + chọn tay | chỉ khi chủ yêu cầu |

**Đã làm (2026-11-06):** ghim hub VĂN THƯ + **20 agent** trên 4 brain chuyên (bỏ qua 2 file conflict-remote).

## 4. Orch có kiểm soát theo phòng

**Hiện trạng:** hầu hết brain `orchestration: off` sau cắt token ([[Cat-cau-hinh-token-dot-3-cong-an-canh-bao-on-tap]]).

**Thiết kế lại (mặc định an toàn):**

| Brain | Orch mặc định | Khi được bật có chủ đích |
|---|---|---|
| Brain Default | **off** | Không bật nền đa việc |
| VĂN THƯ | **off** | Chỉ loop đôn đốc 180 phút; không auto claim Kanban nặng |
| Phân tích / Hồ sơ kiểm tra / Công an / Cảnh báo | **off** | Bật **tạm** khi có chuỗi vụ rõ (1-2 ngày), tắt lại sau |
| Ôn tập thi | **off** | Bật tạm khi chạy supervisor giải đề có mục tiêu số câu |

**Điều kiện trước khi bật orch 1 phòng:**

1. Việc đã có gói bàn giao + đường dẫn hồ sơ.
2. Skill/agent đúng phòng (không còn catalog lệch nặng).
3. Có kênh báo kết quả (`owner_chat` / inbox).
4. Chủ biết tắt lại sau đợt.

Không bật orch "cho chắc" trên Default.

## 5. Archive / gộp brain ít dùng

**Ứng viên (0 agent, chồng skill với phòng chính):**

| Brain | Đề xuất | Lý do |
|---|---|---|
| IFOFFIC | Archive / gộp vào Cảnh báo hoặc VĂN THƯ (skill iOffice) | 0 agent; tra cứu iOffice đã có ở Cảnh báo |
| KIỂM TRA THUẾ | Archive / gộp vào Phân tích hoặc Hồ sơ kiểm tra | Trùng phạm vi tên |
| Brain tính cup | Archive / gộp vào Phân tích (skill CUP) | 0 agent |
| tổng hợp ôn tập | Gộp vào **Ôn tập thi** rồi archive | Trùng ôn tập |
| CONG TAC D?NG (nếu còn) | Archive nếu không dùng | 0 agent |
| Brain Default | **Giữ** làm inbox nhẹ (không xoá được brain mặc định theo docs) | Đã cắt skill; kênh hay rơi vào đây |

**Đã chuẩn bị:** [[Checklist-archive-brain-it-dung]]. **Chưa xoá brain** (chủ làm trên UI, gõ đúng tên).

## 6. Gắn kênh Telegram / Zalo đúng brain

Nguồn: docs `11-telegram.md`, `26-kenh-zalo-bot.md`.

**Cơ chế thật:**

- Bot Telegram/Zalo là **một** bot toàn hệ; brain theo **phiên người dùng**, không phải "mỗi brain một token".
- Chưa từng `/brain` → phiên rơi **brain mặc định** (thường Brain Default) - đúng chỗ hay cháy token trước đây.
- Đổi: gõ `/brain` (bảng nút) hoặc `/brain <tên>` (vd `/brain VĂN THƯ`). Đổi xong hội thoại reset, file inbox theo brain mới.
- Kiểm: `/status` xem brain đang dùng.
- Zalo Bot: lệnh giống (`/brain`, `/status`) nhưng phải gõ tay (không menu).

**Hướng dẫn vận hành đề xuất:**

1. Sau khi bật bot: `/brain VĂN THƯ` làm cửa vào hub (sổ + phân việc).
2. Khi làm vụ chuyên: `/brain Phân tích hồ sơ thuế` (hoặc Công an / Cảnh báo / Hồ sơ kiểm tra / Ôn tập thi) trước khi gửi file.
3. Xong vụ: `/brain VĂN THƯ` lại để đôn đốc / thể thức.
4. Tránh để Default làm nơi chat hàng ngày; Default chỉ inbox khi chưa chọn.
5. Chatbot khách (trang Chatbot) gắn agent/brain riêng - khác kênh chủ nhắn bot cá nhân.

## Việc còn lại

1. ~~Gỡ mirror VĂN THƯ~~ - xong (archive folder, không xoá hẳn).
2. ~~Ghim model brain chuyên~~ - xong (trừ conflict-remote).
3. Chủ **xoá/archive** brain ít dùng theo [[Checklist-archive-brain-it-dung]] (ô checkbox từng brain).
4. Dọn 2 file `*.conflict-remote-*` ở Công an và Cảnh báo.
5. Chốt `/brain VĂN THƯ` làm cửa vào Telegram/Zalo; đo token 7 ngày.

## Liên kết

- [[Ban-do-lien-ket-phong-brain]]
- [[Phuong-phap-lien-ket-phong-uu-nhuoc]]
- [[Cat-cau-hinh-token-Brain-Default]]
- [[Uoc-tinh-tiet-kiem-token-sau-cat]]
