---
type: loop
name: Hồ sơ cung cấp công an
slug: ho-so-cung-cap-cong-an
enabled: false
goal: custom
mode: full
interval_min: 1440
workspace: vault
tools_profile: vault-safe
quiet_hours: ''
max_runs_per_day: 1
owner_chat: web:mwlp8ddhcfj9yn08
notify: true
updated: '2026-11-18T15:00'
---

## Mục tiêu
Mốc quét cố định **12:00 giờ Việt Nam** do reminder cron `r_7c63a0c056` (loop này **tắt**, chỉ giữ bản mô tả quy trình). Khi cron chạy: quét CV MOI, **phân loại file**, xử lý đúng nhánh; soạn thảo vẫn chỉ khi có lệnh riêng.

## Quy trình khi cron 12h chạy
1. **Quét thư mục CV MOI** (`I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI\`)
   - Tìm PDF/Word; đối chiếu sổ `Bao_cao_cong_van.xlsx`
   - Drive sống: Composio `googledrive_ralph-urbane`
   - **Phân loại từng file theo checklist cứng** (mục dưới) trước khi trigger workflow

### Checklist phân loại file trong CV MOI
**Nhánh B - CV trả lời đã ký** nếu có **≥1** dấu hiệu:
- Số hiệu ban hành dạng `…/TCS3-KTr1` hoặc `KTr2` / `KTr3`
- Tên file có `TRA LOI`, `TRA_LOI`, `.signed`, `ky so`, `đã ký`
- OCR: cơ quan ban hành là Thuế cơ sở / Thuế TP; kính gửi Công an; có chữ ký + mộc đỏ; trích yếu kiểu "báo cáo kết quả xử lý theo CV…"

**Nhánh A - CV Công an mới** nếu:
- Cơ quan ban hành là Công an / CQCSĐT / CQANĐT / CSKT / ANĐT (hoặc số hiệu `…/CSKT-…`, `…/ANĐT-…`, `…/YC-…` từ CA)
- **Và** nội dung yêu cầu cung cấp hồ sơ / tài liệu / thông tin NNT

**Không chắc** → báo chat (tên file + lý do), không tự chọn nhánh, không ghi `Đã hoàn thành`.

2. **Nhánh A - Công văn Công an mới** (yêu cầu cung cấp hồ sơ): Trigger **Workflow "Quản lý công văn công an"** (2 bước)
   - **Bước 1:** Agent `khai-thac-du-lieu-ho-so` - Bóc tách DN, tra DB NNT, phân Tổ KT, ghi CBQL
   - **Bước 2:** Agent `Theo dõi & Phân công` - KT1 in (xem trước → chốt in); KT2/KT3 Zalo skill `gui-phan-cong-zalo-kt23` (xem trước → chốt gửi); sau bàn giao thành công mới đổi cột O sang `Đang xử lý`
   - Kiểm soát: `Kiểm soát & Điều phối quy trình`
   - **Dừng tại đây.** Không tự soạn thảo.

3. **Nhánh B - Công văn trả lời đã ký** (bản scan ký / đóng dấu trong CV MOI): Trigger **Workflow `theo-doi-cv-tra-loi-da-ky`**
   - Bóc tách, khớp sổ, cập nhật `Đang xử lý` → `Đã hoàn thành` (hậu tố `-R` nếu cần)
   - Chạy ngay trong cùng lần quét trưa, không chờ lệnh riêng

4. **Khi người dùng yêu cầu soạn thảo:** Workflow/agent `soan-thao-tra-loi-cong-an` (hoặc `soan-thao-quan-ly-ho-so` nếu bật lại)

5. **Khi có lệnh theo dõi ngoài CV MOI** (quét năm 2026 / thư mục đã lưu): vẫn dùng `theo-doi-cv-tra-loi-da-ky`

6. **Báo cáo** về `chat_id` của cron / `owner_chat` (phiên dashboard đang gắn)

## Ghi sổ (nền tảng)
- **Linux/VPS (Javis hiện tại):** Drive Composio `googledrive_ralph-urbane` + vá XML xlsx (giữ DASHBOARD), upload lại, kiểm tra mở được
- **Windows có Excel:** Excel COM (`win32com`)

## Ghi chú
- CV MOI trống → báo "Không có CV mới", không kích hoạt workflow
- Không đổi `Đang xử lý` trước khi chốt in / gửi Zalo thành công (nhánh A)
- Phân loại bắt buộc theo checklist cứng ở trên; không chắc thì báo chat
- Loop `enabled: false` theo chốt 2A (18/11/2026): chỉ cron 12h là kênh quét
- Chốt 18/11/2026: bản ký trong CV MOI → nhánh B xử lý luôn lúc trưa
