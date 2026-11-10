---
type: loop
name: Hồ sơ cung cấp công an
slug: ho-so-cung-cap-cong-an
enabled: true
goal: custom
mode: full
interval_min: 1440
workspace: vault
tools_profile: vault-safe
quiet_hours: ''
max_runs_per_day: 1
owner_chat: web:3e4326e5df604ecabd9c4fec2e2892ca
notify: true
updated: '2026-11-08'
---

## Mục tiêu
Mỗi ngày lúc 12:00 giờ Việt Nam, tự động kiểm thư mục CV MOI để phát hiện công văn Công an mới, sau đó bóc tách dữ liệu, tra cứu DB NNT, soạn dự thảo công văn trả lời và tờ trình lãnh đạo, rồi cập nhật sổ theo dõi.

## Quy trình tự động
1. **Quét thư mục CV MOI** (`I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI\`)
   - Tìm tất cả file PDF/Word có tên công văn (mẫu: `CV XXXX-...` hoặc `CV MOI XXXX`)
   - So với sổ theo dõi `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` để phát hiện CV chưa xử lý

2. **Với mỗi CV mới:** Trigger **Workflow "Quản lý công văn công an"** (4 bước chính)
   - **Bước 1:** Agent `khai-thac-du-lieu-ho-so` - Bóc tách danh sách DN, tra cứu DB NNT, phân Tổ KT, ghi CBQL
   - **Bước 2:** Agent `Theo dõi & Phân công` - Rà soát công văn, in phân công KT1 toàn văn, chốt trạng thái
   - **Bước 3:** Agent `soan-thao-tra-loi-cong-an` - Lập dự thảo công văn trả lời + tờ trình (nếu có KT1)
   - **Bước 4:** Agent `quản-lý-hồ-sơ-công-an` - Lưu trữ PDF, tự tạo dự thảo, rà soát hoàn tất
   - Mỗi bước được kiểm soát bởi: `Kiểm soát & Điều phối quy trình` (xuyên suốt)

3. **Báo cáo kết quả**
   - Sau Workflow xong: liệt kê CV đã xử lý đầy đủ (số hiệu, số DN, Tổ KT, CBQL, ngày)
   - Nếu có lỗi tại bước nào → ghi chú lỗi cụ thể, báo cần xem lại tay
   - Báo về chat này để người dùng xác nhận

## Ghi chú
- Nếu CV MOI trống → loop báo "Không có CV mới" và dừng
- Nếu CV không bóc tách được → Kiểm soát & Điều phối sẽ báo cần xem lại file gốc
- Thư mục CV MOI phải có quyền truy cập (Drive offline sync hoặc đường dẫn mạng có hiệu lực)
