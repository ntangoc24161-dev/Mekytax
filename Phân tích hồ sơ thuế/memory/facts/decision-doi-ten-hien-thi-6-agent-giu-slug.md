---
type: decision
created: 2026-11-08
updated: 2026-11-08
---
Ngày 08/11/2026 người dùng yêu cầu đổi tên 6 agent của brain Phân tích hồ sơ thuế cho ngắn và sát việc. Chỉ đổi field `name` (tên hiển thị), **giữ nguyên toàn bộ slug**.

Bộ tên mới: `thu-ho-so-kiem-ke` = Kiểm kê hồ sơ, `chuyen-vien-du-lieu-dau-vao` = Chuẩn hóa dữ liệu, `phan-tich-vien-cqt` = Đối chiếu dữ liệu thuế, `phan-tich-vien-so-cty` = Soát sổ sách công ty, `phan-tich-vien-so-sach-dn` = Kết luận rủi ro, `truong-phong-phan-tich-ho-so` = Trưởng phòng điều phối.

**Vì sao:** slug được tham chiếu ở `workflows/phan-tich-ho-so-thue-3-giai-doan.md`, `skills/phan-tich-ho-so-thue/SKILL.md` (và bản mirror `.claude/skills/`), `Javis/index.md` và các memory fact. Đổi slug làm workflow gọi vào vai không tồn tại; riêng `phan-tich-vien-so-sach-dn` đã ghi trong thân file là slug khóa.

**Áp dụng:** lần sau muốn đổi tên vai thì sửa `name` và câu mở đầu "Bạn là ..." trong file agent, cộng dòng liệt kê 6 vai ở workflow và ở agent Trưởng phòng. Không chạm vào slug, không đổi tên file .md.
