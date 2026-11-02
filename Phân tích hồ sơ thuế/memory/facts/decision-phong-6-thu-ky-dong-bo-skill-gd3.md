---
type: decision
created: 2026-10-23
updated: 2026-10-23
---
Ngày 23/10/2026 người dùng chốt: đồng bộ thư ký theo skill/biểu mẫu mới và **giảm nhân sự** phòng phân tích hồ sơ thuế từ 11 xuống **6 vai**.

**Vì sao:** Skill GĐ2/GĐ3 và mẫu vòng 2 đã gom đủ sheet theo KQ2.1 / KQ2.2 / KQ3; tách 6 chuyên viên module làm phình workflow và dễ lệch mô tả. Gộp theo tệp đầu ra giữ đúng ranh giới nghiệp vụ, dễ điều phối.

**Áp dụng:**
- `thu-ho-so-kiem-ke` - GĐ1
- `chuyen-vien-du-lieu-dau-vao` - chuẩn hóa IN01/IN02/IN05
- `phan-tich-vien-cqt` - toàn bộ KQ2.1 (GTGT, NCC, TNCN, EBITDA, IN07)
- `phan-tich-vien-so-cty` - toàn bộ KQ2.2 (511-632 tổng+chi tiết, kho, giá thành, công nợ)
- `phan-tich-vien-so-sach-dn` - GĐ3 (KQ3 2 sheet + dấu hiệu + giải trình NNT 10 sheet); tên hiển thị "Phân tích viên kết luận GĐ3"
- `truong-phong-phan-tich-ho-so` - điều phối và chốt
- Workflow `phan-tich-ho-so-thue-3-giai-doan` còn **7 bước** (trưởng phòng 2 lần), `status: active`
- Đã xóa 6 agent module: gtgt-tndn, ncc-hoa-don, tncn, doanh-thu-gia-von, kho-gia-thanh, cong-no-tien-mat
- Đã đồng bộ playbook `playbook-3-giai-doan.md`, `playbook-lenh-theo-buoc.md`, `playbook-tat-ca-skill.md` và skill trục / GĐ2; mirror sang `shared-skills/`
