---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-23
updated: 2026-11-23
---
User chỉ đạo quy trình kiểm soát 2 nhánh:
- Giao agent Kiểm soát chất lượng dữ liệu rà soát file tạm trước khi nạp vào Sổ theo dõi chính.
- **Nếu đúng (Pass):** Tự động cập nhật đồng loạt dữ liệu vào Sổ theo dõi chính, sau đó xóa file tạm.
- **Nếu sai (Fail):** Dừng nạp sổ chính và kích hoạt thực hiện lại đối với dữ liệu/tệp sai sót.
