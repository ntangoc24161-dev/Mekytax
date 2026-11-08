---
type: decision
created: 2026-11-06
updated: 2026-11-06
---
Thiết lập Trợ lý Quản lý Kiểm soát & Điều phối quy trình (`kiem-soat-dieu-phoi-quy-trinh`) đóng vai trò chốt chặn kiểm soát chất lượng (Quality Gatekeeper) và phân phối luồng công việc giữa các trợ lý.

**Vì sao:**
- Tránh trường hợp trợ lý tổng hợp nhầm thông tin công văn (sai số CV, sai MST, nhầm đơn vị hành chính hoặc tự suy diễn số liệu) nhưng không phát hiện ra, dẫn đến sai sót dây chuyền sang khâu in ấn và soạn thảo văn bản.
- Đảm bảo luân chuyển công việc nhịp nhàng: khi một trợ lý hoàn thành nhiệm vụ, kết quả phải được đối chiếu kiểm tra đạt chuẩn trước khi phân phối tiếp cho trợ lý tiếp theo.

**Áp dụng:**
- Áp dụng cơ chế kiểm soát chéo 4 Cổng (Quality Gates):
  1. Cổng 1: Nghiệm thu tiếp nhận và bóc tách công văn (kiểm tra 3 bên: scan gốc <-> file .md <-> sổ NHAP_HO_SO; soát 100% MST và phân Tổ KT đúng địa bàn TCS3).
  2. Cổng 2: Nghiệm thu in ấn phân công KT1 (kiểm tra 100% nội dung, chẵn trang, dòng phân công mép trên, lưu Drive và chỉ đổi trạng thái sau khi chốt in).
  3. Cổng 3: Nghiệm thu dự thảo công văn và tờ trình (kiểm tra kính gửi, chỉ đưa DN thuộc KT1-KT3, trích đúng DB NNT, tô đỏ thông tin chưa xác định).
  4. Cổng 4: Nghiệm thu đóng gói và hoàn tất (ký số, thêm hậu tố -R và đổi trạng thái sổ thành Đã hoàn thành).
- Phân phối công việc: Đạt kiểm tra mới chuyển giao dữ liệu chuẩn cho trợ lý tiếp theo; phát hiện lỗi phải chặn lại và yêu cầu sửa ngay.
