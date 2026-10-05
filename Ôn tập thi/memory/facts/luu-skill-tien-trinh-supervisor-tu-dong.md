---
type: decision
created: 2026-09-26
updated: 2026-09-26
---
Đã đóng gói kiến trúc bộ giám sát tự phục hồi và điều phối thu thập đề thi chạy nền thành skill chuẩn `tien-trinh-supervisor-tu-dong` (Tiến trình supervisor tự động). Kỹ năng bao gồm cơ chế kiểm kê tiến độ thực tế bằng đếm dòng Excel (`ws.max_row - 1`), quản lý vòng đời subprocess `collector_40_de.py`, tự động bắt lỗi gián đoạn mạng và tái kích hoạt worker sau 3 giây cho đến khi hoàn thành 100% mục tiêu 40 đề thi. Skill đã được cập nhật đồng bộ vào brain hiện tại (`skills/`), kho kỹ năng dùng chung (`shared-skills/`) và `.claude/skills/`.

**Vì sao:** Chuẩn hóa quy trình điều phối và giám sát tự động hóa giải đề khảo thí thành kỹ năng tái sử dụng độc lập cho mọi kíp thi và đợt ôn luyện khác mà không lo ngại sự cố rớt mạng.

**Áp dụng:** Tự động kích hoạt khi người dùng yêu cầu chạy supervisor tự động, chạy liên tục đến khi đủ N đề, tự động khôi phục khi gián đoạn hoặc kiểm kê tiến độ thực tế từ file bảng tính.
