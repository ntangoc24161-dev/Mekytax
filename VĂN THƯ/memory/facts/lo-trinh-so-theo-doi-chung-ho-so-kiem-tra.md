---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
source: chat 2026-11-13 (8aa1855e)
---
Người dùng quyết định xây hệ thống theo dõi trước cho **một nhóm kế hoạch kiểm tra** (KHBS-NGOC90). Khi có nhiệm vụ mới, người dùng sẽ tạo thêm **trợ lý theo dõi chung** để **báo cáo trưởng phòng**.

Ngày 2026-11-13, người dùng nói rõ phạm vi khi Javis đề xuất sổ master:
- **KHBS và Kiểm tra quyết toán là CÙNG MỘT công việc**, chỉ chia giai đoạn giao việc. Hai nhóm là các DN khác nhau nhưng cùng nhóm việc, theo dõi theo cùng khuôn của `Danh sach phan cong kiem tra quyet toan.xlsx`. File này đang được hoàn thiện trong brain VĂN THƯ.
- **`So theo doi danh sach Thong bao moi.xlsx`** là danh sách theo dõi chung các DN với nhiều dạng đầu việc khác nhau. Mục đích là gom về một chỗ cho dễ theo dõi và tránh sót việc.
- **Ngừng/đóng mở sử dụng hóa đơn** đã có hồ sơ theo dõi riêng và skill `dong-hoa-don`, nên KHÔNG gộp vào sổ master.

Cấu trúc sổ master Javis đề xuất (Dashboard, các sheet chi tiết, Config, Legend) **người dùng chưa duyệt**.

**Vì sao:** chưa có dữ liệu cho các loại hồ sơ khác, và không muốn dựng một khung tổng quát quá sớm.
**Áp dụng:** không tách sheet KHBS và sheet quyết toán thành hai loại việc riêng, mà gộp thành một nhóm có cột giai đoạn hoặc đợt giao. Không tạo file theo dõi riêng cho KHBS trùng với file phân công. Không đưa phần đóng mở HĐ vào sổ master. Số liệu dashboard phải lấy từ file thật. Chờ người dùng chốt cấu trúc rồi mới dựng file, agent và loop.
