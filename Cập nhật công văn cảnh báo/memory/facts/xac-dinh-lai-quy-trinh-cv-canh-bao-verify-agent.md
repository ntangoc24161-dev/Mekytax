---
type: decision
provenance: user
created: 2026-11-09
updated: 2026-11-09
source: chat 2026-11-09 (2607e2ee)
---
Ngày 09/11/2026 xác định lại quy trình xử lý công văn cảnh báo sau khi lần chạy nền `t_eaa5d3a95f23` vỡ.

**Nguyên nhân vỡ:** workflow cũ có 10 bước, trong đó QC là các bước THƯỜNG. Agent bước 1 trả về "Đã sẵn sàng" thay vì danh sách tệp; agent bước 2 nhận chuỗi đó làm `{{prev}}` nên viết câu hỏi vào đầu ra; agent bước 3 lại nhận câu hỏi đó làm đầu vào. Dây chuyền không dừng mà chạy tiếp trên dữ liệu rỗng, hết 10 bước trong 4 phút (00:38 đến 00:42) và không ra sản phẩm nào.

**Gốc của vấn đề:** trợ lý chạy nền KHÔNG hỏi lại người dùng được. Câu hỏi của nó trở thành dữ liệu đầu vào của agent kế tiếp.

**Thiết kế mới:**
- `day-chuyen-xu-ly-cv-canh-bao` còn 4 bước làm việc thật: thu thập, số hoá, nạp sổ F0/F1, sắp xếp lưu trữ. `status: active`.
- Đối chiếu NCC và soạn thảo tách sang `tra-loi-ho-so-cong-an`, `status: 'off'`, vì điều kiện khởi động khác hẳn.
- Bỏ QC dạng bước thường, chuyển sang `verify_agent: kiem-soat-chat-luong-du-lieu` gắn vào từng bước. QC trả JSON `{"pass": true|false}`; trượt hết `max_retries` thì workflow DỪNG THẬT.
- Agent QC có thêm mục "Chế độ verify_agent": trả JSON thuần, mặc định `pass: false` khi thiếu bằng chứng, không bao giờ pass vì lời hứa thủ tục, không hỏi lại.

**Ba luật bắt buộc trong mọi task của bước:** (1) thiếu đầu vào thì trả một dòng `DUNG: <lý do>` rồi kết thúc, cấm viết câu hỏi; (2) đầu ra phải là bảng hoặc con số cụ thể, cấm "đã sẵn sàng"; (3) `intent` khi giao việc phải ghi đường dẫn tuyệt đối, không để agent tự đi tìm.

**Vì sao:** bịa tiến độ và chạy rỗng nguy hiểm hơn là dừng sớm, vì sổ F0/F1 là dữ liệu nghiệp vụ thật.

**Áp dụng:** giao cả dây chuyền bằng `javis_task route="wf:day-chuyen-xu-ly-cv-canh-bao"` kèm đường dẫn tuyệt đối trong `intent`. Chỉ cần một khâu thì giao `route="auto"` cho một agent, hoặc bảo Javis làm ngay trong chat.

**Điểm kết thúc công việc (Trưởng phòng chốt 09/11/2026, 01:47).** Một lô hồ sơ coi là XONG khi **nạp sổ F0/F1 và sắp xếp lưu trữ đã qua QC đạt**. Đối chiếu NCC và soạn thảo công văn **chỉ thực hiện khi được yêu cầu rõ bằng một lệnh riêng**; không yêu cầu thì không chạy, không tự đề xuất giữa dây chuyền và **không tính là việc còn dở**.

Đã ghi cứng vào ba chỗ: task bước 4 của `day-chuyen-xu-ly-cv-canh-bao` (cấm tự gọi `tra-loi-ho-so-cong-an`), mục "Chỉ chạy KHI ĐƯỢC YÊU CẦU" trong `tra-loi-ho-so-cong-an`, và phạm vi kiểm tra của agent `kiem-soat-chat-luong-du-lieu` (hai mục đối chiếu/soạn thảo ghi "không thuộc phạm vi lô này" chứ không ghi "chờ"). File `PHAN-CONG-CONG-VIEC.md` cũng đã cập nhật theo.
