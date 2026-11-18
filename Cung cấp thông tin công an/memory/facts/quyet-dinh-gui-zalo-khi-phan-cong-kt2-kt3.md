---
type: decision
created: 2026-11-18
updated: 2026-11-18
source: chat 2026-11-18 (mwlp8ddh)
---
Khi phân công công văn cung cấp hồ sơ Công an cho Tổ **KT2** hoặc **KT3**, gửi công văn đó qua Zalo cho đầu mối tổ tương ứng (skill `gui-phan-cong-zalo-kt23`). Tổ **KT1** vẫn đi đường in PDF (`in-phan-cong-kt1`).

**Vì sao:** Người dùng yêu cầu ngày 18/11/2026 để đầu mối KT2/KT3 nhận file ngay trên Zalo thay vì chỉ cập nhật sổ.

**Áp dụng:** Sau khi gán Tổ KT2/KT3 trên sổ (và khi người dùng lệnh gửi / chốt phân công), trợ lý `Theo dõi & Phân công công văn` chạy bước gửi Zalo theo đầu mối đã lưu.

**Quy tắc bổ sung (18/11/2026):**
1. Tin Zalo chỉ gồm: tên công ty, số CV + ngày, hạn thực hiện. Bỏ folder, tình trạng, câu gửi scan, dòng phân tổ.
2. Mọi phân công (Zalo KT2/KT3 và in KT1) phải **xem trước qua Telegram** (`POST /telegram/send-file` kèm caption nháp), rồi người dùng chốt OK mới gửi Zalo / chốt in. Không gửi thẳng đầu mối Zalo khi mới chỉ có xem trước trong khung chat web.
3. Sau khi gửi đầu mối / chốt in thành công, cập nhật cột Tình trạng trên `Bao_cao_cong_van.xlsx` sang **Đang xử lý**.
4. (15:27 18/11/2026) Người dùng chốt lựa chọn "OK quy trình lần sau": giữ luồng Telegram xem trước → OK → mới Zalo KT2/KT3; **không** gửi lại các CV đã bàn giao (1806, 2126, 9372).
