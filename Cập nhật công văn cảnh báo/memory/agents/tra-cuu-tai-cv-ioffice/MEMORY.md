## Nghiệp vụ hiện tại

**Công việc 1: Thu thập & xử lý công văn từ Telegram/Zalo**

- **Nguồn tiếp nhận:** Telegram/Zalo — Javis Core sẽ chuyển dữ liệu trực tiếp cho tôi (không thông qua HO SO MOI)
- **Tệp đầu vào:** PDF/ảnh scan công văn cảnh báo (thường là hỗn hợp)
- **Xử lý:** Đổi tên chuẩn `[Số đến]_<Số ký hiệu>_<tên gốc>`, lưu vào `D:\HO SO CONG VAN CANH BAO\2026\2026-MM`
- **Bàn giao:** Danh sách (số CV, ngày, trích yếu, đường dẫn) → công việc 2 (OCR + nạp sổ)
- **Cấm:** Không OCR, không nạp sổ, không bóc tách tệp

**Quy trình:**
1. Nhận tệp từ Javis Core (từ Telegram/Zalo)
2. Đổi tên tệp theo chuẩn
3. Lưu vào thư mục chuẩn (năm/năm-tháng)
4. Báo danh sách (bảng ngắn: Số CV, ngày, trích yếu, đường dẫn)
5. Bàn giao công việc 2


## Bài học (tự học)
- Trong brain Cập nhật công văn cảnh báo, `ioffice_client.py` bản gốc thiếu `showDivTimKiem()` nên ô lọc ngày không hiển thị và Playwright timeout; dùng bản `ioffice_client.conflict-remote-20261022-155419.py` mới tra được theo khoảng ngày.
- Script iOffice điền khoảng ngày vào cả bộ lọc ngày ban hành và ngày văn bản đến (điều kiện AND), nên khoảng hẹp dễ trả 0 dòng dù thực tế có văn bản.
- Khi không có shell, dùng Grep (-o, multiline) trỏ thẳng vào từng tệp PDF, vì quét cả thư mục sẽ bỏ qua tệp nhị phân. Dung lượng xấp xỉ lấy từ `startxref` cuối tệp. Text hay scan xét theo /FontFile, /Length1 hay /Subtype/Image trong phần gốc trước lần ký số đầu tiên.
