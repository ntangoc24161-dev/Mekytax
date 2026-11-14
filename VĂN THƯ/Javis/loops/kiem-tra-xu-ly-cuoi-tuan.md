---
type: loop
name: Kiểm tra & xử lý cuối mỗi thứ 2 tuần
slug: kiem-tra-xu-ly-cuoi-tuan
enabled: false
mode: full
goal: custom
interval_min: 10080
owner_chat: ""
updated: 2026-11-13
---

## Mục tiêu
Mỗi thứ 2 hàng tuần (chu kỳ 7 ngày), kiểm tra danh sách DN không giải trình sau 2 lần gửi (Deadline L2 đã qua ≥ 7 ngày), sau đó gọi Agent Xử lý cuối để tạo phiếu chuyển hoặc ghi đóng HĐ.

## Quy trình mỗi vòng (7 ngày một lần, mỗi thứ 2 lúc 8h sáng)

1. **Mở bảng**: `QL_Thong_bao_Dong_Mo_HD_2026.xlsx`
2. **Lọc điều kiện**:
   - Deadline L2 ≤ 7 ngày trước (hoặc: Ngày gửi L2 ≤ 7 ngày trước)
   - Giải trình L2 = trống hoặc "Không"
   - Hành động = trống (chưa xử lý)
3. **Gọi Agent Xử lý cuối**: Chuyển danh sách DN cần xử lý
   - Agent sẽ tạo phiếu chuyển (hoặc ghi đóng HĐ nếu có lý do)
   - Agent sẽ cập nhật cột Hành động = "Phiếu chuyển" hoặc "Đóng HĐ"
4. **Lưu tài liệu**: Phiếu chuyển lưu vào `exports/Phieu_chuyen_[YYYY-MM-DD]/`
5. **Báo cáo**:
   - Số DN được xử lý
   - Số phiếu chuyển vs. số đóng HĐ
   - Danh sách DN (MST, Tên)
   - Gửi Telegram tóm tắt kết quả
6. **Ghi log**: `Javis/loops/logs/xu-ly-cuoi-[YYYY-MM-DD].log`

## Điểm chú ý
- Loop chạy mỗi 7 ngày (10080 phút), cố định vào thứ 2 lúc 8h
- Nếu có DN quá hạn nhưng vẫn nằm trong "vùng chờ" (vd: mới L2 được 5 ngày) → chưa xử lý, chỉ nhắc nhở
- Phiếu chuyển là "an toàn" (để tổ QL DN xác minh), đóng HĐ cần thêm chứng cứ
- Nếu Agent Xử lý cuối báo lỗi → Ghi log + đề xuất xử lý thủ công

## Ví dụ báo cáo Telegram
```
📋 **XỬ LÝ CUỐI - Tuần [Ngày bắt đầu]**

✅ Phiếu chuyển: 12 DN
   • CÔNG TY TNHH... (MST: 3603143116)
   • ...

🔴 Đóng HĐ: 2 DN
   • DN XYZ (MST: 3603585123) - Spam

📊 Tổng: 14 DN xử lý
Tài liệu: exports/Phieu_chuyen_2026-11-13/
```
