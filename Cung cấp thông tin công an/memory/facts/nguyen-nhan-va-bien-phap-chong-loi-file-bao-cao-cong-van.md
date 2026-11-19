---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
source: chat 2026-11-18 (3e4326e5)
---
- **Hiện trạng & Nguyên nhân:** File Bao_cao_cong_van.xlsx dễ hỏng cấu trúc gói XML (lệch calcChain.xml, sinh dòng rác cuối bảng làm lệch AutoFilter, hỏng liên kết 2 biểu đồ DASHBOARD) khi tool/script ghi trực tiếp dạng nén XML thuần hoặc xung đột đồng bộ Google Drive khi đang mở.
- **Quy tắc bắt buộc:**
  1. Chỉ đọc/ghi file qua Microsoft Excel COM (win32com) chính thức.
  2. Tuyệt đối không kéo tràn công thức làm sinh dòng rác ở đuôi; giữ chuẩn phạm vi AutoFilter.
  3. Tự động tạo bản sao lưu trước khi can thiệp ghi dữ liệu.
  4. Tránh mở file trên máy tính vào khung giờ tiến trình quét chạy (12:00 hàng ngày) để tránh file lock.
