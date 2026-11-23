---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-23
updated: 2026-11-23
source: [[conversations/2026-11-23]]
---
Mô hình cập nhật dữ liệu 2 pha (Staging/Sổ tạm -> Kiểm soát -> Sổ chính) nhằm bảo đảm an toàn dữ liệu nghiệp vụ [[conversations/2026-11-23]]:
1. **Ghi sổ tạm (Staging):** Dữ liệu sau bóc tách được đổ vào tệp/bảng tạm trước thay vì ghi trực tiếp vào cơ sở dữ liệu hoặc sổ theo dõi chính [[conversations/2026-11-23]].
2. **Kiểm tra độc lập (QC Gate):** Agent kiểm soát độc lập rà soát tính toàn vẹn (định dạng khóa, trường bắt buộc, kiểm tra trùng lặp) [[conversations/2026-11-23]].
3. **Cập nhật đồng loạt (Batch Commit):** Khi kiểm soát đạt chuẩn, hệ thống tự động tạo bản sao lưu (.bak) và cập nhật đồng loạt vào sổ chính, sau đó giải phóng/xóa file tạm [[conversations/2026-11-23]].
4. **Cơ chế Rollback & Retry:** Nếu phát hiện sai lệch, dừng thao tác ghi vào sổ chính, cô lập các bản ghi lỗi và kích hoạt quy trình xử lý lại cho tệp nguồn tương ứng [[conversations/2026-11-23]].
