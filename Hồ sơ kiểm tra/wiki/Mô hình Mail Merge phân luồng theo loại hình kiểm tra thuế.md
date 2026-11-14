---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
source: [[conversations/2026-11-13]]
---
Mô hình Mail Merge trong hệ thống kiểm tra quyết toán phân luồng hồ sơ dựa trên cột loại hình [[conversations/2026-11-13]]:
- Nguồn dữ liệu tập trung: Toàn bộ danh sách hồ sơ ban hành lưu chung tại sheet `01_DA_BAN_HANH$` của sổ `Danh sach kiem tra quyet toan pro.xlsm`, liên kết qua Dòng 1 ẩn chứa 58 trường chuẩn hóa [[conversations/2026-11-13]].
- Phân luồng phôi Word: Phân biệt dựa vào cột `Loai_hinh` gồm 2 mã chính: `KH` (Kiểm tra kế hoạch) và `GT` (Kiểm tra giải thể) [[conversations/2026-11-13]].
- Tách biệt phôi theo nghiệp vụ: Tránh việc lồng câu lệnh điều kiện `IF/THEN/ELSE` phức tạp trong tài liệu Word, hệ thống duy trì 2 phôi riêng gồm `templates/MAUTONG HOP.docx` cho hồ sơ kế hoạch và `templates/MAUTONG HOP GIAI THE.docx` cho hồ sơ giải thể [[conversations/2026-11-13]].
- Quy trình xử lý: Sau bước Mail Merge Tầng 1, hệ thống định tuyến gọi skill chuyên trách tương ứng để tiếp tục đổ dữ liệu chi tiết Tầng 2 từ `Bang tinh.xlsx` (thực tế tính đến 13/11/2026) [[conversations/2026-11-13]].
