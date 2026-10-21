---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-01
updated: 2026-10-01
source: [[conversations/2026-10-01]]
---
Trong quá trình số hóa và bóc tách tự động tài liệu scan văn bản thuế và công văn cảnh báo rủi ro, mô hình OCR thường xuyên gặp lỗi nhận diện sai lệch các ký tự quang học đối với dữ liệu số [[conversations/2026-10-01]]:
- Số `1` thường bị nhận diện nhầm thành chữ cái `l` (thường), chữ `I` (hoa) hoặc số `7` (ví dụ mã số thuế `0108532031` bị chuyển nhầm thành `010853203l`) (thực tế tính đến 10/2026).
- Đối với các tài liệu scan bảng biểu có đường kẻ mờ hoặc chữ bị vỡ hạt, OCR có thể làm đội thêm dòng trùng lặp và sinh lỗi font chữ tiếng Việt [[conversations/2026-10-01]].
- Biện pháp khắc phục: Bắt buộc áp dụng regex kiểm tra độ dài và tính hợp lệ của MST (chỉ chứa ký số 0-9), đối soát chéo danh bạ doanh nghiệp gốc thay vì tin cậy tuyệt đối vào kết quả OCR thô [[conversations/2026-10-01]].
