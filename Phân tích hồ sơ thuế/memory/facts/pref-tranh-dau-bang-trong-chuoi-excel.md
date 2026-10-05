---
type: preference
provenance: user
origin: javis-learned
created: 2026-09-30
updated: 2026-09-30
---
Khi xuất dữ liệu văn bản ra file Excel (đặc biệt các đoạn ghi chú, mũi tên như `=>`), không để ký tự đầu tiên là dấu bằng `=` vì các thư viện (như openpyxl, xlsxwriter) sẽ tự nhận diện nhầm thành công thức (Formula tag `<f>`). Khi Excel mở file sẽ báo lỗi hỏng và tự động bóc tách: 'Removed Records: Formula from /xl/worksheets/sheet1.xml part'. Cần thay bằng `>` hoặc ép kiểu string rõ ràng.
