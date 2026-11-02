---
type: business
provenance: assistant
origin: javis-learned
created: 2026-10-28
updated: 2026-10-28
---
Trên `Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`), cột A (STT) là **công thức**, không ghi số cứng. Sheet có vùng hàng trống/đã định dạng kéo dài tới khoảng hàng 990, nên lấy `max_row` để nối dòng sẽ ghi lệch xuống dưới và người dùng mở file không thấy.

Thực tế tính đến 28/10/2026: folder 271 (9955/CSKT-Đ2) ở hàng 811; folder 272 (12853/CSHS) ở hàng 812; folder 273 (20568/DON-VP) ở hàng 813.

**Vì sao:** ngày 28/10/2026 hai dòng 272-273 bị ghi xuống hàng 990-991 với STT số cứng; người dùng mở sổ vẫn chỉ thấy tới folder 271 và phải báo lại.
**Áp dụng:** tìm dòng cuối thật theo cột Số folder/Số CV có dữ liệu, ghi ngay dòng kế tiếp, chép công thức cột A từ dòng trên.
