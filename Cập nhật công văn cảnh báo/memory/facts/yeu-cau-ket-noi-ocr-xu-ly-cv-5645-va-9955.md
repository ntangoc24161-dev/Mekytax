---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-09
updated: 2026-11-09
---
Đối với các công văn scan chưa có lớp chữ trong hộp thư đến HO SO MOI, người dùng chỉ đạo chỉ kết nối OCR và tiến hành xử lý bóc tách cho hai tệp `5645.pdf` (34 trang) và `cv 9955.pdf` (9 trang).

**Vì sao:** Tệp CV 7268 đã có bản số hóa sạch và file Excel đi kèm nên không cần OCR; các tệp scan cần xử lý OCR trọng tâm là 5645 và cv 9955.
**Áp dụng:** Bật PaddleOCR và chỉ định xử lý OCR cho hai công văn 5645 và 9955, không chạy OCR tràn lan cho các tệp đã có dữ liệu số hóa.
