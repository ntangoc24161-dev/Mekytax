---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-04
updated: 2026-10-04
---
Khi bóc tách danh sách đối tác từ HĐĐT và bảng kê thuế (`excel_hd.py`), không chỉ quét từ khóa `mã số thuế` mà cần hỗ trợ regex nhận diện các biến thể tiêu đề thực tế như `mst người bán/mst người xuất hàng` và `mst người mua/mst người nhận hàng`. Nếu không bao quát, trường `mst_dt` bị rỗng và hệ thống có thể lấy nhầm tên công ty làm mã định danh khóa chính.
