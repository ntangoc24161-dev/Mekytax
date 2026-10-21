---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-09-26
updated: 2026-09-26
---
Khi chạy script bóc tách Ưu tiên 1 và 2 (run_priority_1_and_2.py), việc dùng `ws.cell(row, col)` trên openpyxl với `read_only=True` để nạp 28.163 dòng khiến CPU bị kẹt phân tích lại XML. Đã khắc phục chuyển sang `ws.iter_rows(values_only=True)`, rút ngắn thời gian nạp xuống còn 2,37 giây.
