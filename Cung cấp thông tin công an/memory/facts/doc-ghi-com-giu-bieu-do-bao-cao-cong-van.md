---
type: business
provenance: assistant
origin: javis-learned
created: 2026-11-09
updated: 2026-11-09
---
Khi cần trích số liệu hoặc ghi dữ liệu vào file Bao_cao_cong_van.xlsx (sheet NHAP_HO_SO, có sheet DASHBOARD chứa biểu đồ và công thức liên kết), phải mở và ghi qua COM (win32com Excel), không dùng thư viện đọc/ghi trực tiếp (openpyxl...) vì có thể làm mất hoặc hỏng các biểu đồ nhúng trên sheet DASHBOARD.
