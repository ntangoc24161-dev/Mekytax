---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-17
updated: 2026-11-17
---
- Ngày 17/11/2026, file `Bao_cao_cong_van.xlsx` gặp sự cố mở ra thấy trống: nguyên nhân do bị lưu qua công cụ ngoài Excel làm mất toàn bộ 235 công thức Dashboard (các cột ẩn AM-AZ) và dính AutoFilter lọc cột Tình trạng = trống, khiến con trỏ nhảy xuống cuối bảng.
- Dữ liệu thô 851 MST (dòng cuối MST 0315697994 - Cty TNHH XD Hưng Long Thịnh) vẫn còn nguyên.
- Cách xử lý: sao lưu file hỏng tại `D:\HSO CONG AN FILE MD\Bao_cao_cong_van_truoc_sua_20261117_105309.xlsx`, khôi phục công thức và cột phụ từ bản backup 11/11 qua Excel COM trên ổ D rồi chép đè lại Drive `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx`.
- Tuyệt đối không mở/lưu file này bằng Google Sheets hoặc thư viện ghi xlsx không bảo toàn công thức/biểu đồ.
