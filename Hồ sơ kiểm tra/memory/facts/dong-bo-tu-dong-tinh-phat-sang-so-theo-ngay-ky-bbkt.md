---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
---
Quy chế kiểm tra và đồng bộ số liệu giữa sheet `TINH PHAT + CHAM NOP (C T)` trên Bảng tính và sổ `Danh sach kiem tra quyet toan pro.xlsm`:
- Luôn tự động đồng bộ mỗi lần chạy: cập nhật các trường truy thu TNDN (cột 40), truy thu GTGT (cột 41), chậm nộp GTGT (cột 43), chậm nộp TNDN (cột 44), phạt 20% (cột 46) sang sổ theo dõi.
- Mốc tính phạt chậm nộp 0,03%/ngày lấy theo Ngày ký BBKT trên sổ (cột 36 / AJ) điền vào cột D sheet tính phạt.
- Phạt thủ tục VPHC (cột 47 / AU): Tự động cảnh báo khi có phát sinh (> 0) để người dùng kiểm tra hồ sơ xử phạt riêng.
- Bảo toàn macro VBA, không ghi đè các cột công thức tự tính (AP, AS, AV, AW), và không tạo file `.bak` trong `00. Theo doi chung\`.
