---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-17
updated: 2026-11-17
---
Vụ Kiến Thành Tân (MST 3603216935), Bảng tính tại 0. BIEN BAN:
- Phần thuế GTGT trên sheet DTCHEP: cột Số báo cáo link sang dòng tổng từng năm của sheet GTGT. Chênh lệch GTGT phải nộp = tổng thuế GTGT từng năm trên sheet BANG KE KO HOP LE (thuế GTGT tại đây là phần chênh lệch phải nộp).
- Thuế TNDN giữ nguyên số báo cáo (theo KQKDSXKD), không trừ phần hóa đơn không hợp lệ, nên không phát sinh giảm lỗ.
- Hóa đơn có thuế GTGT = 0 tách sang sheet riêng '1-xem xét không loại'.
- Nguồn hóa đơn loại: sheet 'BK LOẠI HOÁ ĐƠN' của file BẢNG KÊ GTGT KIẾN THÀNH TÂN 2014-2025.xlsx (không dùng sheet 'BK LOẠI HOÁ ĐƠN (2)').
- Bỏ qua sheet BB khi rà số liệu.

**Vì sao:** người dùng chỉ đạo trực tiếp ngày 2026-11-17.
**Áp dụng:** khi sửa DTCHEP, TINH PHAT hoặc lập biên bản vụ này, chỉ truy thu GTGT; không đụng TNDN trừ khi người dùng đổi ý.
