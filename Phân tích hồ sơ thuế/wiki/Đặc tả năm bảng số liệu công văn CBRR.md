---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-09-30
updated: 2026-09-30
source: [[conversations/2026-09-30]]
---
Tài liệu kỹ thuật `BANG_SO_LIEU_CONG_VAN.md` (cập nhật 15/09/2026) mô tả cách điền số lên mẫu `Cong_van_CBRR_*.docx`: mẫu có 9 bảng, trong đó **5 bảng được điền số**. [[BANG_SO_LIEU_CONG_VAN.md]]

Code nhận diện bảng số liệu bằng **chữ tiêu đề 3 hàng đầu**, không hardcode index (trừ kính gửi = `doc.tables[2]`). [[BANG_SO_LIEU_CONG_VAN.md]]

### Bản đồ bảng trên mẫu (thực tế theo tài liệu 15/09/2026)
- Index 0-1: Quốc hiệu / số-ngày-trích yếu - không điền số. [[BANG_SO_LIEU_CONG_VAN.md]]
- Index 2: Kính gửi - thể thức theo `KINH_GUI.md`. [[BANG_SO_LIEU_CONG_VAN.md]]
- Index 3: GTGT theo kỳ - điền số. [[BANG_SO_LIEU_CONG_VAN.md]]
- Index 4: TNDN theo kỳ - điền số. [[BANG_SO_LIEU_CONG_VAN.md]]
- Index 5-6: Đối chiếu bán ra / mua vào - điền số. [[BANG_SO_LIEU_CONG_VAN.md]]
- Index 7: Mục 3.4 người bán rủi ro - điền số hoặc **xóa cả bảng**. [[BANG_SO_LIEU_CONG_VAN.md]]
- Index 8: Nơi nhận / chữ ký - không sửa Nơi nhận. [[BANG_SO_LIEU_CONG_VAN.md]]

Hàng mẫu `«KY»` (hoặc STT `1`) được nhân theo số kỳ; dòng **CỘNG** luôn tính lại từ máy, không copy số trên mẫu. [[BANG_SO_LIEU_CONG_VAN.md]]

### Mapping GTGT / TNDN / đối chiếu
- Bảng GTGT (7 cột): kỳ ← `period.display`; gt mua ← ct23; thuế mua ← ct24; tổng DT ← ct34; thuế đầu ra ← ct35; PS dương ← ct40; còn khấu trừ ← ct43/ct41. [[BANG_SO_LIEU_CONG_VAN.md]]
- Dòng CỘNG GTGT chỉ cộng cột 2-5; cột 6-7 để trống. [[BANG_SO_LIEU_CONG_VAN.md]]
- Nguồn GTGT: sheet `GTGT 01` trong `TONG HOP_TO KHAI.xlsx`, lọc đúng MST, ưu tiên tờ khai loại **C** và lần nộp cao hơn. [[BANG_SO_LIEU_CONG_VAN.md]]
- Bảng TNDN (5 cột): STT, Kỳ, Tổng DT bán ra, Thu nhập chịu thuế, Thuế TNDN phải nộp; nguồn sheet `QT TNDN 03`; không có hàng CỘNG (`cong=None`). [[BANG_SO_LIEU_CONG_VAN.md]]
- Đối chiếu bán ra/mua vào (7 cột): tờ khai GT/thuế → HĐĐT GT/thuế → **chênh lệch = tờ khai − HĐĐT**; CỘNG cộng cả 6 cột số. [[BANG_SO_LIEU_CONG_VAN.md]]

### Hóa đơn hợp lệ trên công văn
`hop_le = not _is_cancelled(...)`. [[BANG_SO_LIEU_CONG_VAN.md]]

**Không cộng** trên công văn (vẫn có thể liệt kê trên bảng kê): hủy/huy; xóa bỏ/xoá bỏ; không đủ điều kiện (kể cả không đủ điều kiện cấp mã); bị thay thế. [[BANG_SO_LIEU_CONG_VAN.md]]

**Vẫn cộng:** “Hóa đơn mới”, “Hóa đơn thay thế” (bản mới). [[BANG_SO_LIEU_CONG_VAN.md]]

Hai lớp tổng hợp: công văn chỉ cộng HĐ hợp lệ; bảng kê Excel phụ lục cộng mọi HĐ ở dòng CỘNG nhưng vẫn tách cột hợp lệ. [[BANG_SO_LIEU_CONG_VAN.md]]

Thiếu dữ liệu thì ghi **0 / trống / `[...]`**, không bịa số minh họa. [[BANG_SO_LIEU_CONG_VAN.md]]
