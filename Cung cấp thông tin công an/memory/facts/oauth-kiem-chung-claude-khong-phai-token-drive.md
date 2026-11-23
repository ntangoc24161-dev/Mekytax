---
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-19
updated: 2026-11-19
source: chat 2026-11-22 (7577ea4a)
---
Ngày 19/11/2026, kiểm chứng Bước 1 (agent khai-thac-du-lieu-ho-so) báo `Failed to authenticate: OAuth session expired and could not be refreshed` sau 3 lượt. Lỗi thuộc **OAuth engine Claude của lượt kiểm chứng**, không phải token Google Drive.

**Vì sao:** Cùng lúc hub Bearer + Composio xác nhận `googledrive_ralph-urbane` ACTIVE (mail `cungcapthongtin.tcs3tdn@gmail.com`); `GOOGLEDRIVE_FIND_FILE` success 4/4; sổ sống dòng 896-898 đúng; CV MOI 0 file; folder 295/296 đủ 4 file mỗi folder; DASHBOARD còn 2 chart.

**Áp dụng:** Khi thấy đúng câu lỗi này lúc kiểm chứng, xác minh Drive bằng hub Bearer + Composio (`account` sibling của `tool_slug`); không kết luận token Drive hết hạn và không dừng nhánh A chỉ vì phản hồi OAuth của engine kiểm chứng.
