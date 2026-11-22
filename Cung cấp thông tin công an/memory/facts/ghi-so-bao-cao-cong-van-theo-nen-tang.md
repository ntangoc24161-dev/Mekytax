---
superseded_by: ghi-so-drive-khoi-phuc-khi-excel-khong-mo
type: decision
provenance: user
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
source: chat 2026-11-19 (309793c0)
---
Khi Javis chạy Linux, ghi sổ sống `Bao_cao_cong_van.xlsx` qua Google Drive Composio `googledrive_ralph-urbane` (fileId `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`): tải về → **`scripts/ghi-so-bao-cao-an-toan.py`** (`set-status` / `set-cells` + `validate` ok) → `GOOGLEDRIVE_UPLOAD_FILE` → tải lại `validate`.

Không `load_workbook` + `save` openpyxl cả workbook. Không kết luận thiếu Drive chỉ vì không thấy ổ `I:\`. Không dùng account mặc định `googledrive_attach-speos` (`ntangoc2416@gmail.com`).

Máy Windows có Excel: vẫn ghi bằng COM (win32com).

**Vì sao:** Người dùng 18/11/2026 lệnh cập nhật qua Drive và bắt kiểm tra file còn mở được sau khi ghi; 20/11 khóa bằng script để khỏi vá XML ad-hoc.
**Áp dụng:** Mọi lần đổi cột O / ghi ô chữ trên sổ sống từ VPS phải qua `ghi-so-bao-cao-an-toan.py`.


## Lịch sử
- [2026-11-19] Bị thay thế bởi [[ghi-so-drive-khoi-phuc-khi-excel-khong-mo]] (user đổi thông tin).
