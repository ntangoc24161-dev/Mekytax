---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-18
updated: 2026-11-18
---
Khi Javis chạy Linux, ghi sổ sống `Bao_cao_cong_van.xlsx` qua Google Drive Composio `googledrive_ralph-urbane` (fileId `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`): tải về, vá ô trong XML của xlsx để giữ chart DASHBOARD, gọi `GOOGLEDRIVE_UPLOAD_UPDATE_FILE`, tải lại kiểm tra mở được và còn biểu đồ.

Không `load_workbook` + `save` openpyxl cả workbook. Không kết luận thiếu Drive chỉ vì không thấy ổ `I:\`. Không dùng account mặc định `googledrive_attach-speos` (`ntangoc2416@gmail.com`).

Máy Windows có Excel: vẫn ghi bằng COM (win32com).

**Vì sao:** Người dùng 18/11/2026 lệnh cập nhật qua Drive và bắt kiểm tra file còn mở được sau khi ghi.
**Áp dụng:** Mọi lần đổi cột O trên sổ sống từ phiên Linux đi đường Drive + vá XML + kiểm chứng mở lại.
