---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-19
updated: 2026-11-19
source: chat 2026-11-19 (309793c0)
---
Khi Javis chạy Linux, ghi sổ sống `Bao_cao_cong_van.xlsx` qua Google Drive Composio `googledrive_ralph-urbane` (fileId `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`): tải về, vá ô trong XML của xlsx để giữ chart DASHBOARD, gọi `GOOGLEDRIVE_UPLOAD_UPDATE_FILE`, tải lại kiểm tra mở được và còn biểu đồ. Không `load_workbook` + `save` openpyxl cả workbook. Máy Windows có Excel: ghi bằng COM (win32com).

Bổ sung 19/11/2026: giữ bản trước khi vá (`ho-so-cv-moi-2026-11/Bao_cao_cong_van_truoc_chot_in_kt1_2026-11-19.xlsx`). Nếu người dùng báo Excel không mở được sau lần ghi: đẩy lại bản mở được lên đúng file sống, vá lại đúng ô với ZIP gần Excel hơn, rồi bảo tải từ Drive web bằng Excel desktop (không Google Sheets); bản local Drive Desktop cũ thì đổi tên/xóa để đồng bộ lại.

**Vì sao:** Người dùng 19/11 báo file bị lỗi không mở được ngay sau chốt in O855/O861.
**Áp dụng:** Mọi lần vá XML sổ sống trên Linux.
