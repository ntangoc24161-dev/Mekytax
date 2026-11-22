---
type: preference
provenance: user
origin: javis-learned
created: 2026-11-20
updated: 2026-11-20
source: chat 2026-11-20 (mwnu8fh5tv6n99vy)
---
Anh chỉ giữ **một file duy nhất** `Bao_cao_cong_van.xlsx` trên Drive (fileId `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`, account `googledrive_ralph-urbane`). Mỗi lần cập nhật phải **ghi đè** file này, không tạo bản mới kiểu `Bao_cao_cong_van_final_fixed_...xlsx` trên Drive.

**Vì sao:** Nhiều bản dated làm anh mở nhầm file hỏng / lệch dữ liệu.

**Áp dụng:** VPS chỉ `GOOGLEDRIVE_UPLOAD_UPDATE_FILE` vào đúng fileId trên. Bản vault local có thể giữ bản sao kỹ thuật để vá XML, nhưng không upload tên mới lên Drive. Không tự tạo thêm file `Bao_cao_cong_van_*.xlsx` trên Drive.
