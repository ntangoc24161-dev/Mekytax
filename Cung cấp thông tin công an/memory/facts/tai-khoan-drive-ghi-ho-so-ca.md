---
type: business
created: 2026-11-12
updated: 2026-11-12
source: chat 2026-11-12 (7a0306fa)
---
Trên Composio, thư mục HS CUNG CAP CA 2026 chỉ ghi được bằng tài khoản `googledrive_ralph-urbane` (cungcapthongtin.tcs3tdn). Tài khoản `googledrive_attach-speos` đọc được nhưng bị 403 khi tạo thư mục. Để thay file sổ Bao_cao_cong_van.xlsx: ghi dòng trong Composio workbench, gọi `upload_local_file` lấy s3key, rồi gọi `GOOGLEDRIVE_UPLOAD_UPDATE_FILE` (fileId 1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph).
**Vì sao:** phiên 12/11/2026 bị 403 khi dùng tài khoản mặc định.
**Áp dụng:** mọi thao tác tạo, chuyển hay ghi trên Drive hồ sơ CA đều dùng `account: googledrive_ralph-urbane`.
