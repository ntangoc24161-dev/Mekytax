---
type: decision
created: 2026-11-19
updated: 2026-11-20
source: chat 2026-11-19 (38b144d6); chat 2026-11-20 (mwnu8fh5tv6n99vy)
---
Khi Javis VPS cập nhật `Bao_cao_cong_van.xlsx` trên Drive, **bắt buộc** chạy `scripts/ghi-so-bao-cao-an-toan.py` (`validate` / `set-status` / `set-cells` / `repair`). Không đóng gói ZIP bằng ZipInfo mới (create_system=0 / attr=0x20) và không `openpyxl.save` cả workbook. Script vá XML sheet2/sharedStrings, giữ ZipInfo Windows (`create_system=3`, `external_attr=0x81a40000`), giữ `count=uniqueCount=số <si>`, đủ namespace sheet2. Anh đóng Excel trước khi VPS ghi đè (cửa kiểm `scripts/cua-kiem-dong-bao-cao.py`). Sau `validate` ok mới `GOOGLEDRIVE_UPLOAD_UPDATE_FILE` đúng fileId sống.

**Vì sao:** Excel Windows từ chối gói ZIP Linux viết lại; Drive sync khi Excel đang mở làm bản local vỡ. Chart DASHBOARD mất nếu save openpyxl. Ad-hoc Python từng làm lệch SST/namespace.

**Áp dụng:** mọi lần ghi sổ sống `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`. Chi tiết `ho-so-cv-moi-2026-11/sua-loi-mo-baocao-2026-11-19.md`, `sua-loi-mo-baocao-2026-11-20.md`.
