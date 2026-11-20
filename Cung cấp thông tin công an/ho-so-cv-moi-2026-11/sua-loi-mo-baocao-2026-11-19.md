---
type: source
source_kind: own-note
created: 2026-11-19
---

# Sửa lỗi Bao_cao không mở được - 19/11/2026 23:48

## Việc vừa làm
Đóng gói lại sổ sống Drive `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph` từ bản restore đã mở được (`e93f10b62c25`):
- Giữ ZipInfo gốc (create_system=3, external_attr=0x81a40000)
- Giữ nguyên XML sharedStrings cũ, chỉ append 24 chuỗi mới, `count=uniqueCount=2838`, declaration một nháy
- Ghi O855/O861 = Đang xử lý
- 7 DN CV 5681 ở dòng 896-902, CV 1141 ở 903
- 2 chart DASHBOARD còn

Tải lại Drive khớp md5 `688d85824864`. Bản vault: `Bao_cao_cong_van_LIVE_after_safe_rebuild_2026-11-19.xlsx`.

Trước khi mở: **đóng Excel** nếu đang mở file Drive, rồi tải lại / mở bản mới.

## Vì sao mỗi lần VPS cập nhật là hỏng

Hai nguyên nhân chồng nhau, không phải "Excel tự hỏng dữ liệu".

1. **VPS đóng gói lại toàn bộ file .xlsx.** Sổ này do Excel Windows tạo, có chart DASHBOARD, calcChain ~50.000 ô, autofilter. Python `zipfile` trên Linux viết lại mọi thành viên ZIP (create_system, external_attr, nén lại sheet2 3.2MB, đôi khi viết lại cả sharedStrings). Excel Windows coi gói đó là "file format is not valid" hoặc hiện hộp Repair. Lần chốt in KT1 và lần bóc 7 DN đều đi đường này.
2. **Excel đang mở file đồng bộ Drive.** VPS ghi đè `Bao_cao_cong_van.xlsx` trong khi máy anh còn mở. Drive for Desktop + file lock Windows làm bản local vỡ. Đóng Excel rồi restore thì mở lại được, dù nội dung XML chỉ khác vài ô.

Không dùng `openpyxl.save` cả workbook: mất chart. Bản ghi XML tại chỗ (đổi `v` của O855/O861) thì XML vẫn lành; hỏng nằm ở **cách đóng gói ZIP** và **tranh chấp file đang mở**.

## Hướng lần sau (VPS)

1. Anh đóng Excel (và Drive sync xong) trước khi bảo ghi sổ.
2. VPS chỉ vá XML `sheet2.xml` / `sharedStrings.xml`, **copy nguyên ZipInfo** từ file đang mở được (`create_system` + `external_attr` + `flag_bits` + `extra`). Không tạo ZipInfo mới (sys=0, attr=0x20). Không viết lại toàn bộ sharedStrings.
3. `count` và `uniqueCount` của sst luôn bằng số thẻ `<si>`. Giữ declaration gốc (`version='1.0'` một nháy).
4. Sau khi đẩy Drive: tải lại, đối md5, `testzip`, còn `xl/charts/chart1.xml` + `chart2.xml`.
5. Đường bền: một máy Windows chạy Excel COM (`win32com`) ghi NHAP_HO_SO, VPS chỉ gửi dữ liệu dòng. Đó là cách Excel tự giữ chart và ZIP.

FileId sổ: `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`. Tài khoản ghi: `googledrive_ralph-urbane`.
