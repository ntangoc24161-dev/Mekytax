---
type: decision
created: 2026-10-16
updated: 2026-10-16
---
Ngày 16/10/2026 người dùng chốt: **giữ file** `Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx` trong `cau-thanh/`. Không xuất `03_Lam_sach_CD1_BT5_BT9` cho từng DN. Không chạy hàng loạt 13 DN. Chạy thử T&T Biên Hòa (MST 3603883505).

**Vì sao:** Hai quyết định 07/10 (giữ mẫu 03) và 11/10 (không xuất 03 tách) đang treo. Anh chọn file chuyên đề BT5-BT9, không chọn mẫu 03 làm kết quả DN.

**Áp dụng:**
- Agent `chuyen-vien-du-lieu-dau-vao` xuất `Doi_chieu_BT5_BT9_<MST>_<KY>.xlsx`.
- Mẫu `03_Lam_sach_CD1_BT5_BT9_MAU.xlsx` chỉ là khung, không copy thành kết quả từng DN.
- Gắn kỹ thuật CD1 vào bảng kê / GTGT / công nợ như quyết định 11/10, kèm file chuyên đề BT5-BT9.
