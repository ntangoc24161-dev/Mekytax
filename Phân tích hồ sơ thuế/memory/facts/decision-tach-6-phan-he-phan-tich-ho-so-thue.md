---
type: decision
created: 2026-10-15
updated: 2026-10-15
---
Tách các phân hệ phân tích còn lại khỏi thân `phan-tich-ho-so-thue` thành 6 skill riêng tại brain này; gắn skill đã có `doi-chieu-gtgt-tndn-cbrr` vào Bước 4.1.

**Vì sao:** Skill trục quá dài, router dễ nhầm module. User yêu cầu gắn các skill đối chiếu đã có và bổ sung skill còn lại tạo tại brain này.

**Áp dụng:**
- Tổ chức folder + checklist A4 → `to-chuc-ho-so-kiem-tra-thue`
- Gộp HĐĐT / bảng kê / NLĐ (mẫu 02) → `tong-hop-dau-vao-hddt-bang-ke-nld`
- Làm sạch MST, HĐ trùng, cụm cùng ngày, BKMV-NKC 133 → `lam-sach-cd1-mst-hoa-don`
- GTGT / TNDN / HĐĐT chuẩn CBRR → `doi-chieu-gtgt-tndn-cbrr`
- NCC vs CVCBRR / Công an / danh bạ → `doi-chieu-ncc-mua-vao-rui-ro`
- QTT TNCN + đa nguồn → `tong-hop-tncn-da-nguon` (chính sách: skill chung `thue-tncn`)
- XNT kho TC3 → `phan-tich-xnt-kho`
- Giá thành / GV thương mại TC28 → `phan-tich-gia-thanh-dinh-muc`
- 511 vs 632 → `doi-chieu-511-632`
- Công nợ / tiền mặt / sao kê → `phan-tich-cong-no-tien-mat`
- Master 8 sheet, 28 tiêu chí, đề cương vẫn nằm trong `phan-tich-ho-so-thue`
- Mirror skill chung tại `shared-skills/` và `.claude/skills`; slug local brain thắng khi trùng.
