---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
---
Người dùng yêu cầu bố cục sheet `Tong_hop_GTGT_TNDN` trong `Phan_tich_ho_so_cong_ty_MAU.xlsx` tách rõ nguồn dữ liệu để đối chiếu tìm sai sót.

**Bốn nguồn độc lập (không gộp chung một cột):**
1. **HĐĐT** - tổng hợp hóa đơn lấy từ dữ liệu cơ quan thuế.
2. **Bảng kê mua vào / bán ra** - do công ty cung cấp, kê khai theo kỳ.
3. **Tờ khai thuế** - do công ty kê khai trên hệ thống cơ quan thuế (GTGT [23] giá trị MV chưa thuế, [24] thuế GTGT MV; bán ra [34][35]).
4. **Sổ kế toán (NKC / Sổ cái)** - số liệu tài khoản ví dụ TK 133 do công ty cung cấp.

**Bố cục sheet:**
- Bỏ cột phụ `DT bảng kê` / `DT hóa đơn điện tử` bên phải PHẦN I.
- PHẦN II (trái) và PHẦN III (phải) nằm ngang, chỉ một khối mỗi phần.
- Mỗi nguồn một cột riêng + cột chênh lệch (CL) giữa các cặp nguồn để tìm lệch.
- PHẦN II: bán ra HĐĐT ↔ Bảng kê ↔ TK GTGT [34][35] ↔ TNDN/BCTC.
- PHẦN III: mua vào giá trị + thuế theo 3 nguồn trên, rồi đối thêm PS Nợ TK 133.

**Vì sao:** Gộp HĐĐT với Bảng kê làm mất dấu vết lệch giữa dữ liệu CQ thuế và dữ liệu công ty cung cấp.
