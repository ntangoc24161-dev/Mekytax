---
type: preference
provenance: user
origin: javis-learned
created: 2026-09-30
updated: 2026-09-30
---
Khi user yêu cầu tổng hợp **đơn vị bên mua** từ `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`:

1. Sheet **`Chi_tiet`:** mỗi dòng sổ đã có tên/MST bên mua, kèm CV + bên bán; sắp theo ngày công văn → MST mua → số CV.
2. Sheet **`Tong_hop`:** gom theo **MST gốc bên mua**; sắp theo số dòng giảm dần.
3. Font **Times New Roman**.

Chỉ lấy dòng **đã có** thông tin bên mua (trên sổ: bên mua chỉ nhập khi công văn nêu). Không suy MST NNT đang kiểm tra có trong cột mua nếu kiểm tra ra 0 dòng.

**Vì sao:** User hỏi riêng tổng hợp bên mua sau khi đã làm đối chiếu bên bán.
**Áp dụng:** Lần xuất `Tong_hop_ben_mua_CVCBRR.xlsx` hoặc file tương đương.
