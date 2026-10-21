---
type: wiki
title: Playbook lệnh - Tổng hợp bên mua từ sổ CVCBRR
updated: 2026-10-15
---

# Playbook `tong-hop-ben-mua-tu-so-cvcbrr`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (B2).

Không dùng để đối chiếu NCC bên bán trên bảng kê 1 DN → `doi-chieu-ncc-mua-vao-rui-ro`.

## Cần

File `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` sheet `THEO_DOI`.

## Lệnh

```
Tổng hợp đơn vị bên mua từ sổ CVCBRR
file SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx sheet THEO_DOI.
Xuất Chi_tiet (chỉ dòng đã có bên mua, sắp ngày CV)
và Tong_hop theo MST mua.
Font Times New Roman. Không bịa doanh thu nếu sổ trống cột tiền.
```

**Siêu ngắn:** `Tổng hợp bên mua sổ CVCBRR`

## File ra

`Tong_hop_ben_mua_CVCBRR.xlsx` (sheet `Chi_tiet` + `Tong_hop`).

## Thiếu thì

Cột tiền trên sổ trống: ghi chú, không bịa. Khớp MST khi hỏi 1 công ty cụ thể, không suy từ tên gần giống.
