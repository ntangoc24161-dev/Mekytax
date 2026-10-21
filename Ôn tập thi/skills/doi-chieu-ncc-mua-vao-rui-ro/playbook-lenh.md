---
type: wiki
title: Playbook lệnh - Đối chiếu NCC mua vào CVCBRR
updated: 2026-10-15
---

# Playbook `doi-chieu-ncc-mua-vao-rui-ro`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A3). Playbook trục Bước 3.

## Cần

Bảng kê mua và/hoặc HĐĐT mua + sổ CVCBRR (F0/F1). Công an + danh bạ nếu có.

## Lệnh

```
Đối chiếu NCC mua vào với CVCBRR và danh bạ cho <FOLDER>
Phần 1: NNT <MST> với sheet F1 và Bao_cao_cong_van.xlsx, ghi số CV + nguồn trích.
Phần 2: người bán sheet mua vao với sheet F0.
Xuất Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx đủ chuẩn Vạn Phát Thành:
- Sheet Tong_hop_doi_chieu: Phần 1 + Bảng 2.1 F0 + Bảng 2.2 ĐỦ 100% NCC + Phần 3 nhận xét
- Sheet Chi_tiet_hoa_don: đủ cột chuẩn, gắn cờ F0/ĐKKD
- Cột Nguồn dữ liệu = Bảng kê / HĐĐT / Chung; có Bảng 2.3 theo kỳ.
Không rút gọn chỉ lấy NCC bất thường.
```

**Siêu ngắn:** `Đối chiếu NCC F0/F1/CA chuẩn VPT <FOLDER>`

## File ra

`Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx`. Nhúng Master `Tong_hop_NCC_CVCBRR`.

## Thiếu thì

Thiếu danh bạ: vẫn đối F0/F1/CA, ghi thiếu ĐKKD. Không nhầm với `tong-hop-ben-mua-tu-so-cvcbrr`.
