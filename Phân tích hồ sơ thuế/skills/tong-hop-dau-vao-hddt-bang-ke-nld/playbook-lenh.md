---
type: wiki
title: Playbook lệnh - Tổng hợp đầu vào HĐĐT bảng kê NLĐ
updated: 2026-10-15
---

# Playbook `tong-hop-dau-vao-hddt-bang-ke-nld`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (A0b). Playbook trục Bước 0b + 1.

## Cần

Folder `1.2` (HĐĐT portal) và/hoặc `2.5` (bảng kê NNT). XML TNCN nếu làm sheet NLĐ.

## Lệnh mẫu 02

```
Tổng hợp dữ liệu đầu vào trước phân tích cho <FOLDER>
MST <MST> kỳ <KY> theo mẫu 02_Tong_hop_dau_vao_HDDT_BangKe_NLD_MAU.xlsx:
1) HĐĐT CQ thuế → 2 sheet HDDT_Mua_vao, HDDT_Ban_ra (lọc HĐ hủy/thay thế).
2) Bảng kê công ty → 2 sheet Bang_ke_Mua_vao, Bang_ke_Ban_ra (bỏ dòng Tổng cộng).
3) NLĐ qua các năm → 1 sheet Tong_hop_NLD; ghi nguồn 05-1 hoặc 05-2.
   Nếu nguồn 05-2 mà công ty không trích 10% → Đề xuất kiểm tra cam kết = Có.
4) Điền Tom_tat_kiem_tra (Đủ/Thiếu/Mỏng).
Lưu Tong_hop_dau_vao_HDDT_BangKe_NLD_<MST>_<KY>.xlsx vào cau-thanh.
Không trộn HĐĐT với Bảng kê chung một sheet chi tiết.
```

## Lệnh gộp bảng kê

```
Tổng hợp bảng kê mua vào bán ra của <FOLDER>
lưu tại 2. HO SO NNT CUNG CAP\2.5. Bang ke mua vao ban ra\
thành Bang_ke_mua_vao_ban_ra_tong_hop_<KY>.xlsx
2 sheet: mua vao, ban ra.
Lọc bỏ dòng Tổng cộng, giữ đối tác tên TỔNG CÔNG TY.
Map cột theo header, khử trùng kỳ bổ sung (lấy bản BS mới nhất).
Sao lưu thêm vào 4. KET QUA PHAN TICH\cau-thanh\
```

**Siêu ngắn:** `Tổng hợp đầu vào HĐĐT-Bảng kê-NLĐ <FOLDER>` hoặc `Gộp bảng kê <FOLDER> vào 2.5`

## File ra

`Tong_hop_dau_vao_HDDT_BangKe_NLD_<MST>_<KY>.xlsx` và/hoặc `Bang_ke_mua_vao_ban_ra_tong_hop_<KY>.xlsx`.

## Thiếu thì

Vẫn xuất sheet có dữ liệu. Tom_tat ghi Thiếu/Mỏng. Bước tiếp: `lam-sach-cd1-mst-hoa-don`.
