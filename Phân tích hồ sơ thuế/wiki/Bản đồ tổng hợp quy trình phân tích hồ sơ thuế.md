---
type: wiki
status: active
tags: [wiki, TTKT, phan-tich-ho-so, quy-trinh, tong-hop]
created: '2026-10-11'
updated: '2026-10-15'
source: '[[phan-tich-ho-so-thue]]'
---
# Bản đồ tổng hợp quy trình phân tích hồ sơ thuế

Trang **chỉ mục vận hành** gom các lớp đã xây: thư mục 4 bước, playbook module, bộ MAU 01→07, skill và chuẩn đầu ra. Chi tiết thao tác nằm ở skill / playbook; trang này chỉ định hướng và liên kết.

Cập nhật: **2026-10-15**. Catalog lệnh mọi skill: [[Playbook lệnh các skill phân tích hồ sơ thuế]] và `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md`.

---

## 1. Ba lớp vận hành

1. **Khung hồ sơ (thư mục 4 bước):** chuẩn hóa chỗ đặt dữ liệu từng DN `<TEN>_<MST>` - skill `to-chuc-ho-so-kiem-tra-thue`; wiki [[Bộ thư mục mẫu hồ sơ kiểm tra thuế và Kịch bản kiểm tra tính đầy đủ]] và [[Quy trình phân tích toàn diện hồ sơ kiểm tra thuế 4 bước và Bộ checklist hồ sơ]].
2. **Playbook chạy module (0→10):** khi chưa đủ 100% dữ liệu, chạy từng phần, một lệnh = một file; thiếu thì ghi thiếu, không để số 0 giả - xem [[Phân tích hồ sơ kiểm tra thuế theo module độc lập]] và `skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md`. Bước 0 (kiểm kê) gọi skill tổ chức hồ sơ.
3. **Bộ file mẫu 01→07:** đánh số theo trình tự làm việc tại `MAU_FILE_TONG_HOP` - xem [[Trình tự thực hiện và bộ file mẫu phân tích hồ sơ kiểm tra thuế]].

Skill khung hồ sơ: `skills/to-chuc-ho-so-kiem-tra-thue/SKILL.md`. Skill chủ phân tích: `skills/phan-tich-ho-so-thue/SKILL.md` (nhóm Tài chính).

---

## 2. Thư mục chuẩn 4 khối

| Khối | Nội dung chính |
|---|---|
| `1. DU LIEU THO NGANH THUE` | XML tờ khai, HĐĐT portal, nghĩa vụ nộp thuế, NLĐ theo QTT TNCN |
| `2. HO SO NNT CUNG CAP` | Sổ kế toán, kho/giá thành, ngân hàng, lao động-TSCĐ, bảng kê công ty |
| `3. DU LIEU BO SUNG & TRA CUU QLT` | CVCBRR/Công an, danh bạ trạng thái NNT, TNCN đa nguồn, giải trình |
| `4. KET QUA PHAN TICH` | Master, checklist, `cau-thanh/`, `sao-luu/`, đề cương |

---

## 3. Playbook module (thứ tự nhanh)

```
0 Kiểm kê → 0b Tổng hợp đầu vào (HĐĐT + Bảng kê + NLĐ)
→ 1 Gộp bảng kê → 1b Làm sạch MST/HĐ trùng (CD1)
→ 2 GTGT/TNDN/HĐĐT → 3 NCC rủi ro → 4 TNCN
→ 5 511-632 → 6 XNT → 7 Giá thành/GV
→ 8 Công nợ/Tiền mặt → 9 Master 28 TC → 10 Đề cương
```

Song song khi đủ đầu vào: **(2)+(3)+(4)** và **(5)+(6)+(8)**.

---

## 4. Bộ MAU 01→07

| Số | File mẫu | Vai trò |
|---|---|---|
| 01 | Checklist tổng hợp hồ sơ | Kiểm kê đủ/thiếu (1 DN) |
| 02 | Tổng hợp đầu vào HĐĐT + Bảng kê + NLĐ | Bóc tách trước phân tích |
| 03 | (Không tách kết quả DN) CD1 BT5-BT9 gắn bảng biểu | BT6→bảng kê MST; BT5→HĐĐT↔BK; BT7→công nợ/TM; BT8/BT9→GTGT mua vào/133 |
| 04 | Danh sách đối tác HĐĐT + Bảng kê | Tra cứu trạng thái ĐKKD |
| 05 | NCC × CVCBRR / trạng thái | Đủ 100% NCC, không rút gọn |
| 06 | Master phân tích hồ sơ | 8 sheet + 28 tiêu chí |
| 07 | Tiến độ đoàn kiểm tra | Theo dõi nhiều DN (khác checklist 1 DN) |

Vị trí mẫu: `thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/` và `TOOL PHAN TICH HO SO THUE\MAU_FILE_TONG_HOP`. Khi áp dụng DN thật: xuất vào `4. KET QUA PHAN TICH`.

---

## 5. Skill liên quan

- `to-chuc-ho-so-kiem-tra-thue` - thư mục 4 bước, phân loại file, 1 checklist A4, gom 1 nơi
- `phan-tich-ho-so-thue` - skill chủ phân tích (Master + CD1 + 28 tiêu chí)
- `doi-chieu-gtgt-tndn-cbrr` - đối chiếu GTGT / TNDN / HĐĐT / bảng kê
- `doi-chieu-ncc-mua-vao-rui-ro` - NCC × F0/F1/Công an/danh bạ
- `doi-chieu-511-632` - doanh thu 511 ↔ giá vốn 632
- `tong-hop-ben-mua-tu-so-cvcbrr` - tổng hợp bên mua từ sổ CVCBRR
- `mo-ho-so-ttkt`, `quan-ly-rui-ro-tuan-thu-thue`, `xu-phat-vphc-thue-hoa-don`

Reference CD1: `skills/phan-tich-ho-so-thue/references/excel-cd1-hv-bt5-bt9.md`.

---

## 6. Chuẩn đầu ra đã chốt

- Master `Phan_tich_ho_so_cong_ty_<TEN>_<MST>.xlsx` đủ **8 sheet**: `Danh_gia_28_tieu_chi_rui_ro`, `Tong_hop_GTGT_TNDN`, `Tong_hop_NCC_CVCBRR`, `Tong_hop_TNCN_va_Da_nguon`, `Tong_hop_XNT_va_Kho`, `Gia_thanh_va_Dinh_muc`, `Doi_chieu_511_632`, `Cong_no_va_Ngan_hang`
- **28/28 tiêu chí**, **7 cột** chuẩn Vạn Phát Thành; không bản rút gọn
- Sheet thiếu nguồn: ghi rõ thiếu; **cấm số 0 giả**
- Số tiền chuẩn VN: hiển thị `1.234.567,89`; Excel `#,##0` / `#,##0.00` - xem [[Quy chuẩn định dạng số tiền trên Excel phục vụ kiểm tra thuế]]

---

## 7. Nguyên tắc vận hành

1. Không chờ đủ mới chạy; module đủ đầu vào thì chạy trước.
2. Một lệnh = một file kết quả.
3. Gộp bảng kê: bỏ dòng `Tổng cộng`, giữ đối tác tên `TỔNG CÔNG TY...`.
4. Checklist 1 DN ≠ file tiến độ đoàn nhiều DN.
5. Lưu một nơi rõ; sau khi chuyển vào hồ sơ DN thì dọn bản nguồn tạm nếu đã chốt.

---

## 8. Liên kết wiki chuyên đề

- [[Bốn Nguồn Dữ Liệu Đối Chiếu Thuế]]
- [[Bóc tách dữ liệu đầu vào HĐĐT Bảng kê và Người lao động trước phân tích]]
- [[Đối chiếu tờ khai GTGT và Quyết toán TNDN chuẩn CBRR gom 1 sheet]]
- [[Đối chiếu NCC mua vào với CVCBRR và danh bạ trạng thái]]
- [[Đối chiếu sổ chi tiết TK 511 với TK 632 bằng Tool_DC_511_632]]
- [[Kỹ thuật kiểm tra hóa đơn trùng và ngưỡng thanh toán tiền mặt trên Excel]]
- [[Đối chiếu mua vào ba nguồn độc lập và xử lý mã số thuế trên Excel]]
- [[Phân định công cụ theo dõi kiểm tra thuế cấp doanh nghiệp và cấp đoàn]]
