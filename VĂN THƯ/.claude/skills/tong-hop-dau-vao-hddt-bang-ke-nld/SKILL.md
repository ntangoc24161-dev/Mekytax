---
name: Tổng hợp đầu vào HĐĐT, bảng kê và NLĐ
description: "Gộp HĐĐT portal, bảng kê NNT (lọc Tổng cộng) và NLĐ QTT TNCN thành file mẫu 02, không trộn nguồn."
group: Tài chính
updated: '2026-10-15'
---
# Tổng hợp đầu vào HĐĐT, bảng kê và NLĐ

Gộp **3 nguồn độc lập** trước mọi đối chiếu: HĐĐT cổng CQT, bảng kê phần mềm kế toán NNT, danh sách NLĐ theo QTT TNCN. Xuất 1 file theo mẫu 02. **Không** trộn HĐĐT với bảng kê chung một sheet chi tiết.

Skill trục: `phan-tich-ho-so-thue` (playbook Bước 0b + Bước 1). Folder lộn xộn thì chạy `to-chuc-ho-so-kiem-tra-thue` trước. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "tổng hợp đầu vào", "gộp bảng kê", "mẫu 02", "HĐĐT và bảng kê", "danh sách NLĐ 05-1".
- Công ty mới, chưa có file `Tong_hop_dau_vao_HDDT_BangKe_NLD_*.xlsx` hoặc `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx`.

Không dùng khi đã có file gộp sạch và user muốn đối chiếu số: chuyển `doi-chieu-gtgt-tndn-cbrr` / `lam-sach-cd1-mst-hoa-don`.

## Chuẩn bị

1. Folder chuẩn `<TEN>_<MST>/`.
2. `1.2. Hoa don dien tu/` = HĐĐT portal (`HDDT.xlsx` hoặc file mua/bán cổng CQT).
3. `2.5. Bang ke mua vao ban ra/` = bảng kê NNT (file rời theo kỳ hoặc đã gộp).
4. XML QTT TNCN / phụ lục 05-1, 05-2 tại `1.1. To khai XML/` nếu làm sheet NLĐ.
5. Mẫu: `thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/02_Tong_hop_dau_vao_HDDT_BangKe_NLD_MAU.xlsx`.

## Quy trình

### 1. HĐĐT cổng CQT

- 2 sheet: `HDDT_Mua_vao`, `HDDT_Ban_ra`.
- Lọc HĐ hủy / thay thế / xóa bỏ (từ khóa: thay thế, xóa bỏ, hủy, bị thay thế).
- Map cột theo **header**, không theo số thứ tự cột.

### 2. Bảng kê NNT

- Gộp các kỳ thành 2 sheet `Bang_ke_Mua_vao`, `Bang_ke_Ban_ra` (hoặc file riêng `Bang_ke_mua_vao_ban_ra_tong_hop_<KY>.xlsx` 2 sheet `mua vao` / `ban ra`).
- Lọc dòng nhãn `Tổng cộng` (thường mỗi quý 1 dòng). Giữ đối tác tên `TỔNG CÔNG TY...`.
- Kỳ bổ sung: lấy bản BS mới nhất, khử trùng kỳ.
- Map header Doanh số / Thuế / Thuế suất; sheet bán ra và mua vào có thể lệch cột.

### 3. NLĐ

- 1 sheet `Tong_hop_NLD`. Ghi nguồn `05-1` hoặc `05-2`.
- Nguồn 05-2 mà công ty không trích 10% → cột Đề xuất kiểm tra cam kết = Có.
- Lưu bản NLĐ theo năm vào `1.4. Tong hop danh sach NLD theo QTT TNCN/` nếu chưa có.

### 4. Tóm tắt và lưu

- Sheet `Tom_tat_kiem_tra`: Đủ / Thiếu / Mỏng từng nguồn. Portal thường **mỏng mua vào** so với bảng kê NNT: ghi rõ độ phủ, không đánh Đủ giả.
- Thiếu 1 nguồn: vẫn xuất sheet có dữ liệu; không điền số 0 giả.
- Lưu `Tong_hop_dau_vao_HDDT_BangKe_NLD_<MST>_<KY>.xlsx` vào `4. KET QUA PHAN TICH/cau-thanh/`.
- File bảng kê 2 sheet (nếu xuất riêng) vào `2.5` và sao `cau-thanh/`.

## Định dạng số tiền

Hiển thị chuẩn VN: `.` hàng nghìn, `,` phần lẻ (`1.234.567,89`). Excel `#,##0` / `#,##0.00`. Font Times New Roman. Cấm ô text bắt đầu bằng `=`.

## Bẫy

- Nhét bảng kê NNT vào `1.2` (folder HĐĐT portal).
- Giữ dòng Tổng cộng khi cộng doanh số.
- Coi HĐĐT portal đủ mua vào vì có file `HDDT.xlsx`.
- Lệch cột bán ra vs mua vào vì map theo index cột.

## Kiểm chứng

- HĐĐT và bảng kê **tách sheet**; Tom_tat ghi Đủ/Thiếu/Mỏng.
- Dòng Tổng cộng đã lọc; đối tác TỔNG CÔNG TY còn.
- File nằm tại `cau-thanh/`. Bước tiếp: `lam-sach-cd1-mst-hoa-don` rồi `doi-chieu-gtgt-tndn-cbrr`.
