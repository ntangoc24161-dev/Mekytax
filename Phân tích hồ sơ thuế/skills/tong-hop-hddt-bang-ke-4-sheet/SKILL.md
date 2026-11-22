---
name: Tổng hợp hóa đơn và bảng kê 4 sheet
description: "Tổng hợp HĐĐT và bảng kê theo 3 giai đoạn: tạo 2 sheet HĐĐT + tổng hợp đối tác, cập nhật bảng kê và phân loại nguồn, bổ sung trạng thái & rủi ro."
group: Tài chính
updated: '2026-11-20'
---

# Lập bảng kê và tổng hợp hóa đơn theo 3 giai đoạn

Skill này tổng hợp hóa đơn và đối tác từ cổng CQT (HĐĐT) và Bảng kê NNT theo quy trình 3 giai đoạn lũy tiến:

---

### Giai đoạn 1: Nguồn HĐĐT & Danh mục đối tác xuất danh bạ
- **2 sheet HĐĐT**: `HDDT_Mua_vao` và `HDDT_Ban_ra` (chuẩn hóa MST theo BT6).
- **1 sheet `Tong_hop_doi_tac`**:
  - Tập hợp 100% MST và Tên công ty mua vào / bán ra từ HĐĐT để phục vụ xuất sang Danh bạ NNT.
  - Cột `Ghi chú nguồn`: Ghi nhận `HDDT_Mua` hoặc `HDDT_Ban`.

---

### Giai đoạn 2: Bổ sung Bảng kê NNT & Đối soát nguồn đối tác
- **2 sheet Bảng kê**: Bổ sung `Bang_ke_Mua_vao` và `Bang_ke_Ban_ra` (chuẩn hóa MST theo BT6, lọc bỏ dòng nhãn kỳ).
- **Cập nhật sheet `Tong_hop_doi_tac`**:
  - Ghép nối dữ liệu đối tác giữa HĐĐT và Bảng kê.
  - Phân loại và cập nhật lại cột `Ghi chú nguồn`:
    - Nếu có sai khác (chỉ xuất hiện ở Bảng kê): ghi `Bang_ke_Mua` hoặc `Bang_ke_Ban`.
    - Nếu khớp cả 2 nguồn (không sai khác): cập nhật thành `Chung_Mua` hoặc `Chung_Ban`.

---

### Giai đoạn 3: Tình trạng hoạt động, CQT quản lý & Phân tích rủi ro
- **Bổ sung cột tra cứu vào các sheet**:
  - `Tình trạng hoạt động`: Lấy từ Danh bạ NNT (cột trạng thái hoạt động).
  - `Cơ quan thuế quản lý`: Lấy từ Danh bạ NNT (cột cơ quan thuế quản lý).
- **Gắn cờ và đối soát rủi ro**:
  - NNT đang kiểm tra: Tra cứu F1, CVCBRR và Sổ Công an.
  - NCC mua vào: Chỉ lọc cờ F0 từ Sổ CVCBRR (ghi trực tiếp số, ngày, cơ quan ban hành công văn vào cột `Cảnh báo rủi ro`).
- **Sheet `Tong_hop_trang_thai_rui_ro`**: Thống kê ĐKKD, thống kê rủi ro F0 và danh sách chi tiết các đơn vị cần kiểm tra trọng tâm.

---

## Quy tắc chuẩn hóa Mã số thuế (BT6 - MST_chuan)

1. **10 chữ số**: Giữ nguyên (VD: `3602996707` → `3602996707`).
2. **9 chữ số**: Thêm 1 số 0 ở đầu để đủ 10 số (VD: `602996707` → `0602996707`).
3. **12 chữ số** (bắt đầu từ 0): Thêm 1 số 0 ở đầu và thêm dấu gạch ngang `-` trước 3 chữ số chi nhánh cuối (VD: `602996707001` → `0602996707-001`).
4. **13 chữ số** (bắt đầu từ 3): Giữ nguyên và thêm dấu gạch ngang `-` trước 3 chữ số chi nhánh cuối (VD: `3602996707001` → `3602996707-001`).
5. **14 ký tự** (đã có sẵn dấu `-`): Giữ nguyên.

## Định dạng & Lưu trữ

- Font: Times New Roman thống nhất toàn bộ các sheet.
- Số tiền định dạng `#,##0`. Dòng tổng cộng dùng công thức `SUM`.
- Lưu tại `ket-qua/cau-thanh/Tong_hop_HDDT_BangKe_<MST>_<KY>.xlsx` và đồng bộ sang thư mục hồ sơ kiểm tra trên máy.
