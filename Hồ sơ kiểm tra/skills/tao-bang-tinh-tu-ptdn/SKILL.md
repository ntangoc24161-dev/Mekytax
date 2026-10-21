---
name: "Tạo Bang tinh từ PTDN"
description: "Tạo Bang tinh.xlsx từ mẫu PTDN01/PTDN02: điền KQKD/CDKT, chỉnh kỳ, đưa lên Drive thư mục vụ."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-10-16
---
## Khi nào dùng

- User yêu cầu tạo hoặc bổ sung `Bang tinh.xlsx` cho một vụ kiểm tra (kế hoạch hoặc giải thể).
- Trong thư mục vụ trên Drive đã có file PTDN (PTDN01 cân đối, PTDN02 kết quả SXKD) và cần bảng tính tỷ lệ/chi phí theo năm.
- Không dùng để xuất hồ sơ Word (`ho-so-kiem-tra-ke-hoach`, `lap-ho-so-kiem-tra`) hay chấm XML/hóa đơn (`phan-tich-ho-so-thue`).

## Chuẩn bị

- Thư mục vụ trên Drive (ví dụ `2026/DANG KIEM TRA/GIAI THE/<TEN CONG TY>` hoặc thư mục kế hoạch tương ứng).
- Mẫu `Bang tinh.xlsx` lấy từ vụ khác đã chuẩn (cùng cấu trúc sheet).
- Nguồn số liệu: `PTDN01_...` (CDKT), `PTDN02_...` (KQKDSXKD). Có thể có nhiều file theo từng đoạn năm.
- MST và tên NNT đã khớp.

## Cách chạy

1. Khớp thư mục Drive và MST.
2. Liệt kê PTDN trong thư mục; xác định khoảng năm có số liệu.
3. Sao mẫu `Bang tinh.xlsx`, xóa số liệu công ty mẫu theo danh tính công ty cũ (tên/MST), không xóa theo mốc năm cố định.
4. Điền sheet **KQKDSXKD** từ PTDN02 (doanh thu, giá vốn, CPBH, CP QLDN…).
5. Điền sheet **CDKT** từ PTDN01 theo các cột năm mà phôi còn hỗ trợ.
6. Đặt kỳ báo cáo đúng yêu cầu user hoặc đúng union các năm PTDN có thật (vd `Từ năm 2014 đến năm 2026`).
7. Lưu bản vault `exports/<slug-mst>/Bang tinh.xlsx` và tải/cập nhật lên đúng folder Drive vụ.
8. Báo link Drive, kỳ đã điền, sheet nào thiếu cột năm, và vài số doanh thu thuần để user đối chiếu.

## Quy trình

### Tạo mới

1. Tải mẫu + toàn bộ PTDN liên quan.
2. Điền KQKDSXKD/CDKT cho đúng các năm có nguồn.
3. Kiểm tra năm đầu kỳ: nếu công ty mới thành lập giữa năm thì doanh thu năm đó có thể = 0 và vẫn giữ cột.

### Bổ sung thêm giai đoạn năm

1. Không tạo file mới nếu đã có `Bang tinh.xlsx` trên Drive/vault: mở file hiện có.
2. Ghép thêm năm từ PTDN đoạn mới; giữ nguyên các năm đã điền.
3. Cập nhật kỳ báo cáo cho khớp toàn bộ năm sau khi ghép.
4. Đẩy lại cùng file Drive (cùng id) và đồng bộ bản vault.

## Bẫy

- Không xóa hàng/cột năm chỉ vì mẫu cũ có số liệu trước 2020: user có thể yêu cầu bổ sung giai đoạn sớm hơn; chỉ xóa dữ liệu công ty mẫu.
- Phôi CDKT có thể không có cột năm rất sớm (đã gặp thiếu 2014-2015): điền các năm còn cột, nói rõ năm nào không có cột.
- Kỳ báo cáo phải khớp union năm user yêu cầu và năm PTDN thật, không mặc định cắt theo một file PTDN đầu tiên.
- Sau khi đẩy Drive, nhắc user F5 nếu sheet còn cache bản cũ.

## Kiểm chứng

- Kỳ trên file khớp yêu cầu user.
- KQKDSXKD có đủ năm đã khai báo (trừ năm không có trên PTDN).
- File nằm đúng folder vụ trên Drive và có bản `exports/<slug-mst>/Bang tinh.xlsx`.
- Vài số doanh thu thuần đối chiếu được với PTDN02.
