# Mẫu văn bản và Excel

Nguồn mã: `D:\SET UP BY AI\HO SO KIEM TRA\app\src\documents.py` và `src\excel_io.py`.
App có 34 mẫu. `TIEN_TRINH.md` ghi 33 là số cũ, trước khi đủ VE-KL.

## Hồ sơ đi (luôn tick)

| Mã | Tên |
|---|---|
| DI-01 | Phân tích hồ sơ (08/QTKT-TC) |
| DI-02 | Tờ trình ban hành QĐ kiểm tra / giám sát |
| DI-03 | Quyết định kiểm tra |
| DI-04 | Giấy xác nhận giao QĐ |
| DI-05 | Quyết định giám sát |
| DI-06 | Kế hoạch giám sát |
| DI-07 | Thư trao đổi thông tin |

## Biên bản (luôn tick)

| Mã | Tên |
|---|---|
| BB-01 | Biên bản công bố QĐ kiểm tra |
| BB-02 | Phiếu phân công công việc |
| BB-03 | BB xác nhận số liệu, thành viên |
| BB-04 | BB xác nhận số liệu, trưởng đoàn |
| BB-05 | Biên bản kiểm tra |
| BB-06 | Biên bản công khai |

## Phát sinh (chỉ khi bật trên form)

| Mã | Nhóm | Tên |
|---|---|---|
| PS-BB-01a | bai_bo | QĐ bãi bỏ quyết định |
| PS-BB-01b | bai_bo | Tờ trình bãi bỏ QĐ |
| PS-BB-02a | tam_hoan | Thông báo tạm hoãn kiểm tra |
| PS-BB-02b | tam_hoan | Tờ trình tạm hoãn |
| PS-BB-03a | tam_dung | Thông báo tạm dừng kiểm tra |
| PS-BB-03b | tam_dung | Tờ trình tạm dừng |
| PS-BB-04a | dieu_chinh | Tờ trình điều chỉnh |
| PS-BB-04b | dieu_chinh | 11.KTT (điều chỉnh) |
| PS-BB-04c | dieu_chinh | 12.KTT (điều chỉnh) |
| PS-BB-04d | dieu_chinh | 13.KTT (điều chỉnh) |
| PS-BB-05a | gia_han | QĐ gia hạn |
| PS-BB-05b | gia_han | Tờ trình gia hạn |

## Hồ sơ về

| Mã | Khi nào có | Tên |
|---|---|---|
| VE-01 | luôn | Báo cáo kết quả kiểm tra (12/QTKT-TC) |
| VE-02 | luôn | Biên bản làm việc giám sát |
| VE-03 | luôn | Báo cáo nội dung và KQ giám sát |
| VE-04 | luôn | Nhật ký thanh tra kiểm tra |
| VE-05 | hướng xử phạt | Văn bản đề nghị ý kiến pháp lý |
| VE-06 | hướng xử phạt | Tờ trình QĐ xử phạt |
| VE-07 | hướng xử phạt | Quyết định xử phạt |
| VE-08 | luôn | Biên bản giao nhận hồ sơ lưu trữ |
| VE-KL | hướng kết luận | Kết luận kiểm tra |

Hướng xử lý trên sheet Vu: `xu_phat` hoặc `ket_luan`.

## Excel

Tải mẫu từ app: `http://127.0.0.1:8000/excel/mau` (sau khi đăng nhập). Ba sheet dữ liệu, dòng 1 là mã cột, không xóa. Điền từ dòng 3 trên sheet Vu. Ngày dạng `dd/mm/yyyy`.

Sheet `Vu`, cột bắt buộc: `ten_cong_ty`, `mst`, `qdkt`. Cột hay dùng thêm: `dia_chi`, `td`, `tv`, `thoi_ky_kiem_tra`, `ngay_qd`, `tu_ngay`, `den_ngay`, `huong_xu_ly`, `thue_tndn`, `thue_gtgt`, `so_qd_xuphat`.

Sheet `SoLieu`, cột bắt buộc: `mst`, `loai_thue`, `chi_tieu`, `so_nnt`, `so_kiem_tra`. Khớp vụ bằng `mst` và `qdkt`. `vai_tro`: `td` (BB-04), `tv` (BB-03), `bb05` (chỉ tiêu thêm trên BB-05), `tong_hop` (3 dòng thuế), `khac`.

Sheet `ChiTiet`: `mst`, `qdkt`, `nam`, `chi_tieu`, `noi_dung`, `so_tien`, `can_cu`. Tổng `so_tien` nên bằng chênh lệch dòng SoLieu. Lệch thì app cảnh báo và vẫn nhập.

File xuất nằm ở `app\output\<MST>\<số quyết định, dấu / đổi thành gạch ngang>\<YYYYMMDD-HHMMSS>\`. Trong thư mục đó có `hoso.zip`.
