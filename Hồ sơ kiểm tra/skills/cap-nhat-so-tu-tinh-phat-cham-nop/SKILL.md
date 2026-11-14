---
name: Cập nhật sổ JAVIS PRO từ sheet TINH PHAT + CHAM NOP
description: Đối chiếu dòng tổng sheet TINH PHAT + CHAM NOP của Bảng tính với sổ Danh sach kiem tra quyet toan pro.xlsm, sửa ô lệch, giữ công thức và VBA.
group: Tài chính
origin: javis-learned
status: active
created: 2026-11-05
updated: '2026-11-12'
learned_updated: '2026-11-12'
---
## Khi nào dùng
- User yêu cầu: "đối chiếu / đồng bộ số liệu phạt và chậm nộp giữa Bảng tính và sổ theo dõi".
- Khi đoàn chốt số biên bản kiểm tra và cần tự động cập nhật số liệu sang sổ theo dõi `Danh sach kiem tra quyet toan pro.xlsm`.
- Chạy trước khi kết thúc khâu soạn biên bản và hồ sơ về.

## Chuẩn bị
- Bảng tính của vụ: `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\<VỤ>\0. BIEN BAN\Bang tinh.xlsx`.
- Sổ theo dõi: `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm`, tìm dòng theo MST/tên vụ tại `01_DA_BAN_HANH` hoặc `02_DANG_THUC_HIEN`.
- Script tự động hóa: `scripts/dong_bo_tinh_phat_sang_so.py`.

## Quy tắc cơ chế chốt
1. **Chế độ vận hành**: Luôn đồng bộ tự động mỗi lần chạy. Khi phát hiện ô số liệu lệch giữa Bảng tính và sổ, tự động ghi đè số mới vào sổ `pro.xlsm` mà không cần xác nhận từng ô.
2. **Mốc thời điểm tính phạt**: Lấy **Ngày ký BBKT** trên sổ (cột 36 / AJ) làm mốc điền vào cột D "Ngày lập biên bản kiểm tra" trên sheet `TINH PHAT + CHAM NOP (C T)` để Excel tính toán chính xác tiền chậm nộp 0,03% đến ngày ký BBKT.
3. **Phạt thủ tục VPHC (cột 47 / AU)**: Cảnh báo riêng khi có phát sinh (> 0) để người dùng kiểm tra lại hồ sơ xử phạt vi phạm hành chính.
4. **Bảo vệ cấu trúc sổ**:
   - Tuyệt đối KHÔNG tạo file `.bak` trong thư mục `00. Theo doi chung\`.
   - Giữ nguyên macro VBA (`keep_vba=True`).
   - KHÔNG ghi đè các cột công thức tự tính: AP (tổng truy thu), AS (tổng chậm nộp), AV (tổng phạt), AW (tổng phát sinh).
   - **Tự động điền công thức chữ `=DocSo(...)`**: Khi cập nhật số liệu hoặc thêm công ty mới vào `01_DA_BAN_HANH`, tự động sao chép/điền công thức đọc số cho 6 cột cuối (cột 53 đến 58 / BA đến BF: `Tong_truy_thu_chu`, `Tong_cham_nop_chu`, `Tong_tien_phat_chu`, `Tong_phat_sinh_chu`, `Giam_lo_chu`, `Giam_khau_tru_chu`).

## Quy trình thực hiện
1. Mở sổ theo dõi `pro.xlsm`, tìm dòng theo MST để đọc **Ngày ký BBKT** (cột 36 / AJ).
2. Mở sheet `TINH PHAT + CHAM NOP (C T)` trên `Bang tinh.xlsx`, kiểm tra cột D xem đã đúng Ngày ký BBKT chưa. Nếu chưa, cập nhật Ngày ký BBKT vào để Excel tính toán đúng tiền chậm nộp 0,03%.
3. Đọc dòng TỔNG CỘNG trên sheet TINH PHAT và so sánh với dòng vụ trên sổ:
   - Cột AN (40): Truy thu TNDN (tổng cột B trên Bảng tính)
   - Cột AO (41): Truy thu GTGT (tổng cột C trên Bảng tính)
   - Cột AQ (43): Chậm nộp GTGT (tổng cột H hoặc I trên Bảng tính)
   - Cột AR (44): Chậm nộp TNDN 0,03% (tổng cột K hoặc M trên Bảng tính)
   - Cột AT (46): Phạt khai sai 20% (tổng cột N trên Bảng tính)
   - Cột AU (47): Phạt thủ tục VPHC (thông báo nếu > 0)
4. Tự động ghi đè các ô lệch vào sổ và lưu file `pro.xlsm`.
5. Báo cáo kết quả dạng bảng: chỉ tiêu, số cũ, số mới, trạng thái và cảnh báo phát sinh nếu có.

## Lệnh thực thi nhanh
```powershell
python "scripts\dong_bo_tinh_phat_sang_so.py" <MST_hoac_Ten_vu>
# Nếu chỉ muốn kiểm tra trước:
python "scripts\dong_bo_tinh_phat_sang_so.py" <MST_hoac_Ten_vu> --dry-run
```
