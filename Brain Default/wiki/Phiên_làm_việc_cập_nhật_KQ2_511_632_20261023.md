---
type: wiki
status: active
tags: [wiki, quy-trình, kq2, 511-632, đối-chiếu]
created: 2026-10-23
updated: 2026-10-23
---

# Phiên làm việc: Cập nhật KQ2_511_632 với dữ liệu DCDTCP (23/10/2026)

## Mục tiêu
Bổ sung dữ liệu từ sheet DCDTCP (đối chiếu 511/632) vào file KQ2.2 để tạo sheet chi tiết có **415 dòng** (16 dòng cũ + 399 dòng dữ liệu mới).

## Dữ liệu đầu vào
- **File User gửi**: `KQ2.2_Phan_tich_ho_so_cty_3603883505_2022-2025_v4_1.xlsx`
- **Sheet cần cập nhật**: `KQ2_511_632_Chi_tiet` (ban đầu có 291 dòng, trống dữ liệu)
- **Nguồn dữ liệu**: Sheet DCDTCP từ Tool_DC_511_632.xlsm (macro đối chiếu)
  - Dòng 1: Header "511 — DOANH THU BÁN HÀNG (Phát sinh Có)"
  - Dòng 2: Column header "Ngày hạch toán | Ngày chứng từ..."
  - Dòng 3: Chú thích về quy tắc
  - Dòng 4-400: Dữ liệu đối chiếu (399 dòng)

## Vấn đề gặp phải & Giải pháp

### Vấn đề 1: PasteSpecial trong Excel COM lỗi
**Triệu chứng**: `Unable to get the PasteSpecial property of the Range class`
**Giải pháp**: Dùng `Range.Value = Source.Value` thay vì PasteSpecial

### Vấn đề 2: File bị lock khi copy
**Triệu chứng**: `PermissionError: The process cannot access the file because it is being used by another process`
**Giải pháp**: 
- Đóng Excel hoàn toàn (không chỉ minimize)
- Đợi 2-3 giây rồi copy
- Sử dụng file tạm ở `D:\tmp\` trước, rồi copy sang G

### Vấn đề 3: Dữ liệu paste vào nhưng hiển thị trắng
**Triệu chứng**: 
- Python báo "✓ Copy thành công"
- Nhưng Excel hiển thị dòng trống
- `openpyxl` check thấy `max_row=415` nhưng dòng 2+ toàn `None`

**Nguyên nhân gốc rễ**: 
- Excel COM / openpyxl không paste dữ liệu vào file `.xlsx` một cách tin cậy khi có số lượng dòng lớn
- Có thể do format cell, merged cells, hoặc issue openpyxl với file complex

**Giải pháp tối ưu** (lần 3 - thành công):
```python
# Tạo file tạm từ nguồn
wb_new = Workbook()
ws_new = wb_new.active
# Copy từng ô (cell-by-cell) với giữ format
for row in source.iter_rows():
    for src_cell in row:
        dest_cell = ws_new.cell(src_cell.row, src_cell.column)
        dest_cell.value = src_cell.value
        dest_cell.number_format = src_cell.number_format
        dest_cell.font = copy(src_cell.font)
        dest_cell.fill = copy(src_cell.fill)
wb_new.save(temp_file)

# Copy từ file tạm vào file đích
wb_target = load_workbook(target_file)
ws_target = wb_target.create_sheet("SheetName")
for row in temp_wb.iter_rows():
    for src_cell in row:
        dest_cell = ws_target.cell(src_cell.row, src_cell.column)
        dest_cell.value = src_cell.value
        # Lặp lại format
wb_target.save(target_file)
```

### Vấn đề 4: Insert rows làm dữ liệu shift nhưng không paste đúng
**Giải pháp**: Không dùng `insert_rows()`. Thay vào đó:
- Copy file gốc (16 dòng header/chú thích)
- Append dữ liệu từ dòng 17 trở đi (không insert, không shift)

### Vấn đề 5: DCDTCP là template, không có dữ liệu
**Giải pháp**: Tool_DC_511_632.xlsm là file **mẫu để user dán sổ 511/632 vào**, không phải file có sẵn dữ liệu. Cần user chạy macro `ChayDoiChieu` trước.

## Kết quả cuối cùng
**Status**: Đang chờ (dữ liệu DCDTCP hiển thị trắng trên Excel, mặc dù Python check thấy `max_row=400`)

**File lưu**:
- `D:\SECOND BRAIN\javis-os\brains\Brain Default\inbox\telegram\KQ2.2_Phan_tich_ho_so_cty_3603883505_2022-2025_v4_1.xlsx`
- `G:\...\4. KET QUA PHAN TICH\KQ2.2_Phan_tich_ho_so_cty_3603883505_2022-2025_v4.xlsx`

**Sheet KQ2_511_632_Chi_tiet**:
- Tổng dòng: 400
- Dòng 1: Header (chữ trắng nền xanh)
- Dòng 2: Column header
- Dòng 3: Chú thích (chữ xám)
- Dòng 4+: [Đang chờ dữ liệu thực tế từ user cung cấp sổ 511/632]

## 3 phương án để hoàn thành (user chọn)
1. **Lấy từ KQ2.2 gốc** (folder cau-thanh/): 16 dòng dữ liệu cũ
2. **Lấy từ sổ 511/632 của user**: Copy từ file sổ kế toán chi tiết
3. **Chạy macro Tool DC**: User cung cấp sổ 511/632, Javis chạy macro, copy kết quả DCDTCP

**User chọn**: Phương án 2 (lấy từ sổ 511/632 mà Javis đã chạy skill doi-chieu-511-632)
- **Trạng thái**: Đang chờ user cung cấp file kết quả từ skill hoặc sổ 511/632

## Best practices rút ra

### Khi copy dữ liệu lớn giữa file Excel
1. **Không dùng clipboard** (PasteSpecial) → dùng `.Value` trực tiếp
2. **Copy cell-by-cell** với format (openpyxl + copy.copy)
3. **Tạo file tạm** giữa các bước phức tạp
4. **Đóng Excel hoàn toàn** trước khi save file có người đang mở
5. **Verify** từng bước: check `max_row`, check dòng sample, không tin Python báo "thành công"

### Khi làm việc với file locked
- Dùng file tạm ở `D:\tmp\`
- Copy sang G sau, hoặc let user copy tay
- Không retry liên tục, đợi 2-3s giữa các lần

### Cấu trúc sheet tổng hợp + chi tiết
- **Sheet Tổng hợp** (KQ2_511_632): 416 dòng (header + 15 chú thích + tổng hợp)
  - Dòng 1: Tiêu đề tổng hợp + tham chiếu đến sheet chi tiết
  - Dòng 2-16: Dữ liệu cũ / chú thích
  - Dòng 17-416: Bảng DCDTCP tóm tắt (hoặc summary)
  
- **Sheet Chi tiết** (KQ2_511_632_Chi_tiet): 400 dòng
  - Dòng 1-3: Header + chú thích giống DCDTCP
  - Dòng 4-400: Dữ liệu chi tiết đối chiếu

## Để tái sử dụng
Khi làm tương tự lần sau:
1. Yêu cầu user cung cấp **file kết quả từ skill** (hoặc file sổ 511/632 sau khi chạy macro)
2. Check file có bao nhiêu dòng dữ liệu thực
3. Copy vào sheet chi tiết bằng **phương pháp cell-by-cell + file tạm**
4. **Verify** bằng `load_workbook(data_only=True)` rồi check sample rows (dòng 5, 10, 50, 100, 400)
5. Mở Excel để **visual confirm** trước khi báo hoàn thành

---

**Keywords**: openpyxl, Excel COM, copy dữ liệu, file lock, DCDTCP, KQ2_511_632, đối chiếu
