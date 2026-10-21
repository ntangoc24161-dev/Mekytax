# Cấu trúc Tool_DC_511_632.xlsm

Nguồn phân tích: file staging `Tool_DC_511_632.xlsm` (53.045 byte), 7 sheet + VBA module `DoiChieu511632`.

## Danh sách sheet

| Sheet | Vai trò |
|---|---|
| `BangDieuKhien` | Hướng dẫn 4 bước vận hành, quy tắc 3 khóa, checklist cài macro |
| `511` | Vùng dán sổ doanh thu (header dòng 1, dữ liệu từ dòng 2) |
| `632` | Vùng dán sổ giá vốn (header dòng 1, dữ liệu từ dòng 2) |
| `DCDTCP` | Bảng đối chiếu chi tiết từng dòng (output chính của macro) |
| `BaoCao` | Tổng hợp theo năm chứng từ + chỉ tiêu lãi gộp / bán dưới giá |
| `CauHinh` | Nhãn trạng thái R và tên sheet nguồn/đích |
| `HuongDan` | Xử lý lỗi macro, cách chạy lại kỳ khác, quy tắc lưu file |

## Macro chính

- Module: `DoiChieu511632`
- Entry: `ChayDoiChieu`
- Output file gợi ý trong VBA: `DoiChieu_511_632_KetQua_...`
- Cảnh báo VBA đã thấy: cần đủ 2 sheet tên đúng `511` và `632`; sheet trống từ dòng 2; thông báo không khớp 632.

Lưu riêng (nếu có macro `LuuKetQua`): Yes = ghi đè tool; No = chọn thư mục file mới; Cancel = không lưu. Nút Save thường của Excel vẫn lưu đè im lặng.

## Nhãn cấu hình (`CauHinh`)

| Khoa | Gia_tri |
|---|---|
| StatusFull | Khớp HĐ + ngày CT + mặt hàng |
| StatusDate | Khớp HĐ + mặt hàng (khác ngày CT) |
| StatusFuzzy | Ghép theo ngày CT (khác HĐ/mặt hàng) |
| StatusNoCogs | Không khớp 632 |
| StatusNoRev | Có giá vốn, không có doanh thu |
| Sheet511 | 511 |
| Sheet632 | 632 |
| SheetOut | DCDTCP |

## Cột `DCDTCP`

### Khối 511 (A–G)

Ngày hạch toán | Ngày chứng từ | Số chứng từ | Số hóa đơn | Diễn giải | Tài khoản | Phát sinh Có

### Khối 632 (I–O)

Ngày hạch toán | Ngày chứng từ | Số chứng từ | Số hóa đơn | Diễn giải | Tài khoản | Phát sinh Nợ

### Khối kết quả (P–S)

| Cột | Ý nghĩa |
|---|---|
| P | Lãi/Lỗ = G − O |
| Q | Khớp diễn giải (TRUE/FALSE) |
| R | Tình trạng đối chiếu (5 nhãn ở trên) |
| S | Năm chứng từ |

## Khối chỉ tiêu `BaoCao`

1. **Số dòng đối chiếu:** khớp 3 khóa / lệch ngày CT / lệch HĐ-tên hàng / có DT không GV / có GV không DT / tổng dòng.
2. **Chỉ tiêu tài chính:** DT 511, GV 632, lãi gộp, tỷ lệ lãi gộp.
3. **PS Nợ 632 theo tình trạng R** + tỷ trọng GV khớp đủ 3 khóa.
4. **Bán dưới giá (P ≤ 0):** số dòng lỗ, hòa vốn, tổng P ≤ 0; DT/GV/ΣP của nhóm; tỷ lệ DT dưới giá / tổng DT.

Cột năm: 2018 → 2030; cột B = tổng mọi năm. Năm báo cáo lấy tự động theo cột S.

## Vận hành 4 bước (từ BangDieuKhien)

1. Dán doanh thu vào sheet `511`.
2. Dán giá vốn vào sheet `632`.
3. Chạy `Alt+F8` → `ChayDoiChieu`.
4. Lọc `DCDTCP` cột R / cột P và đọc `BaoCao`.
