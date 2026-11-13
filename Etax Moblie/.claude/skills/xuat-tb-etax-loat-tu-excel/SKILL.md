---
name: Xuất Thông báo eTax hàng loạt từ Excel
description: Đọc danh sách doanh nghiệp từ Excel, điền mẫu Thông báo/Tờ trình chuẩn eTax, xuất file hàng loạt.
group: Kiểm tra thuế
---

## Khi nào dùng

- Cần phát hành hàng loạt **Thông báo lần 1** hoặc **Thông báo lần 2** cho một đợt kiểm tra eTax Mobile mới.
- Cần xuất **Tờ trình phê duyệt** gom danh sách doanh nghiệp để trình Lãnh đạo Thuế cơ sở ký phê duyệt.
- Muốn tự động hoá việc điền dữ liệu từ file theo dõi mà không copy-paste từng dòng.

**Trigger từ user**: "Xuất thông báo lần 1 hàng loạt", "tạo tờ trình gom", "phát hành TB lần 2 từ file dữ liệu".

---

## Chuẩn bị

### Trước khi bắt đầu

1. **File dữ liệu Excel** (nguồn danh sách):
   - Đường dẫn: `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\3- BAO CAO\4. ETAX MOBLIE\THEO DOI TIEN DO GIAI TRINH ETAX MOBILE.xlsx`
   - Các cột bắt buộc: Họ tên người nộp/doanh nghiệp, Người phản ánh, Năm xác minh, Điều khoản (phản ánh), Ngày phát hiện
   - Dòng nào có cột "Kết quả lần X" đã có dữ liệu thì bỏ qua (đã xuất rồi)

2. **Mẫu chuẩn Word** (template):
   - Thư mục lưu: `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\3- BAO CAO\4. ETAX MOBLIE\MAU CHUAN 2026\`
   - Ba loại mẫu chuẩn hóa (Quyết định 1161/QĐ-CT, Thuế cơ sở 3):
     - `TO TRINH 2026 CHUAN.docx` - Tờ trình gom doanh nghiệp phê duyệt
     - `TB ETAX LAN 1 MAU 2026 CHUAN.docx` - Thông báo lần 1 (mời làm việc lần đầu)
     - `TB ETAX LAN 2 MAU 2026 CHUAN.docx` - Thông báo lần 2 (mời làm việc lần 2)
   - Mẫu này đã chứa: Căn cứ pháp lý, Tên đơn vị Thuế cơ sở 3, mã số `/TB-TCS3`, yêu cầu không ủy quyền

3. **Thư mục đầu ra**:
   - Tạo folder con theo đợt: `4. ETAX MOBLIE\XUAT LAN [N] THANG [M] 2026\` 
   - Ví dụ: `XUAT LAN 1 THANG 11 2026\` cho đợt 1 tháng 11/2026

---

## Quy trình

### Bước 1: Đọc file Excel + sàng lọc

1. Mở file `THEO DOI TIEN DO GIAI TRINH ETAX MOBILE.xlsx`
2. Xác định loại văn bản cần xuất:
   - **Tờ trình**: Đọc toàn bộ dòng chưa có "Số TB lần 1" (chưa xuất lần nào)
   - **Thông báo lần 1**: Các dòng có "Số TB lần 1" rỗng hoặc "(chưa)")
   - **Thông báo lần 2**: Các dòng có "Số TB lần 1" đã có nhưng "Số TB lần 2" rỗng
3. Gom danh sách theo loại (vd 8 doanh nghiệp cho Thông báo lần 1, 5 doanh nghiệp cho lần 2)

### Bước 2: Chuẩn bị dữ liệu chiết xuất từ Excel

Tạo bảng dữ liệu trung gian (có thể ghi vào sheet tạm hoặc giữ ở RAM) với các trường:

| Trường | Lấy từ Excel | Dùng trong mẫu |
|--------|-------------|---|
| Tên doanh nghiệp / người nộp | Cột "Họ tên người nộp" | Tiêu đề thông báo + phần trong thân |
| Mã số thuế | Cột "Mã số thuế" (nếu có) | Phần định danh doanh nghiệp |
| Năm xác minh | Cột "Năm xác minh eTax" | "...chi trả thu nhập năm **2022**..." |
| Ngày phát hiện | Cột "Ngày phát hiện" | Phần "Ngày..." cuối thông báo |
| Người phản ánh | Cột "Người phản ánh" | Ghi chú quản lý nội bộ (nếu cần) |
| Cán bộ liên hệ | Cột "Cán bộ liên hệ" (hoặc mặc định Nguyễn Thị Ánh Ngọc) | Chữ ký cuối thông báo |

### Bước 3: Điền mẫu Word theo loại hồ sơ

#### **Nếu xuất Tờ trình** (`TO TRINH 2026 CHUAN.docx`):
- Mở mẫu -> Tìm bảng "Danh sách người nộp thuế" -> Điền từng dòng: Tên + Mã số thuế + Năm xác minh + Điều khoản phản ánh
- Giữ nguyên: Căn cứ (Quyết định 1161/QĐ-CT), cán bộ đề xuất (Nguyễn Thị Ánh Ngọc), Lãnh đạo Tổ, Lãnh đạo Thuế cơ sở
- Ký tên: Để trống chỗ chữ ký (in ra kỳ vọng ký tay sau)
- Lưu thành: `TO_TRINH_20261113_LOT[N].docx` (ngày+số loạt)

#### **Nếu xuất Thông báo lần 1/2** (`TB ETAX LAN [1/2] MAU 2026 CHUAN.docx`):
- Mở mẫu cho từng doanh nghiệp lần lượt
- **Điểm cốt lõi**: Mẫu đã có danh sách tài liệu bắt buộc + yêu cầu "Người đại diện pháp luật (không ủy quyền) + Kế toán" -> GIỮ NGUYÊN
- Điền các vị trí cần:
  - Tên doanh nghiệp (lần 1-2)
  - Năm xác minh ("năm 2022")
  - Ngày hẹn làm việc (+ 10-15 ngày từ ngày phát hành)
  - Cán bộ tiếp nhận: Nguyễn Thị Ánh Ngọc (hoặc cán bộ trực tiếp khác)
- **Thông báo lần 2 bổ sung**: Viện dẫn Thông báo lần 1 ("Căn cứ Thông báo số ... ngày ... (Lần 1)...")
- Ký tên: Để trống (Phó Giám đốc Thuế cơ sở sẽ ký)
- Lưu thành: `TB_ETAX_LAN[1/2]_[TEN_DN]_20261113.docx` (viết tắt tên DN, ngày xuất)

### Bước 4: Xuất file + lưu vào thư mục đầu ra

1. Đặt tên file chuẩn: `TB_ETAX_LAN[N]_[TEN_DON_VI]_NGAY_[DDMMYYYY].docx`
   - Ví dụ: `TB_ETAX_LAN1_CTY_TAN_DUONG_NGAY_13112026.docx`
   
2. Lưu toàn bộ vào: `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\3- BAO CAO\4. ETAX MOBLIE\XUAT LAN [N] THANG 11 2026\`

3. Kiểm tra từng file:
   - Tên doanh nghiệp, năm, ngày hẹn đúng
   - Không có dòng trống không mong muốn
   - Mã số TB/Tờ trình chuẩn `/TB-TCS3`
   - Chữ ký để trống (sẽ ký tay/ký số sau)

### Bước 5: Cập nhật file Excel theo dõi

Sau khi xuất thành công, ghi lại vào file `THEO DOI TIEN DO GIAI TRINH ETAX MOBILE.xlsx`:
- Cột "Số TB lần 1" / "Số TB lần 2": Ghi số thành công (vd "TB-TCS3-001/2026")
- Cột "Ngày ban hành": Ghi ngày phát hành (vd 13/11/2026)
- Cột "Ngày hẹn": Ghi ngày hẹn làm việc (vd 26/11/2026)
- Cột "Tống đạt": Chuyển bộ PC/cán bộ đưa tay hoặc ghi "Đã tống đạt [ngày]"

---

## Bẫy & xử lý lỗi

### Excel thiếu dữ liệu
- **Nếu thiếu "Năm xác minh"**: Hỏi cán bộ hoặc mặc định năm 2022 (năm áp dụng chính sách eTax)
- **Nếu thiếu "Ngày phát hiện"**: Dùng ngày hôm nay (ngày xuất)
- **Nếu thiếu "Cán bộ liên hệ"**: Mặc định Nguyễn Thị Ánh Ngọc (Tổ KT1)

### Mẫu Word bị thay đổi
- Nếu mẫu mới không khớp cấu trúc cũ (số lượng bảng, vị trí điền): Báo rõ từng vị trí sai, yêu cầu chuẩn hóa lại mẫu trước điền
- Nếu chỉ lỗi chính tả nhỏ: Sửa và cập nhật file mẫu trong thư mục `MAU CHUAN 2026\`

### Tên doanh nghiệp quá dài
- **Nếu tên doanh nghiệp > 80 ký tự**: Tạo tên file viết tắt (vd "CTY XNK ABC PHAT TRIEN" → "CTY_ABC_PT") nhưng điền tên đầy đủ trong nội dung mẫu

### Tạo file hàng loạt quá lâu
- Nếu loạt > 20 doanh nghiệp: Tách thành 2 loạt nhỏ, lưu vào thư mục con riêng (vd `LOT_1/`, `LOT_2/`)

---

## Kiểm chứng

Sau khi xuất xong, trước khi tống đạt:

- [ ] Từng file đã đặt tên chuẩn (mã `/TB-TCS3`, ngày, tên DN)
- [ ] Nội dung đúng với dữ liệu Excel (tên, năm, số loạt)
- [ ] Mẫu lần 2 có viện dẫn Thông báo lần 1 (nếu đó là lần 2)
- [ ] Chữ ký để trống (sẽ ký tay hoặc ký số theo quy trình)
- [ ] File Excel được cập nhật cột "Số TB" + "Ngày ban hành" + "Ngày hẹn"
- [ ] Thư mục lưu file tương ứng với đợt/tháng (vd `XUAT LAN 1 THANG 11 2026`)
- [ ] Không có file trùng tên hoặc sót doanh nghiệp nào từ danh sách gốc

