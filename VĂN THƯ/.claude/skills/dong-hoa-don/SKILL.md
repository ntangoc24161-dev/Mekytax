---
name: Quy trình đóng hóa đơn và theo dõi trạng thái
description: "Quản lý quy trình kiểm tra, đóng, giải trình và mở lại hóa đơn doanh nghiệp theo Luật Quản lý thuế 108/2025 và NĐ 254/2026."
group: Vận hành
---

## Khi nào dùng

- Kiểm tra rủi ro hóa đơn của doanh nghiệp theo các quy tắc thuế
- Chuẩn bị danh sách doanh nghiệp để gửi thông báo đóng hóa đơn
- Theo dõi trạng thái từng doanh nghiệp (Đóng → Giải trình → Mở lại / Ngừng)
- Ban hành thông báo mở lại hoặc ngừng sử dụng hóa đơn
- Tổng hợp báo cáo hoàn thành quy trình năm 2026

---

## Chuẩn bị

**File danh sách chính:**
- `Danh sach theo doi dong mo HD.xlsx` tại: `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\2. DONG MO SU DUNG HDON\DANH SACH - THEO DOI\`

**Cấu trúc Excel (9 cột):**
| STT | Mã số thuế | Tên người nộp thuế | Thông báo đóng | Ngày đóng | Thông báo mở | Nội dung giải trình | Trạng thái | (Ghi chú) |
|---|---|---|---|---|---|---|---|---|
| =IF(B2<>"",COUNTA($B$2:B2),"") | 3602484265 | CÔNG TY TNHH THẾ LY | 5354-01 | 2025-05-23 | 45875 | Đã bổ sung BCTC | **Mở lại** | |

**Công văn template:**
- Thông báo mở: `CONG VAN THONG BAO/2026/3-TB MO HOA DON-MERGE.docx`
- Thông báo ngừng: `CONG VAN THONG BAO/2026/2-TB ngung HOA DON - MERGE.docx`

---

## Quy trình (4 giai đoạn)

### **Giai đoạn 1: Kiểm tra & Chuẩn bị (Tuần 1)**
1. Chạy skill `giai-doan-1-kiem-ke-danh-gia` hoặc `giai-doan-2-phan-tich-du-lieu-co-quan-thue`
2. Xác định 80+ doanh nghiệp rủi ro theo 57 quy tắc thuế
3. Điền danh sách MST vào cột B (Mã số thuế) của Excel
4. Kiểm tra đủ thông tin: tên DN, địa chỉ, ngành nghề, quy mô

### **Giai đoạn 2: Gửi thông báo đóng (Tuần 2-3)**
1. **Chuẩn bị công văn:**
   - Copy template từ `2-TB ngung HOA DON - MERGE.docx` (nếu đóng lâu) hoặc `3-TB MO HOA DON-MERGE.docx`
   - Thay thế thông tin: Số công văn, ngày ban hành, danh sách DN, lý do, căn cứ pháp lý
   - Làm trùng bình công văn (Phần THÔNG BÁO + Phần TỜ TRÌNH)

2. **Ghi chú trong Excel:**
   - Cột D (Thông báo đóng): Ghi số công văn, vd "5354-01"
   - Cột E (Ngày đóng): Ngày ban hành công văn
   - Cột I (Trạng thái): Tự động thành "Đóng"

3. **Gửi:** Qua iOffice hoặc bộ phận công văn, lưu bản gốc vào: `CONG VAN THONG BAO/2026/`

### **Giai đoạn 3: Chờ & Nhận giải trình (Tuần 4-8)**
1. **Ghi nhận giải trình:**
   - DN gửi lại: Báo cáo kế toán bổ sung, giấy tờ minh chứng, kê khai đã điều chỉnh
   - Cột G (Nội dung giải trình): Mô tả ngắn: "Đã bổ sung BCTC", "Kê khai điều chỉnh TNDN"

2. **Kiểm tra đủ điều kiện:**
   - ✅ Nếu đủ: Chuyển sang **Giai đoạn 4a** (Mở lại)
   - ❌ Nếu thiếu/quá hạn: Chuyển sang **Giai đoạn 4b** (Ngừng)

### **Giai đoạn 4a: Ban hành thông báo MỞ LẠI (Tuần 9)**
1. **Chuẩn bị công văn mở:**
   - Copy template `3-TB MO HOA DON-MERGE.docx`
   - Thay thế: Số TB mở, ngày ban hành, danh sách DN được mở
   - Lý do: "Đơn vị đã giải trình, bổ sung thông tin, tài liệu theo quy định"
   - Làm trùng bình công văn (Phần THÔNG BÁO + Phần TỜ TRÌNH)

2. **Cập nhật Excel:**
   - Cột F (Thông báo mở): Ghi số công văn mở, vd "45875"
   - Cột I (Trạng thái): Tự động thành "Mở lại" ✅

3. **Lưu:** `CONG VAN THONG BAO/2026/3-TB MO HOA DON-MERGE.docx` (đã MERGE)

### **Giai đoạn 4b: Ban hành thông báo NGỪNG (Nếu cần)**
1. **Chuẩn bị công văn ngừng:**
   - Copy template `2-TB ngung HOA DON - MERGE.docx`
   - Thay thế: Số TB ngừng, ngày ban hành, danh sách DN bị ngừng
   - Lý do: "Không chấp hành giải trình, bổ sung thông tin tài liệu theo thông báo cơ quan thuế"
   - Làm trùng bình công văn (Phần THÔNG BÁO + Phần TỜ TRÌNH)

2. **Cập nhật Excel:**
   - Cột I (Trạng thái): Chỉnh thành "Ngừng" ⛔
   - Ghi chú: Lý do ngừng, ngày hiệu lực

3. **Lưu:** `CONG VAN THONG BAO/2026/2-TB ngung HOA DON - MERGE.docx` (đã MERGE)

---

## Cách sử dụng file Excel

### **Cột Trạng thái (Cột I) - Tự động hóa:**
```
Logic:
- Nếu (Thông báo mở ≠ trống) → "Mở lại" 🟢
- Nếu (Thông báo đóng ≠ trống) → "Đóng" 🔴
- Nếu cả hai trống → "Chưa xử lý" 🟡
```

### **Cách filter & sort:**
1. **Xem DN chưa xử lý:** Filter Trạng thái = "Chưa xử lý" → xác nhận tại sao (chờ audit? chưa gửi TB?)
2. **Xem DN đang đóng:** Filter Trạng thái = "Đóng", sort Ngày đóng cũ nhất trước
3. **Xem DN đã xử lý:** Filter Trạng thái = "Mở lại" → tính số lượng hoàn thành

---

## Bẫy thường gặp

⚠️ **Số công văn bị trùng:** Kiểm tra iOffice trước khi ghi số công văn mới vào cột D/F

⚠️ **Quá hạn chờ giải trình:** Nếu >30 ngày từ ngày đóng mà chưa có giải trình → chuyển sang ngừng vĩnh viễn

⚠️ **DN giải trình không đủ:** Kiểm tra kỹ: BCTC đã chuẩn không? Kê khai GTGT-TNDN khớp? Chứng từ đầu vào rõ ràng?

⚠️ **Công văn merge bị thay đổi:** Sau khi merge, KHÔNG sửa số công văn hoặc danh sách - nó đã ký số rồi

---

## Kiểm chứng hoàn thành

✅ **Hàng tuần:** Đọc Excel, kiểm tra số lượng:
- Mở lại: bao nhiêu DN?
- Đóng: bao nhiêu DN chờ (lâu hơn 30 ngày)?
- Chưa xử lý: bao nhiêu DN?

✅ **Cuối tháng:** Tổng hợp báo cáo:
- DN được mở lại: XX tờ, lý do chủ yếu (bổ sung BCTC, kê khai điều chỉnh...)
- DN bị ngừng: XX tờ, lý do (quá hạn, không chấp hành giải trình)
- Tiến độ: XX% hoàn thành

✅ **Trước kỳ quý:** Xuất Excel danh sách quyết toán đã xử lý, lưu vào `Bao cao dong mo HD/` để theo dõi năm sau

---

## Liên kết & Tài liệu

- [[Luật Quản lý thuế 108/2025]] - Căn cứ pháp lý chính
- [[Nghị định 254/2026 về hóa đơn]] - Chi tiết về hóa đơn, chứng từ
- **Excel file:** G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\2. DONG MO SU DUNG HDON\DANH SACH - THEO DOI\Danh sach theo doi dong mo HD.xlsx
- **Công văn template:** G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\2. DONG MO SU DUNG HDON\CONG VAN THONG BAO\2026\
