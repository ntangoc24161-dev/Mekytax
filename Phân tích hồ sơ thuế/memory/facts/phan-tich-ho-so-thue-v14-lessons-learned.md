---
type: decision
created: 2026-10-23
updated: 2026-10-23
---

## Bài học rút ra - Phân tích hồ sơ thuế v1.4

### 1. IN04_BCTC_TNDN - Cấu trúc dữ liệu

**Vì sao:** Lần đầu tìm dữ liệu 2022-2024 không được vì tìm sai chỗ

**Bài học:**
- IN04 có cấu trúc **dòng xếp chồng**, không phải cột F (So_nam_truoc)
- Dữ liệu 4 năm (2022-2025) được sắp xếp liên tiếp: Năm 2022 (dòng 3-80), Năm 2023 (81-150), v.v.
- Cách tìm: Quét cột A (năm) + B (loại báo cáo CDKT/KQKD/TNDN) + C (mã số) từ dòng 3 đến hết
- Các chỉ tiêu cần: **KQKD 50** (LN BDKD), **CDKT 141** (Khấu hao), **TNDN ct16** (Lãi vay)

**Áp dụng:** Lần sau lấy dữ liệu IN04, luôn quét dòng (rows) thay vì dựa vào cột, vì cột E-H chỉ chứa dữ liệu năm hiện tại

---

### 2. EBITDA & Khả năng trích lãi vay

**Vì sao:** Công thức EBITDA phức tạp, cần tách rõ logic

**Bài học:**
- **EBITDA Gross = LN BDKD + Khấu hao + Lãi vay - Lãi tiền gửi** (4 thành phần)
- **EBITDA 30% = EBITDA Gross × 30%** (theo Nghị định 20/68/132)
- **Lãi vay không hợp lệ = MAX(Lãi vay - EBITDA 30%, 0)** khi EBITDA > 0
- Nếu EBITDA < 0 → Cảnh báo lỗ, không thể trích lãi

**Trường hợp T&T Biên Hòa:**
- Chi phí lãi vay = 0 đ (không vay nợ)
- → Không có K9 (rủi ro lãi vay không hợp lệ)
- → EBITDA 30% chỉ mục đích so sánh, không ảnh hưởng kết luận

**Áp dụng:** Khi không vay nợ, ghi rõ "Không rủi ro K9" để tránh nhầm lẫn. Nếu có lãi vay, kiểm tra 30% luôn

---

### 3. PDF Sao kê ngân hàng - Định dạng scan vs trích xuất

**Vì sao:** Không thể OCR/trích dòng từ PDF scan, cần biết cách xử lý

**Bài học:**
- **PDF trích xuất** (còn lớp chữ): Có thể `extract_text()` → tên NH + số TK → Đối chiếu công nợ 131/331
- **PDF scan** (toàn ảnh, 0 text): Không thể trích tự động, cần OCR hoặc manual
- T&T Biên Hòa: 5 PDF sổ phụ = scan (28 trang, 5 MB, 0 text) → Ghi "PDF chưa kết luận UNC"

**Lý do ghi "chưa kết luận":** Theo Q7, thiếu UNC = không kết luận TT06 (K6)

**Áp dụng:** 
- Kiểm tra PDF bằng `extract_text()` hoặc `pymupdf.Text()` trước
- Nếu 0 text: `PyTorch/PaddleOCR` hoặc request skill `tram-tai-lieu`
- Nếu có text: Trích tên NH + số TK tự động → Đối chiếu sổ 131/331

---

### 4. NCC rủi ro F0/F1 - Tra cứu danh sách

**Vì sao:** NCC F0 = cảnh báo cao, cần tìm đúng MST

**Bài học:**
- **F0** = Công ty trong CV cảnh báo rủi ro lớn (nợ thuế, hủy đơn, không hoạt động)
- **F1** = Công ty liên quan F0 (cùng địa chỉ, cùng chủ, v.v.)
- Tra từ 3 nguồn:
  1. Sổ CVCBRR (file Excel theo dõi trên Google Drive hoặc brain "Cập nhật công văn cảnh báo")
  2. Sổ Công an (file Excel tại `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx`)
  3. Danh bạ ĐKKD (file Excel tại thư mục hồ sơ `3.2. Danh ba trang thai NNT`)

**Trường hợp T&T Biên Hòa:**
- NCC F0: NNM Group (MST 0317409384, theo CV 6174)
- Phát hiện: 2 HĐ mua từ NNM, tổng 2.442.000 đ
- Kết luận: HD01 vượt ngưỡng (Điểm 3)

**Áp dụng:** 
- Lần sau check **cột MST ở IN07_NCC_CanhBao** (tự động lấy từ sổ CVCBRR + CA)
- Nếu thấy flag F0/F1: Tra ngược danh sách để ghi CV number, ngày cảnh báo, nội dung

---

### 5. HĐ bán dưới giá - Công thức phát hiện

**Vì sao:** Lệch giá vốn-doanh thu nhỏ từng dòng nhưng lũy tích lớn, cần cơ chế phát hiện

**Bài học:**
- **Bán dưới giá = DT < GV** trên 1 HĐ cụ thể
- Công thức: `IF(511_DT < 632_GV, "Có", "Không")` + tính lỗ = GV - DT
- Đối chiếu theo: HĐ + Ngày + Mặt hàng + Số lượng

**Trường hợp T&T Biên Hòa:**
- HĐ số 6, ngày 23/01/2024: DT = 12.500.000 đ, GV = 13.245.000 đ
- Lỗ = 745.000 đ
- Kết luận: DT01 vượt (Điểm 2)

**Áp dụng:**
- Script kiểm tra tự động: So 511 vs 632 từng dòng, flag lỗ → Sheet `KQ2_511_632`
- Khi phát hiện: Ghi rõ HĐ, ngày, dòng hàng, số lượng, giá mua, giá bán, lỗ

---

### 6. Định dạng Excel - Các quy tắc không được quên

**Vì sao:** Lần đầu xuất file nhiều lỗi định dạng, phải tái chạy

**Bài học - Luật Q14-Q19 của skill:**
- **Font:** Times New Roman, size 11, encode Unicode (nếu font sai mã VNI/TCVN3 → fix)
- **Ngày:** `dd/MM/yyyy` (VD: 23/10/2026), KHÔNG hiện giờ phút
- **Số tiền:** Dấu chấm `.` ngăn hàng nghìn, dấu phẩy `,` cho thập phân (VD: 123.456,78)
- **Cột width:** Tự giãn theo nội dung (không để text tràn)
- **Dòng ví dụ:** Dòng đầu tiên dưới header là ví dụ → Xóa khi chạy thật
- **Dòng rỗng:** Nếu năm không phát sinh, xóa dòng đó (không giữ dòng toàn số 0)
- **Merge cell:** Không dùng, ngoại trừ header tiêu đề

**Trường hợp T&T Biên Hòa:**
- Lần v4 có lỗi: Chữ Unicode bị mã TCVN3, dòng tháng rỗng, IN07 trùng
- Lần v5: Sửa lại, định dạng chuẩn

**Áp dụng:** 
- Trước khi save, quét tất cả sheet: Font, ngày, số tiền, width, xóa dòng rỗng
- Dùng openpyxl `set_font()`, `number_format = '#,##0.00'` khi ghi dữ liệu

---

### 7. Tái dùng mẫu - Không tái tính từ đầu

**Vì sao:** 6 mẫu v1.4 đã chuẩn, không cần viết lại

**Bài học:**
- Lấy từ `MAU_FILE_TONG_HOP/`: KQ1.1, KQ1.2, KQ1.3, KQ2.1, KQ2.2, KQ3
- **Mỗi mẫu đã có:**
  - Header (tiêu đề, MST, kỳ, ngày chạy)
  - Sheet layout (cột sắp xếp, công thức, định dạng)
  - Dòng ví dụ (xóa đi khi chạy thật)
- Chỉ **điền dữ liệu** từ IN00-IN14, **không sửa layout**

**Lỗi tránh:**
- Không copy sheet cũ (v1.2, v1.3) vào mẫu mới
- Không xóa dòng ví dụ mà để nó trong file giao
- Không thêm cột riêng của riêng mình, phải dùng sheet có sẵn trong mẫu

**Áp dụng:** 
- Bước 1: Copy 6 mẫu v1.4 sang từng hồ sơ mới
- Bước 2: Nhập dữ liệu vào (không sửa cấu trúc)
- Bước 3: Đặt tên file theo quy chuẩn, lưu vào `4. KET QUA PHAN TICH`

---

### 8. Không có sao kê / GDLK - Cách ghi nhãn

**Vì sao:** Không ghi "thiếu" hoặc "chưa có", cần cách ghi rõ từng trường hợp

**Bài học:**
- **PDF scan chưa OCR:** Ghi "PDF - chưa kết luận" (nhãn Q7)
- **Chưa nộp từ NNT:** Ghi "Chưa nhận" (trong KQ1.3 phần "Hồ sơ NNT còn thiếu")
- **Không có chứng từ:** Ghi "Không phát sinh" (nếu năm đó không có giao dịch)

**Trường hợp T&T Biên Hòa:**
- IN11 (Sao kê): "PDF chưa trích" (5 sổ phụ là scan)
- IN12 (GDLK): "Chưa nộp" (công ty chưa gửi phụ lục)

**Áp dụng:** 
- Khi quyết định không kết luận TT06: `Q7 = "PDF - chưa kết luận"` trong KQ2.2
- Khi hồ sơ cần bổ sung: Liệt kê trong KQ1.3 phần "Đề nghị bổ sung"

---

### Tóm lại - Checklist lần sau

- [ ] **IN04:** Quét dòng (không phải cột) để lấy KQKD 50, CDKT 141, TNDN ct16 cho 4 năm
- [ ] **EBITDA:** Tính từ 4 chỉ tiêu, áp dụng 30%, kiểm tra K9 nếu có lãi vay
- [ ] **PDF sao kê:** Test lớp chữ trước, nếu scan → ghi "PDF chưa kết luận"
- [ ] **NCC F0/F1:** Tra 3 nguồn (CVCBRR, CA, ĐKKD), ghi CV number + ngày
- [ ] **HĐ bán dưới giá:** Quét tất cả từ 511-632, flag khi DT < GV
- [ ] **Định dạng:** Font Times, ngày dd/mm/yyyy, số tiền 123.456,78, tự giãn cột
- [ ] **Mẫu:** Dùng 6 tệp v1.4, chỉ điền dữ liệu, không sửa layout
- [ ] **Ghi nhãn:** "PDF - chưa kết luận" / "Chưa nộp" / "Không phát sinh" (rõ ràng)
