---
type: wiki
status: active
tags: [wiki, quy-trinh, cong-van-cong-an]
created: 2026-11-08
updated: 2026-11-08
source: [[quy-trinh-phoi-hop-va-dieu-phoi-cac-tro-ly-cung-cap-ca]]
---

# Quy trình xử lý Công văn Công an - 4 bước (2026)

**Sơ đồ luồng:**
```
[Công văn Công an đến]
        ↓
[Bước 1: BÓC TÁCH] (khai-thac-du-lieu-ho-so)
        ├→ [KIỂM SOÁT 1: Nghiệm thu bóc tách]
        ↓
[Bước 2: PHÂN CÔNG] (Theo dõi & Phân công công văn)
        ├→ [KIỂM SOÁT 2: Nghiệm thu in phân công]
        ↓
[Bước 3: SOẠN THẢO] (soan-thao-tra-loi-cong-an)
        ├→ [KIỂM SOÁT 3: Nghiệm thu dự thảo Word]
        ↓
[Bước 4: QUẢN LÝ HỒ SƠ] (quản-lý-hồ-sơ-công-an)
        ├→ [KIỂM SOÁT 4: Nghiệm thu hoàn tất]
        ↓
[HỒ SƠ HOÀN THÀNH]
```

**Trợ lý trưởng quy trình:** `Kiểm soát & Điều phối quy trình` (kiểm tra mỗi bước)

---

## Bước 1: BÓC TÁCH (Khai thác dữ liệu hồ sơ)

**Agent:** `khai-thac-du-lieu-ho-so` (Bóc tách dữ liệu)

**Input:** Công văn PDF từ thư mục CV MOI

**Công việc:**
1. Tiếp nhận công văn → xác định số hiệu, ngày, cơ quan ban hành
2. Nếu 1 file PDF có nhiều CV → bóc tách riêng từng cái, cấp STT riêng
3. Trích xuất nội dung → lưu file `.md` tại `D:\HSO CONG AN FILE MD`
4. Tra cứu DB NNT (`DB NNT.xls` sheet `NNT`):
   - Lấy địa chỉ, Phường/Xã, tình trạng pháp lý
   - Nếu MST không có trong DB → tra fallback `masothue.com`, ghi chú "cần xác minh"
5. **Phân Tổ KT:**
   - Chuẩn hóa địa chỉ theo Phường/Xã
   - Vĩnh Cửu → **KT2** (bắt buộc)
   - CQ khác (ngoài TCS3) → ghi **tên CQ**, không phân tổ nội bộ
   - CQ khác (trong TCS3) → **KT1 / KT3** theo địa bàn
6. **Ghi CBQL Tổ KT1:** Tra sheet `KTR1` để lấy tên CBQL (chỉ khi CV có DN thuộc TCS3)
7. **Cập nhật Excel:** Ghi tất cả vào sheet `NHAP_HO_SO` của `Bao_cao_cong_van.xlsx`
8. **Output:** Gói dữ liệu chuẩn → chuyển bước Phân công

**Kiểm soát chéo:** `Kiểm soát & Điều phối quy trình` (6 điểm kiểm tra)
- Số CV + ngày: khớp từng ký tự, không lỗi OCR
- Danh sách MST: đủ, không bỏ sót doanh nghiệp nào
- Phân Tổ KT: đúng địa bàn quản lý (Vĩnh Cửu=KT2 bắt buộc)
- CBQL KT1: tra sheet `KTR1`, không đoán
- Thư mục STT: đúng định dạng `STT. CV SO CV-CQ-TO KT`
- Excel: dòng chính/dòng phụ chuẩn

---

## Bước 2: PHÂN CÔNG (Theo dõi & Phân công công văn)

**Agent:** `Theo dõi & Phân công công văn`

**Input:** Gói dữ liệu bóc tách từ bước 1 → phản hồi xác nhận từ Kiểm soát bước 1

**Công việc:**
1. Rà soát công văn trạng thái "Mới" trên sổ
2. Tra cứu CBQL Tổ KT1 từ sheet `KTR1`
3. In toàn văn công văn **kèm dòng phân công** ở mép trên trang 1:
   - Nội dung dòng: Cán bộ, MST, Tên công ty, Ngày phân công
   - Font chữ co giãn vừa khít 1 dòng, không cụt chữ
4. **Tiêu chuẩn in:**
   - Tổng số trang của mỗi cán bộ = **số chẵn** (để in hai mặt)
   - Không lẫn cán bộ khác
5. Lưu file PDF tổng hợp lên Google Drive (thư mục quy định)
6. **Chốt trạng thái:** Đổi Excel từ "Mới" → "Đang xử lý"
7. **Output:** File in phân công + xác nhận trạng thái

**Kiểm soát chéo:** `Kiểm soát & Điều phối quy trình`
- Toàn văn 100% nội dung gốc
- Dòng phân công không bị cụt chữ hoặc đè
- Tổng trang mỗi cán bộ = số chẵn
- Tên file đúng chuẩn: `Phan cong xu ly KT1 tu <dong dau> den <dong cuoi> ngay in <dd-mm-yyyy>.pdf`
- Lưu trữ đúng Google Drive

---

## Bước 3: SOẠN THẢO (Soạn thảo trả lời Công an)

**Agent:** `soan-thao-tra-loi-cong-an`

**Input:** Gói dữ liệu từ bước 1 + xác nhận phân công từ bước 2

**Công việc:**
1. Lập **2 file Word dự thảo** (1 trang A4 mỗi loại):
   - `TRA LOI CV <Số CV>_<CQ>_du thao.docx` (công văn trả lời)
   - `to trinh <Số CV>.docx` (tờ trình lãnh đạo)
2. **Quyết định hướng kính gửi:**
   - Có công văn Cục Thuế chỉ đạo (DON-KTr1) → **kính gửi Cục Thuế TP**
   - Không → **kính gửi Cơ quan Công an gốc**
3. **Danh sách DN:**
   - Chỉ giữ DN thuộc **KT1, KT2, KT3** quản lý
   - Loại bỏ triệt để DN ngoài địa bàn TCS3
4. **Dữ liệu NNT:** Lấy live từ `DB NNT.xls` sheet `NNT`
   - Tên, trạng thái ĐKKD, số tài khoản, thông báo đóng MST
5. **Xử lý thiếu dữ liệu:**
   - Nếu thiếu nợ thuế, vi phạm, biên bản thanh tra → **tô đỏ** + `[CHƯA XÁC ĐỊNH]`
   - **TUYỆT ĐỐI KHÔNG BỊA** số liệu
6. **Thể thức tờ trình:** Đúng 1 trang A4, thứ tự 3 chữ ký (Ngọc - Giang - Xuân)
7. **Output:** 2 file Word dự thảo → lưu vào thư mục hồ sơ trên Drive

**Kiểm soát chéo:** `Kiểm soát & Điều phối quy trình`
- Hướng kính gửi: có DON-KTr1 hay không?
- Danh sách DN: chỉ KT1/KT2/KT3, loại bỏ ngoài địa bàn
- Dữ liệu NNT: lấy live từ DB, không sai lệch
- Thiếu dữ liệu: có tô đỏ + `[CHƯA XÁC ĐỊNH]` không?
- Thể thức Word: font 14, giãn dòng 1.15, 1 trang A4, 3 chữ ký đúng

---

## Bước 4: QUẢN LÝ HỒ SƠ (Quản lý hồ sơ công an)

**Agent:** `quản-lý-hồ-sơ-công-an`

**Input:** Xác nhận từ Kiểm soát bước 3 + file dự thảo Word từ bước 3

**Công việc:**
1. **Lưu trữ PDF gốc:**
   - Cấp **STT tiếp nối** từ STT cuối cùng trên sổ
   - Tạo thư mục mới tại `1. CUNG CAP CONG AN NAM 2026`:
     - Tên: `STT. CV <SỐ CV>-<CQ>-<TO KT>`
     - Ví dụ: `001. CV 5847-CSDT-KT2`
   - Chuyển PDF gốc vào thư mục (chỉ PDF, không .md)
2. **Tự động tạo dự thảo** (nếu có doanh nghiệp KT1):
   - Kích hoạt skill `tra-loi-cong-van`
   - Tạo 2 file Word chuẩn (công văn + tờ trình)
   - Lưu vào thư mục hồ sơ tương ứng
3. **Xuất danh sách 24 cột** (khi có yêu cầu):
   - Kích hoạt skill `xuat-danh-sach-ho-so-ca`
   - Xuất Excel 24 tiêu chí (tờ khai, BCTC, hóa đơn, nợ, vi phạm...)
   - Lưu tại `CV MOI\Danh_sach_tong_hop_ho_so_cung_cap_CV_MOI.xlsx`
4. **Rà soát & Đồng bộ hoàn tất:**
   - Tìm tệp ký số (`.signed.pdf` hoặc scan ký)
   - Kích hoạt skill `dong-bo-thu-muc-ca`
   - Cập nhật hậu tố `-R` vào tên thư mục: `001. CV 5847-CSDT-KT2-R`
   - Cập nhật Excel: trạng thái → "Đã hoàn thành"
5. **Output:** Hồ sơ hoàn tất, báo cáo xong

**Kiểm soát chéo:** `Kiểm soát & Điều phối quy trình` (Cổng 4)
- PDF gốc có lưu trữ không?
- Thư mục STT có đúng format không?
- Dự thảo có được tạo (nếu KT1)?
- Có tệp ký số không?
- Hậu tố `-R` được cập nhật không?
- Excel cuối cùng: trạng thái "Đã hoàn thành"

---

## Tiêu chí kiểm soát chéo (4 Cổng)

### Cổng 1: Nghiệm thu Bóc tách (Bước 1)
| Tiêu chí | Yêu cầu |
|---|---|
| Số CV + ngày | Trùng từng ký tự, không lỗi OCR, không nhầm số |
| Cơ quan ban hành | Xác định đúng Công an yêu cầu |
| Danh sách MST | Đủ số lượng DN, không bỏ sót |
| Phân Tổ KT | Đúng địa bàn (Vĩnh Cửu=KT2 bắt buộc) |
| CBQL KT1 | Tra sheet `KTR1`, không đoán |
| Thư mục STT | Đúng format `STT. CV SỐ-CQ-TO KT` |

### Cổng 2: Nghiệm thu In phân công (Bước 2)
| Tiêu chí | Yêu cầu |
|---|---|
| Toàn văn | 100% nội dung gốc |
| Dòng phân công | Font vừa 1 dòng, không cụt, không đè |
| Tổng trang | Mỗi cán bộ = số chẵn |
| Tên file | `Phan cong xu ly KT1 tu... den... ngay in dd-mm-yyyy.pdf` |
| Lưu trữ | Đúng Google Drive |

### Cổng 3: Nghiệm thu Dự thảo (Bước 3)
| Tiêu chí | Yêu cầu |
|---|---|
| Hướng kính gửi | Có DON-KTr1 → Cục Thuế TP; không → Công an |
| Danh sách DN | Chỉ KT1/KT2/KT3, loại ngoài TCS3 |
| Dữ liệu NNT | Live từ DB NNT.xls |
| Thiếu dữ liệu | Tô đỏ + `[CHƯA XÁC ĐỊNH]` |
| Thể thức tờ trình | 1 trang A4, 3 chữ ký (Ngọc-Giang-Xuân) |

### Cổng 4: Nghiệm thu Hoàn tất (Bước 4)
| Tiêu chí | Yêu cầu |
|---|---|
| PDF lưu trữ | Có tệp ký số (`.signed.pdf` hoặc scan) |
| Hậu tố `-R` | Thư mục được cập nhật `-R` |
| Excel trạng thái | "Đã hoàn thành" |

---

## Tự động hoá

### Loop tự động
- **Tên:** `ho-so-cung-cap-cong-an`
- **Lịch chạy:** Mỗi 1440 phút (1 ngày)
- **Trigger:** Workflow "Quản lý công văn công an" (4 bước)
- **Báo cáo:** Liệt kê CV xử lý xong (số CV, số DN, Tổ KT, CBQL)

### Workflow chính
- **Tên:** `quản-lý-công-văn-công-an`
- **Trạng thái:** Active
- **4 bước:** Bóc tách → Phân công → Soạn thảo → Quản lý hồ sơ
- **Verify agent:** `Kiểm soát & Điều phối quy trình` (mỗi bước)

---

## Nguyên tắc điều cấm

1. **Không bịa dữ liệu:** Nợ thuế, vi phạm, tài khoản, tình trạng pháp lý, cán bộ
2. **Không che giấu lỗi:** Nếu phát hiện sai lệch → lập danh sách cảnh báo rõ ràng
3. **Không chuyển giao bước sau nếu bước trước chưa vượt cổng:** Kiểm soát phải duyệt xong
4. **Không dùng em dash:** Luôn dùng dấu gạch ngang thường `-`

---

**Cập nhật lần cuối:** 2026-11-08 · [[quy-trinh-phoi-hop-va-dieu-phoi-cac-tro-ly-cung-cap-ca]]
