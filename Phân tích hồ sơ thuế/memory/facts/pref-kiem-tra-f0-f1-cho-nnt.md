---
type: preference
created: 2026-11-11
updated: 2026-11-18
source: chat 2026-11-11 (2b1155ef); chat 2026-11-18 (231315934afc455090dd4603b8b23c3d)
---
## Quy tắc đối chiếu F0/F1 cho người nộp thuế (NNT)

Trong phân tích hồ sơ kiểm tra thuế, **kiểm tra tình trạng rủi ro của chính NNT đang phân tích** là bước đầu tiên, riêng biệt với rà soát NCC.

### Thứ tự kiểm (bắt buộc)

1. **Kiểm F0 trước**: NNT có nằm trong sheet `F0` của `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` không?
   - **Có** → Rủi ro **NNT tự thân cao** (công ty bị cảnh báo). Ghi vào **KQ1.3** (thông báo), **KQ3** (mức rủi zo cao từ đầu), có số công văn cảnh báo.
   - **Không** → Sang bước 2.

2. **Kiểm F1 tiếp**: NNT có nằm trong sheet `F1` của `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` không?
   - **Có** → **HD01 (Mua hóa đơn từ F0)**, ghi vào **KQ2.1** `IN07_NCC_CanhBao` Phần 1 (đánh giá rủi zo NNT).
   - **Không** → NNT an toàn, chuyển sang rà soát NCC của NNT theo thứ tự ưu tiên (CA → CVCBRR → Chung → BK → HĐĐT → ĐKKD).

### Cách ghi kết quả

- **KQ1.3** (`Kiem_ke`): thêm dòng "NNT nằm F0 - Công văn cảnh báo [số]" (nếu có).
- **KQ2.1** `IN07` **Phần 1** (Đánh giá rủi zo NNT): ghi rõ "NNT [tên] (MST [mst]) nằm F1 - đã mua HĐ từ F0 [tên F0]" → phát hiện **HD01** mức **Cao**.
- **KQ3** (Kết luận): nếu NNT là F0, mục "Rủi zo chủ yếu của NNT" ghi ngay "Công ty bị cảnh báo rủi zo thuế" mức **Cao**.

### Nguồn dữ liệu chuẩn

- Sheet `F0`: công ty bị cảnh báo từ sổ CVCBRR hoặc công an.
- Sheet `F1`: công ty mua HĐ từ F0.
- Nếu không thấy NNT trong cả 2 sheet → để trống, NNT không nằm F0/F1.
- Nếu sổ cũ quá **30 ngày** → cảnh báo "nguồn có thể cũ" nhưng vẫn ghi.

### Phân biệt với rà soát NCC

- **F0/F1 của NNT** = kiểm **tình trạng rủi ro của chính doanh nghiệp đang phân tích** → mục tiêu đánh giá tổng thể hồ sơ.
- **F0 của NCC** = bên bán nằm sheet F0 hoặc sổ Công an.
- **F1 của NCC** = chính NCC đó là bên mua trên sheet F1 (tầng hai). Xếp Cao cho rủi ro NCC / HD03. **Không kích HD01** chỉ vì NCC là F1.

**HD01 mức Cao khi:** NNT nằm F1 (đã mua HĐ F0) **hoặc** NNT mua từ NCC F0/CA. Hai nhánh này khác nhau; ghi rõ nhánh nào.

### Ánh xạ cột sổ CVCBRR (bắt buộc)

- Sheet `F0` / `F1`: dữ liệu từ `min_row=8`. Hàng 5-7 là tiêu đề nhóm.
- **F0:** MST bên bán = cột B (index 1), tên bán = C, số CV = D, ngày CV = E.
- **F1:** MST bên bán F0 = cột B (index 1), MST bên mua F1 = cột I (index 8), tên mua = H. Một số dòng nguồn CA để trống MST bán.
- `ncc_flag`: CA → F0 CVCBRR → F1 (keyed theo MST người mua). Vừa F0 vừa F1 thì ghi F0.

### Phần 1 IN07 bắt buộc kể cả kết quả âm

Khi NNT không F0/F1/CA vẫn ghi một dòng Phần 1: đã tra, kết quả âm, nguồn sổ. Không bỏ Phần 1.

### Danh bạ `Danh ba NNT.xls`

File 3.2 có thể là SpreadsheetML (`<?xml` + `mso-application progid="Excel.Sheet"`), không phải BIFF. `xlrd` báo Expected BOF. Đọc bằng xml.etree, tôn trọng `ss:Index`.

Hàng 1 = tiêu đề file, hàng 2 = header. Cột: MST = 4, tên = 5, ngành = 7, địa chỉ = 8, giám đốc = 19, ngày cấp GCN ĐKKD = 33, trạng thái ĐKT tổ chức = 45, lý do đóng tổ chức = 53, ngày đóng tổ chức = 54, ngày bắt đầu KD = 65. **Không lấy cột 3** (mã NNT nội bộ). Chi nhánh 13 số gom về MST 10 số; ưu tiên bản đang hoạt động.

**Vì sao:** NNT là F0 là rủi ro NNT tự thân (mục tiêu kiểm tra lớn), khác HD01 (mua HĐ rủi ro từ F0). Bài học Kiến Thành Tân 18/11/2026: v1 gắn F1 theo MST người mua rồi dán lên NCC bên bán, kích HD01 Cao dù F0 NCC = 0 và NNT không F1.

**Áp dụng:** Mọi hồ sơ sau này kiểm F0 NNT → F1 NNT → CA NNT rồi mới rà NCC. IN07 luôn có Phần 1. HD01 không kích vì NCC F1. Tra danh bạ SpreadsheetML đúng cột. GĐ2 giữ 100% HĐ trên `KQ2_NCC_mua_vao` (Bảng kê / HĐĐT / Chung).
