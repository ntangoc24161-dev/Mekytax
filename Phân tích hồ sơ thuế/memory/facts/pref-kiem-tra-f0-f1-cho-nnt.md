---
type: preference
created: 2026-11-11
updated: 2026-11-11
source: chat 2026-11-11 (2b1155ef)
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

- **F0/F1 của NNT** = kiểm **tình trạng rủi zo của chính doanh nghiệp đang phân tích** → mục tiêu đánh giá tổng thể hồ sơ.
- **F0 của NCC** = rà soát **nhà cung cấp của NNT** có nằm F0/F1 không → mục tiêu kiểm chứng chứng từ, HD01-HD04.

**Vì sao:** NNT là F0 là rủi zo NNT tự thân (mục tiêu kiểm tra lớn), khác với HD01 (mua HĐ rủi zo từ NCC).
