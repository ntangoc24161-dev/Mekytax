---
type: decision
created: 2026-11-09
updated: 2026-11-09
source: chat 2026-11-09 (2607e2ee)
---
Trưởng phòng chốt 8 quy ước vận hành dây chuyền xử lý công văn cảnh báo (02:00 09/11/2026), đã ghi vào workflow `day-chuyen-xu-ly-cv-canh-bao`, skill `theo-doi-cv-canh-bao-rui-ro`, agent `tra-cuu-tai-cv-ioffice` và agent QC:

1. **Hộp thư đến chính thức:** `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\HO SO MOI`. Đường dẫn cũ `01_HO SO CONG AN CHUYEN DEN\2026\HO SO MOI` đã không còn tồn tại.
2. **Phạm vi mặc định:** lấy TẤT CẢ tệp trong thư mục, không lọc theo số CV.
3. **Phân nhánh lưu trữ theo NGUỒN GỐC, không theo cơ quan ký:** hồ sơ có gốc Công an thì vào `01_HO SO CONG AN CHUYEN DEN` kể cả khi cơ quan thuế ký chuyển tiếp. CV 7268/CT-Ktr ngày 30/9/2026 của Cục Thuế có gốc là CV B1246/CV-VPCQCSĐT-Đ2 ngày 29/8/2026 của CSĐT Công an Hà Tĩnh, nên đi nhánh Công an.
4. **Mẫu tên:** `<Năm>_<Số CV>_<CO QUAN BAN HANH>_<SO DN F0>_<HAU TO>`. `SO DN F0` là số lượng doanh nghiệp bên bán (ví dụ 28). Dấu `/` trong số CV đổi thành `_`. Hậu tố bắt buộc để phân biệt tệp cùng công văn: CV, TT, TRIENKHAI, DS, PL.
5. **Xoá tệp nguồn** trên ổ G sau khi copy sang đích và kiểm dung lượng khớp.
6. **Excel là nguồn chuẩn:** lô có tệp .xlsx danh sách doanh nghiệp thì lấy MST, tên, địa chỉ, ngành nghề từ Excel; OCR chỉ dùng lấy SoCV, NgayCV, DonViBH, NoiDung. Lý do: OCR tiếng Việt đang mất dấu nặng.
7. **F1 bỏ qua hóa đơn không có MST bên mua**, không nạp dòng trống MSTMua, phải đếm và báo số dòng đã bỏ.
8. **Khoá lọc trùng:** F0 theo `SoCV + MSTBan`; F1 theo `SoCV + MSTBan + MSTMua`.

**Cấu trúc sổ thật (kiểm trực tiếp):** F0 17.075 dòng / 19 cột; F1 3.761 dòng / 25 cột; tiêu đề dòng 5-6, khoá máy dòng 7, dữ liệu từ dòng 8. Các số 17.333 / 16.952 / 19.638 / 20.050 / 7.282 là số cũ, không dùng làm mốc.

**Chưa chốt:** phương án OCR cho tệp scan không có lớp chữ (kết nối PaddleOCR đang TẮT, bản OCR hiện tại mất dấu).

**Vì sao:** lần chạy 00:38 09/11/2026 vỡ vì đường dẫn sai, OCR rác và các quy ước đặt tên, lọc trùng chưa được định nghĩa, trợ lý nền không có ai để hỏi.
**Áp dụng:** mọi lô hồ sơ mới chạy theo 8 quy ước này, không hỏi lại.
