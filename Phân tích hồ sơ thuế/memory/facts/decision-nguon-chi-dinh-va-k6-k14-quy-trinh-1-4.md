---
type: decision
created: 2026-10-19
updated: 2026-10-19
---
Cán bộ chốt bổ sung các điểm còn thiếu trong "HƯỚNG DẪN THƯỜNG TRỊ" của quy trình 1.4 (`skills/phan-tich-ho-so-thue/SKILL.md` mục 7):

**NGUỒN CHỈ ĐỊNH (Q11):**
- F0/F1 (người bán, NNT): sheet `F0`/`F1` của `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` (kết quả brain Cập nhật công văn cảnh báo). Khớp [[pref-nguon-so-cvcbrr-chuan]].
- CA (công an): `Bao_cao_cong_van.xlsx`, kết quả brain Cung cấp thông tin công an, nằm theo từng hồ sơ.
- Danh bạ ĐKKD: thư mục `3.2. Danh ba trang thai NNT` trong từng hồ sơ.
- Ngày cập nhật sổ gần nhất lấy theo modified date của file CVCBRR; quá 30 ngày (mặc định tạm) thì cảnh báo nguồn có thể cũ.

**K12 (kỳ kiểm tra):** cán bộ chỉ định cùng tên công ty lúc ra lệnh là ưu tiên cao nhất; không nêu thì Javis tự suy theo khoảng thời gian hồ sơ cung cấp, chỉ hỏi khi nhiều kỳ mâu thuẫn không tự chọn được.

**K6/K7 (mẫu BCTC, chế độ kế toán):** xác định bằng cách đọc chính tiêu đề mẫu số in trên BCTC (B01-DN/TT200/TT99 hay B01a-DNN/TT133), không suy đoán theo quy mô DN. BCTC không ghi rõ mẫu số mới hỏi cán bộ.

**K8/K9/K13 (ngưỡng trọng yếu, thuế suất GTGT theo kỳ, số liệu ngành):** thứ tự lấy giá trị là (1) cán bộ chỉ định, (2) tham số sẵn có ở skill/sheet 05, (3) tra cứu trên mạng (WebSearch), (4) mặc định 2%. Riêng K9 (thuế suất GTGT) KHÔNG áp mặc định 2% vì là tỷ lệ luật định, phải tra đúng Nghị định/Thông tư hiệu lực theo kỳ.

**Vì sao:** các placeholder `<tên sổ, đường dẫn>` trong skill để trống khiến Q11 luôn phải ghi KHÔNG_ĐỦ_DỮ_LIỆU; cán bộ điền một lần để Javis chạy được ngay từ hồ sơ tiếp theo mà không hỏi lại.

**Áp dụng:** mọi lệnh phân tích hồ sơ thuế theo quy trình 1.4 (Lệnh 1-4) từ 2026-10-19 trở đi, trừ khi cán bộ chỉ định khác trong lệnh cụ thể.
