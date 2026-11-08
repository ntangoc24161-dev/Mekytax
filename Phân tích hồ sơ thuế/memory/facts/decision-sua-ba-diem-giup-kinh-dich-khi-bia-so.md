---
type: decision
created: 2026-11-08
updated: 2026-11-08
---
Ngày 08/11/2026 người dùng yêu cầu sửa 3 chỗ rủi ro bịa thông tin trong hồ sơ 6 vai:

**1. Nâng model Kết luận rủi ro từ haiku -> sonnet**
- Vai này là duy nhất được quyền kết luận mức rủi ro 80 quy tắc, đáng dùng model mạnh hơn.
- Sửa `agents/phan-tich-vien-so-sach-dn.md`: model haiku → sonnet.

**2. Sửa câu lệnh file giải trình NNT để bỏ áp lực điền cho đầy**
- Câu cũ: "Xuất `KQ3_NNT_Giai_trinh_Chi_tiet_<MST>_<KY>.xlsx` đủ 10 sheet có số liệu thật (không bảng trống)."
- Câu mới: "Xuất `KQ3_NNT_Giai_trinh_Chi_tiet_<MST>_<KY>.xlsx` đủ 10 sheet; sheet nào không có nguồn thì ghi `KHÔNG_ĐỦ_DỮ_LIỆU` chứ không bỏ trống."
- Câu cũ dễ khiến mô hình "bịa số để không bỏ trống"; câu mới rõ là ghi KHÔNG_ĐỦ_DỮ_LIỆU thay vì bỏ.
- Sửa `agents/phan-tich-vien-so-sach-dn.md` mục 6 và bỏ "để file giải trình trống số" ra khỏi phần Cấm.

**3. Thêm bước soát số cho Trưởng phòng điều phối**
- Thêm bước 4 mới (giữa Lệnh 2 và Lệnh 3) trong `agents/truong-phong-phan-tich-ho-so.md`:
- Chọn ngẫu nhiên 5-10 con số trọng yếu trên KQ2.1 (NCC F0/F1, TNCN, EBITDA) và KQ2.2 (511-632 chi tiết, công nợ 131/331, chi tiền mặt 111), dò lại về biểu IN gốc.
- Nếu lệch hay không giải thích được, yêu cầu Phân tích viên làm lại file; không trả Lệnh 3 cho đến xong.
- Lý do: KQ2 đi thẳng vào KQ3 mà chỉ 10 dòng điểm cao nhất được tự kiểm; bước này giúp bắt lỗi sớm trước khi kết luận.

**Vì sao:** ba chỗ này là rủi ro cao nhất theo phân tích hôm trước:
- Lớp chặn mềm (chỉ câu dặn prompt) dễ bị vượt qua.
- Model nhỏ (haiku) + task nặng (80 quy tắc) = dễ bịa.
- File giải trình 10 sheet tạo áp lực "điền cho đầy" → bịa số.

**Áp dụng:** ba sửa này áp dụng ngay với tất cả hồ sơ sắp tới. Thân mệnh của hai vai Đối chiếu dữ liệu thuế (KQ2.1) và Soát sổ sách công ty (KQ2.2) không đổi, nhưng người sử dụng có thêm 1 bước kiểm số ở Trưởng phòng trước khi xuống GĐ3.

**Cần theo dõi:** sau vài hồ sơ, hỏi xem bước soát số này bắt được lỗi không, để đánh giá lại chiến thuật.
