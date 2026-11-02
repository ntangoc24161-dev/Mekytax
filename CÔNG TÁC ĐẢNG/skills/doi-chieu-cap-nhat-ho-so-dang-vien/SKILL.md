---
name: "Đối chiếu cập nhật hồ sơ đảng viên"
description: "Đối chiếu DS đảng viên với DS cập nhật hồ sơ PM 4.0 theo số căn cước, xuất Excel người chưa gửi để đôn đốc."
group: "Công tác Đảng"
origin: javis-learned
status: active
created: 2026-11-02
updated: 2026-11-02
learned_updated: 2026-11-02
---
## Khi nào dùng
- User gửi danh sách đảng viên của chi bộ và danh sách đã cập nhật hồ sơ (file xlsx), rồi nhờ lọc ra người chưa cập nhật để đôn đốc.
- User gửi bản DS cập nhật PM 4.0 mới, nhờ đối chiếu lại tiến độ.
- Ví dụ: "căn cứ danh sách đảng viên và danh sách đã cập nhật hồ sơ, tạo file danh sách đảng viên chưa cập nhật cần đôn đốc".
- Việc lập file quản lý hồ sơ đầy đủ (4 sheet, tổ đảng, số thẻ cũ/mới) thuộc skill `lap-file-quan-ly-ho-so-dang-vien`, không làm trong skill này.

## Chuẩn bị
- File đính kèm: danh sách đảng viên (ví dụ `<ngày>_Danhsachdangvien <chi bộ>.xlsx`) và danh sách yêu cầu cập nhật hồ sơ (ví dụ `DS_Yêu cầu cập nhật hồ sơ đảng viên (n).xlsx`).
- Nếu đã có file quản lý hồ sơ trong `bao-cao/` (ví dụ `Quan-ly-ho-so-dang-vien_CB13_02-11-2026.xlsx`), lấy khung quân số từ sheet DS ĐẢNG VIÊN. Không lấy người ở sheet CHUYỂN SINH HOẠT ĐẢNG.
- Đọc memory/facts của chi bộ: bản ghi trùng, người đã chuyển sinh hoạt, các quyết định user đã chốt.
- Mặc định chỉ đọc file đính kèm. Không lưu vào sources/ khi user chưa yêu cầu.

## Quy trình
1. Đọc file nguồn và xác định cột họ tên, ngày sinh, số định danh/căn cước, chức danh, số điện thoại, ngày vào Đảng, trạng thái hồ sơ, tổ đảng.
2. Lọc danh sách cập nhật theo đúng chi bộ cần xét.
3. Khử trùng danh sách đảng viên: các dòng cùng họ tên và cùng ngày sinh, số định danh chỉ lệch 1 chữ số, gần như chắc là nhập trùng. Tính mỗi người 1 lần và ghi lại để báo user.
4. Ghép theo số căn cước. Dòng nào thiếu số định danh thì ghép theo họ tên (kèm ngày sinh nếu có) và đánh dấu cần xác minh.
5. Loại khỏi danh sách đôn đốc những người đã sang sheet chuyển sinh hoạt. Họ không còn trong quân số đang công tác.
6. Đếm tổng số người đang công tác, số đã gửi, số chưa gửi và tỷ lệ %. Thống kê trạng thái hồ sơ của những người đã gửi (đã duyệt, có kiến nghị điều chỉnh...).
7. Nếu đã có file quản lý hồ sơ: cập nhật cột "Cập nhật PM 4.0" trên sheet DS ĐẢNG VIÊN cho khớp kết quả mới. Không lập file quản lý mới từ đầu. Chỉ ghi sheet NHẬT KÝ HIỆU CHỈNH khi lần này còn sửa trường hồ sơ, không ghi nhật ký chỉ vì đổi trạng thái đã gửi/chưa gửi.
8. Xuất `.xlsx` vào `bao-cao/DS-dang-vien-chua-cap-nhat-ho-so_<chi bộ>_<dd-mm-yyyy>.xlsx`:
   - Sheet "Chua cap nhat": STT, họ tên, tổ đảng (nếu có), chức danh, số điện thoại, ngày sinh, số định danh, ngày vào Đảng, cùng 2 cột trống "Đã đôn đốc (ngày)" và "Ghi chú" để user điền tay.
   - Sheet "Tong hop": bảng tiến độ.
9. Trả lời: con số chính, danh sách cần đôn đốc (ghi theo tổ nếu đã phân tổ), link file tương đối.

## Bẫy
- Số dòng trong danh sách gốc có thể lớn hơn số người thật vì nhập trùng. Báo tổng sau khi khử trùng và nêu rõ các bản ghi trùng.
- Người đã chuyển sinh hoạt vẫn có thể còn tên trong DS gốc hoặc DS phân tổ. Không đưa vào danh sách đôn đốc. Lần xuất đầu của CB13 có 14 người vì còn Trần Văn Đương; sau khi chuyển sinh hoạt còn 13.
- Nhiều người không có số điện thoại trong DS gốc. Lấy thêm từ DS cập nhật PM 3.0 hoặc DS phân tổ nếu có; vẫn trống thì ghi "chưa có".
- Tổ trưởng hoặc phó tổ trưởng có tên trong danh sách chậm thì nên nêu riêng, vì họ còn phải đôn đốc người khác.
- Khi viết số định danh trong câu trả lời, chép đúng từ file, không gõ lại theo trí nhớ.
- Nguồn chuẩn tiến độ là DS xuất từ phần mềm (`DS_Yêu cầu cập nhật hồ sơ đảng viên`), đối chiếu bằng số căn cước, không suy từ file quản lý cũ.
- Mặc định xuất `.xlsx`. Không xuất `.xlsm` cho file đôn đốc.

## Kiểm chứng
- Số đã gửi cộng số chưa gửi phải bằng quân số đang công tác (đã khử trùng, đã loại người chuyển sinh hoạt).
- Mở lại file xlsx vừa ghi, kiểm số dòng ở sheet "Chua cap nhat" khớp với số chưa gửi trước khi dẫn link.
- Nếu đã sửa file quản lý hồ sơ, số "CHƯA GỬI" ở sheet TỔNG HỢP phải bằng số dòng sheet đôn đốc.

## Lịch sử (tự học)
- [2026-11-02] User chốt lấy DS cập nhật PM 4.0 làm chuẩn, đối chiếu bằng số căn cước. Bổ sung bước loại người chuyển sinh hoạt, bước cập nhật cột PM 4.0 trên file quản lý nếu file đó đã có, và cột tổ đảng trên sheet đôn đốc.
