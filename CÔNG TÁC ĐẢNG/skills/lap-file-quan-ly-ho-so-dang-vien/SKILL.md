---
name: Lập file quản lý hồ sơ đảng viên
description: 'Hợp nhất nhiều danh sách đảng viên thành file .xlsx quản lý hồ sơ chi bộ: sheet tổng, tổ đảng, chuyển sinh hoạt, nhật ký hiệu chỉnh.'
group: Công tác Đảng
origin: javis-learned
status: active
created: 2026-11-02
updated: '2026-11-02'
learned_updated: '2026-11-02'
---
## Khi nào dùng
- User gửi nhiều danh sách đảng viên của cùng một chi bộ, lập ở các thời điểm khác nhau (DS xuất từ hệ thống, DS đề nghị cấp thẻ, file sai sót cần hiệu chỉnh, DS cập nhật phần mềm, DS phân tổ quản lý), rồi nhờ hiệu chỉnh và xuất một file quản lý hồ sơ chung.
- Ví dụ: "điều chỉnh thông tin đảng viên bị sai sót theo file mới, xuất danh sách quản lý hồ sơ, 1 sheet tổng, sheet chuyển sinh hoạt đảng".
- User báo không mở được file quản lý đã xuất và nhờ xuất lại bằng định dạng khác.

## Chuẩn bị
- Đọc các file đính kèm. Mặc định chỉ đọc, không lưu vào sources/.
- Đọc memory/facts của chi bộ để biết các bản ghi trùng, các hiệu chỉnh và quyết định user đã chốt trước đó.
- Nếu đã có file quản lý cũ trong `bao-cao/`, cập nhật file đó thay vì lập file mới từ đầu.

## Quy trình
1. Lấy danh sách mới nhất làm khung quân số. Khử trùng các dòng cùng họ tên, cùng ngày sinh mà số định danh chỉ lệch 1 chữ số.
2. Ghép các nguồn theo số định danh. Dòng nào thiếu số định danh thì ghép theo họ tên kèm ngày sinh, rồi đánh dấu cần xác minh.
3. Thứ tự ưu tiên khi các nguồn lệch nhau: file sai sót hoặc hiệu chỉnh mới nhất trước, sau đó đến DS cập nhật phần mềm, cuối cùng là DS cũ. Ghi giá trị của nguồn kia vào cột Ghi chú, không xoá đi.
4. Lập danh sách người có trong DS cũ mà không còn trong DS mới, đưa sang sheet chuyển sinh hoạt. Hỏi user lý do rời nếu chưa rõ.
5. Lấy tổ đảng, tổ chuyên môn, địa chỉ thường trú, số điện thoại từ DS phân tổ và DS cập nhật phần mềm.
6. Số thẻ đảng viên để 2 cột: số thẻ CŨ (theo quyết định cấp thẻ trước đây) và số thẻ MỚI (bằng số định danh cá nhân).
7. Xuất `.xlsx` vào `bao-cao/` gồm 4 sheet. Chỉ xuất `.xlsm` khi user cần macro thật:
   - TỔNG HỢP: quân số, phân bổ tổ đảng, tiến độ cập nhật, chất lượng dữ liệu, tính bằng công thức.
   - DS ĐẢNG VIÊN: đủ trường hồ sơ, bật lọc tự động, cố định dòng tiêu đề, sắp theo tổ đảng, tô màu ô thiếu dữ liệu, cột Ghi chú nêu việc cần làm cho từng người.
   - CHUYỂN SINH HOẠT ĐẢNG: kèm cột trống và danh sách thả xuống cho lý do, ngày chuyển, nơi đến, số quyết định.
   - NHẬT KÝ HIỆU CHỈNH: mỗi dòng ghi sửa gì và căn cứ vào nguồn nào.
8. Trả lời: số hiệu chỉnh đã áp dụng, quân số chốt, các khoảng trống dữ liệu còn lại, việc cần user xác minh, link file tương đối.

## Bẫy
- Cột số thẻ trong DS xuất từ hệ thống có thể bị điền nhầm bằng số định danh ở mọi dòng. Kiểm định dạng trước khi dùng.
- Số thẻ cũ không phải lúc nào cũng có cùng tiền tố (phần lớn 42xxxxxx, nhưng cũng có 40..., 45..., 66...). Giữ nguyên như nguồn, đừng tự sửa về một dạng.
- Cùng một người có thể bị ghi tên khác dấu (ví dụ Đương và Đượng) hoặc có hậu tố năm sinh để phân biệt người trùng tên. Ghép theo số định danh, không chỉ theo tên.
- Ngày vào Đảng lệch giữa các nguồn thường do đảo ngày với tháng. Ca lệch năm thì nêu riêng để user tra quyết định gốc.
- Người mới kết nạp có thể chưa được cấp thẻ cũ và chưa có ngày chính thức. Đánh dấu, đừng bịa ra.
- File .xlsm Javis tự tạo (không nhúng macro) đã có lần user không mở được trên máy (02/11/2026). Vì vậy mặc định xuất .xlsx. Khi xuất lại từ file .xlsm cũ, giữ đủ 4 sheet, công thức, lọc, tô màu và danh sách thả xuống.

## Kiểm chứng
- Tổng số người theo các tổ đảng cộng với số chưa phân tổ phải bằng quân số chốt.
- Quân số DS cũ trừ số người rời, cộng số người mới phải bằng quân số chốt.
- Mở lại file vừa ghi, kiểm số dòng từng sheet trước khi dẫn link.

## Lịch sử (tự học)
- [2026-11-02] Skill cũ bắt xuất .xlsm và ghi 'xlsm không macro vẫn hợp lệ', nhưng user báo không mở được file Quan-ly-ho-so-dang-vien_CB13_02-11-2026.xlsm và yêu cầu xuất lại .xlsx. Đổi định dạng mặc định sang .xlsx, thay bẫy cũ bằng bẫy mới.
