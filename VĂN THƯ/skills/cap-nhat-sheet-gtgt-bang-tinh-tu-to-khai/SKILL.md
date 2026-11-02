---
name: "Cập nhật sheet GTGT Bảng tính từ tờ khai GTGT quý"
description: "Đổ số tờ khai GTGT quý vào sheet GTGT của Bảng tính vụ kiểm tra; khớp số dư khấu trừ chuyển kỳ, xóa quý chưa đến hạn, vá merge cell."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-10-27
---
## Khi nào dùng
- Người dùng gửi tờ khai GTGT mới (vd file BCT-<mã DN> theo quý) và yêu cầu cập nhật lại sheet GTGT của `Bang tinh.xlsx` hồ sơ vụ kiểm tra.
- Sau khi đã thêm khung năm mới cho sheet GTGT nhưng chưa có số khai thật.
- Việc này chạy thẳng trong lượt, không đẩy chạy nền, với vai trò trợ lý hồ sơ kiểm tra thuế (không qua Phòng Văn Thư vì không phải công văn).

## Chuẩn bị
- Tìm thư mục vụ của DN trên Drive (hồ sơ vụ nằm ở Drive, không có trong vault), lấy `Bang tinh.xlsx` và các tờ khai GTGT từng quý.
- Tải về thư mục tạm riêng; đọc cấu trúc sheet GTGT (dòng từng quý, cột chỉ tiêu, ô số dư khấu trừ chuyển kỳ).

## Quy trình
1. Với mỗi quý có tờ khai: ghi giá trị HHDV bán ra, thuế đầu ra, thuế GTGT đầu vào được khấu trừ, điều chỉnh tăng [38] và điều chỉnh giảm [37] đúng theo tờ khai.
2. Quý không phát sinh doanh thu: vẫn ghi khấu trừ đầu vào nếu có, để công thức chuyển kỳ chạy đúng.
3. Quý chưa đến hạn nộp (vd Q4 trước hạn 20/01 năm sau): nếu đang mang số copy từ năm trước thì xóa, để TRỐNG, không điền 0 giả.
4. Kiểm cột phải nộp / khấu trừ chuyển kỳ của từng dòng: nếu ô bị merge làm mất công thức, gỡ merge và khôi phục công thức theo mẫu các quý khác.
5. Tính lại workbook, so số dư khấu trừ cuối mỗi quý với tờ khai; phải khớp tuyệt đối.
6. Ghi đè file lên Drive, xác minh bằng metadata/md5 mới thay vì tin tool báo thành công.

## Bẫy
- Khung năm mới hay được copy từ năm trước nên mang số giả ở quý chưa khai.
- Merge cell trong dòng quý làm đứt chuỗi khấu trừ sang quý sau mà không báo lỗi.
- Số dư đầu kỳ Q1 kế thừa từ chuỗi công thức năm trước; không có tờ khai Q4 năm trước thì ghi rõ là chưa xác minh.
- Không động tới sheet lương/TNCN khi người dùng chưa yêu cầu.

## Kiểm chứng
- Số dư khấu trừ cuối từng quý khớp 100% tờ khai.
- Checksum file trên Drive khớp bản local sau khi ghi.
- Báo rõ điểm chưa xác minh (đầu kỳ Q1) và nguồn còn thiếu.
