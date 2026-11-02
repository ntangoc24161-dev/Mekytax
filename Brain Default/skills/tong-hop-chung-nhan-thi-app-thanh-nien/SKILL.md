---
name: Tổng hợp chứng nhận thi app Thanh niên Việt Nam
description: 'Tổng hợp ảnh chứng nhận app Thanh niên Việt Nam thành báo cáo đoàn viên đã thi: gộp ảnh trùng người, đếm phần thi, đối chiếu danh sách.'
group: Năng suất
origin: javis-learned
status: active
created: 2026-10-27
updated: '2026-10-27'
learned_updated: '2026-10-27'
---
# Tổng hợp chứng nhận thi app Thanh niên Việt Nam

## Khi nào dùng
- Người dùng gửi một loạt ảnh chứng nhận (thường qua Telegram, lưu ở `inbox/telegram/photo_*.jpg`) kèm câu như "Tổng hợp các hình trên thành báo cáo các đoàn viên đã hoàn thành phần thi trên app Thanh niên Việt Nam năm 2026".
- Người dùng gửi thêm ảnh không kèm lời nhắn sau một lượt tổng hợp trước đó: bổ sung vào báo cáo đang có.
- Cần đối chiếu người đã thi với danh sách tổng đoàn viên để biết ai còn thiếu.
- Người dùng yêu cầu "Xuất ra file excel", "Xuất excel đủ 27 ng", "Cập nhật theo mẫu trên" cho báo cáo đã lập.

## Chuẩn bị
- Liệt kê lại thư mục `inbox/telegram/` ngay lúc đó để biết chính xác các ảnh của đợt này. Đừng dựa vào trí nhớ lượt trước.
- Đọc fact `doan-vien-thi-app-thanh-nien-viet-nam-2026` TRƯỚC KHI báo cáo: nó giữ danh sách đầy đủ nhất đã tổng hợp (tính đến 27/10/2026 là 27 người, 26 đã thi, 99 lượt chứng nhận).
- Kiểm tra file báo cáo đã xuất gần nhất ở vault: `Bao_cao_Doan_vien_thi_Thanh_nien_Viet_Nam_2026_2026-10-27.xlsx` (3 sheet: Tong_quan, Chi_tiet_doan_vien, Can_bo_sung). Đây là bản mẫu và dữ liệu nền để cập nhật tiếp.
- File `inbox/telegram/Bao_cao_Doan_vien_thi_Thanh_nien_Viet_Nam_2026.xlsx` và các bản trùng ở `inbox/khach/` là bản CŨ lập 23/10/2026, chỉ 17 người. "Mục I" trong file đó là phần TỔNG QUAN, KHÔNG phải roster 32 đoàn viên. Chưa có file roster chính thức trong vault.

## Quy trình
1. Mở TỪNG ảnh của đợt, ghi lại: họ tên đoàn viên, tên từng phần thi/chứng nhận trên ảnh.
2. Gộp các ảnh của cùng một người (cùng một người có thể chụp ở 2 thời điểm khác nhau). Bỏ các chứng nhận trùng lặp.
3. Chuẩn hóa tên phần thi để so được với nhau, ví dụ:
   - Nghị quyết Đại hội Đoàn toàn quốc lần thứ XIII
   - Nghị quyết Đại hội Đoàn toàn quốc lần thứ XII (chính thức)
   - Nghị quyết Đại hội XIV của Đảng
   - Chuyên đề tư tưởng, đạo đức, phong cách Hồ Chí Minh 2026 (ĐVTN)
   - Chuyên đề tư tưởng, đạo đức, phong cách Hồ Chí Minh 2026 (CB Đoàn)
   - Chuyên đề tư tưởng Hồ Chí Minh 2024 (đoàn viên, thanh niên)
   - Các cuộc thi tuần (ví dụ Tuần 3 - tìm hiểu đồng chí Phạm Hữu Lầu)
4. Ảnh app báo "Không có dữ liệu hiển thị": vẫn ghi tên người đó vào báo cáo với 0 phần (chưa có dữ liệu), không bỏ qua.
5. GỘP kết quả đợt mới vào danh sách đầy đủ nhất đã có (fact + file đã xuất), không lập lại từ riêng ảnh của phiên hiện tại.
6. Lập báo cáo theo từng đoàn viên: tên, số phần đã hoàn thành, danh sách phần thi.
7. Cuối báo cáo ghi tổng: số đoàn viên có dữ liệu, số người chưa có dữ liệu, tổng lượt chứng nhận, số người đủ 4 phần cốt lõi (ĐVTN + CB Đoàn + NQ XIV Đảng + NQ Đoàn XIII).
8. Nếu có danh sách tổng: đối chiếu tên (xét cả sai khác dấu), chia nhóm đã đủ / còn thiếu phần / chưa làm.
9. Xuất Excel theo cấu trúc file đã xuất trước (Tong_quan, Chi_tiet_doan_vien, Can_bo_sung) khi người dùng yêu cầu.

## Bẫy
- Đếm sai số ảnh: ngày 27/10/2026 người dùng gửi ảnh trong nhiều tin nhắn, có lượt trả lời chỉ thấy 1 ảnh, lượt khác 4 ảnh, lượt khác ra 18 người, cho ra nhiều báo cáo lệch nhau. Luôn liệt kê lại thư mục và đọc đủ mọi ảnh của đợt trước khi báo cáo.
- Khi người dùng gửi lại đúng câu yêu cầu cũ, đó là yêu cầu làm lại cho đủ, không phải chỉ đọc ảnh mới nhất.
- Một đợt tổng hợp mới phải gộp với báo cáo đầy đủ nhất đã có, không báo cáo lại từ một tập con nhỏ hơn.
- Phiên khách/phiên nền (Telegram khách) chỉ thấy ảnh gửi trong phiên đó. Ngày 27/10/2026 phiên này nhiều lần trả lời "mới có 9 người, thiếu 18" và đòi người dùng gửi lại ảnh, trong khi danh sách 27 người và file Excel đã có sẵn trong vault. Luôn đọc fact và file đã xuất trước khi kết luận thiếu dữ liệu.
- Đừng tin "32 đoàn viên ở mục I" trong file Excel cũ: đó là tiêu đề phần tổng quan, không phải roster.
- Đừng báo "3/3 phần", "4/4", "5/5" nếu không biết tổng số phần bắt buộc. Ghi "hoàn thành 3 phần" thay vì đặt mẫu số tự đoán.
- Tên gần giống nhưng là 2 người khác nhau (ví dụ Bùi Thị Thủy và Bùi Thị Thúy Nga): không gộp.
- Tên khác dấu nhưng cùng người (ví dụ "Anh Ngọc" và "Ánh Ngọc"): gộp khi đối chiếu.
- Phiên bị khóa quyền ghi file không xuất được Excel. Khi đó nói rõ chưa xuất được, CHỈ RA file đầy đủ đã xuất trước đó (nếu có) để người dùng dùng ngay, đưa bảng dữ liệu đầy đủ (không phải tập con) dạng 1 dòng mỗi phần thi để dán vào Excel, và chỉ người dùng làm lại ở phiên chat chính.

## Kiểm chứng
- Số ảnh đã đọc bằng số ảnh trong thư mục của đợt.
- Số đoàn viên trong báo cáo không nhỏ hơn bản đầy đủ nhất đã lưu trong fact.
- Mỗi đoàn viên chỉ xuất hiện một lần trong báo cáo.
- Tổng lượt chứng nhận bằng tổng số phần thi cộng của từng người.
- Người "không có dữ liệu" có mặt trong báo cáo với số 0.

## Lịch sử (tự học)
- [2026-10-27] Hội thoại 27/10/2026 lộ thêm 2 chỗ skill cũ thiếu: (1) yêu cầu 'Xuất ra file excel' bị chặn vì phiên khóa quyền ghi, skill chưa có hướng xử lý; (2) ảnh Nguyễn Văn Kiên báo 'Không có dữ liệu hiển thị' cần ghi 0 phần, và các lượt vẫn báo cáo từ tập con (1, 4 người) dù đã có bản 18 người, lại tiếp tục
- [2026-10-27] Batch 27/10/2026 cho thấy làm theo skill cũ ra kết quả sai: (1) mục Chuẩn bị ghi file inbox/khach có '32 đoàn viên ở mục I', nhưng đã xác minh đó là bản cũ 17 người, mục I chỉ là tổng quan; (2) phiên khách bị khóa ghi chỉ dựng báo cáo từ ảnh trong phiên, lặp lại 'mới có 9 người, thiếu 18' khi người
