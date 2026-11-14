---
name: Hồ sơ kiểm tra giải thể
description: 'Lập và chuẩn hóa bộ hồ sơ kiểm tra quyết toán giải thể: bộ văn bản, nộp thừa, bù trừ, thể thức số ngày và văn bản DN.'
group: Tài chính
updated: '2026-11-11'
learned_updated: '2026-11-11'
---
## Khi nào dùng
Áp dụng khi lập, rà soát và chuẩn hóa bộ hồ sơ kiểm tra quyết toán giải thể doanh nghiệp (file `TONG HOP.docx`).

## Quy trình & Quy chuẩn văn bản
1. **Tiêu đề và dẫn chiếu cuộc kiểm tra:** Toàn bộ tiêu đề Quyết định kiểm tra và các câu dẫn chiếu căn cứ phải ghi đồng nhất là "về việc kiểm tra quyết toán giải thể tại [Tên DN]".
2. **Thời hạn kiểm tra:** Ghi rõ số ngày kèm đơn vị "ngày làm việc".
3. **Số và ngày quyết định khi chưa ban hành:**
   - Trên Giấy xác nhận giao nhận quyết định: bỏ hẳn năm, không điền gì, không dùng dấu ba chấm `...`.
   - Định dạng canh lề: chữ "số:" canh thẳng hàng nằm dưới chữ "sở" của cơ quan ban hành (ví dụ: Thuế cơ sở 3).
4. **Văn bản của doanh nghiệp kèm theo:**
   - Văn bản từ chối nhận lại tiền nộp thừa và Công văn xin quyết toán thuế giải thể phải thay đúng tên, mã số thuế, đại diện pháp luật của doanh nghiệp đang giải thể.
   - Không kèm các cam kết đặc thù ngành nghề (như cam kết kinh doanh vàng) cho doanh nghiệp không kinh doanh ngành đó.
5. **Nguồn dữ liệu:** Luôn căn cứ và đối chiếu theo file `Danh sach kiem tra quyet toan pro.xlsm`.

## Bẫy cần tránh
- Không copy chéo thông tin chi nhánh/đơn vị thành viên từ hồ sơ doanh nghiệp khác.
- Tránh lỗi gõ lặp năm (ví dụ: `20262026`).
- Không để sót văn bản mang tên giám đốc hoặc doanh nghiệp của phôi mẫu cũ.

## Lịch sử (tự học)
- [2026-10-30] Vụ giải thể Sơn Tiền Hải chuyển từ phôi cũ sang phôi Mo Fa để cập nhật đồng bộ Luật Thanh tra 2025, QĐ 1614 ngày 13/10/2020, QĐ 1377 ngày 30/06/2025; Mục III Nghĩa vụ NSNN bắt buộc điền đủ số phát sinh 5 năm theo công thức suy luận `(3)=0, (6)=(4), (7)=(3)+(4)+(5)-(6)`.
- [2026-10-30] Phát hiện lỗi bên nhận hồ sơ trùng chính bên giao tại Biên bản giao nhận trong phôi gốc; đã sửa bên nhận thành Nguyễn Thành Đồng (Kiểm tra viên thuế).
- [2026-10-30] Phát hiện lỗi sót MST của vụ khác (3600275611) trên Văn bản từ chối nhận lại số tiền nộp thừa của phôi Mo Fa; đã chuẩn hóa điền đúng MST của doanh nghiệp theo từng ô lưới ký tự.
- [2026-10-31] Chuẩn hóa tên gọi chỉ tiêu: quét regex thay thế triệt để 18 vị trí từ "chi phí quản lý kinh doanh" sang "chi phí quản lý doanh nghiệp" trong hồ sơ kiểm tra.
- [2026-10-31] Đóng gói bộ hồ sơ đề xuất hoàn thuế kiêm bù trừ thu NSNN giải thể (mẫu MS-03-DCS và Tờ trình đề xuất hoàn giải thể) áp dụng Điều 32 NĐ 181/2025/NĐ-CP.
- [2026-11-11] Quy tắc kiểm tra số gõ cứng trên DTCHEP hoạt động dạng cảnh báo mềm (non-blocking warning), ghi nhận vào open-questions mà không làm gián đoạn luồng tự động đổ số vào biên bản.
- [2026-11-11] Đổi quy tắc gọi tên cuộc kiểm tra: tiêu đề Quyết định kiểm tra và các câu "Căn cứ... về việc..." dẫn chiếu lại phải ghi "kiểm tra quyết toán giải thể" (không còn dùng "kiểm tra thuế" chung như quy tắc cũ); phát hiện khi đối chiếu vụ Kiến Thành Tân với mẫu Thúy Hồng, sửa cả tiêu đề và 8 đoạn dẫn chiếu bị sót.
- [2026-11-11] Phát hiện vụ Thúy Hồng gõ sai năm căn cứ QĐ 1377/QĐ-CT thành 30/06/2026 (đúng là 30/06/2025); đã sửa và bổ sung cảnh báo đối chiếu bảng căn cứ pháp lý khi lập hồ sơ mới.
- [2026-11-11] Thời hạn kiểm tra phải ghi đơn vị "ngày làm việc" (không ghi bare "ngày"); áp dụng sửa cho vụ Kiến Thành Tân theo mẫu Thúy Hồng.
- [2026-11-11] Cập nhật bổ sung quy chuẩn trình bày số/ngày chưa ban hành (bỏ năm, bỏ chấm, canh dưới chữ sở), chuẩn hóa văn bản nộp kèm của DN và xác định nguồn chuẩn từ sổ JAVIS PRO theo phản hồi của người dùng.
