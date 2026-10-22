---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
source: [[conversations/2026-10-19]]
---
### 1. Căn cứ đánh giá kết quả phân tích hồ sơ thuế
Kết quả phân tích trong bộ 6 tệp chuẩn v1.4 được xác lập dựa trên 4 trụ cột kỹ thuật và pháp lý [[Quy_trinh_3_giai_doan_va_mau_bieu_chot_v1.4.md]]:
- **Hệ thống 80 quy tắc rủi ro định chuẩn:** Được lượng hóa trên sheet duy nhất 39 cột của tệp `KQ3`, chia thành 44 quy tắc trực tiếp, 21 quy tắc từ dữ liệu cơ quan thuế (`KQ2.1`) và 15 quy tắc từ sổ sách nội bộ doanh nghiệp (`KQ2.2`) [[Quy_trinh_3_giai_doan_va_mau_bieu_chot_v1.4.md]].
- **Nguyên tắc đối chiếu độc lập 4 nguồn:** So khớp chéo giữa Hóa đơn điện tử portal, Bảng kê mua vào/bán ra của NNT, Tờ khai thuế GTGT/Quyết toán TNDN/TNCN và Sổ sách kế toán chi tiết (NKC, Sổ cái, Thẻ kho, Sổ 511/632/111/112/131/331) [[Bốn Nguồn Dữ Liệu Đối Chiếu Thuế]].
- **Bộ chỉ số và danh mục rủi ro tuân thủ:** Dựa trên chỉ số cảnh báo rủi ro của Tổng cục Thuế theo Thông tư 31/2021/TT-BTC, danh sách doanh nghiệp không hoạt động tại địa chỉ đăng ký/bỏ trốn (F0), chuỗi hóa đơn liên quan (F1) và danh sách cảnh báo từ cơ quan Công an [[Quản lý rủi ro tuân thủ thuế (QLRR)]].
- **Quy tắc phân tầng kết luận:** Tệp `KQ2.1` và `KQ2.2` chỉ phản ánh số liệu chênh lệch, tỷ lệ và cờ máy; toàn bộ nhận xét, mức độ rủi ro (Cao/Trung bình/Thấp), điểm rủi ro và đề cương kiểm tra tại trụ sở NNT chỉ được phép kết luận tập trung tại tệp `KQ3` [[decision-quy-trinh-1-4-sau-tep-p7.md]].

### 2. Cơ chế chống xuất mẫu biểu trùng lắp
- **Quy chuẩn 6 tệp duy nhất:** Quy trình 1.4 xóa bỏ toàn bộ các file chuyên đề riêng lẻ trước đây (như file đối chiếu 511-632 tách rời, file NCC tách rời, file XNT riêng). Toàn bộ các skill vệ tinh A0 đến B3 chỉ được phép ghi dữ liệu vào đúng các sheet thành phần của 6 tệp `KQ1.1`, `KQ1.2`, `KQ1.3`, `KQ2.1`, `KQ2.2` và `KQ3` [[decision-quy-trinh-1-4-sau-tep-p7.md]].
- **Lưu trữ một nguồn:** 11 tệp mẫu cũ từ phiên bản v1.2 trở về trước đã được di chuyển cô lập vào thư mục lưu trữ `cu-v1.2`, loại bỏ hoàn toàn khả năng script tự động gọi nhầm template cũ [[conversations/2026-10-19]].

### 3. Những hạn chế và phạm vi Javis chưa thực hiện được
- **Chưa kiểm thử hành vi runtime trên dữ liệu thật (thực tế tính đến 19/10/2026):** Bộ 6 tệp v1.4 mới chỉ nghiệm thu phần cấu trúc tĩnh trên mẫu trống; các hành vi khử trùng mã W/V, tự động tăng phiên bản và cơ chế từ chối mẫu cũ khi chạy dữ liệu phát sinh thật cần có hồ sơ kiểm thử để xác nhận [[conversations/2026-10-19]].
- **Dữ liệu chứng từ phi cấu trúc/chưa số hóa:** Javis chưa tự động bóc tách hóa đơn giấy scan mờ, sổ sách viết tay chưa qua OCR chuẩn hóa, hoặc hợp đồng/biên bản nghiệm thu dạng file ảnh chụp (cần xác minh).
- **Xác minh thực tế ngoài hồ sơ:** Javis không thể thay thế cán bộ thuế trong việc kiểm tra hiện trường nhà xưởng, xác minh dòng tiền qua tài khoản cá nhân không đăng ký, hoặc kết luận ý chí chủ quan về hành vi trốn thuế ngoài bằng chứng số liệu định lượng [[Quy_trinh_3_giai_doan_va_mau_bieu_chot_v1.4.md]].
