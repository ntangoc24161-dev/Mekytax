# Bộ nhớ Javis - Index

> Chỉ mục bộ nhớ dài hạn của Javis. Mỗi dòng = 1 ký ức, trỏ tới file trong `facts/`.
> Nội dung file này được nạp vào đầu mỗi câu hỏi để Javis nhớ ngữ cảnh.

- [Nhu cầu kiểm soát văn bản và kết nối các phòng qua Phòng Văn Thư](facts/yeu-cau-he-thong-phong-van-thu.md) - Người dùng yêu cầu lập Phòng Văn Thư làm đầu mối kiểm soát văn bản phát hành và kết quả đầu ra.
- [Chỉ đạo tự động hóa điều phối công việc qua Phòng Văn Thư](facts/yeu-cau-tu-dong-dieu-phoi-phong-van-thu.md) - Người dùng yêu cầu Phòng Văn Thư tự động tiếp nhận việc và điều phối cho các phòng khác
- [Điều chỉnh Phòng Văn Thư chuẩn nghiệp vụ Thuế cơ sở 3 Đồng Nai](facts/yeu-cau-viet-lai-agent-phong-van-thu-chuan-tcs3.md) - Người dùng duyệt viết lại Agent Phòng Văn Thư đúng thực tế nghiệp vụ thuế thay vì doanh nghiệp chung.
- [Điều chỉnh Agent Yêu cầu Tài liệu chuẩn nghiệp vụ TCS3 Đồng Nai](facts/dieu-chinh-agent-yeu-cau-tai-lieu-chuan-tcs3.md) - Viết lại agent Yêu cầu Tài liệu làm cửa bàn giao văn bản giữa Văn thư và các phòng chuyên môn TCS3 thay vì doanh nghiệp chung.
- [Sắp xếp lớp Văn thư gọn, rẻ](facts/sap-xep-lop-van-thu-gon.md) - Gộp bàn giao vào Phòng Văn Thư, tắt workflow trùng 6 bước, không tạo trợ lý mới; xóa hẳn khi người dùng gật.
- [Xóa agent Yêu cầu Tài liệu khỏi lớp Văn thư](facts/xoa-agent-yeu-cau-tai-lieu-tcs3.md) - Người dùng quyết định xóa hẳn agent Yêu cầu Tài liệu, tinh gọn lớp Văn thư thành một cửa duy nhất.
- [Quy tắc mặc định xử lý file PDF và ảnh scan (OCR)](facts/preference-mac-dinh-skill-ocr.md) - Mặc định dùng paddleocr (đọc ra chat), pdf-sang-md (ghi sources/) và tram-tai-lieu (Office/searchable PDF/thao tác trang).
- [Quyết định giãn chu kỳ loop Văn thư điều phối lên 180 phút](facts/gian-loop-van-thu-180-phut.md) - Người dùng quyết định giãn chu kỳ loop Văn thư điều phối và đôn đốc lên 180 phút để tiết kiệm token.
- [Phân bổ hồ sơ Công ty CP CTL Investment trên các Brain Javis](facts/ho-so-ctl-phan-bo-brain.md) - Hồ sơ Công ty CP CTL Investment (MST 3602996707) được xử lý phân tách theo hai brain chuyên môn: Cập nhật công văn cảnh báo và Phân tích hồ sơ thuế.
- [Thị hiếu lập file Excel danh sách kiểm tra NNT](facts/preference-format-excel-danh-sach-kiem-tra.md) - Người dùng chuộng định dạng Excel chuyên nghiệp, chuẩn chỉnh, không cần dashboard.
- [Hồ sơ kiểm tra bổ sung KHBS-NGOC90 năm 2026](facts/ho-so-kiem-tra-bo-sung-khbs-ngoc90.md) - 4 doanh nghiệp kiểm tra bổ sung năm 2026 được phân công cho cán bộ Ngọc 90
- [Nguồn dữ liệu theo dõi kiểm tra quyết toán NNT](facts/nguon-du-lieu-theo-doi-kiem-tra-quyet-toan.md) - File theo dõi tổng hợp và phân công kiểm tra quyết toán năm 2026
- [Quy cách số ký hiệu và địa danh ngày tháng dự thảo thông báo](facts/quy-cach-tieu-ngu-ngay-thang-du-thao-thong-bao.md) - Chuẩn trình bày số ký hiệu và địa danh ngày tháng trong văn bản dự thảo
- [Thông tin liên hệ cán bộ kiểm tra Nguyễn Thị Ánh Ngọc](facts/thong-tin-can-bo-kiem-tra-ngoc90.md) - Thông tin đầu mối công chức phụ trách kiểm tra thuế
- [Cập nhật tiến độ dự thảo thông báo và cảnh báo quá hạn giải trình](facts/cap-nhat-tien-do-thong-bao-len-danh-sach-phan-cong.md) - Dự thảo thông báo mời lần 1/2 đồng thời cập nhật file phân công và cảnh báo quá hạn giải trình
- [Mẫu và căn cứ thông báo mời cung cấp hồ sơ kiểm tra quyết toán TCS3](facts/mau-thong-bao-cung-cap-ho-so-quyet-toan-tcs3.md) - Quy cách và cấu trúc văn bản thông báo mời cung cấp tài liệu theo mẫu gốc TCS3
- [Thư mục theo dõi đóng mở sử dụng hóa đơn](facts/thu-muc-theo-doi-dong-mo-hoa-don.md) - Thư mục lưu trữ danh sách theo dõi đóng mở hóa đơn và văn bản merge liên quan
- [Phân loại trạng thái theo dõi đóng mở hóa đơn](facts/quy-trinh-phan-loai-trang-thai-dong-mo-hoa-don.md) - Quy ước 3 trạng thái theo dõi đóng mở hóa đơn trong file Excel
- [Quy trình xử lý thông báo mời giải trình và đóng mở hóa đơn](facts/quy-trinh-xu-ly-thong-bao-moi-giai-trinh-nnt.md) - Quy trình xử lý người nộp thuế: gửi thông báo lần 1, lần 2, nếu không giải trình thì đóng hóa đơn hoặc lập phiếu chuyển xác minh hoạt động.
- [Thời hạn giải trình thông báo và kênh nhắc nhở quy trình đóng mở hóa đơn](facts/preference-thoi-han-giai-trinh-thong-bao-dong-mo-hd.md) - Thời hạn giải trình thông báo lần 1/2 là 7 ngày làm việc, cán bộ tự xử lý, nhận nhắc nhở qua Telegram.
- [Đường dẫn thư mục Hành chính trong công tác kiểm tra](facts/thu-muc-luu-tru-hanh-chinh-kiem-tra.md) - Thư mục hành chính lưu tại G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\4- HANH CHINH và được tổ chức theo loại công việc.
- [Thư mục lưu phôi mẫu văn bản kiểm tra](facts/thu-muc-phoi-mau-kiem-tra.md) - Phôi mẫu tờ trình, thông báo mời và các mẫu tái dùng lưu tại D:\01. Cong tac kiem tra\PHOIMAU
- [Tờ trình ban hành thông báo mời đi kèm bộ thông báo](facts/to-trinh-ban-hanh-thong-bao-moi-khbs.md) - Tờ trình ban hành thông báo mời căn cứ QĐ 1412/QĐ-DON, lưu chung thư mục THONG BAO MOI
- [Lộ trình mở rộng hệ thống theo dõi hồ sơ kiểm tra](facts/lo-trinh-so-theo-doi-chung-ho-so-kiem-tra.md) - Làm trước cho nhóm KHBS; khi có nhiệm vụ mới mới thêm trợ lý theo dõi chung báo cáo trưởng phòng
- [Sổ theo dõi danh sách thông báo mời giải trình](facts/so-theo-doi-thong-bao-moi-giai-trinh.md) - Các trường hợp mời giải trình theo dõi tại G:\...\1. CONG TAC KIEM TRA\So theo doi danh sach Thong bao moi.xlsx
