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
