---
type: workflow
name: Dây chuyền hồ sơ kiểm tra
slug: day-chuyen-ho-so-kiem-tra
status: 'off'
group: Tài chính
description: "Nối 3 trợ lý văn bản qua Quản lý kiểm soát: khởi tạo dữ liệu, soạn hồ sơ đi, và soạn biên bản hồ sơ về, tự động thẩm tra chống nhầm lẫn số liệu."
steps:
  - agent: khoi-tao-ho-so
    task: "Đọc {{input}}. Rút MST, tên DN, kỳ kiểm tra, đường dẫn XML tờ khai, PTDN02, BCTC và thư mục hồ sơ nếu có. Thiếu MST thì dừng và hỏi đúng một câu. Cập nhật thông tin công ty vào sổ/app, điền Bảng tính từ PTDN02 (năm hoặc ô thiếu thì bù BCTC sau cùng), bóc tách KQKD và điền phân tích Mục III. Đầu ra: các trường đã cập nhật, đường dẫn Bảng tính đã điền, và một câu bàn giao cho bước soạn hồ sơ đi. Không bịa số liệu, không ký số, không cấp số quyết định."
    verify_agent: quan-ly-ho-so-kiem-tra
    max_retries: 1
  - agent: soan-ho-so-di
    task: "Từ {{input}} và {{prev}}, lập bộ hồ sơ đi (mặc định chỉ TONG HOP.docx), tờ trình kiểm tra, quyết định kiểm tra và quyết định giám sát. Dùng dữ liệu công ty và đường dẫn Bảng tính từ bước trước. Thiếu Số QĐKT thì ghi Số:              /QĐ-TCS3 (không dùng DU_THAO-TRINH-KY, không bịa số chính thức). Khối chữ ký quyết định cách chức danh ít nhất 5 dòng trống, giữ cùng trang. Văn bản hành chính theo NĐ 30/2020; biểu mẫu nghiệp vụ giữ nguyên phôi gốc, không chạy script chuẩn hóa hàng loạt. Đầu ra: danh sách file kèm đường dẫn tương đối trong vault và chỗ cần điền tay (chữ ký, số quyết định thật)."
    verify_agent: quan-ly-ho-so-kiem-tra
    max_retries: 1
  - agent: soan-bien-ban-ho-so-ve
    task: "Từ {{input}} và {{prev}}, chỉ chạy khi Bảng tính đã chốt số (DTCHEP) hoặc {{input}} yêu cầu rõ lập biên bản/hồ sơ về. Thiếu số chốt thì dừng, nêu thiếu gì, không tự chốt thay đoàn. Điền BBXN/BBKT từ Bảng tính, xác định hành vi và mức phạt theo NĐ 125/310/291/68, soạn quyết định xử phạt hoặc kết luận theo NĐ 30/2020. Đầu ra: danh sách file biên bản và quyết định kèm đường dẫn tương đối, tổng truy thu + phạt nếu đã chốt, và chỗ cần điền tay. Không bịa số tiền phạt hay căn cứ pháp lý."
    verify_agent: quan-ly-ho-so-kiem-tra
    max_retries: 1
updated: '2026-11-06'
---

Quy trình tắt. Mở trang Quy trình của brain Hồ sơ kiểm tra, xem chuỗi 3 bước và vai trò kiểm duyệt của Quản lý, rồi bật khi muốn Javis tự chạy.

Thứ tự bắt buộc: khởi tạo hồ sơ -> Quản lý thẩm định -> soạn hồ sơ đi -> Quản lý thẩm định -> soạn biên bản và hồ sơ về -> Quản lý thẩm định chốt. Mỗi bước đều có Quản lý (`quan-ly-ho-so-kiem-tra`) soi lỗi và chặn dữ liệu sai. Bước 3 dừng nếu Bảng tính chưa chốt số.

Không đụng workflow cũ `lap-ho-so-kiem-tra` (chạy app xuất Word). Trợ lý tri thức gọi riêng, không nằm trong dây chuyền này.
