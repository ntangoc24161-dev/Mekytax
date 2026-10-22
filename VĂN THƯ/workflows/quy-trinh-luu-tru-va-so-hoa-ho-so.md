---
type: workflow
name: Quy trình lưu trữ và số hóa hồ sơ
slug: quy-trinh-luu-tru-va-so-hoa-ho-so
status: active
group: Văn thư
description: Thu hồi, kiểm kê, số hóa OCR và lưu trữ hồ sơ tài liệu vào Second Brain phục vụ tra cứu
steps:
  - agent: phong-van-thu
    task: "Tiếp nhận hồ sơ kết thúc từ các phòng ban/đoàn kiểm tra theo {{input}}. Kiểm tra tính toàn vẹn, đầy đủ của tài liệu theo Danh mục thành phần hồ sơ. Lập biên bản giao nhận hồ sơ lưu trữ."
    max_retries: 1
  - agent: so-hoa-ocr-cong-van
    task: "Từ hồ sơ {{prev}}: Thực hiện quét scan, chuyển đổi toàn bộ tài liệu hồ sơ giấy thành file PDF số hóa và file .md trích xuất qua công cụ pdf-sang-md / paddleocr. Đặt tên file theo cấu trúc chuẩn: <Năm>_<MãHS>_<TênTaiLieu>."
    max_retries: 1
  - agent: tri-thuc-second-brain
    task: "Nạp tài liệu đã số hóa từ {{prev}} vào hệ thống tri thức Second Brain: Cập nhật thẻ metadata, phân loại vào thư mục lưu trữ tương ứng trong sources/, lập liên kết wiki và mục lục tra cứu nhanh để phục vụ cán bộ truy xuất thông tin tức thì."
    max_retries: 1
updated: 2026-10-18
---

## Mô tả quy trình
Quy trình thu thập, phân loại, số hóa và bảo quản hồ sơ tài liệu văn thư:
1. **Kiểm kê & Tiếp nhận (phong-van-thu)**: Thu hồi hồ sơ từ đơn vị giải quyết, kiểm soát thành phần tài liệu.
2. **Số hóa tài liệu (so-hoa-ocr-cong-van)**: Scan, OCR toàn văn và chuẩn hóa cấu trúc tệp tin điện tử.
3. **Lưu trữ Second Brain (tri-thuc-second-brain)**: Chỉ mục hóa, chưng cất tri thức và tạo khả năng tra cứu thông minh.
