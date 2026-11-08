---
type: workflow
name: Lập hồ sơ kiểm tra
slug: lap-ho-so-kiem-tra
status: off
group: Tài chính
description: Nhập hoặc import một vụ kiểm tra thuế rồi xuất bộ Word hồ sơ đi, biên bản, phát sinh và hồ sơ về.
steps:
  - agent: lap-ho-so-kiem-tra
    task: "Đọc {{input}}. Rút MST, số quyết định kiểm tra, đường dẫn file Excel nếu có, và các mã mẫu (DI-03, BB-05, ...). Thiếu MST hoặc số quyết định thì dừng và hỏi đúng một câu. Không bịa số liệu."
    max_retries: 0
  - agent: lap-ho-so-kiem-tra
    task: "Từ {{prev}}, chạy python \"D:\\SECOND BRAIN\\javis-os\\shared-skills\\lap-ho-so-kiem-tra\\scripts\\chay_vu.py\" status. Nếu in chua chay thì khởi động nền \"D:\\SET UP BY AI\\HO SO KIEM TRA\\app\\chay.bat\" rồi gọi lại status đến khi in dang chay. Giữ cổng 8000. Không cài thêm thư viện."
    max_retries: 0
  - agent: lap-ho-so-kiem-tra
    task: "Nếu {{input}} có file Excel thì chạy script trên với lệnh import và đường dẫn đó. Báo tao=, cap_nhat= và mọi dòng CANH BAO. Không có file thì nói chưa import, không tạo vụ trắng. Cùng MST và số quyết định là cập nhật vụ cũ."
    max_retries: 0
  - agent: lap-ho-so-kiem-tra
    task: "Chạy lệnh xuat của script với --mst và --qd từ bước 1. Có mã mẫu thì thêm --docs. Không chỉ mã thì bỏ --docs để lấy nhóm đang tick. Nếu script in LOI, CANH BAO hoặc CHUA XUAT thì đưa nguyên chữ đó và dừng. Không chạy lại, không tự xác nhận."
    max_retries: 0
  - agent: lap-ho-so-kiem-tra
    task: "Từ {{prev}}, trả thư mục output và đường dẫn hoso.zip. Nhắc mở Word để điền chỗ ký và số quyết định xử phạt nếu chưa cấp. Không ký số, không xuất PDF, không sửa HO SO NGUON."
    max_retries: 0
updated: 2026-10-01
---

Quy trình tắt. Mở trang Quy trình của brain Hồ sơ kiểm tra, xem chuỗi, rồi bật khi muốn Javis tự chạy.
