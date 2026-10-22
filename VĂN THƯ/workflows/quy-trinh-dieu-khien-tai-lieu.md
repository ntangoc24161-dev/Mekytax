---
type: workflow
name: Quy trình điều khiển tài liệu
slug: quy-trinh-dieu-khien-tai-lieu
status: off
group: Văn thư
description: TẮT. Trùng ban hành văn bản đi. Dùng quy-trinh-soan-thao-va-ban-hanh-van-ban-di.
steps:
  - agent: yeu-cau-tai-lieu
    task: "Phòng {{input}} chuẩn bị dự thảo tài liệu loại {{doc_type}}, gửi yêu cầu tới Phòng Văn Thư kèm thông tin: mục đích {{purpose}}, người nhận {{recipient}}, người ký {{signatory}}, hạn chót {{deadline}}"
    max_retries: 1

  - agent: phong-van-thu
    task: "Kiểm duyệt dự thảo từ phòng {{prev}}: rà soát thể thức VB theo NĐ 30/2020, kiểm tra chữ ký số, trình ký, định dạng. Nếu hợp lệ ghi 'PHÊDUYỆT', nếu sai ghi 'CẦN SỬA' kèm danh sách lỗi cụ thể."
    verify_agent: phong-van-thu
    max_retries: 1

  - agent: yeu-cau-tai-lieu
    task: "Nhận kết quả kiểm duyệt: {{prev}}. Nếu PHÊDUYỆT thì chuyển sang bước 4. Nếu CẦN SỬA thì sửa theo danh sách lỗi rồi gửi lại Phòng Văn Thư."
    max_retries: 2

  - agent: phong-van-thu
    task: "Phát hành tài liệu: cấp phó, ký số TCS3, ghi số hiệu VB, ghi ngày phát hành. Lưu bản chính vào hồ sơ, gửi bản sao cho phòng {{yeu-cau-tai-lieu}}. Ghi thông tin: số hiệu, ngày, người ký, người nhận."
    max_retries: 1

  - agent: yeu-cau-tai-lieu
    task: "Sử dụng tài liệu số {{prev.so_hieu}} theo mục đích {{doc_type}}. Sau khi sử dụng (hoặc hết hạn), báo cáo kết quả tới Phòng Văn Thư: người nhận, trạng thái (thành công/chậm/vấn đề), ghi chú."
    max_retries: 1

  - agent: phong-van-thu
    task: "Nhận báo cáo kết quả từ phòng {{prev}}: ghi nhận số hiệu, ngày phát hành, trạng thái sử dụng, ghi chú. Cập nhật hồ sơ tài liệu. Nếu có vấn đề thì liên hệ phòng để xử lý."
    max_retries: 1

updated: 2026-10-19
---

## Mô tả

**Đã tắt 19/10/2026.** Trùng `quy-trinh-soan-thao-va-ban-hanh-van-ban-di` và tốn 6 lượt agent. Không bật lại. Xóa file khi người dùng gật.

Workflow cũ (giữ để đối chiếu):

**Các bước:**
1. **Gửi yêu cầu**: Phòng làm việc chuẩn bị dự thảo + gửi yêu cầu chi tiết tới Phòng Văn Thư
2. **Kiểm duyệt**: Phòng Văn Thư rà soát thể thức theo NĐ 30/2020 (chữ ký số, trình ký, định dạng...)
3. **Sửa lỗi (nếu cần)**: Phòng sửa theo yêu cầu rồi gửi lại
4. **Phát hành**: Phòng Văn Thư ký số, ghi số hiệu, phát hành bản chính
5. **Sử dụng**: Phòng sử dụng tài liệu theo mục đích
6. **Báo cáo**: Phòng báo lại kết quả (thành công / chậm / vấn đề); Phòng Văn Thư ghi nhận để báo cáo lãnh đạo

**Lợi ích:**
- Tài liệu không vỡ quy chuẩn được phát hành
- Phòng Văn Thư có bằng chứng kiểm duyệt đầy đủ
- Báo cáo kết quả giúp lãnh đạo biết tài liệu có hiệu quả hay không
- Tuân thủ quy trình quản lý tài liệu hành chính

**Khi nào dùng:**
- Mỗi lần phòm làm việc cần phát hành công văn, hợp đồng, biên bản, báo cáo chính thức
- Cần theo dõi chất lượng + hiệu quả của tài liệu
- Muốn ghi nhận lịch sử điều chỉnh + báo cáo kết quả cho lãnh đạo
