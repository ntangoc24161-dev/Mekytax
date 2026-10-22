---
type: workflow
name: Quy trình soạn thảo và ban hành văn bản đi
slug: quy-trinh-soan-thao-va-ban-hanh-van-ban-di
status: active
group: Văn thư
description: Kiểm thể thức NĐ 30, chuẩn hóa Word TCS3, cấp số và phát hành văn bản đi
steps:
  - agent: phong-van-thu
    task: "Từ {{input}}: lập gói ban hành (loại VB, trích yếu, nơi nhận, phòng dự thảo, người ký, file Word, tờ trình nếu có). Ghi sổ `06 - Văn thư/so-van-ban-di.md` trạng thái `Chờ kiểm thể thức`. Rà 9 thành phần NĐ 30/2020 (Quốc hiệu, tên cơ quan TCS3, số-ký hiệu, địa danh-ngày, tên loại-trích yếu, thẩm quyền ký, chữ ký số tổ chức TCS3, nơi nhận, lề-phông). Đạt thì ghi `ĐẠT THỂ THỨC`. Chưa đạt thì liệt kê từng điểm sai. Không cấp số ở bước này. Không gọi agent Yêu cầu Tài liệu."
    verify_agent: phong-van-thu
    max_retries: 2
  - agent: soan-thao-van-ban
    task: "Từ kết quả {{prev}}: căn chỉnh dự thảo đúng thể thức NĐ 30/2020 và ký hiệu TCS3, xuất Word sạch trình ký. Không soạn phúc đáp Công an (việc của soan-thao-tra-loi-cong-an)."
    max_retries: 1
  - agent: phong-van-thu
    task: "Chỉ khi thể thức đã đạt: cấp số đi trên `06 - Văn thư/so-van-ban-di.md`, gắn ngày phát hành, hỗ trợ ký số tổ chức TCS3, ghi nơi nhận, lưu bản gốc, trạng thái `Đã phát hành`. Sai thể thức hoặc sai thẩm quyền thì từ chối cấp số, ghi rõ lý do."
    max_retries: 1
updated: 2026-10-19
---

## Mô tả quy trình

Ba bước, một cửa Văn thư:

1. **Gói và kiểm thể thức** (`phong-van-thu`)
2. **Chuẩn hóa Word** (`soan-thao-van-ban`)
3. **Cấp số và phát hành** (`phong-van-thu`)
