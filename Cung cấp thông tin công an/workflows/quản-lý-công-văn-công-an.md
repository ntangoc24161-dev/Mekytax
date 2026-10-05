---
type: workflow
name: Quản lý công văn công an
slug: quản-lý-công-văn-công-an
status: active
group: Tài chính
description: Tiếp nhận, bóc tách công văn công an, tra cứu NNT, cập nhật sổ theo dõi, lập dự thảo trả lời Tổ KT1, xuất Excel chi tiết và đồng bộ thư mục.
steps:
- agent: quản-lý-hồ-sơ-công-an
  task: 'Tiếp nhận công văn từ thư mục CV MOI hoặc Zalo; trích xuất nội dung file
    .md lưu tại D:\HSO CONG AN FILE MD; tra cứu CSDL nội bộ DB NNT.xls và cổng thông
    tin để phân Tổ KT1-KT2-KT3; tra cứu CBQL Tổ KT1; cập nhật file Bao_cao_cong_van.xlsx
    (sheet NHAP_HO_SO); lưu PDF vào thư mục năm 2026; tự động lập dự thảo công văn và
    tờ trình cho Tổ KT1; sẵn sàng xuất Excel tổng hợp chi tiết 24 cột tiêu chí và rà
    soát đồng bộ hậu tố -R cho các thư mục hoàn tất.'
  verify_agent: quản-lý-hồ-sơ-công-an
  max_retries: 1
updated: '2026-09-26'
---
