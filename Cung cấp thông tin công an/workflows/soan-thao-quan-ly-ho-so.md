---
type: workflow
name: Soạn thảo & Quản lý hồ sơ công an
slug: soan-thao-quan-ly-ho-so
status: 'off'
group: Vận hành
description: Soạn dự thảo công văn trả lời và tờ trình lãnh đạo theo yêu cầu, lưu file và cập nhật sổ theo dõi.
steps:
  - agent: soan-thao-tra-loi-cong-an
    task: "Soạn dự thảo công văn trả lời (TRA_LOI_CV..._du_thao.docx) và tờ trình lãnh đạo (to_trinh....docx) theo mẫu quy chuẩn. Bám sát yêu cầu trong công văn và dữ liệu từ phân công. Đầu vào: {{input}}"
    verify_agent: Kiểm soát & Điều phối quy trình
    max_retries: 1
  - agent: quản-lý-hồ-sơ-công-an
    task: "Lưu dự thảo công văn + tờ trình vào thư mục công văn trên Drive. Cập nhật sổ theo dõi Bao_cao_cong_van.xlsx với thông tin công văn, danh sách DN, trạng thái. Dữ liệu từ bước trước: {{prev}}"
    verify_agent: Kiểm soát & Điều phối quy trình
    max_retries: 1
updated: '2026-11-08'
---

## Mục tiêu
Khi bạn yêu cầu "soạn thảo công văn trả lời cho những DN này", workflow tự động soạn dự thảo công văn + tờ trình theo quy chuẩn, lưu file vào Drive, và cập nhật sổ theo dõi. `Kiểm soát & Điều phối quy trình` sẽ kiểm tra chất lượng từng bước.

## Trigger
- **Bạn yêu cầu trực tiếp** (qua chat, Telegram, Zalo): "soạn thảo công văn trả lời cho CV XXX"
- **Hoặc** sau Workflow "Bóc tách & Phân công" hoàn tất, bạn giao CV có thêm yêu cầu mới

## Đầu vào
- Danh sách doanh nghiệp cần soạn thảo (từ Workflow trước hoặc yêu cầu bạn)
- Dữ liệu từ DB NNT (số MST, tình trạng, cơ quan quản lý)
- Mẫu công văn và tờ trình có sẵn

## Đầu ra
- Dự thảo công văn trả lời (docx) - chờ ký
- Dự thảo tờ trình lãnh đạo (docx) - chờ ký
- Sổ theo dõi cập nhật với CV mới
- Báo cáo về chat này: số hiệu CV, danh sách DN, đường dẫn file, trạng thái

## Ghi chú
- Phần chưa xác định (nợ/VPHC chưa có bằng chứng) → tô chữ đỏ
- Tờ trình phải có đủ ý kiến ký của Tổ KT1 và TCS3
- File tên theo quy chuẩn: `TRA_LOI_CV_[số]_du_thao.docx`
