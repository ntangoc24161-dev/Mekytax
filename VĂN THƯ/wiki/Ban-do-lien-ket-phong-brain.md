---
type: wiki
status: active
tags: [wiki, van-hanh, dieu-phoi]
created: 2026-10-23
updated: 2026-11-06
source: Javis/index + agents/workflows các brain + loop van-thu-dieu-phoi-don-doc
---
# Bản đồ liên kết phòng / brain

Nguyên tắc: **VĂN THƯ = cửa vào và kiểm soát sổ**; việc chuyên môn chạy ở **brain chuyên** (hoặc workflow nội bộ nếu đã đủ agent trong VĂN THƯ). Javis không gọi agent chéo brain native; bàn giao bằng sổ + checklist [[ban-giao-lien-phong]] + đổi brain / chạy workflow tay hoặc Kanban đúng brain.

Liên quan: [[Cat-cau-hinh-token-dot-3-cong-an-canh-bao-on-tap]], quyết định hub Văn thư trong memory facts.

## Bảng map việc → brain → agent / workflow

| Loại việc | Brain đích | Agent chính | Workflow (nếu có) | Ghi chú |
|---|---|---|---|---|
| Công văn Công an, cung cấp hồ sơ NNT | **Cung cấp thông tin công an** | `quản-lý-hồ-sơ-công-an` → `soan-thao-tra-loi-cong-an` | `quản-lý-công-văn-công-an` (active) | Hỗ trợ: `khai-thac-du-lieu-ho-so`, `phap-che-chinh-sach-thue`. Mirror VĂN THƯ đã archive 2026-11-06. |
| Công văn cảnh báo rủi ro HĐ/NCC; đối chiếu mặt hàng (vd CTL) | **Cập nhật công văn cảnh báo** | `tra-cuu-tai-cv-ioffice` → `so-hoa-ocr-cong-van` → `cập-nhật-công-văn-cảnh-báo-rủi-ro` → `doi-chieu-rui-ro-ncc` → `soan-thao-van-ban` | `day-chuyen-xu-ly-cv-canh-bao` (active) | Phân tích mặt hàng CTL đã chạy ở brain này. |
| Phân tích hồ sơ thuế 3 giai đoạn (KQ1-KQ3) | **Phân tích hồ sơ thuế** | `thu-ho-so-kiem-ke` / `truong-phong-phan-tich-ho-so` + chuyên viên | `phan-tich-ho-so-thue-3-giai-doan` (active) | Cần folder tuyệt đối + MST + kỳ. Orch Kanban đang off. |
| Lập / soạn hồ sơ kiểm tra đi-về-biên bản | **Hồ sơ kiểm tra** | `khoi-tao-ho-so`, `soan-ho-so-di`, `soan-bien-ban-ho-so-ve`, `lap-ho-so-kiem-tra` | `day-chuyen-ho-so-kiem-tra` / `lap-ho-so-kiem-tra` (**off**, bật tay) | Không lẫn với phân tích 80 quy tắc. |
| Tiếp nhận CV đến, phân việc, thể thức VB đi, lưu trữ OCR | **VĂN THƯ** | `phong-van-thu`, `so-hoa-ocr-cong-van`, `soan-thao-van-ban`, `tri-thuc-second-brain` | `quy-trinh-tiep-nhan-va-xu-ly-cong-van-den`, `quy-trinh-soan-thao-va-ban-hanh-van-ban-di`, `quy-trinh-luu-tru-va-so-hoa-ho-so` | Loop `van-thu-dieu-phoi-don-doc` (180 phút). |
| Ôn / giải đề khảo thí | **Ôn tập thi** | (chưa có agent; dùng skill) | - | Skill giải đề + cheat-sheet; orch off. |
| Inbox / việc chưa rõ phòng | **Brain Default** (nhẹ) hoặc hỏi lại | - | - | Default đã cắt skill chuyên; chỉ OCR/wiki hệ thống. |

## Cách bàn giao (tóm tắt)

1. Ghi sổ Văn thư: Phòng thụ lý + Brain đích + hạn.
2. Chạy skill `ban-giao-lien-phong` hoặc điền đủ checklist trong tin nhắn / Kanban.
3. Đổi sang brain đích trên dashboard → mở agent hoặc workflow tương ứng.
4. Khi xong: ghi `Kết quả báo lại` trên sổ + đường dẫn file đầu ra.

## Mâu thuẫn / hiện trạng

- **Nguồn chân lý = brain chuyên.** VĂN THƯ chỉ còn 4 agent hub; mirror agent/workflow chuyên đã chuyển `_archive-mirror-2026-11-06/` (2026-11-06).
- Không tự gọi chéo brain: loop Văn thư chỉ ghi hướng dẫn + phân sổ; người đổi brain rồi chạy agent/workflow tại brain đích.
- Handoff xác nhận, ghim model, orch, archive, kênh: [[Backlog-van-hanh-lien-phong]]; checklist xoá brain ít dùng: [[Checklist-archive-brain-it-dung]].
