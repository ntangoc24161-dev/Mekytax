---
type: wiki
status: active
tags: [wiki, van-hanh, token]
created: 2026-10-23
updated: 2026-10-23
source: usage_index tháng 10/2026 + cấu hình sau cắt
---
# Ước tính tiết kiệm token sau cắt cấu hình

Căn cứ số đo tháng 10/2026 (đến trước/giữa đợt cắt) trong `usage_index` và các thay đổi đã ghi ở [[Cat-cau-hinh-token-Brain-Default]], [[Cat-cau-hinh-token-Phan-tich-ho-so-thue]], [[Cat-cau-hinh-token-Ho-so-kiem-tra]], [[Cat-cau-hinh-token-dot-3-cong-an-canh-bao-on-tap]].

## Cơ sở đo được

- Tổng ~**186.3M** token; **background 86.3%**, chat 10.6%, subagent 3.1%.
- Cache hit ~**93%**: tổng token nhìn lớn; phần đắt hơn nằm ở lượt chạy + output + model nặng.
- Tắt orchestration **không xoá hết** nhãn background: phiên agent/Claude Code dài vẫn được ghi là nền (ví dụ Hồ sơ kiểm tra orch đã off vẫn có hàng trăm lượt nền Sonnet ngày 23/10).
- Catalog `SKILL.md` 6 brain đã cắt: giảm ~**40%** dung lượng mô tả skill (ước bytes), chủ yếu nhẹ prompt chat/router.

## Ba kịch bản (ước %)

| Kịch bản | Ước giảm tổng token | Điều kiện |
|---|---:|---|
| A. Chỉ cấu hình | **15-30%** | Orch off + loop 180 phút + tắt skill/reminder; vẫn chat và chạy phân tích tay như cũ |
| B. Cấu hình + thói quen | **40-60%** | Thêm: không để dây chuyền tự chạy hàng loạt, đúng brain, Haiku cho việc nhẹ |
| C. Ít chạy nền phân tích | **50-65%** | Tháng sau gần như không lặp các đợt nền Sonnet/Opus kiểu 19-23/10 |

Phần **chắc từ cấu hình thuần**:

- Loop VĂN THƯ 60→180 phút: phòng này ~3.7% tổng và 100% nền → giảm khoảng **~2% tổng hệ** (giảm ~1/2 đến 2/3 số vòng loop).
- Huỷ reminder kể chuyện + task blocked: nhỏ, thường **&lt;1%**.
- Orch off: cắt việc Kanban tự claim/retry; tiết kiệm lớn khi trước đây hay để auto chạy lại dây chuyền, nhưng **không** bằng 86% nền lịch sử.

Phần **phụ thuộc anh dùng thế nào** (lớn nhất): các đợt nền Phân tích / Default / Hồ sơ kiểm tra (~88% tổng) có lặp lại hay không.

## Cách kiểm chứng sau 7 ngày

Trên trang Mức dùng, so kỳ 7 ngày sau cắt với 7 ngày trước (hoặc cùng số ngày làm việc):

1. Tổng token và token/lượt
2. Tỷ trọng background
3. Xếp hạng project: Phân tích, Default, Hồ sơ kiểm tra

Nếu background vẫn cao trong khi orch off: nguyên nhân là phiên chat/agent dài, không phải dispatcher.
