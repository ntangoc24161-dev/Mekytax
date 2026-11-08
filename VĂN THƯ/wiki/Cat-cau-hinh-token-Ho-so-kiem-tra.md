---
type: wiki
status: active
tags: [wiki, van-hanh, token]
created: 2026-10-23
updated: 2026-10-23
source: usage_index + cấu hình brain Hồ sơ kiểm tra
---
# Cắt cấu hình token - Hồ sơ kiểm tra

## Hiện trạng đo được (tháng 10/2026)

Theo chỉ mục `usage_index`:

- Brain **Hồ sơ kiểm tra** xếp **hạng 3** token tháng này: ~**40.8M** token, ~**444** lượt, quy đổi ~**$27.53** (đắt hơn Brain Default dù token thấp hơn, vì Opus/Sonnet).
- Ngày 15/10: nền **Opus** 70 lượt; ngày 23/10: nền **Sonnet** 232 lượt + chat Sonnet 66 lượt.
- Có **5 agent** (khởi tạo, lập hồ sơ, soạn hồ sơ đi, biên bản/hồ sơ về, tri thức); **2 workflow** đều `status: off`; **0 loop**, **0 reminder**.
- Trước cắt thêm: Kanban đã **orchestration=off**; **26 skill** (~232 KB), lẫn skill phân tích GĐ1-3 / NCC / ôn thi / iOffice / công an vốn thuộc phòng khác.

Liên quan: [[Cat-cau-hinh-token-Brain-Default]], [[Cat-cau-hinh-token-Phan-tich-ho-so-thue]].

## Quyết định đã áp (2026-10-23)

1. **Giữ / xác nhận** Kanban `orchestration: off` (đã off sẵn; gọi lại API để chốt).
2. **Không đụng** 2 workflow (đã `off`): `day-chuyen-ho-so-kiem-tra`, `lap-ho-so-kiem-tra`. Bật tay khi cần lập bộ hồ sơ.
3. **Tắt (không xoá)** 11 skill lệch phòng:
   - Phân tích sang phòng Phân tích: `giai-doan-1/2/3`, `phan-tich-ho-so-thue`, `doi-chieu-ncc-mua-vao-rui-ro`, `phan-tich-cong-no-tien-mat`
   - Ngoài phạm vi: `html-to-webcake`, `phan-tich-cau-hoi-sinh-doi-cheat-sheet`, `tien-trinh-supervisor-tu-dong`, `tra-cuu-ioffice-dong-nai`, `tra-loi-cong-van`
4. **Giữ** skill gắn agent lập/soạn hồ sơ: `lap-ho-so-kiem-tra`, `nhap-bang-tinh-tu-bao-cao-hdkd`, `boc-tach-bang-tinh-phan-tich`, `ho-so-kiem-tra-ke-hoach`, `dien-bien-ban-tu-bang-tinh`, `soan-thao-van-ban`, `xu-phat-vphc-thue-hoa-don`, cộng hệ thống/OCR/wiki.

## Mục tiêu vận hành sau cắt

- Phòng này = **lập và soạn hồ sơ kiểm tra** (đi/về/biên bản), không chạy phân tích 80 quy tắc.
- Phân tích số liệu để ở brain **Phân tích hồ sơ thuế**.
- Token nền ngày 23/10 xảy ra dù orchestration đã off: nhiều khả năng là phiên chat/agent trên đúng brain này; sau cắt skill, catalog nhẹ hơn; khi không soạn hồ sơ nên chuyển brain khác.

## Việc còn mở

- Ghim model nhẹ cho khởi tạo/điền bảng; Opus chỉ soạn luận xử phạt khó.
- Đo lại token 7 ngày sau cắt.
- Cân nhắc cắt tiếp các phòng còn lại có orchestration=auto: Cung cấp thông tin công an, Cập nhật công văn cảnh báo, Ôn tập thi.
