---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-17
updated: 2026-11-20
source: [[conversations/2026-11-20]]
---
Sổ Master Quản lý công tác kiểm tra là mô hình cơ sở dữ liệu bảng tính tập trung nhằm theo dõi đa dạng các đầu việc kiểm tra thuế (kiểm tra bổ sung KHBS, kiểm tra quyết toán tại trụ sở NNT, thông báo mời giải trình, chuyên đề trường công lập...) [[conversations/2026-11-17]].

### Cấu trúc chuẩn (từ 20/11/2026)
1. **DASHBOARD**: Tổng hợp nhanh số DN theo loại và trạng thái [[conversations/2026-11-20]].
2. **DANH_SACH_THEO_DOI_CHUNG**: **Sheet theo dõi duy nhất** - gom toàn bộ NNT và loại công việc (địa chỉ, SĐT, cán bộ, giai đoạn, L1/L2, trạng thái, cấp độ, hành động) [[conversations/2026-11-20]].
3. **CONFIG**: Tham số loại việc, hạn 7 ngày làm việc, hành động cuối, cán bộ mặc định [[conversations/2026-11-17]].
4. **LEGEND**: Hướng dẫn mã trạng thái và quy ước cảnh báo [[conversations/2026-11-17]].

Sheet `KT_TONG_HOP` đã được gộp vào `DANH_SACH_THEO_DOI_CHUNG` và xóa theo yêu cầu theo dõi trên một sheet chung [[conversations/2026-11-20]].
