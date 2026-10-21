---
name: "Lập hồ sơ giải thể tại bàn"
description: "Lập hồ sơ giải thể tại bàn từ folder NNT: điền Word theo Bang tinh và TB ngừng MST, sửa sót mẫu, lưu vault/Drive."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-10-16
---
## Khi nào dùng
- User yêu cầu lập hồ sơ giải thể / giải thể tại bàn.
- Path folder nằm dưới `giai the` (ví dụ `2026/dang kiem tra/giai the/<ten-dn>`).
- Đã có hoặc sẽ có `Bang tinh.xlsx` và Thông báo ngừng MST (TB-TCS).

Không dùng skill này cho phân tích Master 8 sheet hay đối chiếu GTGT-TNDN thường.

## Chuẩn bị
- Path folder tuyệt đối hoặc tương đương Drive của DN.
- File nguồn tối thiểu: `Bang tinh.xlsx`, Thông báo ngừng MST (số TB, ngày TB, ngày ngừng MST).
- Mẫu Word hồ sơ giải thể tại bàn (tên kiểu `HO_SO_GIAI_THE_TAI_BAN_*.docx`).
- Định danh DN: tên đầy đủ, MST, địa chỉ, người đại diện, KTT, ngày thành lập/cấp GCN, ngành nghề (chỉ điền khi có nguồn).

## Quy trình
1. **Nhận diện loại việc:** nếu path có `giai the` hoặc user nói giải thể thì chạy nhánh này; ghi rõ MST và tên DN.
2. **Kiểm kê folder:** liệt kê Word/Excel/TB; ưu tiên `Bang tinh.xlsx` và mẫu/hồ sơ giải thể có sẵn.
3. **Bóc định danh:** tên, MST, địa chỉ, GĐ, KTT, ngày thành lập/cấp GCN, số TB, ngày TB, ngày ngừng MST, cán bộ thụ lý.
4. **Bóc số liệu:** lấy các năm có trong `Bang tinh.xlsx`; năm thành lập nếu không phát sinh thì ghi 0 hoặc để trống đúng mẫu, không bịa năm trước thành lập.
5. **Điền Word:** thay toàn bộ placeholder/mẫu cũ; xuất bản vault tại `exports/<slug-mst>/` kèm `Bang tinh.xlsx`, Word, zip nếu có.
6. **Rà sót mẫu:** tìm-thay mọi ngày/MST/tên còn của DN mẫu (ví dụ ngày thành lập mẫu).
7. **Upload:** đưa Word (và file kèm nếu user yêu cầu) lên đúng folder Drive của DN; trả link file + link folder.
8. **Báo còn trống:** liệt kê mục chưa có nguồn (ngành nghề, CCCD, giờ ký, doanh thu năm ngoài bảng…) để user bổ sung tay.

## Bẫy
- Nhầm sang playbook TTKT/Master 8 sheet chỉ vì user nói “lập hồ sơ kiểm tra”.
- Để sót ngày/MST/tên của bản mẫu trong Word (đã gặp sót `15/08/2011`).
- Bịa ngành nghề khi MST không khớp Danh bạ NNT - để placeholder và nói rõ còn thiếu.
- Ghi doanh thu các năm trước ngày thành lập như thể có phát sinh.
- Hứa “sẽ báo lại” mà không làm ngay trong lượt hoặc không xếp việc nền.

## Kiểm chứng
- Word trên vault và Drive cùng MST/tên DN đang làm.
- Grep/search trong Word: không còn ngày thành lập/MST/tên mẫu cũ.
- Số TB và ngày ngừng MST khớp Thông báo nguồn.
- Các năm số liệu khớp `Bang tinh.xlsx`; mục thiếu nguồn được liệt kê công khai.
