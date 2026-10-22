---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
source: [[conversations/2026-10-07]]
---
Tài liệu bồi dưỡng của Hội Kế toán Thuế TP. Hồ Chí Minh (tháng 06/2026, 55 trang) chuẩn hóa các thao tác xử lý dữ liệu kế toán bằng Excel phục vụ công tác thanh tra kiểm tra thuế [[sources/2026_CD1.pdf]].

Nội dung trọng tâm gồm 5 khối kỹ thuật xử lý dữ liệu trên bảng tính [[sources/2026_CD1.pdf]]:
- **Pivot Table trên Nhật ký chung (NKC) và Sổ cái**: Chuẩn hóa tiêu đề dòng, thiết lập ma trận tài khoản Nợ đối ứng tài khoản Có để lập chữ T tài khoản tiền mặt (TK 111), công nợ (TK 131), doanh thu (TK 511) và chi phí quản lý (TK 642) [[sources/2026_CD1.pdf]].
- **Kiểm soát chuỗi chi phí sản xuất và giá thành**: Lọc đối ứng các tài khoản chi phí `62*` (621, 622, 627) sang dở dang TK 154 và kết chuyển giá vốn TK 632 theo tháng phục vụ rà soát doanh nghiệp sản xuất [[sources/2026_CD1.pdf]].
- **Chuỗi đối chiếu doanh thu 4 lớp**: Thực hiện so sánh đối ứng giữa 4 nguồn dữ liệu độc lập: HĐĐT bán ra ↔ Bảng kê bán ra (BKBR) ↔ Sổ Nhật ký chung (TK 511, TK 3331) ↔ Tờ khai thuế GTGT (mẫu 01/GTGT) theo từng tháng hoặc quý (sử dụng hàm `ROUNDUP` để gom quý từ cột ngày tháng) [[sources/2026_CD1.pdf]].
- **Kiểm soát hóa đơn trùng và ngưỡng thanh toán**: Rà soát hóa đơn mua vào trùng lặp và các giao dịch có giá trị từ 20 triệu đồng trở lên nhằm kiểm soát điều kiện chứng từ thanh toán không dùng tiền mặt [[sources/2026_CD1.pdf]].
- **Chuẩn hóa khóa liên kết mã số thuế**: Xử lý lỗi định dạng MST (mất số 0 đầu, khoảng trắng, ký tự đặc biệt) để làm khóa liên kết khi đối chiếu chéo dữ liệu mua vào với NKC và TK 133 [[sources/2026_CD1.pdf]].
