---
type: decision
created: 2026-10-11
updated: 2026-10-11
---
Không xuất riêng file kết quả `03_Lam_sach_CD1_BT5_BT9` cho từng DN. Gắn kỹ thuật CD1 vào bảng biểu nghiệp vụ: BT6 chuẩn hóa MST trên bảng kê/hồ sơ có MST; BT5 đối chiếu HĐĐT CQ thuế ↔ bảng kê NNT (cột khớp/không khớp + thống kê); BT7 cụm cùng ngày vào sheet công nợ/ngân hàng; BT8/BT9 thuế GTGT mua vào theo tháng giữa bảng kê, HĐĐT (loại hủy/thay thế) và PS Nợ TK 133.

**Vì sao:** THP bị quá nhiều mẫu/kết quả/backup trùng; mẫu 03 tách làm phình thư mục mà nội dung đã nằm trong đối chiếu GTGT và bảng kê.

**Áp dụng:** Master 06 giữ sheet GTGT/NCC/TNCN theo mẫu mới; các sheet còn lại giữ form đối chiếu cũ Vạn Phát Thành. Kết quả DN: 1 Master ở `4. KET QUA PHAN TICH/` + chuyên đề ở `cau-thanh/` + json kỹ thuật ở `sao-luu/`.
