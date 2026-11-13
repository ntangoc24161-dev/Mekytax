---
type: preference
provenance: user
origin: javis-learned
created: 2026-11-12
updated: 2026-11-12
---
Người dùng yêu cầu: trước khi in hoặc chốt hồ sơ, Javis đối chiếu từng con số trên sheet `DuLieuMerge` của sổ theo dõi với dòng tổng sheet `TINH PHAT + CHAM NOP` của `Bang tinh.xlsx`. Lệch dù chỉ 1 đồng thì chỉ rõ ô bị lệch để người dùng sửa trước khi merge.

**Vì sao:** tiền chậm nộp đổi theo từng ngày, nên phải chốt số đúng trước khi đưa vào văn bản.
**Áp dụng:** khi ngày ký biên bản hoặc ngày nộp tiền thực tế đổi, cập nhật ngày đó vào Bảng tính trước, sau đó mới kiểm tra chéo và đẩy số mới sang sổ. Trên phôi, trường tiền dùng switch `\# "#,##0"`. Biên bản kiểm tra và QĐ xử phạt phải có số tiền bằng số kèm bằng chữ.
