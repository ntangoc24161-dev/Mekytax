---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
---
Khi đối chiếu nhà cung cấp mua vào với sổ cảnh báo rủi ro (CVCBRR F0/F1, danh bạ trạng thái), phải đối chiếu chính xác theo đúng mã số thuế bao gồm cả mã số thuế 13 số (chi nhánh). Không tự ý gom cảnh báo của chi nhánh này gán cho chi nhánh khác chỉ vì chung 10 số đầu của Tổng công ty mẹ.

**Vì sao:** Trường hợp Công ty Dầu khí TP.HCM có chi nhánh cửa hàng số 9 (`0300507707-030`) bị cảnh báo F0 nhưng chi nhánh Đồng Nai (`0300507707-006`) bán xăng xe bình thường; nếu chỉ đối chiếu 10 số đầu sẽ gây cảnh báo oan và sai lệch đánh giá chi phí đầu vào của NNT.
**Áp dụng:** Thuật toán và công cụ đối chiếu rủi ro hóa đơn phải so khớp đầy đủ 13 ký tự nếu là đơn vị trực thuộc.
