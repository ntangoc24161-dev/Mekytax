---
name: "Cập nhật sổ JAVIS PRO từ sheet TINH PHAT + CHAM NOP"
description: "Đối chiếu dòng tổng sheet TINH PHAT + CHAM NOP của Bảng tính với sổ Danh sach kiem tra quyet toan pro.xlsm, sửa ô lệch, giữ công thức và VBA."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-11-05
---
## Khi nào dùng
- User yêu cầu: "căn cứ file bang tinh.xlsx / sheet TINH PHAT + CHAM NOP (C T) trong thư mục <VỤ>/0. BIEN BAN cập nhật lại sổ Danh sach kiem tra quyet toan pro.xlsm".
- Sau khi đoàn chốt số truy thu, phạt, chậm nộp trên Bảng tính và cần đồng bộ sang sổ theo dõi.

## Chuẩn bị
- Bảng tính của vụ: Drive `ntangoc2416.1@gmail.com`, thư mục `1. TD/<số>. <TEN VU>/0. BIEN BAN/Bang tinh.xlsx`.
- Sổ theo dõi: `Danh sach kiem tra quyet toan pro.xlsm` (thư mục `00. Theo doi chung` trên Drive), sheet `01_DA_BAN_HANH`, tìm dòng của vụ theo MST.
- Đọc Drive qua Composio, PHẢI truyền account của user (xem fact doc-file-google-drive-qua-composio).

## Quy trình
1. Lấy metadata cả Bảng tính và sổ trên Drive, so ngày sửa với bản trong vault `exports/<vu>-<mst>/`. Bản Drive mới hơn thì tải lại và thay bản vault; không dùng bản vault cũ làm nguồn.
2. Nếu thấy file lock `~$` mới tạo trên Drive thì file đang mở trong Excel của user: vẫn làm nhưng báo user số có thể còn chạy và việc lưu đè từ Excel sẽ làm mất ô vừa sửa.
3. Mở sheet `TINH PHAT + CHAM NOP (C T)`, đọc dòng TỔNG CỘNG.
4. Tìm dòng của vụ trong `01_DA_BAN_HANH` theo MST, đối chiếu từng cột:
   - AJ: ngày ký BBKT
   - AN: truy thu TNDN; AO: truy thu GTGT
   - AQ: chậm nộp GTGT; AR: chậm nộp TNDN (0,03%)
   - AT: phạt khai sai 20%; AU: phạt thủ tục
   - AX: giảm lỗ
5. Chỉ ghi các ô lệch. Ô khớp giữ nguyên.
6. Các ô AP, AS, AV, AW là công thức: KHÔNG ghi đè, Excel tự tính lại khi mở.
7. Lưu sổ giữ nguyên VBA (định dạng .xlsm), dọn file tạm tải về.
8. Kiểm tổng: truy thu + chậm nộp + phạt phải khớp ô tổng phát sinh trên sheet TINH PHAT (sai lệch chỉ do làm tròn).
9. Báo user dạng bảng: cột, chỉ tiêu, số cũ, số mới; liệt kê ô giữ nguyên; tổng sau cập nhật.

## Bẫy
- Sheet TINH PHAT không có chỉ tiêu giảm lỗ: không tự chốt cột AX, hỏi user hoặc lấy từ DTCHEP khi user đồng ý.
- Chưa ghi được ngược lên Drive (Composio upload chỉ nhận s3key trong kho của nó). Báo user sửa tay ô lệch trong sổ đang mở, hoặc đóng Excel rồi tải bản vault lên thay.
- Không ghi đè ô công thức, kể cả khi giá trị cache trong file trông sai.

## Kiểm chứng
- Tổng phát sinh của dòng vụ khớp ô tổng trên sheet TINH PHAT.
- VBA của sổ còn nguyên sau khi lưu.
- Không còn file tạm/bak trong thư mục export.
