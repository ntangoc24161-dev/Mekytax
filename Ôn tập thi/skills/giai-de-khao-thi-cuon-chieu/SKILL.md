---
name: Giải đề khảo thí cuốn chiếu
description: "Giải đề trắc nghiệm taxedu, dò đáp án chuẩn 100%, tra cứu 2 tầng ngân hàng và AI, lưu cuốn chiếu liên tục ra Excel."
group: Năng suất
---

# Giải đề khảo thí cuốn chiếu - tự động hóa thi trắc nghiệm và bảo toàn dữ liệu

Kỹ năng điều phối kịch bản tự động truy cập cổng khảo thí taxedu (api-e3.khaothi.online), bóc tách đề thi nguyên bản 90 câu, tra cứu đáp án 2 tầng (Ngân hàng Master và AI Solver), dò đáp án chuẩn 100% từ máy chủ với giãn cách an toàn, nhận diện đề cũ để làm tiếp câu còn thiếu, tự động bỏ qua đề bị tạm dừng để sang đề tiếp theo, và thu thập tự động đến khi đủ 40 đề thi.

## Khi nào dùng

- Người dùng yêu cầu: "thu thập đề thi cuốn chiếu", "giải đề lưu cuốn chiếu", "chạy collector 40 đề", "tiếp tục thực hiện đề mới đến khi đủ 40 đề", "nếu tạm dừng thì tự động bỏ qua", "nhận diện đề cũ thì làm tiếp câu còn thiếu".
- Có kíp thi trực tuyến trên cổng `taxedu.gdt.gov.vn` hoặc `api-e3.khaothi.online` (ví dụ kíp `P3-39917`).
- Cần tự động hóa làm bài thi với cơ chế bảo toàn tiến độ tối đa: lưu file Excel liên tục sau từng câu để không mất dữ liệu khi gián đoạn mạng hoặc ngắt phiên.
- Cần giải đề kết hợp 2 tầng tri thức: tầng 1 tra cứu ngân hàng câu hỏi có sẵn, tầng 2 nạp AI chuyên sâu về thuế để dự đoán phương án đúng ngay lượt đầu.
- Cần kết xuất bộ đề chuẩn hóa 14 cột tại thư mục `40 DE THI` với đáp án đúng được máy chủ khảo thí xác nhận.

## Chuẩn bị

1. Thư mục lưu trữ học tập của người dùng: `D:\HOC TAP\ON THI 2026.09`.
2. Kịch bản thực thi chính: `D:\HOC TAP\ON THI 2026.09\collector_40_de.py`.
3. Bộ giải thuật toán AI Thuế: `D:\HOC TAP\ON THI 2026.09\tax_solver_engine.py`.
4. Kịch bản giải mã payload máy chủ: `D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\skills\lay-va-tong-hop-de-thi\scripts\decrypt_dcv2.js`.
5. Ngân hàng câu hỏi đối soát: `D:\HOC TAP\ON THI 2026.09\Tong_Hop_Ngan_Hang_1247_Cau_Kip_P3_39917.xlsx`.
6. Thư mục kết xuất đề thi: `D:\HOC TAP\ON THI 2026.09\40 DE THI`.
7. Tệp đồng bộ tiến độ: `D:\HOC TAP\ON THI 2026.09\checkpoint_P3_39917.json`.
8. Tệp nhật ký chi tiết: `D:\HOC TAP\ON THI 2026.09\collector_40_de.log`.
9. Môi trường thực thi: Python 3 (thư viện `openpyxl`, `urllib`), Node.js runtime (giải mã AES-256-CBC).

## Cách chạy

### 1. Chạy tự động batch 40 đề (Mặc định)

Mở terminal hoặc PowerShell trên máy tính và chạy script:

```powershell
python "D:\HOC TAP\ON THI 2026.09\collector_40_de.py" --max-exams 40
```

Kịch bản sẽ tự động:
- Quét các đề đã có trong thư mục `40 DE THI`.
- Mở phiên thi mới từ máy chủ.
- Nhận diện nếu là đề cũ chưa xong: đồng bộ nhanh các câu đã có và làm tiếp từ câu còn thiếu đến câu 90.
- Nhận diện nếu là đề mới: gán số thứ tự đề tiếp theo và quét từ câu 1 đến 90.
- Nếu gặp lỗi hoặc bị tạm dừng ở đề nào: lưu cuốn chiếu dữ liệu hiện tại, tự động bỏ qua đề đó và chuyển ngay sang đề tiếp theo.
- Lặp lại cho đến khi đủ 40 đề thi hoàn tất 100%.

### 2. Chạy đơn lẻ từng đề chỉ định

```powershell
python "D:\HOC TAP\ON THI 2026.09\collector_40_de.py" --exam-num 2
```

## Quy trình

1. **Khởi tạo định danh và thiết bị ảo**:
   - Tự động sinh tên thí sinh ngẫu nhiên từ danh sách họ và tên tiếng Việt.
   - Tạo mã `device_id` và thông số trình duyệt Chrome trên Windows để đảm bảo phiên kết nối hợp lệ.
2. **Đăng ký và mở phiên làm bài thi**:
   - Gửi yêu cầu tới endpoint `guest-register`, nhận payload mã hóa và giải mã bằng `decrypt_dcv2.js` để lấy mã dự thi `code`.
   - Gọi endpoint `start`, giải mã bộ đề thi để lấy token phiên cùng trọn vẹn 90 câu hỏi và 4 phương án lựa chọn nguyên bản.
3. **Đối chiếu và nhận diện mẫu đề**:
   - So sánh chữ ký câu hỏi (3 câu đầu hoặc độ trùng khớp >80 câu) với tất cả các đề đã có trong thư mục `40 DE THI`.
   - Nếu trùng đề cũ chưa xong: nạp toàn bộ câu đã xác thực, gửi đồng bộ nhanh lên máy chủ, tiếp tục giải cuốn chiếu từ câu còn thiếu đến câu 90.
   - Nếu trùng đề cũ đã xong 90/90: tự động bỏ qua để bốc đề khác.
   - Nếu là đề mới: gán số thứ tự tiếp theo (Đề 03, Đề 04... Đề 40) và quét từ câu 1 đến câu 90.
4. **Cơ chế tra cứu và giải đề 2 tầng**:
   - Tầng 1 (Ngân hàng đề): Đối chiếu `item_id` hoặc nội dung câu hỏi với ngân hàng 1.247 câu. Nếu đã có đáp án xác thực, chọn ngay phương án chuẩn và ghi chú nguồn "Từ ngân hàng đề".
   - Tầng 2 (AI Solver): Nếu là câu mới hoặc chưa xác thực, gọi `tax_solver_engine.py` kết hợp tri thức từ các skill thuế để suy luận phương án tối ưu và ghi chú nguồn "AI đề xuất".
5. **Dò đáp án chuẩn 100% và giãn cách an toàn**:
   - Áp dụng thời gian trễ ngẫu nhiên từ 2.0s đến 4.0s trước mỗi lần gửi phương án để tăng tốc xử lý nhưng vẫn bảo đảm an toàn kết nối.
   - Gửi phương án qua endpoint `submit-item-answer`. Nếu phản hồi `is_correct: true`, khóa đáp án đúng.
   - Nếu phương án ban đầu chưa đúng, tự động gửi tiếp các phương án còn lại (kèm giãn cách an toàn 2.0s - 4.0s) cho đến khi máy chủ phản hồi đúng.
   - Trích xuất nội dung `instruction` trả về từ máy chủ để bổ sung vào cột Căn cứ pháp lý.
6. **Lưu cuốn chiếu liên tục sau từng câu**:
   - Ghi đè file Excel ngay sau khi mỗi câu hỏi có kết quả xác thực.
   - Bắt ngoại lệ `PermissionError` an toàn: nếu người dùng đang mở xem file Excel, script tạm bỏ qua lượt ghi đó và tiếp tục làm bài mà không bị dừng tiến trình.
   - Cập nhật song song file `checkpoint_P3_39917.json` và ghi nhật ký vào `collector_40_de.log`.
7. **Tự động phục hồi và chuyển tiếp (Auto-skip)**:
   - Nếu một phiên thi bị mất kết nối hoặc tạm dừng: dữ liệu đã ghi cuốn chiếu được bảo toàn 100%. Tiến trình ghi nhận log, tự động bỏ qua đề đó và mở phiên thi mới cho đề tiếp theo.

## Bẫy

- **Bẫy khóa file khi đang mở**: Người dùng mở file Excel trong lúc script đang chạy sẽ gây xung đột ghi tệp. Luôn phải bọc lệnh ghi file trong khối `try...except PermissionError` để bảo vệ luồng chạy chính.
- **Bẫy tần suất thao tác**: Tuyệt đối không hạ độ trễ gửi đáp án xuống dưới 5.1 giây. Cần duy trì ngẫu nhiên 5.1s đến 8.0s giữa các lần chọn để bảo vệ phiên kết nối.
- **Bẫy ngắt kết nối giữa chừng**: Cơ chế lưu cuốn chiếu đảm bảo dữ liệu luôn được lưu. Khi xảy ra lỗi mạng, cơ chế try/except tự động bắt lỗi và chuyển sang đề tiếp theo mà không làm chết toàn bộ chuỗi 40 đề.
