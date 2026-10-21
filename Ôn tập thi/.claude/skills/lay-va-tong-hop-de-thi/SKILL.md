---
name: Truy cập và tổng hợp đề thi
description: "Truy cập hệ thống khảo thí thuế, tự động lấy đề, dò đáp án chuẩn và tổng hợp ngân hàng câu hỏi vào bảng tính Excel."
group: Năng suất
---

# Truy cập và tổng hợp đề thi - tự động hóa khảo thí và ngân hàng câu hỏi

Chạy bộ công cụ tự động hóa khảo thí thuế để kết nối cổng khảo thí trực tuyến (taxedu.gdt.gov.vn / api-e3.khaothi.online), bóc tách đề thi, giải đề đối soát đáp án chuẩn 100% và tổng hợp vào ngân hàng câu hỏi Master.

## Khi nào dùng

- Người dùng yêu cầu: "truy cập trang taxedu lấy đề", "giải đề kíp thi", "tổng hợp bộ đề thi trắc nghiệm", "lấy đề từ link khảo thí", "bổ sung câu hỏi vào ngân hàng đề Master".
- Có đường link kíp thi dạng `https://taxedu.gdt.gov.vn/t/<subscription>-<group_id>` (ví dụ `https://taxedu.gdt.gov.vn/t/P3-36751`) hoặc mã kíp thi trực tiếp.
- Cần tự động hóa quá trình làm bài và dò tìm đáp án chuẩn với tỷ lệ chính xác 100%, có độ trễ an toàn để tránh bị khóa phiên thi.
- Cần nạp, lọc trùng và gộp câu hỏi mới vào bảng tính Master `Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.xlsx`.
- Không dùng cho việc thi thật có giám thị hoặc các kỳ thi sát hạch chính thức.

## Chuẩn bị

1. Thư mục gốc Javis: `D:\SECOND BRAIN\javis-os`.
2. Thư mục lưu trữ học tập của người dùng: `D:\HOC TAP\ON THI 2026.09`.
3. Bảng tính Master Bank: `D:\HOC TAP\ON THI 2026.09\Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.xlsx`.
4. Môi trường thực thi:
   - Python 3 có sẵn thư viện `openpyxl` và `urllib`.
   - Node.js runtime (để chạy bộ giải mã `scripts/decrypt_dcv2.js` với thuật toán OpenSSL EVP AES-256-CBC, khóa `jcjioe92ujdkod021`).
5. Thư mục tài liệu đồng bộ trong Second Brain: `D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\sources`.

## Cách chạy

### 1. Truy cập, lấy đề và giải đề tự động từ hệ thống khảo thí

Chạy kịch bản `khao_thi_client.py` nằm trong thư mục scripts của skill:

```powershell
python "D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\skills\lay-va-tong-hop-de-thi\scripts\khao_thi_client.py" --url "https://taxedu.gdt.gov.vn/t/P3-36751"
```

Hoặc chỉ định trực tiếp mã kíp thi và mã phân hệ:

```powershell
python "D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\skills\lay-va-tong-hop-de-thi\scripts\khao_thi_client.py" --group-id 36751 --sub P3
```

Các tham số tùy chọn:
- `--name "Tên Thí Sinh"`: Tên thí sinh đăng ký tự do (mặc định lấy ngẫu nhiên).
- `--master "<đường_dẫn_file_master>"`: Đường dẫn file Master Bank để nạp cache đối chiếu trước khi gửi đáp án.

### 2. Gộp đề thi vào Bảng tính Ngân hàng Master

Chạy kịch bản `bank_aggregator.py` để khử trùng lặp và bổ sung câu hỏi mới:

```powershell
python "D:\SECOND BRAIN\javis-os\brains\Ôn tập thi\skills\lay-va-tong-hop-de-thi\scripts\bank_aggregator.py" --source "D:\HOC TAP\ON THI 2026.09\Bo_De_On_Thi_Kiem_Tra_Thue_P3_36751_Chuan.xlsx" --label "Kíp thi 36751"
```

## Quy trình

1. **Khởi tạo định danh thiết bị ảo**: Sinh `device_id` và thông tin trình duyệt Chrome trên Windows để đáp ứng tiêu chuẩn bảo mật của hệ thống khảo thí.
2. **Đăng ký thí sinh tự do**: Gửi request POST tới endpoint `guest-register`, nhận payload mã hóa và giải mã bằng `decrypt_dcv2.js` để lấy mã dự thi `code`.
3. **Mở phiên thi và nạp câu hỏi**: Gọi endpoint `start`, giải mã bộ đề để lấy token phiên làm việc cùng danh sách toàn bộ câu hỏi và các phương án lựa chọn.
4. **Tận dụng Cache Master Bank để tối ưu**:
   - Đối chiếu nội dung từng câu hỏi với ngân hàng Master đã giải trước đó.
   - Nếu câu hỏi đã có đáp án chuẩn trong Master Bank: Chọn ngay đáp án chuẩn ngay lượt gửi đầu tiên, giúp tiết kiệm thời gian và không tốn token suy luận.
5. **Dự đoán và dò đáp án chuẩn**:
   - Nếu câu mới: Áp dụng heuristic nhận diện (dấu chấm phẩy đơn, lựa chọn Tất cả các phương án, trích dẫn văn bản quy phạm pháp luật) để chọn phương án tối ưu ban đầu.
   - Áp dụng độ trễ an toàn ngẫu nhiên từ 5.1 đến 7.3 giây trước khi gửi.
   - Gửi kiểm tra qua endpoint `submit-item-answer`. Nếu phản hồi `is_correct: true`, xác nhận đáp án; nếu chưa đúng, tự động thử tiếp các phương án còn lại cho đến khi đạt đáp án đúng 100%.
6. **Xuất bảng tính và đồng bộ Second Brain**:
   - Ghi file Excel chi tiết của đề thi tại `D:\HOC TAP\ON THI 2026.09`.
   - Ghi file Markdown tổng hợp câu hỏi có định dạng vào `sources/` của brain `Ôn tập thi`.
   - Tự động nạp vào trang tính `Ngan_Hang_Cau_Hoi_Chuan` của Master Bank và cập nhật tỷ trọng chuyên đề trong trang `Thong_Ke_Chuyen_De`.

## Bẫy

- **Bẫy giãn cách tần suất**: Tuyệt đối không hạ độ trễ gửi đáp án xuống dưới 5 giây. Hệ thống khảo thí giám sát tần suất tự động, nếu gửi dồn dập sẽ bị khóa phiên thi hoặc chặn IP.
- **Bẫy giải mã dữ liệu**: Dữ liệu phản hồi từ máy chủ khảo thí được mã hóa theo chuẩn OpenSSL EVP AES-256-CBC (`Salted__`). Cần đảm bảo Node.js hoạt động để chạy tệp giải mã đi kèm.
- **Bẫy so khớp trùng lặp**: Nội dung câu hỏi trong đề thi thường chứa thẻ HTML thừa (`<p>`, `<br>`, `&nbsp;`), khoảng trắng kép hoặc dấu câu khác biệt. Cần chạy qua hàm chuẩn hóa văn bản trước khi so sánh với Master Bank.
- **Bẫy định dạng bảng tính**: Trang tính Master Bank có 2 sheet (`Ngan_Hang_Cau_Hoi_Chuan` và `Thong_Ke_Chuyen_De`). Khi thêm dòng mới, phải tô màu xanh lá cho ô đáp án chuẩn và tính toán lại bảng tỷ trọng chuyên đề.

## Kiểm chứng

- File Excel đề thi sinh ra mở được bình thường, đủ các cột A, B, C, D và cột Đáp Án Chuẩn.
- Tỷ lệ đáp án đúng từ hệ thống đạt 100% cho toàn bộ các câu hỏi.
- Master Bank sau khi tổng hợp không phát sinh câu hỏi trùng lặp nội dung.
- File Markdown xuất hiện trong `sources/` và có thể đọc trực tiếp từ Second Brain.
