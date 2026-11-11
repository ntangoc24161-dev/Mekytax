---
name: Kiểm tra tại bàn giải thể / chấm dứt MST
description: "Kiểm tra tại bàn giải thể/chấm dứt MST: đối chiếu số liệu, cập nhật Bảng tính, biên bản, checklist giải thể."
group: Tài chính
---

## Khi nào dùng

Kích hoạt khi xử lý hồ sơ kiểm tra TẠI BÀN phục vụ giải thể doanh nghiệp hoặc chấm dứt hiệu lực
mã số thuế (đóng MST) - khác kiểm tra theo kế hoạch (ra trụ sở NNT, có QĐ kiểm tra/giám sát, dùng
skill `ho-so-kiem-tra-ke-hoach`). Ví dụ: "kiểm tra tại bàn giải thể công ty X", "đóng MST, cập
nhật hồ sơ kiểm tra", "công ty ngừng hoạt động, rà hồ sơ trước khi trình", "bổ sung năm cuối vào
Bảng tính rồi cập nhật biên bản giải thể".

## Chuẩn bị

- Xác định: MST, thời kỳ kiểm tra (từ năm đầu chưa kiểm tới năm ngừng hoạt động), căn cứ ngừng
  hoạt động (số/ngày thông báo của Chi cục Thuế).
- Định vị: hồ sơ gốc DN theo từng năm (sổ kế toán, tờ khai), Bảng tính phân tích (thường
  `0.Bang tinh.xlsx` hoặc tương tự), hồ sơ Word mẫu/đang soạn (biên bản, tờ trình, phiếu nhận
  xét, phiếu chuyển nội bộ).
- Chưa có Bảng tính -> dùng skill `tao-bang-tinh-tu-ptdn` để dựng mới trước khi làm tiếp.

## Quy trình

### 1. Nhận diện bộ văn bản, lọc bản trùng, rà lỗi thể thức TRƯỚC khi điền số

Hồ sơ tại bàn/giải thể thường gồm 4 loại: Biên bản kiểm tra tại bàn, Tờ trình lãnh đạo xin duyệt,
Phiếu chuyển nội bộ, Phiếu nhận xét hồ sơ rủi ro. Nhiều loại có cả "bản trình" và "bản lưu" giống
hệt nhau - xác định đâu là bản chính trước khi sửa, đừng sửa hai bản riêng rẽ rồi để lệch nhau.

Trước khi điền số, đọc lướt cả 4 văn bản tìm:
- Khung biên bản sai (dùng mẫu xử phạt/VPHC thay vì mẫu BBKT/BBXN thuế).
- Placeholder còn sót (`<...>`), thiếu số/ký hiệu văn bản, lỗi chính tả rõ.
- Mâu thuẫn nội dung GIỮA các văn bản (ví dụ tờ trình ghi "nộp đầy đủ đúng hạn" nhưng biên bản
  ghi "chưa nộp") - đây là dấu hiệu một bên đang sai; đối chiếu nguồn thật ở bước 3 rồi mới sửa,
  không tự chọn bên nào đúng theo cảm tính.

### 2. Checklist nghiệp vụ giải thể bắt buộc (luôn kiểm, không phụ thuộc DT năm đó)

Dù hồ sơ được đánh giá rủi ro thấp, case giải thể/đóng MST phải kiểm đủ các mục sau, không bỏ
qua chỉ vì doanh thu năm cuối bằng 0:
- Công nợ phải thu/phải trả còn tồn tại thời điểm giải thể (131/331...).
- Hàng tồn kho còn lại, TSCĐ chưa thanh lý.
- Hóa đơn còn tồn/chưa sử dụng, hóa đơn đã hủy (đối chiếu thông báo hủy hóa đơn).
- Hoàn thuế GTGT (nếu có số dư khấu trừ cuối kỳ).
- Nghĩa vụ TNCN của người lao động năm cuối.
- QTT TNDN, BCTC, QTT TNCN năm cuối - xem bước 3, không suy đoán.

### 3. Đối chiếu số liệu từng năm từ NHIỀU nguồn, không tin một nguồn duy nhất

Với mỗi năm trong thời kỳ:
- Lấy đủ: sổ Nhật ký chung, sổ chi tiết tài khoản (đặc biệt 515 doanh thu tài chính, 6422 chi
  phí quản lý - hay bị bỏ sót), CĐKT, KQHDKD/PTDN02, tờ khai GTGT/TNCN theo quý, lệ phí môn bài.
- P&L (KQHDKD) ghi doanh thu = 0 KHÔNG đồng nghĩa năm đó không phát sinh gì: báo cáo năm chưa
  khóa sổ thường CHƯA kết chuyển 515/6422 sang 911. Luôn đối chiếu sổ chi tiết cấp tài khoản để
  bắt phát sinh thật (lãi tiền gửi, phí ngân hàng...).
- Có thể dùng skill `nhap-bang-tinh-tu-bao-cao-hdkd` để đổ KQKDSXKD từ PTDN02, và
  `boc-tach-bang-tinh-phan-tich` để tính tỷ lệ chi phí/doanh thu nếu cần phân tích thêm.
- Đối chiếu nghĩa vụ kê khai THẬT qua file nghĩa vụ/TMS (XML tờ khai) trước khi viết kết luận
  "đã nộp" hay "chưa nộp" trong biên bản. Không thấy trong hồ sơ hiện có thì ghi đúng "chưa thấy
  trong hồ sơ", KHÔNG khẳng định "chưa nộp" nếu chưa chắc.

### 4. Cập nhật Bảng tính - BẮT BUỘC backup và kiểm cân đối

- Backup file trước MỖI lần ghi (đặt tên có timestamp, lưu cùng thư mục vụ).
- Ghi KQKDSXKD và CĐKT theo đúng cột năm; cột năm ngoài kỳ kiểm tra (vd cột mẫu còn để năm cũ)
  phải đổi lại cho khớp thời kỳ thật trước khi ghi.
- Sau khi ghi CĐKT mỗi năm, kiểm NGAY Tổng tài sản = Tổng nguồn vốn. Lệch thì tìm nguyên nhân
  (thường do sổ chưa kết chuyển 515/6422 sang 911, hoặc chỉ tiêu lợi nhuận chưa cập nhật) và nói
  rõ số lệch, không ghi đè cho khớp.
- Rà sheet xem có phôi của công ty khác lẫn vào (tên/MST khác do copy từ vụ trước) - xóa phôi
  lạ, không đụng số liệu vụ đang làm.
- Rà toàn bộ `#REF!`: sửa bằng cách suy chiếu từ công thức hàng xóm còn sống (cùng hàng/cùng
  khối năm). Khi không có bản gốc để soi (ô đích cũ đã mất, phải dựng lại công thức theo cấu
  trúc dữ liệu) - phải NÓI RÕ đây là công thức dựng lại theo suy luận, không phải khớp nguyên
  bản, để người kiểm biết mức tin cậy.

### 5. Cập nhật biên bản/hồ sơ Word

- Backup trước khi sửa.
- Đổ bảng theo 3 cột Số báo cáo / Số kiểm tra / Chênh lệch cho từng năm có phát sinh thật (lấy
  từ Bảng tính đã cập nhật ở bước 4).
- Viết phần đánh giá sai sót: PHÂN BIỆT RÕ "sai sót sổ sách" (vd chưa kết chuyển 515/6422, không
  ảnh hưởng số thuế phải nộp) với "sai phạm thuế" (ảnh hưởng số thuế phải nộp, phải truy thu/xử
  phạt). Không gộp hai loại lại thành một kết luận.
- Sửa câu kết luận cho khớp kết quả đối chiếu TMS ở bước 3 (ví dụ bỏ câu "chưa nộp QTT/BCTC/TNCN"
  nếu thực tế đã nộp).
- Có thể dùng skill `soan-thao-van-ban` để chuẩn hóa thể thức theo NĐ 30/2020 sau khi đã điền số
  xong.

## Bẫy

- Sửa cả bản trình và bản lưu riêng rẽ rồi quên đồng bộ - hai bản lệch số.
- Kết luận "đã/chưa nộp" chỉ dựa vào bảng số liệu kiểm tra, không đối chiếu TMS - dễ sai (từng
  xảy ra: biên bản ghi "chưa nộp" trong khi DN đã nộp đầy đủ).
- Doanh thu = 0 nhưng vẫn có phát sinh thật (lãi tiền gửi, phí ngân hàng, biến động công nợ) -
  đừng kết luận "không phát sinh" chỉ vì P&L chưa khóa sổ.
- Ghi đè Bảng tính không backup, hoặc ghi mà không kiểm TS = NV ngay sau đó.
- Sửa `#REF!` mà không nói rõ chỗ nào là suy luận dựng lại, chỗ nào khớp công thức gốc còn sống.
- Bỏ qua checklist công nợ/tồn kho/TSCĐ/hóa đơn tồn/hoàn thuế chỉ vì hồ sơ được đánh giá rủi ro
  thấp.

## Kiểm chứng

- Mọi năm trên CĐKT: Tổng tài sản = Tổng nguồn vốn (hoặc nêu đúng số lệch và nguyên nhân nếu
  có).
- Không còn `#REF!` nào trên các sheet đã sửa (quét lại sau khi sửa).
- Số trên biên bản khớp đúng số trên Bảng tính đã cập nhật (không còn số mẫu/số vụ khác lẫn
  vào).
- Mọi câu kết luận "đã nộp/chưa nộp" trong biên bản có đối chiếu TMS làm căn cứ, không suy đoán.
- Checklist giải thể (công nợ, tồn kho, TSCĐ, hóa đơn, hoàn thuế, TNCN năm cuối) đã được nêu ý
  kiến, không để trống.
