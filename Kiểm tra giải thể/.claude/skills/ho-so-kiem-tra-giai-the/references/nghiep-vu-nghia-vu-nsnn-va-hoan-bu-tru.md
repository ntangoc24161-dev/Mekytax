# Nghiệp vụ Mục III Nghĩa vụ NSNN và Hoàn thuế bù trừ giải thể

## 1. Phương pháp suy luận và điền Mục III Nghĩa vụ NSNN

Trong hồ sơ kiểm tra theo kế hoạch thông thường, Mục III thường để trống hoặc gạch ngang do chỉ kiểm tra theo chuyên đề và việc theo dõi nợ thuộc bộ phận Quản lý nợ. Tuy nhiên, đối với **hồ sơ kiểm tra giải thể**, CQT phải chốt toàn diện nghĩa vụ thuế trước khi đóng mã số thuế (chấm dứt hiệu lực MST), do đó **Mục III bắt buộc phải điền đầy đủ số phát sinh các năm**.

### Công thức suy luận số liệu (khi chưa có sao kê chi tiết từ TMS)
Số liệu Mục III được suy ra từ Kết luận kiểm tra (Mục 1.2 về tình hình kê khai, nộp thuế và Mục 1.5 về tình trạng nợ NSNN):
- Cột (1): Năm kiểm tra.
- Cột (2): Sắc thuế (Thuế GTGT, Thuế TNDN, Thuế TNCN, Lệ phí môn bài...).
- Cột (3) - Số nợ kỳ trước chuyển sang: Thường bằng `0` (khi doanh nghiệp không có nợ đọng cũ).
- Cột (4) - Số phát sinh phải nộp theo kê khai: Lấy từ số liệu quyết toán của doanh nghiệp (hoặc ô tổng nộp).
- Cột (5) - Số phát sinh tăng/giảm qua kiểm tra: Lấy đúng số thuế truy thu kiểm tra xác định từng năm.
- Cột (6) - Số đã thực nộp trong kỳ: Bằng số phát sinh theo kê khai `(6) = (4)` (vì doanh nghiệp đã nộp đủ số kê khai trước giải thể).
- Cột (7) - Số còn phải nộp cuối kỳ: Tính theo công thức chuẩn:
  $$\text{Cột (7)} = \text{Cột (3)} + \text{Cột (4)} + \text{Cột (5)} - \text{Cột (6)}$$
  Vì $\text{Cột (3)} = 0$ và $\text{Cột (6)} = \text{Cột (4)}$, nên $\text{Cột (7)} = \text{Cột (5)}$ (chính là số thuế truy thu còn phải nộp qua kiểm tra).

### Ví dụ thực tế vụ Sơn Tiền Hải (5 năm):
- Năm 2022, 2023: Không phát sinh truy thu $\rightarrow$ Cột (7) = 0.
- Năm 2024: Đã nộp 11.591.829 đ, không truy thu $\rightarrow$ Cột (7) = 0.
- Năm 2025: Đã nộp 5.814.545 đ, truy thu TNDN 10.180.180 đ $\rightarrow$ Cột (7) = 10.180.180 đ.
- Năm 2026: Kê khai 0 đ, truy thu TNDN sau giảm lỗ 2.872.243 đ $\rightarrow$ Cột (7) = 2.872.243 đ.

---

## 2. Thủ tục hoàn thuế kiêm bù trừ thu NSNN giải thể (Điều 32 Nghị định 181/2025/NĐ-CP)

Khi doanh nghiệp làm thủ tục giải thể, nếu tại tờ khai GTGT kỳ cuối cùng có số thuế GTGT chưa khấu trừ hết (chuyển kỳ sau):

### Cơ chế áp dụng:
1. **Không áp dụng điều kiện hoàn xuất khẩu thông thường**: Kể cả trường hợp doanh nghiệp có doanh thu xuất khẩu nhưng số thuế chưa khấu trừ dưới 300 triệu đồng (không đủ điều kiện hoàn theo khoản 1 Điều 15 Luật Thuế GTGT 48/2024/QH15), đoàn kiểm tra chuyển căn cứ hoàn thuế sang **khoản 4 Điều 15 Luật 48/2024/QH15 và Điều 32 Nghị định 181/2025/NĐ-CP (diện giải thể)**.
2. **Cơ chế bù trừ tự động**:
   - Số thuế GTGT được hoàn sẽ được ưu tiên **bù trừ trực tiếp** với:
     + Số thuế TNDN, GTGT truy thu qua kiểm tra.
     + Tiền phạt vi phạm hành chính về thuế, hóa đơn.
     + Tiền chậm nộp tiền thuế tính đến ngày ký biên bản/quyết định.
   - Phần chênh lệch còn lại (nếu số hoàn lớn hơn số nợ): Ngân sách hoàn trả cho doanh nghiệp.
   - Nếu số hoàn nhỏ hơn: Doanh nghiệp nộp thêm phần thiếu vào NSNN trước khi khóa mã số thuế.
3. **Bộ hồ sơ trình ký**:
   - Tờ trình đề xuất hoàn thuế kiêm bù trừ thu NSNN.
   - Phiếu đề xuất hoàn thuế (mẫu `MS-03-DCS`).
   - Dự thảo Quyết định hoàn thuế kiêm bù trừ thu NSNN.

---

## 3. Văn bản từ chối nhận lại số tiền nộp thừa

- **Mục đích**: Trường hợp doanh nghiệp có số tiền thuế nộp thừa nhỏ (vài chục ngàn đến vài trăm ngàn đồng) hoặc doanh nghiệp không có nhu cầu nhận lại để đẩy nhanh tiến độ đóng MST.
- **Yêu cầu kỹ thuật**:
  + Thể thức: Văn bản theo mẫu của cơ quan thuế hoặc văn bản cam kết của doanh nghiệp.
  + **Mã số thuế**: Trình bày dưới dạng bảng ô lưới rời rạc, mỗi ô tương ứng 1 chữ số của MST (10 chữ số chính + 3 chữ số chi nhánh nếu có). Khi điền phải điền từng ô, không dồn cả chuỗi vào ô đầu tiên.
  + Kiểm tra kỹ MST trên văn bản này, tránh để sót MST của phôi vụ cũ (như từng phát hiện ở bộ Mo Fa).
