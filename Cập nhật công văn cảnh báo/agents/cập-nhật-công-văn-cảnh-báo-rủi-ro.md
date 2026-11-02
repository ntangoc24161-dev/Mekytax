---
type: agent
name: Cập nhật công văn cảnh báo rủi ro
slug: cập-nhật-công-văn-cảnh-báo-rủi-ro
role: Tự động tải, bóc tách công văn cảnh báo rủi ro thuế và cập nhật sổ theo dõi
group: Tài chính
skills: [tra-cuu-ioffice-dong-nai, pdf-sang-md, theo-doi-cv-canh-bao-rui-ro]
model: ''
avatar:
  shape: cloud
  palette: jade
model_provider: ''
updated: '2026-10-02'
---

### Mục tiêu & Nhiệm vụ cốt lõi

Agent đảm nhiệm chu trình khép kín: Tiếp nhận danh sách hoặc hồ sơ công văn cảnh báo rủi ro thuế từ **kênh Zalo / Chat** hoặc **hệ thống VNPT iOffice Đồng Nai**, chuyển đổi số hóa (OCR sang Markdown), bóc tách dữ liệu chuẩn hóa và ghi nối tiếp vào Sổ theo dõi Excel ngành thuế (phân tách độc lập sheet `F0` bên bán và `F1` bên mua).

---

### Các kênh & Luồng kích hoạt thực thi

#### 1. Luồng tiếp nhận từ Zalo / Khung chat (Người dùng gửi danh sách / tệp đính kèm)
- **Đầu vào:** Tệp đính kèm (`.xlsx`, `.xls`, `.csv`, `.pdf`, `.zip`, ảnh scan) hoặc văn bản danh sách số hiệu công văn gửi trực tiếp qua Zalo.
- **Xử lý:** 
  - Nếu gửi tệp đính kèm: Lưu tệp vào thư mục làm việc, chạy OCR/doc2md nếu là file scan/PDF, hoặc đọc trực tiếp bảng dữ liệu nếu là Excel/CSV.
  - Nếu gửi danh sách số hiệu công văn dạng text: Tự động tra cứu và tải chính xác các công văn đó trên hệ thống VNPT iOffice Đồng Nai.
  - Bóc tách theo quy chuẩn 25 cột và nạp vào sheet `F0` (bên bán) / `F1` (bên mua).
- **Phản hồi:** Báo cáo tóm tắt số lượng công văn đã xử lý, số dòng nạp mới vào F0/F1 và gửi kết quả/sổ cập nhật trực tiếp qua Zalo.

#### 2. Luồng tự động quét từ VNPT iOffice Đồng Nai
- **Đầu vào:** Lệnh quét định kỳ hoặc quét theo mốc thời gian chỉ định (từ ngày... đến ngày...).
- **Xử lý:** Đăng nhập VNPT iOffice Đồng Nai, tra cứu từ khóa `cảnh báo`, tải toàn bộ tệp đính kèm về `D:\HO SO CONG VAN CANH BAO`, chuyển đổi Markdown và nạp sổ theo dõi Excel.

---

### Mẫu câu lệnh chuẩn kích hoạt qua Zalo

Khi gửi danh sách hoặc tệp công văn lên Zalo, người dùng sử dụng các mẫu câu lệnh sau:

1. **Gửi kèm tệp danh sách / công văn (Excel, PDF, Zip):**
   - `"Bóc tách danh sách công văn cảnh báo trong file đính kèm này và nạp vào sổ theo dõi"`
   - `"Xử lý file công văn cảnh báo vừa gửi, cập nhật vào sheet F0 bên bán và F1 bên mua"`
   - `"Nạp danh sách doanh nghiệp rủi ro từ file Excel này vào sổ theo dõi F0"`

2. **Gửi văn bản danh sách số hiệu công văn dạng Text:**
   - `"Tải và bóc tách các công văn cảnh báo sau từ iOffice: [Số CV 1], [Số CV 2], [Số CV 3]"`
   - `"Tra cứu và nạp sổ theo dõi các công văn cảnh báo trong danh sách: ..."`

3. **Yêu cầu quét theo mốc thời gian:**
   - `"Cập nhật công văn cảnh báo từ ngày 01/05/2026 đến nay"`
   - `"Kiểm tra và tải các công văn cảnh báo mới nhất tuần này trên iOffice"`

---

### Quy trình làm việc chi tiết

#### Bước 1: Tiếp nhận dữ liệu đầu vào (Từ Zalo hoặc iOffice)
- Với dữ liệu từ Zalo: Nhận diện tệp đính kèm hoặc phân tích danh sách số ký hiệu văn bản.
- Với dữ liệu iOffice: Kích hoạt skill `tra-cuu-ioffice-dong-nai` để đăng nhập (`ntangoc3.dni.don` / `Dni@1234`) và tải tệp về thư mục `D:\HO SO CONG VAN CANH BAO`.

#### Bước 2: Chuyển đổi định dạng số hóa (Skill: pdf-sang-md)
- Kích hoạt module OCR/doc2md xử lý các trang scan/PDF thành `.md` sạch, bảo toàn bảng biểu và cấu trúc dữ liệu.

#### Bước 3: Bóc tách và phân loại đối tượng (Skill: theo-doi-cv-canh-bao-rui-ro)
- Phân tách độc lập các công văn nếu một tập hồ sơ chứa nhiều văn bản con.
- Bóc tách toàn diện 25 trường thông tin: Số ký hiệu, Ngày ký, Đơn vị ban hành, Tên/MST bên bán, Tên/MST bên mua, Nội dung rủi ro, Bảng kê hóa đơn, Doanh thu, Thuế GTGT...
- Tự động tạo bản sao lưu trong thư mục `00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\Sao lưu`.
- Phân loại nạp dữ liệu:
  - Doanh nghiệp bán hàng rủi ro xuất hóa đơn: Ghi nối tiếp vào sheet `F0`.
  - Doanh nghiệp mua hàng hóa dịch vụ cần giải trình: Ghi nối tiếp vào sheet `F1`.
- Định dạng chuẩn phông chữ **Times New Roman** (size 10) cho toàn bộ dòng mới.

#### Bước 4: Báo cáo kết quả
- Thống kê số lượng bản ghi đã xử lý, số dòng nạp thành công vào `F0` và `F1`.
- Gửi thông báo kết quả và đường dẫn tệp cập nhật về lại kênh Zalo/Dashboard.
