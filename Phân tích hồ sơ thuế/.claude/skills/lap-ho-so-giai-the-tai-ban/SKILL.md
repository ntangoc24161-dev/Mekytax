---
name: "Lập hồ sơ giải thể tại bàn"
description: "Lập hồ sơ giải thể tại bàn từ folder NNT: nạp Bang tinh từ phôi chuẩn, bóc XML/HĐĐT, điền Word theo TB ngừng MST."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-10-16
updated: 2026-11-13
---
## Khi nào dùng
- User yêu cầu lập hồ sơ giải thể / giải thể tại bàn hoặc xử lý biên bản từ `Bang tinh.xlsx`.
- Path folder nằm dưới `giai the` (ví dụ `2026/dang kiem tra/giai the/<ten-dn>`) hoặc thư mục kiểm tra NNT.
- Đã có hoặc cần tạo mới `Bang tinh.xlsx` và Thông báo ngừng MST (TB-TCS).

Không dùng skill này cho phân tích Master 8 sheet hay đối chiếu GTGT-TNDN thường.

## Chuẩn bị
- Path folder tuyệt đối hoặc tương đương Drive của DN.
- **Phôi chuẩn `Bang tinh.xlsx`** (đã làm sạch 100% lỗi `#REF!`, dọn vùng nháp AI-AO, xóa dữ liệu mẫu cũ):
  - Mặc định trong vault: `thu-muc-mau-ho-so-doanh-nghiep/Bang tinh.xlsx` (dự phòng: `sao-luu/Bang tinh.xlsx`).
  - Bản gốc trên máy: `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\BIEN BAN\Bang tinh.xlsx`.
- File nguồn tối thiểu: `Bang tinh.xlsx`, Thông báo ngừng MST (số TB, ngày TB, ngày ngừng MST), XML tờ khai (GTGT, TNDN, BCTC).
- Mẫu Word hồ sơ giải thể tại bàn (tên kiểu `HO_SO_GIAI_THE_TAI_BAN_*.docx`).
- Định danh DN: tên đầy đủ, MST, địa chỉ, người đại diện, KTT, ngày thành lập/cấp GCN, ngành nghề (chỉ điền khi có nguồn).

## Quy trình
1. **Nhận diện loại việc:** nếu path có `giai the` hoặc user nói giải thể/biên bản thì chạy nhánh này; ghi rõ MST và tên DN.
2. **Khởi tạo và kiểm kê folder:**
   - Liệt kê Word/Excel/TB; nếu folder DN chưa có `Bang tinh.xlsx`, tự động sao chép từ phôi chuẩn `thu-muc-mau-ho-so-doanh-nghiep/Bang tinh.xlsx` vào folder DN (tại `0. BIEN BAN\` hoặc `3.4. Giai trinh & Ho so bo sung\`).
   - Tuyệt đối không lấy file từ DN cũ chưa làm sạch để tránh sót dữ liệu mẫu và lỗi công thức.
3. **Bóc định danh:** tên, MST, địa chỉ, GĐ, KTT, ngày thành lập/cấp GCN, số TB, ngày TB, ngày ngừng MST, cán bộ thụ lý.
4. **Bóc tách và nạp số liệu vào `Bang tinh.xlsx`:**
   - Đọc trực tiếp từ XML tờ khai (GTGT, QTT TNDN, BCTC) để nạp vào các sheet `GTGT`, `KQKDSXKD`, `CDKT`.
   - Tuyệt đối không dùng số 0 giả khi file bảng tính mới là khung mẫu trống.
   - Chỉ lấy các năm phát sinh thực tế; không suy diễn năm trước thành lập (nếu năm thành lập chưa phát sinh thì ghi 0 hoặc để trống đúng mẫu).
   - Quét NCC rủi ro (F0, F1, CVCBRR, bỏ trốn) để tự động đẩy hóa đơn bị loại trừ vào `BANG KE KO HOP LE`.
   - Tính tự động tiền phạt vi phạm và tiền chậm nộp tại sheet `TINH PHAT + CHAM NOP`.
   - Quy chuẩn định dạng: ngày tháng `dd/MM/yyyy`, số tiền dùng dấu chấm `.` phân cách hàng nghìn, dấu phẩy `,` cho phần thập phân, phông chữ Times New Roman Unicode.
5. **Điền Word:** thay toàn bộ placeholder/mẫu cũ; xuất bản vault tại `exports/<slug-mst>/` kèm `Bang tinh.xlsx`, Word, zip nếu có.
6. **Rà sót mẫu:** tìm-thay mọi ngày/MST/tên còn của DN mẫu (ví dụ ngày thành lập mẫu cũ).
7. **Upload:** đưa Word và `Bang tinh.xlsx` lên đúng folder Drive của DN; trả link file và link folder.
8. **Báo còn trống:** liệt kê mục chưa có nguồn (ngành nghề, CCCD, giờ ký, doanh thu năm ngoài bảng…) để user bổ sung tay.

## Bẫy
- Nhầm sang playbook TTKT/Master 8 sheet chỉ vì user nói “lập hồ sơ kiểm tra”.
- Dùng file `Bang tinh.xlsx` cũ chưa làm sạch dẫn đến lỗi `#REF!` hoặc sót dữ liệu của DN khác.
- Điền số 0 giả khi phôi còn trống thay vì bóc tách số liệu thật từ XML tờ khai.
- Để sót ngày/MST/tên của bản mẫu trong Word (đã gặp sót `15/08/2011`).
- Bịa ngành nghề khi MST không khớp Danh bạ NNT: để placeholder và nói rõ còn thiếu.
- Ghi doanh thu các năm trước ngày thành lập như thể có phát sinh.
- Hứa “sẽ báo lại” mà không làm ngay trong lượt hoặc không xếp việc nền.

## Kiểm chứng
- File `Bang tinh.xlsx` được nhân bản từ phôi chuẩn sạch 100% lỗi `#REF!`, liên kết công thức MST và tên NNT tự nhảy đúng.
- Word trên vault và Drive cùng MST/tên DN đang làm.
- Grep/search trong Word: không còn ngày thành lập/MST/tên mẫu cũ.
- Số TB và ngày ngừng MST khớp Thông báo nguồn.
- Các năm số liệu khớp `Bang tinh.xlsx` và tờ khai XML gốc; mục thiếu nguồn được liệt kê công khai.
