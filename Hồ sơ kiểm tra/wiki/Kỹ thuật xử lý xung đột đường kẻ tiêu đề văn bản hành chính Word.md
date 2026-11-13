---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
source: [[conversations/2026-11-13]]
---
Trong quá trình tự động hóa lập hồ sơ kiểm tra thuế bằng Microsoft Word, hiện tượng tiêu đề văn bản bị gạch dưới 2 lần (double underline) xuất hiện khi cấu trúc bảng tiêu đề của phôi mẫu chứa hàng đệm thừa [[conversations/2026-11-13]].

### Nguyên nhân kỹ thuật
1. **Phôi gốc có hàng đệm chứa shape:** Một số biểu mẫu Word cũ thiết kế bảng tiêu đề 2 cột, 2 hàng. Hàng 1 chứa chữ (Quốc hiệu, Tiêu ngữ, Tên cơ quan, Trích yếu), hàng 2 chứa đường kẻ đồ họa `v:line` [[conversations/2026-11-13]].
2. **Xung đột chèn đường kẻ tự động:** Các script chuẩn hóa thể thức theo Nghị định 30/2020/NĐ-CP kiểm tra đoạn văn tại hàng 1, nhận diện chưa có đường kẻ gạch chân nên tự động tạo thêm một thẻ `w:drawing` hoặc `v:line` ngay dưới đoạn văn đó [[conversations/2026-11-13]]. Khi render trên Word, cả 2 hàng đều có nét kẻ, tạo thành 2 đường gạch song song chồng lên nhau [[conversations/2026-11-13]].

### Quy chuẩn xử lý trên phôi gốc
- Xóa bỏ triệt để hàng số 2 trống trong bảng tiêu đề của phôi tổng hợp `templates/MAUTONG HOP.docx` [[conversations/2026-11-13]].
- Quét và loại bỏ các đối tượng shape `v:line` hoặc drawing nằm ẩn trong các ô ngày tháng hoặc ô đệm [[conversations/2026-11-13]].
- Giữ cấu trúc bảng tiêu đề 1 hàng duy nhất cho mỗi phần chữ và gắn trực tiếp đường kẻ dưới đoạn văn bản tương ứng theo chuẩn Nghị định 30/2020/NĐ-CP [[conversations/2026-11-13]].
