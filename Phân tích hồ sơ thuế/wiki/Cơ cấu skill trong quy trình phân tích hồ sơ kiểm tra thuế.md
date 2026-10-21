---
type: wiki
status: active
tags: [wiki]
origin: javis-learned
created: 2026-10-15
updated: 2026-10-15
source: [[conversations/2026-10-15]]
---
Quy trình phân tích hồ sơ kiểm tra thuế được vận hành thông qua 1 skill trục chính và hệ thống các skill vệ tinh chuyên trách [[conversations/2026-10-15]].

### 1. Skill trục điều phối
- `phan-tich-ho-so-thue`: Đóng vai trò skill trục điều phối toàn bộ chu trình phân tích từ tổng hợp dữ liệu đầu vào HĐĐT/bảng kê/NLĐ (mẫu 02), làm sạch dữ liệu CD1 (BT5-BT9), rà soát thuế TNCN đa nguồn, đối chiếu XNT kho, giá thành, công nợ/tiền mặt, đến dựng Master 8 sheet, đánh giá 28 tiêu chí rủi ro và lập dự thảo đề cương kiểm tra [[conversations/2026-10-15]].

### 2. Các skill vệ tinh đã tách riêng
- `to-chuc-ho-so-kiem-tra-thue`: Chuyên trách Bước 0, tổ chức cây thư mục chuẩn 4 bước (`<TEN>_<MST>`), phân loại và map file dữ liệu, xuất 1 file checklist A4 tổng hợp tình trạng hồ sơ, và xóa sạch nguồn cũ [[conversations/2026-10-15]].
- `doi-chieu-ncc-mua-vao-rui-ro`: Chuyên trách Bước 3 (hoặc 4.2), đối chiếu nhà cung cấp mua vào từ Bảng kê và HĐĐT với sổ CVCBRR F0/F1, danh sách Công an, phân loại 3 nguồn (Bảng kê/HĐĐT/Chung) và dựng Bảng 2.3 theo kỳ [[conversations/2026-10-15]].
- `doi-chieu-511-632`: Chuyên trách Bước 5 (hoặc 4.7), đối chiếu sổ chi tiết doanh thu TK 511 với giá vốn TK 632 theo công thức Tool_DC_511_632, phân tích tỷ lệ lãi gộp và bóc tách các giao dịch bán dưới giá vốn [[conversations/2026-10-15]].
- `doi-chieu-gtgt-tndn-cbrr`: Chuyên trách Bước 2 (hoặc 4.1), đối chiếu Bảng kê - Tờ khai GTGT - Quyết toán TNDN - HĐĐT chuẩn 2 sheet CBRR [[conversations/2026-10-15]].

### 3. Các phân hệ còn nằm trong thân skill trục
Các phân hệ chưa tách thành skill độc lập mà chạy trực tiếp trong `phan-tich-ho-so-thue` gồm: tổng hợp mẫu 02 đầu vào; kiểm tra BT5-BT9 Excel; rà soát TNCN đa nguồn; kiểm tra XNT kho 152/156 và dở dang 154; kiểm tra công nợ và ngưỡng tiền mặt TK 111; dựng Master 8 sheet và chấm điểm 28 tiêu chí rủi ro [[conversations/2026-10-15]].
