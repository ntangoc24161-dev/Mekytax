---
type: decision
created: 2026-10-16
updated: 2026-10-16
---
Phòng làm việc cung cấp hồ sơ Công an tách theo 2 trục chính + 2 trợ lý hỗ trợ.

**Trục chính (chạy workflow `quản-lý-công-văn-công-an`):**
1. Tiếp nhận & Điều phối công văn (`quản-lý-hồ-sơ-công-an`) - nhận CV, bóc tách, tra NNT, phân tổ, ghi sổ, đóng thư mục `-R`.
2. Soạn thảo trả lời Công an (`soan-thao-tra-loi-cong-an`) - dự thảo CV + tờ trình cho DN thuộc KT1+KT2+KT3.

**Hỗ trợ (không nằm trong workflow mặc định):**
3. Khai thác dữ liệu hồ sơ (`khai-thac-du-lieu-ho-so`) - Excel 24 cột, NCC/công nợ; chỉ soi sâu GĐ1-3 khi user yêu cầu.
4. Pháp chế & chính sách thuế (`phap-che-chinh-sach-thue`) - tra điều khoản khi vướng căn cứ; từ 2026-10-16 cũng là chỗ tra cứu chính sách cho brain Cập nhật công văn cảnh báo (8 skill sắc thuế/HĐĐT/xử phạt/quy trình/PCTN chuyển về đây).

**Vì sao:** Một agent cũ ôm cả điều phối lẫn soạn thảo lẫn xuất Excel nên loãng vai. Việc hàng ngày của brain này là tiếp nhận và trả lời công văn, không phải mở đoàn kiểm tra hay ôn thi. Pháp chế dùng chung cho phúc đáp Công an và dẫn chiếu F0/F1 nên để một chỗ.

**Áp dụng:** Workflow 2 bước Điều phối → Soạn thảo. Không tạo trợ lý ôn thi trong phòng này. Giữ nguyên slug agent cũ để không gãy lịch sử chạy.
