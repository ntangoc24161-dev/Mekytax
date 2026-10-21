---
name: "Tổng hợp bên mua từ sổ CVCBRR"
description: "Tổng hợp đơn vị bên mua từ sổ CVCBRR: xuất Chi_tiet theo ngày CV và Tong_hop theo MST mua."
group: "Tài chính"
origin: javis-learned
status: active
created: 2026-09-30
updated: '2026-10-09'
---
## Khi nào dùng

- User hỏi tổng hợp / liệt kê **đơn vị bên mua** trong sổ `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`.
- Câu kiểu: "tổng hợp bên mua trong sổ", "các công ty mua trên CVCBRR".
- Không dùng để đối chiếu NCC bên bán trên bảng kê mua vào của một hồ sơ TTKT → `doi-chieu-ncc-mua-vao-rui-ro`.

Playbook lệnh: `playbook-lenh.md`.

## Chuẩn bị

1. File sổ: `...\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` (sheet `THEO_DOI`).
2. Xác định cột tên/MST bên mua trên sổ (chỉ các dòng đã nhập).
3. Thư mục xuất cùng nhánh sổ hoặc nơi user chỉ định.

## Cách chạy

1. Đọc sheet `THEO_DOI`; lọc dòng **đã có** thông tin bên mua.
2. Xuất `Chi_tiet`: giữ CV + bên bán kèm bên mua; sắp ngày công văn → MST mua → số CV.
3. Xuất `Tong_hop`: gom theo **MST gốc bên mua** (số dòng, số CV, danh sách CV, doanh thu/thuế nếu sổ có).
4. Font **Times New Roman**.
5. Nếu user hỏi một MST cụ thể có nằm cột bên mua không: đếm đúng MST đó, không suy từ tên gần giống.

## Quy trình

1. Đếm tổng dòng sổ, dòng có bên mua, số MST mua khác nhau trước khi xuất.
2. `Tong_hop` sắp theo số dòng giảm dần trừ khi user chỉ tiêu chí khác.
3. Ghi chú nếu doanh thu/thuế trên sổ gần như trống: không bịa số mua.
4. Lưu ví dụ `Tong_hop_ben_mua_CVCBRR.xlsx` cạnh sổ hoặc đúng path user yêu cầu.

## Định dạng số tiền (bắt buộc)

- **Hiển thị** (chat, Markdown, báo cáo, ô text): chuẩn Việt Nam - dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (openpyxl `number_format`): dùng mã invariant `#,##0` (nguyên) hoặc `#,##0.00` (có phần lẻ). Excel locale VN sẽ hiện thành dấu `.` / `,`. Không mô tả nhầm là "dấu phẩy ngăn cách hàng nghìn".
- **Tỷ lệ %**: hiển thị `12,34%` (dấu `,` phần lẻ).
- Cấm dạng Mỹ `1,234,567.89` trong câu trả lời / báo cáo tiếng Việt.

## Bẫy

- Coi đối chiếu NCC mua vào của hồ sơ TTKT như đã làm xong phần bên mua trên sổ.
- Kết luận NNT đang kiểm tra có trong cột mua chỉ vì tên chứa cụm từ giống; phải khớp MST.
- Xuất cả dòng sổ trống bên mua làm phình `Chi_tiet`.
- Quên font Times New Roman.

## Kiểm chứng

- `Chi_tiet` chỉ gồm dòng có bên mua; sắp theo ngày CV.
- `Tong_hop` theo MST gốc; số MST và số dòng khớp bộ lọc.
- Trả lời rõ MST được hỏi có/không có trong cột bên mua kèm số dòng.

## Lịch sử (tự học)
- [2026-10-09] Chốt định dạng số tiền chuẩn VN đồng loạt: hiển thị dấu `.` hàng nghìn và `,` phần lẻ; Excel `#,##0` / `#,##0.00`.
