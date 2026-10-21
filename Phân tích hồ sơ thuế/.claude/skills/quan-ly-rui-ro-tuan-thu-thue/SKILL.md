---
name: Quản lý rủi ro tuân thủ thuế
description: "QLRR tuân thủ thuế: Luật 108/2025, TT 31/2021, phân loại rủi ro, bộ chỉ số TCT, 5 bước xử lý rủi ro."
group: Tài chính
updated: '2026-10-09'
---

# Quản lý rủi ro tuân thủ thuế (QLRRTT)

Khung chọn đối tượng và phân bổ biện pháp quản lý thuế theo mức tuân thủ / rủi ro. Nguồn chính: slide QLRR 29-30/12/2025, TT 31/2021/TT-BTC, Luật QLT 38/2019 và Luật QLT 108/2025. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "quản lý rủi ro thuế", "QLRR", "chọn doanh nghiệp kiểm tra", "phân loại tuân thủ", "bộ chỉ số rủi ro hóa đơn/hoàn thuế".
- Lập hoặc giải thích kế hoạch thanh tra - kiểm tra, hoàn thuế, giám sát hóa đơn, hộ KD.
- Không dùng để mổ chi tiết một bộ BCTC/tờ khai cụ thể → `mo-ho-so-ttkt` hoặc `phan-tich-ho-so-thue`.
- Không dùng để tính mức phạt VPHC → `xu-phat-vphc-thue-hoa-don`.

## Chuẩn bị

- Xác định nghiệp vụ đang làm: kế hoạch TTKT / hoàn GTGT / hóa đơn / kiểm tra tại CQT / hộ KD.
- Lấy đúng bộ chỉ số - quy trình TCT của nghiệp vụ đó (xem bảng trong `references/bo-chi-so-tct.md`).
- Có thông tin ngoài hệ thống (tố cáo, CQĐT, ngân hàng…) thì thủ trưởng có thể đổi mức rủi ro theo thẩm quyền và chịu trách nhiệm.

## Cách chạy

1. Thu thập - làm sạch dữ liệu trong/ngoài ngành.
2. Đánh giá tuân thủ (4 mức) rồi rủi ro tổng thể (5 hạng với DN).
3. Chọn biện pháp tương ứng: hỗ trợ → giám sát → kiểm tra.
4. Ghi nhận kết quả xử lý để cập nhật tiêu chí chu kỳ sau.

## Quy trình (5 bước)

1. **Bối cảnh & dữ liệu**: mục tiêu CQT; khoảng trống dòng tiền - DT kê khai; TMĐT; TK cá nhân dùng cho KD.
2. **Tiêu chí & nhận diện**: bộ chỉ số theo nghiệp vụ; dấu hiệu bất thường HĐĐT vs tờ khai.
3. **Phân tích & xếp hạng**: ưu tiên chạy tự động trên ứng dụng; thủ công khi sự cố / dữ liệu chưa kịp cập nhật.
4. **Kế hoạch xử lý**: khớp biện pháp với mức tuân thủ + rủi ro + dấu hiệu thời điểm quyết định.
5. **Xử lý - theo dõi - đánh giá lại**: cập nhật kết quả; điều chỉnh chỉ số; kế hoạch nâng cao tuân thủ.

## Nguyên tắc

- Khuyến khích tuân thủ tự nguyện; can thiệp sớm; giảm tiếp xúc trực tiếp không cần thiết.
- Đánh giá phân loại thực hiện tự động; công chức làm đúng quy định được miễn trừ trách nhiệm cá nhân.
- Kết quả dùng để chọn biện pháp nghiệp vụ và xây kế hoạch nâng cao tuân thủ, không phải công bố xếp hạng cho NNT theo mặc định.

## Định dạng số tiền (bắt buộc)

- **Hiển thị** (chat, Markdown, báo cáo, ô text): chuẩn Việt Nam - dấu `.` phân cách hàng nghìn, dấu `,` phần lẻ. Ví dụ: `1.234.567` hoặc `1.234.567,89`.
- **Excel số thật** (openpyxl `number_format`): dùng mã invariant `#,##0` (nguyên) hoặc `#,##0.00` (có phần lẻ). Excel locale VN sẽ hiện thành dấu `.` / `,`. Không mô tả nhầm là "dấu phẩy ngăn cách hàng nghìn".
- **Tỷ lệ %**: hiển thị `12,34%` (dấu `,` phần lẻ).
- Cấm dạng Mỹ `1,234,567.89` trong câu trả lời / báo cáo tiếng Việt.

## Bẫy

- Chọn DN kiểm tra cảm tính, bỏ qua kết quả QLRR khi không có thông tin mới.
- Nhầm Luật 38 Điều 9 với Luật 108 Điều 32-34 khi trích dẫn sau 01/07/2026.
- Áp bộ chỉ số hoàn thuế cho việc chọn kế hoạch TTKT (sai nghiệp vụ).
- Quên cập nhật kết quả xử lý → vòng QLRR sau bị lệch.

## Kiểm chứng

- Nêu đúng nghiệp vụ + số quyết định bộ chỉ số/quy trình đang dùng.
- Có mức tuân thủ/rủi ro (hoặc nói rõ hệ thống chưa trả mức).
- Biện pháp đề xuất khớp mức rủi ro; có bước cập nhật ngược.

## Lịch sử (tự học)
- [2026-10-09] Chốt định dạng số tiền chuẩn VN đồng loạt: hiển thị dấu `.` hàng nghìn và `,` phần lẻ; Excel `#,##0` / `#,##0.00`.
