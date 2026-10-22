---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
---
Nguyên tắc trình bày và xử lý dữ liệu hồ sơ kiểm tra thuế:
- Định dạng chung: Font Times New Roman, cỡ 11, ngày dd/MM/yyyy KHÔNG hiện giờ, số phân cách hàng nghìn dấu chấm '.', thập phân dấu phẩy ',', tự động giãn cột theo nội dung dài nhất (kể cả tiêu đề).
- Font sai bảng mã cũ (VNI, TCVN3/ABC): tự chuyển về Unicode dựng sẵn trước khi ghi, không copy nguyên chữ lỗi mã.
- Dòng ví dụ ngay dưới tiêu đề mẫu: XÓA khi chạy hồ sơ thật, không trộn với dữ liệu thật.
- Dòng/năm/kỳ hoàn toàn không phát sinh (ngoài phạm vi kiểm tra, không phải thiếu nguồn): XÓA, không giữ dòng toàn số 0.
- IN01_HDDT: Tách riêng và tính tổng theo Mua vào và Bán ra; hóa đơn thay thế vẫn tính vào tổng.
- IN02_BangKe: Tách riêng và tính tổng theo Mua vào và Bán ra (lỗi đã gặp 19/10: chỉ lọc được mua vào, đã sửa skill `giai-doan-1-kiem-ke-danh-gia`). Khối phân đoạn "I. BẢNG KÊ MUA VÀO"/"II. BẢNG KÊ BÁN RA": chữ đứng, đậm - không in nghiêng.
- IN05_NLD_TNCN: nhóm theo năm, trong mỗi năm dòng QTT (quyết toán) đứng TRƯỚC các dòng NLĐ của chính năm đó.
- IN07_CVCBRR: Đối chiếu rủi ro bên bán (hóa đơn mua vào) với sổ F0 tại `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx`, kiểm tra công ty đang kiểm tra với F0, F1 và sổ Công an tại `Bao_cao_cong_van.xlsx`. Tiêu đề cột đậm, chữ đứng.
