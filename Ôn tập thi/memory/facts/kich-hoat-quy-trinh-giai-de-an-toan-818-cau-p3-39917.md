---
type: decision
provenance: user
origin: javis-learned
created: 2026-09-24
updated: 2026-09-24
---
Người dùng yêu cầu kích hoạt quy trình giải đề an toàn và dò đáp án chuẩn 100% từ máy chủ cho 818 câu hỏi chưa hoàn thiện (trên tổng số 1.247 câu hỏi độc bản qua 40 đề thi) của kíp thi P3-39917.

- File dữ liệu đích: `D:\HOC TAP\ON THI 2026.09\Tong_Hop_Ngan_Hang_1247_Cau_Kip_P3_39917.xlsx` và Master Bank `Tong_Hop_Ngan_Hang_De_On_Thi_Kiem_Tra_Thue_2026.xlsx`.
- Cơ chế vận hành: Giãn cách an toàn 5.1s-7.3s giữa các câu, nghỉ giữa phiên 10s-15s, tự động thử các phương án còn lại cho đến khi máy chủ phản hồi `is_correct: true`.
