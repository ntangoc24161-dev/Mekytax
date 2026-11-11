---
type: wiki
status: active
tags: [wiki, Tài chính, Quy trình]
created: 2026-11-11
updated: 2026-11-11
source: [[conversations/2026-11-11]]
---

Quy trình chuẩn hóa lập và hoàn thiện bộ hồ sơ kiểm tra quyết toán thuế phục vụ giải thể doanh nghiệp (đóng mã số thuế) trong hệ thống Second Brain [[conversations/2026-11-11]]:

## 1. Bản chất và Phân loại cuộc kiểm tra
- **Mục tiêu**: Chốt toàn diện nghĩa vụ ngân sách nhà nước (NSNN) phát sinh trong suốt quá trình hoạt động của người nộp thuế để cơ quan thuế ban hành văn bản xác nhận hoàn thành nghĩa vụ thuế, phục vụ đóng mã số thuế theo quy định của Luật Quản lý thuế [[conversations/2026-11-11]].
- **Phân loại 3 diện thực tế**:
  + **Diện theo thông báo ngừng hoạt động (như bộ Mo Fa)**: Căn cứ thông báo đóng mã số thuế / giải thể; thời hạn kiểm tra rút gọn (thường 03 ngày làm việc) [[conversations/2026-11-11]].
  + **Diện giải thể trên nền kế hoạch năm (như bộ Sơn Tiền Hải)**: Doanh nghiệp thuộc diện giải thể nhưng cuộc kiểm tra được phê duyệt theo kế hoạch thanh kiểm tra năm (Quyết định 1412/QĐ-DON); thời hạn kiểm tra 20 ngày [[conversations/2026-11-11]].
  + **Diện giải thể kiêm hoàn thuế GTGT (như bộ Đại Phát Đạt)**: Doanh nghiệp vừa giải thể vừa có số thuế GTGT chưa khấu trừ hết đề nghị hoàn; áp dụng hoàn theo Điều 32 Nghị định 181/2025/NĐ-CP và bù trừ trực tiếp nghĩa vụ tài chính [[conversations/2026-11-11]].

## 2. Cấu trúc bộ hồ sơ giải thể chuẩn (Bộ 22 văn bản)
Bộ hồ sơ tổng hợp `MAUTONG HOP GIAI THE.docx` bao gồm 22 biểu mẫu, mở rộng so với hồ sơ kế hoạch thông thường [[Chuẩn định dạng TONG HOP hồ sơ kiểm tra thuế]]:
1. Tờ trình ban hành Quyết định kiểm tra và Quyết định giám sát.
2. Quyết định kiểm tra thuế (QĐKT).
3. Quyết định giám sát hoạt động Đoàn kiểm tra (QĐGS).
4. Phân tích rủi ro / Kế hoạch kiểm tra.
5. Biên bản công bố Quyết định kiểm tra.
6. Giấy xác nhận giao Quyết định kiểm tra.
7. Biên bản xác nhận số liệu kiểm tra của Thành viên (BBXN TV).
8. Biên bản xác nhận số liệu kiểm tra của Trưởng đoàn (BBXN TD).
9. Biên bản kiểm tra thuế (BBKT Giải thể) - bao gồm Mục III Nghĩa vụ NSNN.
10. Biên bản công khai và bàn giao dự thảo BBKT.
11. Báo cáo kết quả kiểm tra.
12. Kế hoạch giám sát hoạt động của Đoàn kiểm tra.
13. Biên bản làm việc của Người giám sát (Mẫu 04-TTKT).
14. Báo cáo nội dung và kết quả giám sát hoạt động của Đoàn kiểm tra (checkbox Unicode `☐`/`☑` sau chữ).
15. Nhật ký thanh tra kiểm tra.
16. Biên bản phân công công việc.
17. Quyết định xử phạt vi phạm hành chính (QĐXP).
18. Tờ trình ban hành Quyết định xử phạt vi phạm hành chính.
19. Biên bản giao nhận hồ sơ, tài liệu cuộc kiểm tra (bên nhận khác bên giao).
20. Văn bản từ chối nhận lại số tiền nộp thừa (điền MST theo từng ô lưới).
21. File xuất riêng: `QD Kiem tra tai cty...docx` (chuẩn 1 trang A4).
22. File xuất riêng: `QDgiam sat tai cty...docx` (chuẩn 1 trang A4).

## 3. Các điểm kỹ thuật nghiệp vụ bắt buộc
- **Mục III Nghĩa vụ NSNN**: Bắt buộc phải điền đủ số phát sinh các năm. Trường hợp chưa có sao kê chi tiết, áp dụng công thức suy luận: `(3)=0`, `(6)=(4)`, `(7)=(3)+(4)+(5)-(6)` (số còn phải nộp chính bằng số truy thu qua kiểm tra) [[conversations/2026-11-11]].
- **Xử lý số gõ cứng trên DTCHEP**: Cơ chế cảnh báo mềm (non-blocking warning), không chặn tiến trình đổ số vào Word, ghi nhận cảnh báo xác minh chứng từ gốc vào `wiki/_open-questions.md` [[conversations/2026-11-11]].
- **Định danh xử phạt VPHC**: Khai sai dẫn đến thiếu thuế phạt 20% theo Điều 16 Nghị định 125/2020/NĐ-CP; trường hợp khai sai chi phí nhưng sau bù trừ lỗ vẫn không làm phát sinh thuế thiếu thì phạt tiền cố định 6.500.000 đồng theo Nghị định 310/2025/NĐ-CP [[conversations/2026-11-11]].
- **Bù trừ hoàn thuế giải thể**: Trường hợp có số thuế GTGT chưa khấu trừ hết tại thời điểm giải thể, lập hồ sơ đề xuất hoàn thuế kiêm bù trừ thu NSNN (mẫu MS-03-DCS và dự thảo QĐ hoàn thuế) theo Điều 32 Nghị định 181/2025/NĐ-CP [[Hồ sơ đề xuất hoàn thuế kiêm bù trừ thu ngân sách đối với doanh nghiệp giải thể]].
