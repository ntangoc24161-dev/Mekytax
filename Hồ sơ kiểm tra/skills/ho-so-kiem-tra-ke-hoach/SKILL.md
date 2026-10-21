---
name: Hồ sơ kiểm tra theo kế hoạch
description: 'Lập hồ sơ kiểm tra kế hoạch từ phôi MAUTONG HOP: điền số liệu, lưu TONG HOP toàn bộ và chỉ xuất riêng các Quyết định.'
group: Tài chính
updated: '2026-10-16'
learned_updated: '2026-10-16'
---
## Khi nào dùng
Dùng khi lập hoặc xuất hồ sơ đi kiểm tra thuế theo kế hoạch năm từ phôi `MAUTONG HOP.docx`, ví dụ: tạo hồ sơ đi theo tên công ty/MST, xuất trình ký, xuất TONG HOP và hai Quyết định, hoặc cập nhật Số QĐKT/QĐGS thật vào bộ đã xuất mã tạm.

Không dùng cho hồ sơ giải thể, chấm XML/hóa đơn (`phan-tich-ho-so-thue`), hay chỉ ghi sổ Excel (`lap-ho-so-kiem-tra`).

## Chuẩn bị
- Phôi chuẩn `MAUTONG HOP.docx` trong thư mục vụ hoặc `templates/` của vault.
- Dữ liệu định danh NNT: sổ `Danh sach kiem tra quyet toan - JAVIS PRO.xlsm` trên ổ `G:` nếu gắn được; nếu không, bản sao DB NNT trong vault hoặc sổ trên Drive user chỉ định.
- `Bang tinh.xlsx` nếu có (tỷ lệ GVHB, CP QLDN, CPBH, CPTC trên BCTC). Thiếu file này thì vẫn xuất, không dừng.
- Số QĐKT thật nếu đã cấp; chưa có thì mã tạm `DU_THAO-TRINH-KY`.

## Quy trình
1. Khớp MST và tên NNT. Tra DB NNT: địa chỉ, ngành, GCN ĐKKD, GĐ, KTT, chi nhánh.
2. Thời kỳ, nội dung kiểm tra, thành phần đoàn: lấy từ sổ Excel JAVIS PRO. Ổ `G:` không đọc được thì lấy theo vụ kế hoạch gần nhất cùng chuyên đề và nói rõ đó là giả định.
3. Dữ liệu tài chính: có `Bang tinh.xlsx` thì điền tỷ lệ GVHB/CP theo năm. Không có thì điền số hóa đơn/TNCN đã có (bán ra, mua vào, TNCT), để trống tỷ lệ trên BCTC, báo user gửi bảng tính để điền lại.
4. Điền trực tiếp vào phôi `MAUTONG HOP.docx` (thay trường trộn, đúng vị trí). Không dựng Word từ trang trắng.
5. Lưu toàn bộ biểu mẫu vào `TONG HOP.docx`.
6. Trích xuất riêng hai Quyết định (`QD Kiem tra...`, `QDgiam sat...`). Không xuất lẻ DI-01 đến DI-07.
7. Đọc lại hai file Quyết định ngay sau khi xuất. Nếu cắt nhầm đoạn (dính tờ trình, dính điều khoản file kia, thiếu chữ ký) thì dò lại ranh giới trong phôi đã điền rồi xuất lại.
8. Gói `hoso.zip`. Ổ vụ trên Windows/`G:` không gắn được thì ghi vào vault `exports/<slug-mst>/`.
9. Báo file đã xuất, trường còn trống (ngày ký, ngày cấp CCCD, tỷ lệ BCTC, Số QĐKT), nhờ user mở hai Quyết định điền ngày ký.
10. **Khi đã có Số QĐKT/QĐGS thật** (thay `DU_THAO-TRINH-KY`): chỉ thay số quyết định và ngày ban hành **trên đúng hai Quyết định** (và chỗ tương ứng của hai QĐ trong `TONG HOP.docx`). Không điền lan ngày/số sang các biểu mẫu sau QĐ giám sát. Đồng bộ sổ Excel JAVIS PRO; nếu user chỉ Drive thì đẩy Word lên thư mục vụ rồi nhắc: PDF cũ trên Drive chưa mang số mới, Excel đang mở cần đóng mở lại mới thấy số.

## Bẫy
- Không dựng file Word từ trang trắng: mất tab stop, lề và đường kẻ theo Nghị định 30/2020/NĐ-CP.
- Không xuất lẻ DI-01 đến DI-07 trừ khi user yêu cầu cụ thể.
- Không bịa Số QĐKT. Trình ký khi chưa có số thì dùng `DU_THAO-TRINH-KY`.
- Sau khi điền phôi, ranh giới đoạn Quyết định có thể lệch so với phôi gốc: bắt buộc đối chiếu nội dung QĐ trước khi giao file.
- Ngày cấp CCCD không có thì để trống, không bịa.
- Thiếu `Bang tinh.xlsx` thì không bịa tỷ lệ GVHB/CP trên BCTC.
- Cập nhật số QĐ thật: **chỉ sửa số và ngày trên hai Quyết định**. Điền lan ngày sang biểu mẫu phía sau QĐ giám sát là sai (đã lộ ở vụ Tâm Hoàng Phát 16/10/2026).
- Agent không xuất PDF; PDF cũ trên Drive có thể lệch Word mới — nhắc user xuất lại từ Word nếu cần PDF.

## Kiểm chứng
- Hai Quyết định đúng 1 trang A4, đúng tên NNT/MST, không dính đoạn của biểu mẫu khác.
- `TONG HOP.docx` không còn trường trộn của công ty khác; khi đã cấp số thì không còn `DU_THAO-TRINH-KY` trên hai QĐ.
- Thư mục xuất có `TONG HOP.docx`, hai file Quyết định và `hoso.zip`.
- Câu trả lời nêu giả định (thời kỳ, đoàn) và phần còn thiếu.
- Sau cập nhật số QĐ: đối chiếu đúng QĐKT/QĐGS/ngày; các biểu mẫu sau QĐGS không bị đổi ngày ngoài ý muốn.

## Lịch sử (tự học)
- [2026-10-02] Cập nhật bẫy định dạng phôi Word: không dựng văn bản từ trang trắng làm mất định dạng Nghị định 30/2020/NĐ-CP, cần thao tác trực tiếp trên phôi MAUTONG HOP.docx và đồng bộ đầy đủ 3 chuyên đề kiểm tra.
- [2026-10-02] Cập nhật quy cách xuất tài liệu mới từ phản hồi người dùng: không xuất lẻ từng file hồ sơ DI, chỉ xuất các file Quyết định và gom tất cả biểu mẫu còn lại vào file tổng hợp TONG HOP.docx.
- [2026-10-13] Thêm bước đối chiếu ranh giới hai file Quyết định sau khi xuất (vụ Tâm Hoàng Phát bị cắt nhầm đoạn lần đầu); thêm phương án khi thiếu `Bang tinh.xlsx` hoặc không gắn được ổ `G:` (xuất vào vault `exports/`, điền số hóa đơn/TNCN, để trống tỷ lệ BCTC).
- [2026-10-13] Lượt 2026-10-13 làm theo skill cũ đã xuất hai Quyết định cắt nhầm đoạn, phải dò lại ranh giới phôi rồi xuất lại; skill chưa có bước kiểm chứng đó. Cùng lượt thiếu Bang tinh.xlsx và không đọc được sổ trên G: nhưng skill cũ coi hai nguồn này là bắt buộc, không có phương án xuất thiếu tỷ lệ BCTC và ghi
- [2026-10-16] Khi điền số QĐ thật vào bộ đã xuất `DU_THAO`, lần đầu điền nhầm ngày sang biểu mẫu sau QĐ giám sát; phải làm lại chỉ sửa số/ngày trên hai Quyết định. Thêm bước 10 và bẫy tương ứng; nhắc PDF Drive có thể còn bản cũ và Excel đang mở cần mở lại.
- [2026-10-16] Lượt 16/10/2026 cập nhật QĐ thật cho Tâm Hoàng Phát: skill chưa có bước/bẫy khi thay DU_THAO bằng số thật; làm theo cách cũ đã điền nhầm ngày sang biểu mẫu sau QĐ giám sát, phải xuất lại chỉ sửa số/ngày trên hai Quyết định.
