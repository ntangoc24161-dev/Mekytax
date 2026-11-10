---
name: Xuất danh sách tổng hợp chi tiết hồ sơ Công an
description: Xuất Excel tổng hợp chi tiết hồ sơ công văn Công an theo từng cột tiêu chí tờ khai, BCTC, hóa đơn, vi phạm, nợ thuế tại thư mục CV MOI.
group: Tài chính
updated: '2026-09-26'
learned_updated: '2026-09-26'
---
## Khi nào dùng
- Khi cán bộ hoặc Lãnh đạo yêu cầu xuất bảng kê / danh sách tổng hợp chi tiết hồ sơ tài liệu cần cung cấp theo các công văn Công an (đặc biệt là các công văn trạng thái "Mới" và "Đang xử lý" của Tổ KT1 hoặc toàn bộ các công văn).
- Cần phân tách chi tiết từng loại tài liệu, tờ khai, báo cáo thành các cột độc lập kèm thời gian/kỳ thuế cụ thể để phục vụ việc trích xuất và bàn giao hồ sơ.

## Chuẩn bị
- File sổ theo dõi: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`).
- Thư mục chứa file Markdown trích xuất công văn: `D:\HSO CONG AN FILE MD`.
- Thư mục lưu trữ công văn năm 2026: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026`.
- File CSDL NNT: `I:\My Drive\DB NNT.xls` (hoặc `.xlsx`).
- Vị trí tệp đích xuất ra:
  + Tệp chính thức: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI\Danh_sach_tong_hop_ho_so_cung_cap_CV_MOI.xlsx`
  + Tệp đồng bộ tên ngắn: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI\CV MOI.xlsx` và tại thư mục gốc năm 2026.

## Quy trình
1. **Thu thập và lọc dữ liệu công văn:**
   - Đọc dữ liệu từ sheet `NHAP_HO_SO` của `Bao_cao_cong_van.xlsx`.
   - Lọc danh sách công văn theo yêu cầu (mặc định lấy các công văn trạng thái "Mới" và "Đang xử lý" có phân công cho Tổ KT1, hoặc toàn bộ theo chỉ đạo).
   - Đọc nội dung yêu cầu cụ thể từ các file `.md` tương ứng tại `D:\HSO CONG AN FILE MD`.

2. **Phân tách các cột tiêu chí tài liệu và mốc thời gian:**
   Mỗi tiêu chí nghiệp vụ được tách thành một cột độc lập và điền rõ khoảng thời gian/kỳ thuế cần cung cấp:
   - **Tờ khai thuế GTGT:** Ghi rõ kỳ tháng/quý/năm yêu cầu (ví dụ: *Các kỳ trong năm 2022*, *Từ Quý 1/2021 đến Quý 4/2024*, *Điều chỉnh giảm thuế GTGT 03 HĐ khống*).
   - **Tờ khai Quyết toán thuế TNDN:** Ghi rõ năm tài chính yêu cầu quyết toán (ví dụ: *Năm 2021 - 2023*, *Năm 2022 - 2024*, *Năm 2025*).
   - **Tờ khai Quyết toán / Khấu trừ thuế TNCN:** Ghi rõ năm quyết toán hoặc kỳ khấu trừ định kỳ (ví dụ: *Năm 2021 - 2024*, *Khấu trừ Quý 1/2025*, *Không yêu cầu*).
   - **Báo cáo tài chính (BCTC):** Ghi rõ năm lập BCTC cần sao lục (ví dụ: *Năm 2020 - 2024*, *Năm 2022*, *Năm 2025*).
   - **Hóa đơn điện tử / Bảng kê HĐ:** Ghi rõ yêu cầu xuất bảng kê HĐĐT mua vào - bán ra kèm file Excel gửi email Công an, hoặc nội dung xác minh/bóc tách hóa đơn cụ thể.
   - **Quyết định Xử phạt VPHC / Cưỡng chế nợ:** Ghi rõ giai đoạn rà soát quyết định xử phạt VPHC về thuế/hóa đơn, quyết định cưỡng chế tài khoản/hóa đơn hoặc phương án xử phạt trốn thuế.
   - **Thông báo Đóng MST / Bỏ địa chỉ:** Ghi rõ số và ngày thông báo NNT không hoạt động tại địa chỉ đã đăng ký hoặc tình trạng tạm ngừng kinh doanh.
   - **Thông tin Định danh & Pháp lý NNT:** Ghi rõ thông tin Giấy phép ĐKKD, ĐDPL, CCCD, Kế toán, CBQL, Số tài khoản ngân hàng (chỉ lập khi công văn Công an có yêu cầu).
   - **Nghĩa vụ nộp NSNN / Nợ thuế:** Ghi rõ tình trạng nợ thuế tính đến thời điểm hiện tại hoặc số tiền thuế đã nộp vào NSNN.
   - **Hồ sơ tài liệu khác / Nội dung đặc thù:** Hồ sơ hoàn thuế GTGT, biên bản thanh tra/kiểm tra thuế, phiếu chuyển Tổ HKD, phương án chuyển hồ sơ xử lý hành vi trốn thuế...

3. **Khởi tạo và định dạng sổ tính Excel nhiều Sheet:**
   - **Sheet 1 (`TONG_HOP_CHI_TIET`):** Chi tiết từng Doanh nghiệp / Mã số thuế, gồm 26 cột tiêu chí: STT, Số CV, Ngày CV, Cơ quan ban hành, Điều tra viên/Cán bộ thụ lý, Tên NNT, Mã số thuế, Địa chỉ, Tổ KT, Cán bộ xử lý, Trạng thái, Tờ khai GTGT, Quyết toán TNDN, Quyết toán TNCN, BCTC, Hóa đơn điện tử, Xử phạt VPHC & Cưỡng chế, Thông báo Đóng MST, Thông tin Định danh DN, Nghĩa vụ NSNN & Nợ thuế, Hồ sơ khác, Tệp dự thảo CV trả lời, Tệp Tờ trình, Thư mục lưu trữ trên Drive, **Ngày giao công văn (Q)**, **Hạn thực hiện (R)**.
   - **Sheet 2 (`TONG_HOP_THEO_CV`):** Tổng hợp theo từng Công văn (STT, Số CV, Ngày CV, Cơ quan, Số lượng DN, Tóm tắt yêu cầu hồ sơ, Trạng thái xử lý, Tệp dự thảo, Thư mục Drive).
   - **Sheet 3 (`TOAN_BO_HO_SO_DU_THAO`):** Danh mục toàn bộ các hồ sơ đang lưu dự thảo Word trong thư mục năm 2026.
   - **Sheet 4 (`DANH_MUC_TIEU_CHI`):** Bảng giải thích chi tiết mục đích và hướng dẫn trích xuất dữ liệu của từng cột tiêu chí.

4. **Quy chuẩn định dạng trang tính Excel:**
   - Sử dụng phông chữ chuẩn Arial hoặc Segoe UI, cỡ chữ 10-11pt.
   - Hàng tiêu đề: Nền màu xanh đậm (Navy `#1B365D` hoặc `#2F5597`), chữ trắng, in đậm, căn giữa.
   - Cố định dòng tiêu đề (**Freeze Panes**) và bật bộ lọc tự động (**AutoFilter**) trên tất cả các sheet.
   - Tô màu nhẹ phân biệt trạng thái: "Mới" (vàng nhạt), "Đang xử lý" (cam nhạt), "Đang trình ký" (xanh dương nhạt), "Đã hoàn thành" (xanh lá nhạt).
   - Tự động căn chỉnh độ rộng cột (`auto-fit column width`) đảm bảo nội dung không bị che khuất.

## Bẫy
- Gộp chung nhiều loại tờ khai/báo cáo vào 1 cột văn bản duy nhất khiến cán bộ không thể lọc theo từng loại nghiệp vụ.
- Không ghi rõ kỳ kê khai / mốc thời gian cụ thể (chỉ ghi chung chung "có yêu cầu") dẫn đến việc cán bộ phải mở lại từng file PDF để tra cứu.
- Quên không cập nhật các công văn mới phát sinh vào danh sách tổng hợp.

## Kiểm chứng
- Mở tệp Excel kiểm tra: Tất cả các sheet hiển thị đầy đủ, không có ô lỗi `#VALUE!`, `#REF!`.
- Bộ lọc AutoFilter hoạt động chuẩn xác trên từng cột tiêu chí.
- Các mốc thời gian và kỳ thuế khớp đúng 100% với nội dung công văn gốc.

## Lịch sử (tự học)
- [2026-09-26] Khởi tạo skill xuất danh sách tổng hợp chi tiết hồ sơ cung cấp công an ra file Excel tại thư mục CV MOI, tự động phân tách độc lập 24 cột tiêu chí nghiệp vụ kèm kỳ thuế và thời gian yêu cầu.
