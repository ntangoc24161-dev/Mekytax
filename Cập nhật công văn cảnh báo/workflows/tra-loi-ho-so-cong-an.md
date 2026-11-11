---
type: workflow
name: Trả lời hồ sơ Công an
slug: tra-loi-ho-so-cong-an
status: 'off'
group: Tài chính
description: Đối chiếu NCC mua vào rủi ro rồi soạn công văn phúc đáp và tờ trình cho hồ sơ Công an chuyển đến.
steps:
  - agent: doi-chieu-rui-ro-ncc
    task: |-
      ĐỐI CHIẾU NCC. Đầu vào: {{input}} phải nêu rõ MST hoặc tên doanh nghiệp cần đối chiếu,
      kỳ đối chiếu, và đường dẫn Bảng kê hoặc dữ liệu HĐĐT nếu có.

      Nếu {{input}} không nêu được MST hoặc tên doanh nghiệp, trả về đúng một dòng
      "DUNG: chua xac dinh duoc doanh nghiep can doi chieu" rồi kết thúc. KHÔNG hỏi lại.

      Sổ theo dõi dùng để tra F0 và F1: G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx

      Xuất đủ các phần: Phần 1, Bảng 2.1, Bảng 2.2 phủ 100% NCC, Bảng 2.3 theo kỳ, Phần 3.
      Mỗi dòng NCC phải ghi nguồn là Bảng kê, HĐĐT, hoặc Chung.

      ĐẦU RA bắt buộc kèm các con số: tổng số NCC, số NCC trùng danh sách rủi ro,
      tổng giá trị và tổng thuế GTGT theo kỳ. CẤM bịa số, thiếu dữ liệu thì ghi "không có dữ liệu".
    verify_agent: kiem-soat-chat-luong-du-lieu
    max_retries: 1
  - agent: soan-thao-van-ban
    task: |-
      SOẠN THẢO. Dùng {{input}} và kết quả đối chiếu ở {{prev}}.

      Nếu {{prev}} không chứa bảng đối chiếu có số liệu, trả về đúng một dòng
      "DUNG: khong co ket qua doi chieu de soan thao" rồi kết thúc.

      Soạn hai văn bản đúng thể thức NĐ 30/2020 và ký hiệu trình ký TCS3:
      1. TRA LOI CV <Số CV>_<Cơ quan>_du thao.docx
      2. TO TRINH <Số CV>.docx

      Mọi con số trong văn bản phải lấy từ {{prev}} hoặc từ sổ theo dõi. Cần dẫn chiếu điều khoản
      cụ thể thì để chỗ trống đánh dấu [CAN PHAP CHE] và ghi vào mục cần bổ sung, chuyển trợ lý
      phap-che-chinh-sach-thue ở brain Cung cấp thông tin công an.

      ĐẦU RA bắt buộc: đường dẫn hai tệp .docx đã ghi, kèm danh sách chỗ trống [CAN PHAP CHE].
      CẤM bịa số liệu, CẤM bịa số hiệu văn bản.
    verify_agent: kiem-soat-chat-luong-du-lieu
    max_retries: 1
updated: "2026-11-09"
---

Quy trình phúc đáp, chỉ chạy khi có hồ sơ Công an hoặc cơ quan khác yêu cầu cung cấp thông tin. Không phải việc thường ngày.

Tách khỏi `day-chuyen-xu-ly-cv-canh-bao` ngày 09/11/2026, vì hai khâu này có điều kiện khởi động khác hẳn: dây chuyền trục chạy mỗi khi có công văn cảnh báo mới, còn phúc đáp chỉ chạy khi có yêu cầu cung cấp hồ sơ.

### Chỉ chạy KHI ĐƯỢC YÊU CẦU

Quy trình này **không bao giờ tự khởi động**. Nó cần một lệnh riêng của Trưởng phòng nêu rõ doanh nghiệp hoặc số công văn cần phúc đáp, ví dụ "đối chiếu NCC của MST 3703435585 kỳ 2026" hoặc "soạn công văn trả lời CV 7268".

Dây chuyền trục `day-chuyen-xu-ly-cv-canh-bao` chạy xong và QC đạt là **hết việc**. Không có lệnh yêu cầu đối chiếu hay soạn thảo thì tuyệt đối không gọi sang đây, và cũng không ghi nhận đó là việc còn dở.

### Điều kiện trước khi chạy
- Doanh nghiệp cần đối chiếu đã có dòng trong F0 hoặc F1 của sổ theo dõi.
- Có Bảng kê hoặc dữ liệu HĐĐT của kỳ cần đối chiếu, nếu không thì đối chiếu chỉ dựa trên sổ và phải ghi rõ giới hạn đó.

### Gác cổng
Cả hai bước có `verify_agent: kiem-soat-chat-luong-du-lieu`. QC trả JSON `{"pass": true|false}`; trượt là quy trình dừng, không soạn thảo trên dữ liệu rỗng.

### Trạng thái
Đang `status: 'off'`. Bật ở trang Workflows khi Trưởng phòng xem xong nội dung các bước.
